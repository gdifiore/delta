from docx import Document
import os
import json
import hashlib
from datetime import datetime
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional, Tuple
import difflib
from pathlib import Path
import shutil

@dataclass
class Delta:
    timestamp: str
    parent_hash: str
    content: List[str]
    diff_from_parent: List[str]
    message: str
    hash: str

class DocxVersionStore:
    def __init__(self, path: str = ".docx-versions"):
        self.root_path = Path(path)
        self.objects_path = self.root_path / "objects"
        self.refs_path = self.root_path / "refs"
        self.head_path = self.root_path / "HEAD"
        if self.root_path.exists():
            shutil.rmtree(self.root_path)

        self.initialize_store()

    def initialize_store(self):
        if not self.root_path.exists():
            self.objects_path.mkdir(parents=True)
            self.refs_path.mkdir()
            self.head_path.write_text("")
            (self.refs_path / "master").write_text("")

    def _extract_content(self, docx_path: str) -> List[str]:
        doc = Document(docx_path)
        return [para.text for para in doc.paragraphs]

    def _calculate_hash(self, content: List[str]) -> str:
        hasher = hashlib.sha256()
        for line in content:
            hasher.update(line.encode())
        return hasher.hexdigest()[:8]

    def _calculate_diff(self, old_content: List[str], new_content: List[str]) -> List[str]:
        return list(difflib.unified_diff(old_content, new_content, fromfile='previous', tofile='current', lineterm=''))

    def _store_object(self, delta: Delta):
        object_path = self.objects_path / delta.hash
        with open(object_path, 'w') as f:
            json.dump(asdict(delta), f, indent=2)

    def _load_object(self, hash_id: str) -> Optional[Delta]:
        object_path = self.objects_path / hash_id
        if not object_path.exists():
            return None
        with open(object_path, 'r') as f:
            data = json.load(f)
            return Delta(**data)

    def _update_ref(self, ref_name: str, hash_id: str):
        (self.refs_path / ref_name).write_text(hash_id)

    def _get_ref(self, ref_name: str) -> Optional[str]:
        ref_path = self.refs_path / ref_name
        return ref_path.read_text().strip() if ref_path.exists() else None

    def _update_head(self, ref_name: str):
        self.head_path.write_text(f"ref: {ref_name}")

    def _get_head(self) -> Optional[str]:
        if not self.head_path.exists():
            return None
        head_content = self.head_path.read_text().strip()
        return self._get_ref(head_content[5:]) if head_content.startswith("ref: ") else head_content

    def commit(self, docx_path: str, message: str) -> str:
        content = self._extract_content(docx_path)
        head = self._get_head()
        diff = []
        if head:
            parent_delta = self._load_object(head)
            if parent_delta:
                diff = self._calculate_diff(parent_delta.content, content)
        
        content_hash = self._calculate_hash(content)
        delta = Delta(
            timestamp=datetime.now().isoformat(),
            parent_hash=head or "",
            content=content,
            diff_from_parent=diff,
            message=message,
            hash=content_hash
        )
        self._store_object(delta)
        self._update_ref("master", content_hash)
        self._update_head("master")
        return content_hash

    def get_history(self) -> List[Tuple[str, str, datetime]]:
        history = []
        current = self._get_head()
        while current:
            delta = self._load_object(current)
            if delta:
                history.append((delta.hash, delta.message, datetime.fromisoformat(delta.timestamp)))
                current = delta.parent_hash
            else:
                break
        return history

    def export_version(self, version_hash: str, output_path: str):
        delta = self._load_object(version_hash)
        if delta:
            doc = Document()
            for paragraph_text in delta.content:
                doc.add_paragraph(paragraph_text)
            doc.save(output_path)

    def show_diff(self, version_hash: str) -> Optional[List[str]]:
        delta = self._load_object(version_hash)
        return delta.diff_from_parent if delta else None

    def _generate_changes_summary(self, old_content, new_content) -> str:
        """Generates a human-readable summary of changes"""
        old_len = len(old_content)
        new_len = len(new_content)

        added = len([1 for line in new_content if line not in old_content])
        removed = len([1 for line in old_content if line not in new_content])

        summary_parts = []
        if added:
            summary_parts.append(f"{added} line{'s' if added != 1 else ''} added")
        if removed:
            summary_parts.append(f"{removed} line{'s' if removed != 1 else ''} removed")

        if not summary_parts:
            if old_len != new_len:
                return "Content rearranged"
            return "No text changes (perhaps formatting only)"

        return "Changes: " + ", ".join(summary_parts)

    def show_friendly_diff(self, version_hash: str) -> str:
        """Returns a more readable version of the diff with context"""
        delta = self._load_object(version_hash)
        if not delta or not delta.parent_hash:
            return "Initial version"
        
        parent_delta = self._load_object(delta.parent_hash)
        if not parent_delta:
            return "Parent version not found"

        old_content = parent_delta.content
        new_content = delta.content

        # Use differ to get detailed comparison
        differ = difflib.Differ()
        diff = list(differ.compare(old_content, new_content))

        # Format the output nicely with line numbers
        result = []
        current_section = []
        showing_context = False
        old_line_num = 1
        new_line_num = 1

        for line in diff:
            marker = line[0]
            content = line[2:]

            if marker == ' ':  # Context line
                if not showing_context:
                    if current_section:
                        result.extend(current_section)
                        current_section = []
                    result.append("...")
                    showing_context = True
                result.append(f"{' ':3}{old_line_num:4} {new_line_num:4}  {content}") # Both line numbers
                old_line_num += 1
                new_line_num += 1

            else:
                showing_context = False
                if marker == '-':
                    current_section.append(f"-  {old_line_num:4}{' ':8}{content}")  # Only old line number

                    old_line_num += 1

                elif marker == '+':
                    current_section.append(f"+  {new_line_num:4}{' ':8}{content}")  # Only new line number

                    new_line_num += 1


        if current_section:
            result.extend(current_section)

        # Add a summary at the top
        changes_summary = self._generate_changes_summary(old_content, new_content)
        result.insert(0, changes_summary)

        return "\n".join(result)

    def compare_versions(self, hash1: str, hash2: str) -> List[str]:
        content1 = self._load_object(hash1).content if self._load_object(hash1) else []
        content2 = self._load_object(hash2).content if self._load_object(hash2) else []
        return self._calculate_diff(content1, content2)
