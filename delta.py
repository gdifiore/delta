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
from dataclasses import dataclass
from enum import Enum
from typing import List, Tuple
import difflib
import re
from collections import defaultdict
import gzip

@dataclass
class Delta:
    timestamp: str
    parent_hash: str
    content: List[str]
    diff_from_parent: List[str]
    message: str
    hash: str

class ChangeType(Enum):
    ADDED = "+"
    REMOVED = "-"
    MOVED_FROM = "←"
    MOVED_TO = "→"
    UNCHANGED = " "

@dataclass
class TextBlock:
    content: str
    old_position: int = -1
    new_position: int = -1

@dataclass
class WordChange:
    type: ChangeType
    content: str
    position: int

class CombinedDiffer:
    def __init__(self):
        self.COLORS = {
            'GREEN': '\033[92m',    # for additions
            'RED': '\033[91m',      # for deletions
            'BLUE': '\033[94m',     # for moved text
            'YELLOW': '\033[93m',   # for move destinations
            'GREY': '\033[90m',     # for context
            'RESET': '\033[0m'
        }

    def _split_into_words(self, text: str) -> List[str]:
        """Split text into words while preserving whitespace and punctuation"""
        return re.findall(r'\S+|\s+', text)

    def _normalize_text(self, text: str) -> str:
        """Normalize text for comparison (strip whitespace, lowercase)"""
        return re.sub(r'\s+', ' ', text.strip().lower())

    def _find_moved_blocks(self, old_paragraphs: List[str], new_paragraphs: List[str],
                          min_block_length: int = 5) -> Dict[str, TextBlock]:
        """Identify blocks of text that have been moved"""
        blocks = {}

        old_normalized = [self._normalize_text(p) for p in old_paragraphs]
        new_normalized = [self._normalize_text(p) for p in new_paragraphs]

        for old_idx, old_text in enumerate(old_paragraphs):
            old_words = old_text.split()
            if len(old_words) < min_block_length:
                continue

            old_norm = old_normalized[old_idx]
            for new_idx, new_norm in enumerate(new_normalized):
                if old_norm == new_norm and old_idx != new_idx:
                    block = TextBlock(
                        content=old_text,
                        old_position=old_idx,
                        new_position=new_idx
                    )
                    blocks[old_norm] = block

        return blocks

    def format_word_diff(self, old_text: str, new_text: str) -> str:
        """Generate a formatted word-level diff"""
        old_words = self._split_into_words(old_text)
        new_words = self._split_into_words(new_text)

        matcher = difflib.SequenceMatcher(None, old_words, new_words)

        result = []
        for op, i1, i2, j1, j2 in matcher.get_opcodes():
            if op == 'equal':
                result.append(''.join(old_words[i1:i2]))
            elif op == 'delete':
                result.append(self._color_text(''.join(old_words[i1:i2]), 'RED'))
            elif op == 'insert':
                result.append(self._color_text(''.join(new_words[j1:j2]), 'GREEN'))
            elif op == 'replace':
                result.append(self._color_text(''.join(old_words[i1:i2]), 'RED'))
                result.append(self._color_text(''.join(new_words[j1:j2]), 'GREEN'))

        return ''.join(result)

    def format_combined_diff(self, old_paragraphs: List[str], new_paragraphs: List[str]) -> str:
        """Format diff showing both moved blocks and word-level changes"""
        result = []
        moved_blocks = self._find_moved_blocks(old_paragraphs, new_paragraphs)

        moved_from_positions = {block.old_position for block in moved_blocks.values()}
        moved_to_positions = {block.new_position for block in moved_blocks.values()}

        max_len = max(len(old_paragraphs), len(new_paragraphs))
        for i in range(max_len):
            #result.append(f"\n=== Paragraph {i+1} ===")

            # Handle moves first
            if i in moved_from_positions:
                result.append(self._color_text(f"MOVED FROM: {old_paragraphs[i]}", 'BLUE'))
                continue

            if i in moved_to_positions:
                result.append(self._color_text(f"MOVED TO: {new_paragraphs[i]}", 'YELLOW'))
                continue

            # For non-moved paragraphs, show word-level diff
            if i < len(old_paragraphs) and i < len(new_paragraphs):
                if old_paragraphs[i] != new_paragraphs[i]:
                    result.append(self.format_word_diff(old_paragraphs[i], new_paragraphs[i]))
                else:
                    result.append(old_paragraphs[i])
            elif i < len(old_paragraphs):
                result.append(self._color_text(f"{old_paragraphs[i]}", 'RED'))
            elif i < len(new_paragraphs):
                result.append(self._color_text(f"{new_paragraphs[i]}", 'GREEN'))

        return '\n'.join(result)

    def _color_text(self, text: str, color: str) -> str:
        """Wrap text in color codes"""
        return f"{self.COLORS[color]}{text}{self.COLORS['RESET']}"

class DocxVersionStore:
    def __init__(self, path: str = ".docx-versions"):
        self.root_path = Path(path)
        self.objects_path = self.root_path / "objects"
        self.refs_path = self.root_path / "refs"
        self.head_path = self.root_path / "HEAD"

        if not self.root_path.exists():
            self.initialize_store()
        else:
            # Ensure expected subdirectories exist if the folder was manually created
            self._validate_store()

    def _validate_store(self):
        """Ensures the store structure is intact if the folder already exists."""
        missing_dirs = []
        if not self.objects_path.exists():
            missing_dirs.append(self.objects_path)
        if not self.refs_path.exists():
            missing_dirs.append(self.refs_path)
        if not self.head_path.exists():
            self.head_path.write_text("")

        # Create any missing directories
        for directory in missing_dirs:
            directory.mkdir(parents=True)

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
        # Convert to JSON and get uncompressed size
        json_data = json.dumps(asdict(delta), indent=2)
        uncompressed_size = len(json_data.encode('utf-8'))

        # Compress and store
        with gzip.open(object_path, 'wt', encoding='utf-8') as f:
            f.write(json_data)
        '''
        # Get compressed size
        compressed_size = object_path.stat().st_size

        print(f"Delta {delta.hash} sizes:")
        print(f"  Uncompressed: {uncompressed_size:,} bytes")
        print(f"  Compressed:   {compressed_size:,} bytes")
        print(f"  Ratio:        {compressed_size/uncompressed_size:.2%}")
        '''

    def _load_object(self, hash_id: str) -> Optional[Delta]:
        object_path = self.objects_path / hash_id
        if not object_path.exists():
            return None
        try:
            with gzip.open(object_path, 'rt', encoding='utf-8') as f:
                data = json.load(f)
                return Delta(**data)
        except (gzip.BadGzipFile, json.JSONDecodeError) as e:
            print(f"Error loading compressed delta: {e}")
            return None

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
        delta = self._load_object(version_hash)
        if not delta or not delta.parent_hash:
            return "Initial version"

        parent_delta = self._load_object(delta.parent_hash)
        if not parent_delta:
            return "Parent version not found"

        differ = CombinedDiffer()

        # Add summary header
        result = []
        changes_summary = self._generate_changes_summary(parent_delta.content, delta.content)
        result.append(f"\n=== {changes_summary} ===\n")

        # Add the combined diff
        result.append(differ.format_combined_diff(parent_delta.content, delta.content))

        return '\n'.join(result)

    def compare_versions(self, hash1: str, hash2: str) -> List[str]:
        content1 = self._load_object(hash1).content if self._load_object(hash1) else []
        content2 = self._load_object(hash2).content if self._load_object(hash2) else []
        return self._calculate_diff(content1, content2)
