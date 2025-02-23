from docx import Document
import os
import json
import hashlib
from datetime import datetime
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional, Tuple
import difflib
from pathlib import Path
import gzip
import re

@dataclass
class DeltaOperation:
    """Represents a single incremental change operation."""
    op: str  # 'insert', 'delete', 'replace'
    position: int  # Line number where the operation applies
    # New content for insert/replace, None for delete
    content: Optional[List[str]] = None


@dataclass
class Delta:
    timestamp: str
    parent_hash: str
    baseline_hash: str  # Reference to the nearest full content snapshot
    # Incremental changes from baseline or parent
    operations: List[DeltaOperation]
    message: str
    hash: str


class CombinedDiffer:
    def __init__(self):
        self.COLORS = {
            'GREEN': '\033[92m', 'RED': '\033[91m', 'BLUE': '\033[94m',
            'YELLOW': '\033[93m', 'GREY': '\033[90m', 'RESET': '\033[0m'
        }

    def _split_into_words(self, text: str) -> List[str]:
        """Split text into words while preserving whitespace and punctuation."""
        return re.findall(r'\S+|\s+', text)

    def _normalize_text(self, text: str) -> str:
        """Normalize text for comparison (strip whitespace, lowercase)."""
        return re.sub(r'\s+', ' ', text.strip().lower())

    def _find_moved_blocks(self, old_paragraphs: List[str], new_paragraphs: List[str],
                           min_block_length: int = 5) -> Dict[str, 'TextBlock']:
        """Identify blocks of text that have been moved."""
        from dataclasses import dataclass

        @dataclass
        class TextBlock:
            content: str
            old_position: int = -1
            new_position: int = -1

        blocks = {}
        old_normalized = [self._normalize_text(p) for p in old_paragraphs]
        new_normalized = [self._normalize_text(p) for p in new_paragraphs]

        for old_idx, old_text in enumerate(old_paragraphs):
            old_words = old_text.split()
            if len(old_words) < min_block_length:
                continue
            old_norm = old_normalized[old_idx]
            if old_norm in blocks:
                continue
            for new_idx, new_norm in enumerate(new_normalized):
                if old_norm == new_norm and old_idx != new_idx:
                    blocks[old_norm] = TextBlock(
                        content=old_text,
                        old_position=old_idx,
                        new_position=new_idx
                    )
                    break

        return blocks

    def format_word_diff(self, old_text: str, new_text: str) -> str:
        """Generate a formatted word-level diff."""
        old_words = self._split_into_words(old_text)
        new_words = self._split_into_words(new_text)
        matcher = difflib.SequenceMatcher(None, old_words, new_words)
        result = []
        for op, i1, i2, j1, j2 in matcher.get_opcodes():
            if op == 'equal':
                result.append(''.join(old_words[i1:i2]))
            elif op == 'delete':
                result.append(self._color_text(
                    ''.join(old_words[i1:i2]), 'RED'))
            elif op == 'insert':
                result.append(self._color_text(
                    ''.join(new_words[j1:j2]), 'GREEN'))
            elif op == 'replace':
                result.append(self._color_text(
                    ''.join(old_words[i1:i2]), 'RED'))
                result.append(self._color_text(
                    ''.join(new_words[j1:j2]), 'GREEN'))
        return ''.join(result)

    def format_combined_diff(self, old_paragraphs: List[str], new_paragraphs: List[str]) -> str:
        """Format diff showing both moved blocks and word-level changes."""
        result = []
        moved_blocks = self._find_moved_blocks(old_paragraphs, new_paragraphs)
        moved_from_positions = {
            block.old_position for block in moved_blocks.values()}
        moved_to_positions = {
            block.new_position for block in moved_blocks.values()}

        max_len = max(len(old_paragraphs) if old_paragraphs else 0,
                      len(new_paragraphs) if new_paragraphs else 0)
        for i in range(max_len):
            if i in moved_from_positions:
                result.append(self._color_text(
                    f"MOVED FROM: {old_paragraphs[i]}", 'BLUE'))
                continue
            if i in moved_to_positions:
                result.append(self._color_text(
                    f"MOVED TO: {new_paragraphs[i]}", 'YELLOW'))
                continue
            if i < len(old_paragraphs) and i < len(new_paragraphs):
                if old_paragraphs[i] != new_paragraphs[i]:
                    result.append(self.format_word_diff(
                        old_paragraphs[i], new_paragraphs[i]))
                else:
                    result.append(old_paragraphs[i])
            elif i < len(old_paragraphs):
                result.append(self._color_text(f"{old_paragraphs[i]}", 'RED'))
            elif i < len(new_paragraphs):
                result.append(self._color_text(
                    f"{new_paragraphs[i]}", 'GREEN'))

        return '\n'.join(result)

    def _color_text(self, text: str, color: str) -> str:
        """Wrap text in color codes."""
        return f"{self.COLORS[color]}{text}{self.COLORS['RESET']}"


class DocxVersionStore:
    SNAPSHOT_INTERVAL = 10  # Create a full snapshot every 10 commits

    def __init__(self, path: str = ".docx-versions"):
        self.root_path = Path(path)
        self.objects_path = self.root_path / "objects"
        self.refs_path = self.root_path / "refs"
        self.head_path = self.root_path / "HEAD"
        if not self.root_path.exists():
            self.initialize_store()
        else:
            self._ensure_store_structure()

    def _ensure_store_structure(self) -> None:
        """Ensure the store structure is intact or create it if missing."""
        for directory in [self.objects_path, self.refs_path]:
            directory.mkdir(parents=True, exist_ok=True)
        if not self.head_path.exists():
            self.head_path.touch()

    def initialize_store(self) -> None:
        """Initialize a new version store."""
        self.objects_path.mkdir(parents=True)
        self.refs_path.mkdir(parents=True)
        self.head_path.touch()
        (self.refs_path / "master").touch()

    def _extract_content(self, docx_path: str) -> List[str]:
        """Extract paragraph text from a DOCX file."""
        doc = Document(docx_path)
        return [para.text for para in doc.paragraphs]

    def _calculate_hash(self, content: List[str]) -> str:
        """Calculate a hash for a list of content."""
        return hashlib.sha256(''.join(content).encode('utf-8')).hexdigest()[:8]

    def _compute_operations(self, old_content: List[str], new_content: List[str]) -> List[DeltaOperation]:
        """Compute incremental operations between two versions."""
        matcher = difflib.SequenceMatcher(None, old_content, new_content)
        operations = []
        for op, i1, i2, j1, j2 in matcher.get_opcodes():
            if op == 'equal':
                continue
            elif op == 'delete':
                operations.append(DeltaOperation('delete', i1, None))
            elif op == 'insert':
                operations.append(DeltaOperation(
                    'insert', i1, new_content[j1:j2]))
            elif op == 'replace':
                operations.append(DeltaOperation(
                    'replace', i1, new_content[j1:j2]))
        return operations

    def _store_object(self, delta: Delta) -> None:
        """Store a Delta object as compressed JSON."""
        object_path = self.objects_path / delta.hash
        json_data = json.dumps(asdict(delta), indent=2)
        with gzip.open(object_path, 'wt', encoding='utf-8') as f:
            f.write(json_data)

    def _load_object(self, hash_id: str) -> Optional[Delta]:
        """Load a Delta object from storage, converting operations to DeltaOperation instances."""
        object_path = self.objects_path / hash_id
        if not object_path.exists():
            return None
        try:
            with gzip.open(object_path, 'rt', encoding='utf-8') as f:
                data = json.load(f)
                # Convert operations from dicts to DeltaOperation instances
                operations = [DeltaOperation(**op)
                              for op in data['operations']]
                data['operations'] = operations
                return Delta(**data)
        except (gzip.BadGzipFile, json.JSONDecodeError, FileNotFoundError, KeyError) as e:
            raise ValueError(
                f"Failed to load delta {hash_id}: {str(e)}") from e

    def _update_ref(self, ref_name: str, hash_id: str) -> None:
        """Update a reference to point to a hash."""
        (self.refs_path / ref_name).write_text(hash_id)

    def _get_ref(self, ref_name: str) -> Optional[str]:
        """Get the hash a reference points to."""
        ref_path = self.refs_path / ref_name
        return ref_path.read_text().strip() if ref_path.exists() else None

    def _update_head(self, ref_name: str) -> None:
        """Update HEAD to point to a reference."""
        self.head_path.write_text(f"ref: {ref_name}")

    def _get_head(self) -> Optional[str]:
        """Get the current HEAD hash."""
        if not self.head_path.exists():
            return None
        head_content = self.head_path.read_text().strip()
        return self._get_ref(head_content[5:]) if head_content.startswith("ref: ") else head_content

    def _get_content(self, version_hash: str) -> List[str]:
        """Reconstruct content for a given version hash."""
        delta = self._load_object(version_hash)
        if not delta:
            raise ValueError(f"Version {version_hash} not found")

        # If this is a baseline, return its content directly
        if delta.baseline_hash == delta.hash:
            if not delta.operations or delta.operations[0].op != 'insert':
                return []
            # Full content stored in initial operation
            return delta.operations[0].content

        # Get baseline content and apply operations
        baseline_content = self._get_content(delta.baseline_hash)
        if not baseline_content:
            raise ValueError(f"Baseline {delta.baseline_hash} not found")

        content = baseline_content.copy()
        for op in delta.operations:
            if op.op == 'insert':
                content[op.position:op.position] = op.content
            elif op.op == 'delete':
                if op.position < len(content):
                    del content[op.position]
            elif op.op == 'replace':
                if op.position < len(content):
                    content[op.position:op.position + 1] = op.content
        return content

    def commit(self, docx_path: str, message: str) -> str:
        """Commit a new version of a DOCX file."""
        content = self._extract_content(docx_path)
        head = self._get_head()
        content_hash = self._calculate_hash(content)
        commit_count = len(self.get_history()) + 1

        # Determine if this should be a snapshot
        is_snapshot = (head is None) or (commit_count %
                                         self.SNAPSHOT_INTERVAL == 0)

        if is_snapshot:
            # Store full content as a baseline
            operations = [DeltaOperation('insert', 0, content)]
            delta = Delta(
                timestamp=datetime.now().isoformat(),
                parent_hash=head or "",
                baseline_hash=content_hash,  # Self-referential baseline
                operations=operations,
                message=message,
                hash=content_hash
            )
        else:
            # Compute incremental diff from parent
            parent_content = self._get_content(head) if head else []
            operations = self._compute_operations(parent_content, content)
            parent_delta = self._load_object(head) if head else None
            baseline_hash = parent_delta.baseline_hash if parent_delta else content_hash
            delta = Delta(
                timestamp=datetime.now().isoformat(),
                parent_hash=head or "",
                baseline_hash=baseline_hash,
                operations=operations,
                message=message,
                hash=content_hash
            )

        self._store_object(delta)
        self._update_ref("master", content_hash)
        self._update_head("master")
        return content_hash

    def get_history(self) -> List[Tuple[str, str, datetime]]:
        """Get the commit history."""
        history = []
        current = self._get_head()
        while current:
            delta = self._load_object(current)
            if delta:
                history.append((delta.hash, delta.message,
                               datetime.fromisoformat(delta.timestamp)))
                current = delta.parent_hash
            else:
                break
        return history

    def export_version(self, version_hash: str, output_path: str) -> None:
        """Export a version to a DOCX file."""
        content = self._get_content(version_hash)
        doc = Document()
        for paragraph_text in content:
            doc.add_paragraph(paragraph_text)
        doc.save(output_path)

    def show_diff(self, version_hash: str) -> Optional[List[str]]:
        """Show unified diff for a version (for compatibility)."""
        delta = self._load_object(version_hash)
        if not delta or not delta.parent_hash:
            return None
        parent_content = self._get_content(delta.parent_hash)
        current_content = self._get_content(version_hash)
        return list(difflib.unified_diff(parent_content, current_content, fromfile='previous', tofile='current', lineterm=''))

    def _generate_changes_summary(self, old_content: List[str], new_content: List[str]) -> str:
        """Generate a human-readable summary of changes."""
        old_len = len(old_content)
        new_len = len(new_content)
        added = len([line for line in new_content if line not in old_content])
        removed = len(
            [line for line in old_content if line not in new_content])
        summary_parts = []
        if added:
            summary_parts.append(
                f"{added} line{'s' if added != 1 else ''} added")
        if removed:
            summary_parts.append(
                f"{removed} line{'s' if removed != 1 else ''} removed")
        if not summary_parts:
            return "No text changes" if old_len == new_len else "Content rearranged"
        return "Changes: " + ", ".join(summary_parts)

    def show_friendly_diff(self, version_hash: str) -> str:
        """Show a human-readable diff for a version."""
        delta = self._load_object(version_hash)
        if not delta or not delta.parent_hash:
            return "Initial version"

        parent_content = self._get_content(delta.parent_hash)
        current_content = self._get_content(version_hash)
        differ = CombinedDiffer()
        result = [
            f"\n=== {self._generate_changes_summary(parent_content, current_content)} ===\n"]
        result.append(differ.format_combined_diff(
            parent_content, current_content))
        return '\n'.join(result)

    def compare_versions(self, hash1: str, hash2: str) -> List[str]:
        """Compare two versions and return a unified diff."""
        content1 = self._get_content(hash1)
        content2 = self._get_content(hash2)
        return list(difflib.unified_diff(content1, content2, fromfile='version1', tofile='version2', lineterm=''))
