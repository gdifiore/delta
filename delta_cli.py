import argparse
import sys
from delta import DocxVersionStore


def init_repo(store: DocxVersionStore) -> None:
    #store = DocxVersionStore()
    print("Initialized new Docx version repository.")


def commit_version(store: DocxVersionStore, docx_path: str, message: str) -> None:
    try:
        commit_hash = store.commit(docx_path, message)
        print(f"Committed new version: {commit_hash}")
    except Exception as e:
        print(f"Error committing version: {e}", file=sys.stderr)
        sys.exit(1)


def show_history(store: DocxVersionStore) -> None:
    try:
        history = store.get_history()
        if not history:
            print("No history available.")
            return
        for hash_id, msg, timestamp in history:
            print(f"{timestamp}: {hash_id} - {msg}")
    except Exception as e:
        print(f"Error showing history: {e}", file=sys.stderr)
        sys.exit(1)


def export_version(store: DocxVersionStore, version_hash: str, output_path: str) -> None:
    try:
        store.export_version(version_hash, output_path)
        print(f"Exported version {version_hash} to {output_path}")
    except Exception as e:
        print(f"Error exporting version: {e}", file=sys.stderr)
        sys.exit(1)


def show_diff(store: DocxVersionStore, version_hash: str) -> None:
    try:
        diff = store.show_friendly_diff(version_hash)
        print(diff)
    except Exception as e:
        print(f"Error showing diff: {e}", file=sys.stderr)
        sys.exit(1)


def main() -> None:
    store = DocxVersionStore()
    parser = argparse.ArgumentParser(description="Docx Version Control CLI")
    subparsers = parser.add_subparsers(dest="command")

    subparsers.add_parser("init", help="Initialize a new repository")

    commit_parser = subparsers.add_parser(
        "commit", help="Commit a new version")
    commit_parser.add_argument("docx_path", help="Path to the DOCX file")
    commit_parser.add_argument(
        "-m", "--message", required=True, help="Commit message")

    subparsers.add_parser("history", help="Show version history")

    export_parser = subparsers.add_parser("export", help="Export a version")
    export_parser.add_argument("version_hash", help="Version hash to export")
    export_parser.add_argument("output_path", help="Output file path")

    diff_parser = subparsers.add_parser("diff", help="Show differences")
    diff_parser.add_argument("version_hash", help="Version hash to compare")

    args = parser.parse_args()

    commands = {
        "init": lambda: init_repo(store),
        "commit": lambda: commit_version(store, args.docx_path, args.message),
        "history": lambda: show_history(store),
        "export": lambda: export_version(store, args.version_hash, args.output_path),
        "diff": lambda: show_diff(store, args.version_hash),
    }

    if args.command in commands:
        commands[args.command]()
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
