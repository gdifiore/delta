import argparse
from delta import DocxVersionStore

def init_repo():
    store = DocxVersionStore()
    print("Initialized new Docx version repository.")

def commit_version(docx_path, message):
    store = DocxVersionStore()
    commit_hash = store.commit(docx_path, message)
    print(f"Committed new version: {commit_hash}")

def show_history():
    store = DocxVersionStore()
    history = store.get_history()
    print(history)
    for hash_id, msg, timestamp in history:
        print(f"{timestamp}: {hash_id} - {msg}")

def export_version(version_hash, output_path):
    store = DocxVersionStore()
    store.export_version(version_hash, output_path)
    print(f"Exported version {version_hash} to {output_path}")

def show_diff(version_hash):
    store = DocxVersionStore()
    diff = store.show_friendly_diff(version_hash)
    print(diff)

def main():
    parser = argparse.ArgumentParser(description="Docx Version Control CLI")
    subparsers = parser.add_subparsers(dest="command")

    subparsers.add_parser("init", help="Initialize a new repository")

    commit_parser = subparsers.add_parser("commit", help="Commit a new version")
    commit_parser.add_argument("docx_path", help="Path to the DOCX file")
    commit_parser.add_argument("-m", "--message", required=True, help="Commit message")

    subparsers.add_parser("history", help="Show version history")

    export_parser = subparsers.add_parser("export", help="Export a version")
    export_parser.add_argument("version_hash", help="Version hash to export")
    export_parser.add_argument("output_path", help="Output file path")

    diff_parser = subparsers.add_parser("diff", help="Show differences")
    diff_parser.add_argument("version_hash", help="Version hash to compare")

    args = parser.parse_args()

    if args.command == "init":
        init_repo()
    elif args.command == "commit":
        commit_version(args.docx_path, args.message)
    elif args.command == "history":
        show_history()
    elif args.command == "export":
        export_version(args.version_hash, args.output_path)
    elif args.command == "diff":
        show_diff(args.version_hash)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
