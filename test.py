from delta import DocxVersionStore

repo = DocxVersionStore()

# Make some commits
hash1 = repo.commit("v1.docx", "Initial version")
hash2 = repo.commit("v2.docx", "Added new section")

# View history
print("Commit History:")
print(repo.get_history())

# See detailed changes
print("\nChanges in latest commit (Unified Diff):")
print(repo.show_diff(hash2))

# See friendly diff
print("\nChanges in latest commit (Friendly Diff):")
print(repo.show_friendly_diff(hash2))

# Export a version
repo.export_version(hash1, "document_v1_restored.docx")
print("\nVersion exported: document_v1_restored.docx")