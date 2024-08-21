import os

attachments_dir = "C:\\temp\\attachments"

def remove_new_files(directory):
    """Remove all files that have '_new' in their filenames."""
    for root, dirs, files in os.walk(directory):
        for file in files:
            if '_new' in file:
                file_path = os.path.join(root, file)
                print(f"Removing file: {file_path}")
                os.remove(file_path)

if __name__ == "__main__":
    remove_new_files(attachments_dir)
    print("All '_new' files have been removed.")
