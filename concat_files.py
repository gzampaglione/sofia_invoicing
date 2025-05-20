import os

# Set the root directory you want to scan (current folder or adjust as needed)
ROOT_DIR = "."

# Set output filename
OUTPUT_FILE = "combined_output.txt"

# File extensions to include
INCLUDE_EXTENSIONS = {".json", ".gs"}

with open(OUTPUT_FILE, "w", encoding="utf-8") as outfile:
    for root, _, files in os.walk(ROOT_DIR):
        for filename in files:
            ext = os.path.splitext(filename)[1]
            if ext in INCLUDE_EXTENSIONS:
                filepath = os.path.join(root, filename)
                # Write separator and filename
                relative_path = os.path.relpath(filepath, ROOT_DIR)
                outfile.write(f"==== {relative_path}\n")
                with open(filepath, "r", encoding="utf-8") as infile:
                    outfile.write(infile.read())
                outfile.write("\n\n")
