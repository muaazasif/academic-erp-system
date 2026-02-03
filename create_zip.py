import os
import zipfile
from pathlib import Path

def create_project_zip():
    """Create a zip file of the entire project for manual GitHub upload"""
    
    # Define the source directory (current directory)
    source_dir = Path.cwd()
    
    # Define the output zip file name
    zip_filename = "academic-erp-system.zip"
    
    # Files and directories to include
    include_patterns = [
        "*.py",           # Python files
        "requirements.txt",
        "README.md",
        "LICENSE",
        ".gitignore",
        "Dockerfile",
        "railway.json",
        "static/*",       # Static files
        "templates/*",    # Template files
        "*.md",           # Documentation files
    ]
    
    # Files and directories to exclude
    exclude_patterns = [
        ".git/",
        "__pycache__/",
        "*.pyc",
        ".vscode/",
        ".idea/",
        "*__pycache__*",
        "venv/",
        "env/",
        "*.log",
        "credentials.json",
        "token.pickle",
        "cookies.txt",
        "student_cookies.txt",
        ".DS_Store",
        "Thumbs.db"
    ]
    
    print(f"Creating zip file: {zip_filename}")
    
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zipf:
        for root, dirs, files in os.walk(source_dir):
            # Skip excluded directories
            dirs[:] = [d for d in dirs if not any(exclude in os.path.join(root, d) for exclude in exclude_patterns)]
            
            for file in files:
                file_path = Path(root) / file
                
                # Skip excluded files
                if any(exclude in str(file_path) for exclude in exclude_patterns):
                    continue
                
                # Include only the files we want
                if any(str(file_path).endswith(pattern.lstrip("*")) or 
                       pattern.startswith("*") and str(file_path).endswith(pattern[1:]) or
                       str(file_path).endswith('.py') or
                       file in ['requirements.txt', 'README.md', 'LICENSE', '.gitignore', 'Dockerfile', 'railway.json']):
                    
                    # Create archive name (relative to source directory)
                    archive_name = file_path.relative_to(source_dir)
                    
                    # Add to zip
                    zipf.write(file_path, archive_name)
                    print(f"Added: {archive_name}")
    
    print(f"\nZip file '{zip_filename}' created successfully!")
    print(f"File size: {os.path.getsize(zip_filename)} bytes")
    print("\nTo upload to GitHub:")
    print("1. Go to https://github.com/muaazasif/academic-erp-system")
    print("2. Click 'Add file' → 'Upload files'")
    print("3. Drag and drop the zip file or click to select it")
    print("4. Commit the changes")

if __name__ == "__main__":
    create_project_zip()