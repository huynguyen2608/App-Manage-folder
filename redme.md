// build
python -m PyInstaller --onefile --windowed --add-data "folder_icon.png;." main.py --name "Manage Folder"
python -m pyinstaller --onefile --windowed --add-data "openai_key.txt:." main.py