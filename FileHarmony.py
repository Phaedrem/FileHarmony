######################
# Author: Darren Bowers
# Program: FileHarmony
# Purpose: Choose folders containing mp3/wav/flac files or select specific files to rename to the title in the metadata if necessary and available. 
######################


import os  # Provides functions to interact with the operating system
from PyQt6.QtGui import QIcon  # Imports QIcon for setting the app window icon
from PyQt6.QtWidgets import (  # Core PyQt widgets for building the UI
    QApplication, QMainWindow, QWidget, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QPushButton, QMessageBox)
from PyQt6.QtCore import Qt  # Qt enums used for item flags and states
from mutagen.easyid3 import EasyID3  # Reads ID3 metadata from MP3 files
from mutagen.flac import FLAC  # Reads metadata from FLAC files
from mutagen.wave import WAVE  # Reads metadata from WAV files
import sys  # Access to runtime system info and arguments
import ctypes  # Used to interact with Windows API for drive listing

SUPPORTED_EXTS = ('.mp3', '.flac', '.wav')  # Supported audio file extensions

class AudioRenamer(QMainWindow):  # Main window class inheriting from QMainWindow
    def __init__(self):
        super().__init__()  # Call base class constructor
        self.setWindowTitle("FileHarmony")  # Set window title

        # Set application icon depending on execution context
        import sys, os
        if getattr(sys, 'frozen', False):  # If running as a PyInstaller EXE
            base_path = sys._MEIPASS  # Use temp dir where bundled files are extracted
        else:
            base_path = os.path.abspath(".")  # Otherwise use current dir

        icon_path = os.path.join(base_path, "rename.ico")  # Build icon path
        self.setWindowIcon(QIcon(icon_path))  # Set the window icon

        self.tree = QTreeWidget()  # Tree view widget for folders/files
        self.tree.setHeaderLabels(["Drives and Folders"])  # Header for the tree
        self.tree.setColumnCount(1)  # Only one column needed
        self.tree.itemChanged.connect(self.on_item_changed)  # Connect check change event
        self.tree.itemExpanded.connect(self.on_item_expanded)  # Connect expand event

        self.rename_button = QPushButton("Rename Checked Files")  # Set button name
        self.rename_button.clicked.connect(self.rename_files)  # Connect button to action

        layout = QVBoxLayout()  # Create vertical layout
        layout.addWidget(self.tree)  # Add tree to layout
        layout.addWidget(self.rename_button)  # Add button to layout

        container = QWidget()  # Create central container widget
        container.setLayout(layout)  # Set layout on container
        self.setCentralWidget(container)  # Set container as main content

        self.load_all_drives()  # Load initial drive list into tree

    def load_all_drives(self):  # Populate the tree with available drives
        drives = self.list_drives()
        for drive in drives:
            drive_item = QTreeWidgetItem([drive])  # Create item for drive
            drive_item.setData(0, Qt.ItemDataRole.UserRole, drive)  # Store full path
            drive_item.setFlags(drive_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)  # Make item checkable
            drive_item.setCheckState(0, Qt.CheckState.Unchecked)  # Default unchecked
            drive_item.addChild(QTreeWidgetItem(["Loading..."]))  # Placeholder child
            self.tree.addTopLevelItem(drive_item)  # Add to top level in tree

    def list_drives(self):  # Get all logical drives using Windows API
        drives = []
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()  # Get bitmask of drives
        for i in range(26):  # Iterate over letters A-Z
            if bitmask & (1 << i):  # If bit is set for this drive
                drives.append(f"{chr(65 + i)}:\\")  # Add drive letter
        return drives

    def on_item_expanded(self, item):  # Triggered when a folder item is expanded
        path = item.data(0, Qt.ItemDataRole.UserRole)  # Get full path
        if not path or item.childCount() > 0 and item.child(0).text(0) != "Loading...":
            return  # Skip if already loaded or no path

        item.takeChildren()  # Remove placeholder
        self.populate_tree(path, item)  # Populate with actual folder contents

        parent_state = item.checkState(0)  # Get parent item state
        if parent_state in (Qt.CheckState.Checked, Qt.CheckState.PartiallyChecked):
            for i in range(item.childCount()):  # Propagate check state to children
                child = item.child(i)
                child.setCheckState(0, parent_state)

    def populate_tree(self, path, parent_item):  # Populate folder contents into tree
        try:
            for entry in sorted(os.scandir(path), key=lambda e: (not e.is_dir(), e.name.lower())):
                if entry.is_file() and not entry.name.lower().endswith(SUPPORTED_EXTS):
                    continue  # Skip unsupported files

                item = QTreeWidgetItem([entry.name])  # Create item for file/folder
                item.setData(0, Qt.ItemDataRole.UserRole, entry.path)  # Store full path
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)  # Make checkable
                item.setCheckState(0, Qt.CheckState.Unchecked)  # Default unchecked
                parent_item.addChild(item)  # Add as child

                if entry.is_dir():  # Add lazy load placeholder if folder
                    item.addChild(QTreeWidgetItem(["Loading..."]))
        except Exception as e:
            print(f"Error scanning {path}: {e}")  # Print error if scan fails

    def on_item_changed(self, item, column):  # Sync child checkboxes when parent is toggled
        state = item.checkState(0)
        for i in range(item.childCount()):
            child = item.child(i)
            child.setCheckState(0, state)

    def find_item_by_path(self, path):  # Locate a tree item by its stored full path
        def recurse(item):  # Recursive helper function
            if item.data(0, Qt.ItemDataRole.UserRole) == path:
                return item
            for i in range(item.childCount()):
                found = recurse(item.child(i))
                if found:
                    return found
            return None

        for i in range(self.tree.topLevelItemCount()):
            top_item = self.tree.topLevelItem(i)
            found = recurse(top_item)
            if found:
                return found
        return None

    def rename_files(self):  # Rename all checked audio files to match their title metadata
        to_rename = []  # List of (item, path) pairs to rename
        folders_to_refresh = set()  # Folders to refresh in the tree view

        def gather_checked(item):  # Recursively collect checked files/folders
            path = item.data(0, Qt.ItemDataRole.UserRole)
            if path is None:
                return

            if item.checkState(0) == Qt.CheckState.Checked:
                if os.path.isfile(path) and path.lower().endswith(SUPPORTED_EXTS):
                    to_rename.append((item, path))
                elif os.path.isdir(path):
                    for root, _, files in os.walk(path):
                        for name in files:
                            if name.lower().endswith(SUPPORTED_EXTS):
                                full_path = os.path.join(root, name)
                                tree_item = self.find_item_by_path(full_path)
                                to_rename.append((tree_item, full_path))
                    return

            for i in range(item.childCount()):
                gather_checked(item.child(i))

        for i in range(self.tree.topLevelItemCount()):
            gather_checked(self.tree.topLevelItem(i))

        renamed = 0  # Counter for successful renames
        for item, path in to_rename:
            folder, file_name = os.path.split(path)
            ext = os.path.splitext(file_name)[1].lower()

            try:
                if ext == '.mp3':
                    audio = EasyID3(path)
                    title = ''.join(audio.get("title", [file_name[:-4]]))
                elif ext == '.flac':
                    audio = FLAC(path)
                    title = ''.join(audio.get("title", [file_name[:-5]]))
                elif ext == '.wav':
                    audio = WAVE(path)
                    title = ''.join(audio.tags.get("TIT2", [file_name[:-4]]))
                else:
                    continue
                new_path = os.path.join(folder, title + ext)
            except:
                continue  # Skip file if metadata can't be read

            try:
                if not os.path.exists(new_path):  # Only rename if target doesn't exist
                    os.rename(path, new_path)
                    folders_to_refresh.add(folder)  # Mark folder for refresh
                    if item is not None:
                        item.setText(0, os.path.basename(new_path))  # Update tree item label
                        item.setData(0, Qt.ItemDataRole.UserRole, new_path)  # Update stored path
                    renamed += 1
            except:
                continue  # Skip file if rename fails

        QMessageBox.information(self, "Done", f"Renamed {renamed} file(s).")  # Notify user

        for folder in folders_to_refresh:
            item = self.find_item_by_path(folder)
            if item:
                item.takeChildren()  # Clear folder contents in tree
                self.populate_tree(folder, item)  # Reload with updated contents

if __name__ == '__main__':  # Main entry point for the script
    app = QApplication(sys.argv)  # Initialize Qt application
    window = AudioRenamer()  # Create main window
    window.resize(800, 600)  # Set initial size
    window.show()  # Show the window
    sys.exit(app.exec())  # Start the Qt event loop