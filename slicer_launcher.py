import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import sys
import subprocess
import platform

class SlicerLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("Slicer Launcher")
        self.root.geometry("500x400")
        self.root.resizable(False, False)
        
        # Get the directory where the executable/script is located
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            self.app_dir = os.path.dirname(os.path.abspath(sys.executable))
        else:
            # Running as script
            self.app_dir = os.path.dirname(os.path.abspath(__file__))
        
        self.config_file = os.path.join(self.app_dir, "slicer_config.json")
        self.config = self.load_config()
        self.target_file = None
        
        # Get target file from command line if provided
        if len(sys.argv) > 1:
            self.target_file = sys.argv[1]
        
        self.setup_ui()
        
    def load_config(self):
        """Load configuration from JSON file"""
        default_config = {
            "slicers": [
                {
                    "name": "Bambu Studio",
                    "path": "C:\\Program Files\\Bambu Studio\\bambu-studio.exe"
                },
                {
                    "name": "PrusaSlicer",
                    "path": "C:\\Program Files\\Prusa3D\\PrusaSlicer\\prusa-slicer.exe"
                },
                {
                    "name": "Orca Slicer",
                    "path": "C:\\Program Files\\OrcaSlicer\\orca-slicer.exe"
                }
            ],
            "file_extensions": [".3mf", ".stl", ".stp", ".step"]
        }
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    return json.load(f)
            else:
                # Create default config file
                with open(self.config_file, 'w') as f:
                    json.dump(default_config, f, indent=4)
                return default_config
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load config: {str(e)}")
            return default_config
    
    def setup_ui(self):
        """Setup the user interface"""
        # Menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Associate File Types", command=self.associate_files)
        file_menu.add_command(label="Restore Previous Associations", command=self.unassociate_files)
        file_menu.add_separator()
        file_menu.add_command(label="Launch", command=self.launch_slicer)
        file_menu.add_command(label="Close", command=self.root.quit)
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title label
        title_label = ttk.Label(main_frame, text="Select a Slicer", font=('Arial', 14, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 10))
        
        # Show target file if provided
        if self.target_file:
            file_label = ttk.Label(main_frame, text=f"File: {os.path.basename(self.target_file)}", 
                                  font=('Arial', 9))
            file_label.grid(row=1, column=0, columnspan=2, pady=(0, 10))
        
        # Listbox for slicers
        listbox_frame = ttk.Frame(main_frame)
        listbox_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL)
        self.slicer_listbox = tk.Listbox(listbox_frame, height=12, width=60, 
                                         yscrollcommand=scrollbar.set, font=('Arial', 10))
        scrollbar.config(command=self.slicer_listbox.yview)
        
        self.slicer_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Populate listbox
        for slicer in self.config.get("slicers", []):
            self.slicer_listbox.insert(tk.END, slicer["name"])
        
        # Select first item by default
        if self.slicer_listbox.size() > 0:
            self.slicer_listbox.selection_set(0)
        
        # Double-click to launch
        self.slicer_listbox.bind('<Double-Button-1>', lambda e: self.launch_slicer())
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=(10, 0))
        
        launch_btn = ttk.Button(button_frame, text="Launch", command=self.launch_slicer, width=15)
        launch_btn.pack(side=tk.LEFT, padx=5)
        
        close_btn = ttk.Button(button_frame, text="Close", command=self.root.quit, width=15)
        close_btn.pack(side=tk.LEFT, padx=5)
        
        # Info label
        info_label = ttk.Label(main_frame, 
                              text=f"Config file: {self.config_file}\nEdit this file to add/remove slicers",
                              font=('Arial', 8), foreground='gray')
        info_label.grid(row=4, column=0, columnspan=2, pady=(20, 0))
        
    def launch_slicer(self):
        """Launch the selected slicer"""
        selection = self.slicer_listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a slicer to launch")
            return
        
        selected_index = selection[0]
        slicer = self.config["slicers"][selected_index]
        slicer_path = slicer["path"]
        
        if not os.path.exists(slicer_path):
            messagebox.showerror("Error", 
                               f"Slicer not found at:\n{slicer_path}\n\nPlease update the path in {self.config_file}")
            return
        
        try:
            if self.target_file:
                # Launch slicer with the target file
                if platform.system() == "Windows":
                    subprocess.Popen([slicer_path, self.target_file])
                else:
                    subprocess.Popen([slicer_path, self.target_file])
            else:
                # Launch slicer without a file
                if platform.system() == "Windows":
                    subprocess.Popen([slicer_path])
                else:
                    subprocess.Popen([slicer_path])
            
            self.root.quit()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch slicer:\n{str(e)}")
    
    def associate_files(self):
        """Associate file types with this application (Windows only)"""
        if platform.system() != "Windows":
            messagebox.showinfo("Info", 
                              "File association is currently only supported on Windows.\n" +
                              "On Linux/Mac, please use your system's file association settings.")
            return
        
        try:
            import winreg
            
            # Get the path to this executable
            if getattr(sys, 'frozen', False):
                # Running as compiled executable
                app_path = os.path.abspath(sys.executable)
            else:
                # Running as script (for testing)
                app_path = os.path.abspath(sys.argv[0])
                messagebox.showwarning("Warning", 
                                     "You are running as a Python script.\n" +
                                     "File associations work best with the compiled .exe version.\n\n" +
                                     "Proceeding anyway for testing purposes.")
            
            extensions = self.config.get("file_extensions", [])
            associated = []
            failed = []
            
            for ext in extensions:
                try:
                    # Ensure extension starts with a dot
                    if not ext.startswith('.'):
                        ext = '.' + ext
                    
                    # Create ProgID
                    prog_id = "SlicerLauncher.File"
                    
                    # Set default program for extension
                    ext_key_path = f"Software\\Classes\\{ext}"
                    ext_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, ext_key_path)
                    winreg.SetValueEx(ext_key, "", 0, winreg.REG_SZ, prog_id)
                    winreg.CloseKey(ext_key)
                    
                    # Create OpenWithProgids key for better Windows 10/11 compatibility
                    open_with_path = f"Software\\Classes\\{ext}\\OpenWithProgids"
                    open_with_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, open_with_path)
                    winreg.SetValueEx(open_with_key, prog_id, 0, winreg.REG_NONE, b'')
                    winreg.CloseKey(open_with_key)
                    
                    # Create ProgID key with description
                    prog_key_path = f"Software\\Classes\\{prog_id}"
                    prog_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, prog_key_path)
                    winreg.SetValueEx(prog_key, "", 0, winreg.REG_SZ, "3D Model File for Slicer Launcher")
                    winreg.SetValueEx(prog_key, "FriendlyTypeName", 0, winreg.REG_SZ, "3D Model File")
                    winreg.CloseKey(prog_key)
                    
                    # Set default icon (uses first icon in exe)
                    icon_path = f"Software\\Classes\\{prog_id}\\DefaultIcon"
                    icon_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, icon_path)
                    winreg.SetValueEx(icon_key, "", 0, winreg.REG_SZ, f"{app_path},0")
                    winreg.CloseKey(icon_key)
                    
                    # Create shell command
                    command_path = f"Software\\Classes\\{prog_id}\\shell\\open\\command"
                    command_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, command_path)
                    command = f'"{app_path}" "%1"'
                    winreg.SetValueEx(command_key, "", 0, winreg.REG_SZ, command)
                    winreg.CloseKey(command_key)
                    
                    # Add to user's choice list
                    user_choice_path = f"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\{ext}\\OpenWithProgids"
                    try:
                        user_choice_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, user_choice_path)
                        winreg.SetValueEx(user_choice_key, prog_id, 0, winreg.REG_NONE, b'')
                        winreg.CloseKey(user_choice_key)
                    except:
                        pass  # This key might not exist, which is okay
                    
                    associated.append(ext)
                    
                except Exception as e:
                    failed.append(f"{ext}: {str(e)}")
                    print(f"Failed to associate {ext}: {str(e)}")
            
            # Show results
            message = ""
            if associated:
                message += f"Successfully associated:\n{', '.join(associated)}\n\n"
            if failed:
                message += f"Failed to associate:\n" + "\n".join(failed) + "\n\n"
            
            message += "Changes should take effect immediately.\n"
            message += "If file icons don't update, try:\n"
            message += "1. Restart File Explorer (Task Manager → Restart 'Windows Explorer')\n"
            message += "2. Log out and log back in\n"
            message += "3. Restart your computer"
            
            if associated:
                messagebox.showinfo("File Association Complete", message)
            else:
                messagebox.showerror("File Association Failed", message)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create file associations:\n{str(e)}\n\n" +
                               "Try running the application as Administrator.")
    
    def unassociate_files(self):
        """Remove file type associations with this application (Windows only)"""
        if platform.system() != "Windows":
            messagebox.showinfo("Info", 
                              "File association management is currently only supported on Windows.")
            return
        
        result = messagebox.askyesno("Confirm Unassociation", 
                                     "This will remove Slicer Launcher associations for all configured file types.\n\n" +
                                     "The files will revert to their previous default programs.\n\n" +
                                     "Continue?")
        if not result:
            return
        
        try:
            import winreg
            
            extensions = self.config.get("file_extensions", [])
            removed = []
            failed = []
            prog_id = "SlicerLauncher.File"
            
            for ext in extensions:
                try:
                    # Ensure extension starts with a dot
                    if not ext.startswith('.'):
                        ext = '.' + ext
                    
                    # Remove the default association if it points to our ProgID
                    try:
                        ext_key_path = f"Software\\Classes\\{ext}"
                        ext_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, ext_key_path, 0, winreg.KEY_READ | winreg.KEY_WRITE)
                        current_prog_id, _ = winreg.QueryValueEx(ext_key, "")
                        
                        if current_prog_id == prog_id:
                            winreg.DeleteValue(ext_key, "")
                            removed.append(ext)
                        
                        winreg.CloseKey(ext_key)
                    except WindowsError:
                        pass  # Key doesn't exist or no permission
                    
                    # Remove from OpenWithProgids
                    try:
                        open_with_path = f"Software\\Classes\\{ext}\\OpenWithProgids"
                        open_with_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, open_with_path, 0, winreg.KEY_WRITE)
                        winreg.DeleteValue(open_with_key, prog_id)
                        winreg.CloseKey(open_with_key)
                    except WindowsError:
                        pass  # Value doesn't exist
                    
                    # Remove from user's choice list
                    try:
                        user_choice_path = f"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\{ext}\\OpenWithProgids"
                        user_choice_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, user_choice_path, 0, winreg.KEY_WRITE)
                        winreg.DeleteValue(user_choice_key, prog_id)
                        winreg.CloseKey(user_choice_key)
                    except WindowsError:
                        pass  # Value doesn't exist
                    
                except Exception as e:
                    failed.append(f"{ext}: {str(e)}")
            
            # Try to remove the ProgID entirely
            try:
                prog_key_path = f"Software\\Classes\\{prog_id}"
                self._delete_registry_key(winreg.HKEY_CURRENT_USER, prog_key_path)
            except:
                pass  # Key might not exist or have subkeys
            
            # Show results
            message = ""
            if removed:
                message += f"Removed associations for:\n{', '.join(removed)}\n\n"
            if failed:
                message += f"Failed to remove:\n" + "\n".join(failed) + "\n\n"
            
            if not removed and not failed:
                message = "No Slicer Launcher associations were found.\n\n"
            
            message += "Changes should take effect immediately.\n"
            message += "You may need to restart File Explorer for icons to update."
            
            messagebox.showinfo("Unassociation Complete", message)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to remove file associations:\n{str(e)}")
    
    def _delete_registry_key(self, hkey, key_path):
        """Recursively delete a registry key and all its subkeys"""
        try:
            key = winreg.OpenKey(hkey, key_path, 0, winreg.KEY_ALL_ACCESS)
            
            # Delete all subkeys first
            subkeys = []
            i = 0
            while True:
                try:
                    subkeys.append(winreg.EnumKey(key, i))
                    i += 1
                except WindowsError:
                    break
            
            for subkey in subkeys:
                self._delete_registry_key(hkey, f"{key_path}\\{subkey}")
            
            winreg.CloseKey(key)
            winreg.DeleteKey(hkey, key_path)
        except WindowsError:
            pass  # Key doesn't exist or can't be deleted

def main():
    root = tk.Tk()
    app = SlicerLauncher(root)
    root.mainloop()

if __name__ == "__main__":
    main()