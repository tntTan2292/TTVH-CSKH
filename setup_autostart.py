import os
import winshell
from win32com.client import Dispatch

def create_startup_shortcut():
    startup_path = winshell.startup()
    project_path = r"d:\Antigravity - Project - TTVH\CSKH"
    vbs_path = os.path.join(project_path, "silent_start.vbs")
    
    shortcut_path = os.path.join(startup_path, "VNPostHueVIPDashboard.lnk")
    
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = "wscript.exe"
    shortcut.Arguments = f'"{vbs_path}"'
    shortcut.WorkingDirectory = project_path
    shortcut.IconLocation = "wscript.exe, 0"
    shortcut.Description = "VNPost Hue VIP Dashboard Autostart"
    shortcut.save()
    
    print(f"Successfully added to Startup: {shortcut_path}")

if __name__ == "__main__":
    try:
        create_startup_shortcut()
    except Exception as e:
        print(f"Error: {e}")
