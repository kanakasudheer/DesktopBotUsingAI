import os
import json
import datetime
import subprocess
import getpass
import threading
import webbrowser
import ctypes
import time
import math
import requests
import random
import queue
import traceback
import pygame
from pygame.locals import *
from OpenGL.GL import *
from OpenGL.GLU import *
import numpy as np
from matplotlib import animation
from effects import window_shake_effect

import speech_recognition as sr
r = sr.Recognizer()

import pyttsx3
import wikipedia
import google.generativeai as genai
import tkinter as tk
from tkinter import scrolledtext, simpledialog, filedialog, messagebox, ttk
import pyautogui
import psutil
import win32gui
import win32con
import winsound
from io import BytesIO
from PIL import Image, ImageTk
import cv2

# ---------- 3D Circle Animation Class ----------
class Circle3D:
    def __init__(self, parent):
        self.parent = parent
        self.canvas = parent.canvas
        self.center_x = parent.center_x
        self.center_y = parent.center_y
        self.radius = 80
        self.segments = 36
        self.rotation_x = 0
        self.rotation_y = 0
        self.rotation_z = 0
        self.circle_points = []
        self.circle_lines = []
        self.animation_speed = 0.005  # Slower animation speed
        
        # Create 3D circle points
        self.create_circle_points()
        
        # Start animation
        self.animate()
    
    def create_circle_points(self):
        """Create 3D circle points"""
        self.circle_points = []
        for i in range(self.segments):
            angle = (2 * math.pi * i) / self.segments
            x = self.radius * math.cos(angle)
            y = self.radius * math.sin(angle)
            z = 0
            self.circle_points.append([x, y, z])
    
    def rotate_point(self, x, y, z, rx, ry, rz):
        """Rotate a point in 3D space"""
        # Rotate around X axis
        cos_x = math.cos(rx)
        sin_x = math.sin(rx)
        y1 = y * cos_x - z * sin_x
        z1 = y * sin_x + z * cos_x
        
        # Rotate around Y axis
        cos_y = math.cos(ry)
        sin_y = math.sin(ry)
        x2 = x * cos_y + z1 * sin_y
        z2 = -x * sin_y + z1 * cos_y
        
        # Rotate around Z axis
        cos_z = math.cos(rz)
        sin_z = math.sin(rz)
        x3 = x2 * cos_z - y1 * sin_z
        y3 = x2 * sin_z + y1 * cos_z
        
        return x3, y3, z2
    
    def project_point(self, x, y, z):
        """Project 3D point to 2D screen coordinates"""
        # Simple perspective projection
        distance = 200
        factor = distance / (distance + z)
        screen_x = self.center_x + x * factor
        screen_y = self.center_y + y * factor
        return screen_x, screen_y
    
    def draw_circle(self):
        """Draw the 3D circle"""
        try:
            # Clear previous circle
            for line in self.circle_lines:
                try:
                    self.canvas.delete(line)
                except:
                    pass  # Ignore errors when deleting lines
            self.circle_lines = []
        
            # Project and draw points
            projected_points = []
            for point in self.circle_points:
                # Rotate the point
                rx, ry, rz = self.rotate_point(point[0], point[1], point[2], 
                                             self.rotation_x, self.rotation_y, self.rotation_z)
                # Project to 2D
                screen_x, screen_y = self.project_point(rx, ry, rz)
                projected_points.append((screen_x, screen_y))
            
            # Draw circle lines
            for i in range(len(projected_points)):
                start_point = projected_points[i]
                end_point = projected_points[(i + 1) % len(projected_points)]
                
                # Calculate line color based on Z position (depth)
                z1 = self.rotate_point(self.circle_points[i][0], self.circle_points[i][1], self.circle_points[i][2],
                                     self.rotation_x, self.rotation_y, self.rotation_z)[2]
                z2 = self.rotate_point(self.circle_points[(i + 1) % len(self.circle_points)][0], 
                                     self.circle_points[(i + 1) % len(self.circle_points)][1], 
                                     self.circle_points[(i + 1) % len(self.circle_points)][2],
                                     self.rotation_x, self.rotation_y, self.rotation_z)[2]
                
                avg_z = (z1 + z2) / 2
                # Color gradient from cyan to blue based on depth
                intensity = max(0, min(255, int(255 * (avg_z + self.radius) / (2 * self.radius))))
                color = f'#{0:02x}{intensity:02x}{255:02x}'  # Cyan to blue gradient
                
                line = self.canvas.create_line(
                    start_point[0], start_point[1], end_point[0], end_point[1],
                    fill=color, width=2, tags='hud'
                )
                self.circle_lines.append(line)
            
            # Draw center point
            center_line = self.canvas.create_oval(
                self.center_x - 3, self.center_y - 3,
                self.center_x + 3, self.center_y + 3,
                fill='#00FFFF', outline='#00FFFF', tags='hud'
            )
            self.circle_lines.append(center_line)
        except Exception as e:
            print(f"Error in draw_circle: {e}")
            pass
    
    def animate(self):
        """Animate the 3D circle"""
        try:
            # Update rotations
            self.rotation_x += self.animation_speed
            self.rotation_y += self.animation_speed * 0.7
            self.rotation_z += self.animation_speed * 0.3
            
            # Draw the circle
            self.draw_circle()
        except Exception as e:
            print(f"Error in 3D circle animation: {e}")
            # Stop animation on error
            return
        
        # Schedule next animation frame using the canvas's after method
        if hasattr(self.canvas, 'after'):
            self.canvas.after(100, self.animate)  # Slower refresh rate
        else:
            # Fallback: use the parent's root window
            if hasattr(self.parent, 'root'):
                self.parent.root.after(100, self.animate)  # Slower refresh rate
            else:
                # If no after method available, just continue animation
                import threading
                import time
                def delayed_animate():
                    time.sleep(0.1)  # Slower sleep
                    self.animate()
                threading.Thread(target=delayed_animate, daemon=True).start()

# ---------- System File Explorer Integration ----------
class SystemFileExplorer:
    def __init__(self):
        self.current_path = os.path.expanduser("~")
        
    def open_explorer(self, path=None):
        """Open Windows Explorer at specified path"""
        try:
            if path is None:
                path = self.current_path
            if not os.path.exists(path):
                path = os.path.expanduser("~")
            
            # Use subprocess to open Windows Explorer
            subprocess.run(['explorer', path], check=True)
            self.current_path = path
            return f"Opened File Explorer at: {path}"
        except Exception as e:
            return f"Error opening File Explorer: {str(e)}"
            
    def open_desktop(self):
        """Open File Explorer at Desktop"""
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        return self.open_explorer(desktop_path)
        
    def open_documents(self):
        """Open File Explorer at Documents"""
        documents_path = os.path.join(os.path.expanduser("~"), "Documents")
        return self.open_explorer(documents_path)
        
    def open_downloads(self):
        """Open File Explorer at Downloads"""
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        return self.open_explorer(downloads_path)
        
    def open_pictures(self):
        """Open File Explorer at Pictures"""
        pictures_path = os.path.join(os.path.expanduser("~"), "Pictures")
        return self.open_explorer(pictures_path)
        
    def open_music(self):
        """Open File Explorer at Music"""
        music_path = os.path.join(os.path.expanduser("~"), "Music")
        return self.open_explorer(music_path)
        
    def open_videos(self):
        """Open File Explorer at Videos"""
        videos_path = os.path.join(os.path.expanduser("~"), "Videos")
        return self.open_explorer(videos_path)
        
    def open_drive(self, drive_letter):
        """Open File Explorer at specific drive"""
        drive_path = f"{drive_letter.upper()}:\\"
        if os.path.exists(drive_path):
            return self.open_explorer(drive_path)
        else:
            return f"Drive {drive_letter.upper()}: not found"
            
    def open_path(self, path):
        """Open File Explorer at custom path"""
        if os.path.exists(path):
            return self.open_explorer(path)
        else:
            return f"Path not found: {path}"
            
    def create_folder(self, folder_name, path=None):
        """Create a new folder using Windows Explorer"""
        try:
            print(f"Creating folder: {folder_name} in path: {path}")  # Debug log
            
            if path is None:
                # Ask user where to create the folder
                try:
                    import tkinter as tk
                    from tkinter import simpledialog
                    
                    # Create a simple dialog to ask for location
                    root = tk.Tk()
                    root.withdraw()  # Hide the main window
                    
                    # Show options dialog
                    options = [
                        "Desktop",
                        "Documents", 
                        "Downloads",
                        "Pictures",
                        "Music",
                        "Videos",
                        "Current location"
                    ]
                    
                    location = simpledialog.askstring(
                        "Create Folder", 
                        f"Where would you like to create '{folder_name}'?\n\n"
                        f"Options: Desktop, Documents, Downloads, Pictures, Music, Videos, or Current location",
                        initialvalue="Desktop"
                    )
                    
                    root.destroy()
                    
                    if not location:
                        return "Folder creation cancelled"
                    
                    location = location.lower()
                    
                    if location == "desktop":
                        path = os.path.join(os.path.expanduser("~"), "Desktop")
                    elif location == "documents":
                        path = os.path.join(os.path.expanduser("~"), "Documents")
                    elif location == "downloads":
                        path = os.path.join(os.path.expanduser("~"), "Downloads")
                    elif location == "pictures":
                        path = os.path.join(os.path.expanduser("~"), "Pictures")
                    elif location == "music":
                        path = os.path.join(os.path.expanduser("~"), "Music")
                    elif location == "videos":
                        path = os.path.join(os.path.expanduser("~"), "Videos")
                    elif location == "current location":
                        path = self.current_path
                    else:
                        return f"Unknown location: {location}. Please choose from: Desktop, Documents, Downloads, Pictures, Music, Videos, or Current location"
                        
                except Exception as dialog_error:
                    print(f"Dialog error: {dialog_error}")  # Debug log
                    # Fallback to desktop if dialog fails
                    path = os.path.join(os.path.expanduser("~"), "Desktop")
                    print(f"Using fallback path: {path}")  # Debug log
            
            # Validate path exists
            if not os.path.exists(path):
                print(f"Path does not exist: {path}")  # Debug log
                return f"Path does not exist: {path}"
            
            new_folder_path = os.path.join(path, folder_name)
            print(f"Creating folder at: {new_folder_path}")  # Debug log
            
            # Create the folder
            os.makedirs(new_folder_path, exist_ok=True)
            
            # Verify folder was created
            if os.path.exists(new_folder_path):
                print(f"Folder created successfully: {new_folder_path}")  # Debug log
                # Open the parent folder to show the new folder
                try:
                    self.open_explorer(path)
                except Exception as explorer_error:
                    print(f"Explorer error: {explorer_error}")  # Debug log
                    # Continue even if explorer fails
                
                return f"Created folder '{folder_name}' in {os.path.basename(path)}"
            else:
                print(f"Folder creation failed: {new_folder_path}")  # Debug log
                return f"Failed to create folder '{folder_name}'"
                
        except PermissionError as pe:
            print(f"Permission error: {pe}")  # Debug log
            return f"Permission denied: Cannot create folder '{folder_name}' in {os.path.basename(path) if path else 'unknown location'}"
        except OSError as ose:
            print(f"OS error: {ose}")  # Debug log
            return f"OS Error creating folder '{folder_name}': {str(ose)}"
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Unexpected error in create_folder: {error_details}")  # Debug log
            return f"Error creating folder: {str(e)}"
            
    def delete_folder(self, folder_name, path=None):
        """Delete a folder using Windows Explorer"""
        try:
            if path is None:
                # Ask user where the folder is located
                import tkinter as tk
                from tkinter import simpledialog
                
                root = tk.Tk()
                root.withdraw()
                
                location = simpledialog.askstring(
                    "Delete Folder", 
                    f"Where is the folder '{folder_name}' located?\n\n"
                    f"Options: Desktop, Documents, Downloads, Pictures, Music, Videos, or Current location",
                    initialvalue="Desktop"
                )
                
                root.destroy()
                
                if not location:
                    return "Folder deletion cancelled"
                
                location = location.lower()
                
                if location == "desktop":
                    path = os.path.join(os.path.expanduser("~"), "Desktop")
                elif location == "documents":
                    path = os.path.join(os.path.expanduser("~"), "Documents")
                elif location == "downloads":
                    path = os.path.join(os.path.expanduser("~"), "Downloads")
                elif location == "pictures":
                    path = os.path.join(os.path.expanduser("~"), "Pictures")
                elif location == "music":
                    path = os.path.join(os.path.expanduser("~"), "Music")
                elif location == "videos":
                    path = os.path.join(os.path.expanduser("~"), "Videos")
                elif location == "current location":
                    path = self.current_path
                else:
                    return f"Unknown location: {location}"
            
            folder_path = os.path.join(path, folder_name)
            
            if not os.path.exists(folder_path):
                return f"Folder '{folder_name}' not found in {os.path.basename(path)}"
            
            # Confirm deletion
            import tkinter as tk
            from tkinter import messagebox
            
            root = tk.Tk()
            root.withdraw()
            
            result = messagebox.askyesno(
                "Confirm Delete", 
                f"Are you sure you want to delete the folder '{folder_name}'?\n\nThis action cannot be undone!"
            )
            
            root.destroy()
            
            if result:
                import shutil
                shutil.rmtree(folder_path)
                self.open_explorer(path)
                return f"Deleted folder '{folder_name}' from {os.path.basename(path)}"
            else:
                return "Folder deletion cancelled"
                
        except Exception as e:
            return f"Error deleting folder: {str(e)}"
            
    def copy_folder(self, source_folder, destination=None):
        """Copy a folder to another location"""
        try:
            if destination is None:
                # Ask user for destination
                import tkinter as tk
                from tkinter import simpledialog
                
                root = tk.Tk()
                root.withdraw()
                
                destination = simpledialog.askstring(
                    "Copy Folder", 
                    f"Where would you like to copy '{source_folder}' to?\n\n"
                    f"Options: Desktop, Documents, Downloads, Pictures, Music, Videos, or Current location",
                    initialvalue="Desktop"
                )
                
                root.destroy()
                
                if not destination:
                    return "Folder copy cancelled"
                
                destination = destination.lower()
                
                if destination == "desktop":
                    dest_path = os.path.join(os.path.expanduser("~"), "Desktop")
                elif destination == "documents":
                    dest_path = os.path.join(os.path.expanduser("~"), "Documents")
                elif destination == "downloads":
                    dest_path = os.path.join(os.path.expanduser("~"), "Downloads")
                elif destination == "pictures":
                    dest_path = os.path.join(os.path.expanduser("~"), "Pictures")
                elif destination == "music":
                    dest_path = os.path.join(os.path.expanduser("~"), "Music")
                elif destination == "videos":
                    dest_path = os.path.join(os.path.expanduser("~"), "Videos")
                elif destination == "current location":
                    dest_path = self.current_path
                else:
                    return f"Unknown destination: {destination}"
            else:
                dest_path = destination
            
            # Find source folder in common locations
            source_path = None
            common_locations = [
                os.path.join(os.path.expanduser("~"), "Desktop"),
                os.path.join(os.path.expanduser("~"), "Documents"),
                os.path.join(os.path.expanduser("~"), "Downloads"),
                os.path.join(os.path.expanduser("~"), "Pictures"),
                os.path.join(os.path.expanduser("~"), "Music"),
                os.path.join(os.path.expanduser("~"), "Videos"),
                self.current_path
            ]
            
            for location in common_locations:
                potential_path = os.path.join(location, source_folder)
                if os.path.exists(potential_path):
                    source_path = potential_path
                    break
            
            if not source_path:
                return f"Folder '{source_folder}' not found in common locations"
            
            # Copy the folder
            import shutil
            dest_folder_path = os.path.join(dest_path, source_folder)
            
            # Add copy suffix if destination already exists
            if os.path.exists(dest_folder_path):
                dest_folder_path = dest_folder_path + "_copy"
            
            shutil.copytree(source_path, dest_folder_path)
            self.open_explorer(dest_path)
            return f"Copied folder '{source_folder}' to {os.path.basename(dest_path)}"
            
        except Exception as e:
            return f"Error copying folder: {str(e)}"
            
    def move_folder(self, source_folder, destination=None):
        """Move a folder to another location"""
        try:
            if destination is None:
                # Ask user for destination
                import tkinter as tk
                from tkinter import simpledialog
                
                root = tk.Tk()
                root.withdraw()
                
                destination = simpledialog.askstring(
                    "Move Folder", 
                    f"Where would you like to move '{source_folder}' to?\n\n"
                    f"Options: Desktop, Documents, Downloads, Pictures, Music, Videos, or Current location",
                    initialvalue="Desktop"
                )
                
                root.destroy()
                
                if not destination:
                    return "Folder move cancelled"
                
                destination = destination.lower()
                
                if destination == "desktop":
                    dest_path = os.path.join(os.path.expanduser("~"), "Desktop")
                elif destination == "documents":
                    dest_path = os.path.join(os.path.expanduser("~"), "Documents")
                elif destination == "downloads":
                    dest_path = os.path.join(os.path.expanduser("~"), "Downloads")
                elif destination == "pictures":
                    dest_path = os.path.join(os.path.expanduser("~"), "Pictures")
                elif destination == "music":
                    dest_path = os.path.join(os.path.expanduser("~"), "Music")
                elif destination == "videos":
                    dest_path = os.path.join(os.path.expanduser("~"), "Videos")
                elif destination == "current location":
                    dest_path = self.current_path
                else:
                    return f"Unknown destination: {destination}"
            else:
                dest_path = destination
            
            # Find source folder in common locations
            source_path = None
            common_locations = [
                os.path.join(os.path.expanduser("~"), "Desktop"),
                os.path.join(os.path.expanduser("~"), "Documents"),
                os.path.join(os.path.expanduser("~"), "Downloads"),
                os.path.join(os.path.expanduser("~"), "Pictures"),
                os.path.join(os.path.expanduser("~"), "Music"),
                os.path.join(os.path.expanduser("~"), "Videos"),
                self.current_path
            ]
            
            for location in common_locations:
                potential_path = os.path.join(location, source_folder)
                if os.path.exists(potential_path):
                    source_path = potential_path
                    break
            
            if not source_path:
                return f"Folder '{source_folder}' not found in common locations"
            
            # Move the folder
            import shutil
            dest_folder_path = os.path.join(dest_path, source_folder)
            
            # Add move suffix if destination already exists
            if os.path.exists(dest_folder_path):
                dest_folder_path = dest_folder_path + "_moved"
            
            shutil.move(source_path, dest_folder_path)
            self.open_explorer(dest_path)
            return f"Moved folder '{source_folder}' to {os.path.basename(dest_path)}"
            
        except Exception as e:
            return f"Error moving folder: {str(e)}"
            
    def rename_folder(self, old_name, new_name, location=None):
        """Rename a folder"""
        try:
            if location is None:
                # Ask user where the folder is located
                import tkinter as tk
                from tkinter import simpledialog
                
                root = tk.Tk()
                root.withdraw()
                
                location = simpledialog.askstring(
                    "Rename Folder", 
                    f"Where is the folder '{old_name}' located?\n\n"
                    f"Options: Desktop, Documents, Downloads, Pictures, Music, Videos, or Current location",
                    initialvalue="Desktop"
                )
                
                root.destroy()
                
                if not location:
                    return "Folder rename cancelled"
                
                location = location.lower()
                
                if location == "desktop":
                    path = os.path.join(os.path.expanduser("~"), "Desktop")
                elif location == "documents":
                    path = os.path.join(os.path.expanduser("~"), "Documents")
                elif location == "downloads":
                    path = os.path.join(os.path.expanduser("~"), "Downloads")
                elif location == "pictures":
                    path = os.path.join(os.path.expanduser("~"), "Pictures")
                elif location == "music":
                    path = os.path.join(os.path.expanduser("~"), "Music")
                elif location == "videos":
                    path = os.path.join(os.path.expanduser("~"), "Videos")
                elif location == "current location":
                    path = self.current_path
                else:
                    return f"Unknown location: {location}"
            else:
                path = location
            
            old_path = os.path.join(path, old_name)
            new_path = os.path.join(path, new_name)
            
            if not os.path.exists(old_path):
                return f"Folder '{old_name}' not found in {os.path.basename(path)}"
            
            if os.path.exists(new_path):
                return f"Folder '{new_name}' already exists in {os.path.basename(path)}"
            
            os.rename(old_path, new_path)
            self.open_explorer(path)
            return f"Renamed folder '{old_name}' to '{new_name}' in {os.path.basename(path)}"
            
        except Exception as e:
            return f"Error renaming folder: {str(e)}"
            
    def show_folder_properties(self, folder_name, location=None):
        """Show properties of a folder"""
        try:
            if location is None:
                # Ask user where the folder is located
                import tkinter as tk
                from tkinter import simpledialog
                
                root = tk.Tk()
                root.withdraw()
                
                location = simpledialog.askstring(
                    "Folder Properties", 
                    f"Where is the folder '{folder_name}' located?\n\n"
                    f"Options: Desktop, Documents, Downloads, Pictures, Music, Videos, or Current location",
                    initialvalue="Desktop"
                )
                
                root.destroy()
                
                if not location:
                    return "Properties cancelled"
                
                location = location.lower()
                
                if location == "desktop":
                    path = os.path.join(os.path.expanduser("~"), "Desktop")
                elif location == "documents":
                    path = os.path.join(os.path.expanduser("~"), "Documents")
                elif location == "downloads":
                    path = os.path.join(os.path.expanduser("~"), "Downloads")
                elif location == "pictures":
                    path = os.path.join(os.path.expanduser("~"), "Pictures")
                elif location == "music":
                    path = os.path.join(os.path.expanduser("~"), "Music")
                elif location == "videos":
                    path = os.path.join(os.path.expanduser("~"), "Videos")
                elif location == "current location":
                    path = self.current_path
                else:
                    return f"Unknown location: {location}"
            else:
                path = location
            
            folder_path = os.path.join(path, folder_name)
            
            if not os.path.exists(folder_path):
                return f"Folder '{folder_name}' not found in {os.path.basename(path)}"
            
            # Get folder properties
            import stat
            st = os.stat(folder_path)
            
            # Calculate folder size
            total_size = 0
            file_count = 0
            folder_count = 0
            
            for dirpath, dirnames, filenames in os.walk(folder_path):
                folder_count += len(dirnames)
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    try:
                        total_size += os.path.getsize(filepath)
                        file_count += 1
                    except:
                        pass
            
            # Format size
            def format_size(size_bytes):
                if size_bytes == 0:
                    return "0 B"
                size_names = ["B", "KB", "MB", "GB", "TB"]
                i = 0
                while size_bytes >= 1024 and i < len(size_names) - 1:
                    size_bytes /= 1024.0
                    i += 1
                return f"{size_bytes:.1f} {size_names[i]}"
            
            # Show properties dialog
            import tkinter as tk
            from tkinter import messagebox
            
            root = tk.Tk()
            root.withdraw()
            
            properties_text = f"""Folder Properties: {folder_name}

Location: {folder_path}
Size: {format_size(total_size)}
Files: {file_count}
Subfolders: {folder_count}
Created: {datetime.datetime.fromtimestamp(st.st_ctime).strftime('%Y-%m-%d %H:%M:%S')}
Modified: {datetime.datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M:%S')}
Accessed: {datetime.datetime.fromtimestamp(st.st_atime).strftime('%Y-%m-%d %H:%M:%S')}"""
            
            messagebox.showinfo("Folder Properties", properties_text)
            root.destroy()
            
            return f"Showed properties for folder '{folder_name}'"
            
        except Exception as e:
            return f"Error showing folder properties: {str(e)}"
            
    # ========== OS-BASED FOLDER OPERATIONS ==========
    
    def create_folder_os(self, folder_name, path=None):
        """Create folder using direct OS operations"""
        try:
            if path is None:
                path = os.path.expanduser("~")  # Default to user home directory
                
            folder_path = os.path.join(path, folder_name)
            
            # Check if folder already exists
            if os.path.exists(folder_path):
                return f"Folder '{folder_name}' already exists in {os.path.basename(path)}"
            
            # Create the folder using os.makedirs
            os.makedirs(folder_path, exist_ok=False)
            
            # Verify folder was created
            if os.path.exists(folder_path) and os.path.isdir(folder_path):
                return f"Successfully created folder '{folder_name}' in {os.path.basename(path)}"
            else:
                return f"Failed to create folder '{folder_name}'"
                
        except PermissionError:
            return f"Permission denied: Cannot create folder '{folder_name}' in {os.path.basename(path)}"
        except OSError as e:
            return f"OS Error creating folder '{folder_name}': {str(e)}"
        except Exception as e:
            return f"Error creating folder '{folder_name}': {str(e)}"
            
    def delete_folder_os(self, folder_name, path=None):
        """Delete folder using direct OS operations"""
        try:
            if path is None:
                path = os.path.expanduser("~")  # Default to user home directory
                
            folder_path = os.path.join(path, folder_name)
            
            # Check if folder exists
            if not os.path.exists(folder_path):
                return f"Folder '{folder_name}' not found in {os.path.basename(path)}"
            
            if not os.path.isdir(folder_path):
                return f"'{folder_name}' is not a folder in {os.path.basename(path)}"
            
            # Get folder size for confirmation
            total_size = 0
            file_count = 0
            
            try:
                for dirpath, dirnames, filenames in os.walk(folder_path):
                    for filename in filenames:
                        filepath = os.path.join(dirpath, filename)
                        try:
                            total_size += os.path.getsize(filepath)
                            file_count += 1
                        except:
                            pass
            except:
                pass
            
            # Format size for display
            def format_size(size_bytes):
                if size_bytes == 0:
                    return "0 B"
                size_names = ["B", "KB", "MB", "GB", "TB"]
                i = 0
                while size_bytes >= 1024 and i < len(size_names) - 1:
                    size_bytes /= 1024.0
                    i += 1
                return f"{size_bytes:.1f} {size_names[i]}"
            
            # Show confirmation dialog
            import tkinter as tk
            from tkinter import messagebox
            
            root = tk.Tk()
            root.withdraw()
            
            confirm_text = f"""Are you sure you want to delete the folder '{folder_name}'?

Location: {folder_path}
Size: {format_size(total_size)}
Files: {file_count}

This action cannot be undone!"""
            
            result = messagebox.askyesno("Confirm Delete", confirm_text)
            root.destroy()
            
            if result:
                # Use shutil.rmtree for recursive deletion
                import shutil
                shutil.rmtree(folder_path)
                
                # Verify folder was deleted
                if not os.path.exists(folder_path):
                    return f"Successfully deleted folder '{folder_name}' from {os.path.basename(path)}"
                else:
                    return f"Failed to delete folder '{folder_name}'"
            else:
                return "Folder deletion cancelled"
                
        except PermissionError:
            return f"Permission denied: Cannot delete folder '{folder_name}' from {os.path.basename(path)}"
        except OSError as e:
            return f"OS Error deleting folder '{folder_name}': {str(e)}"
        except Exception as e:
            return f"Error deleting folder '{folder_name}': {str(e)}"
            
    def list_folders_os(self, path=None):
        """List all folders in a directory using OS operations"""
        try:
            if path is None:
                path = os.path.expanduser("~")  # Default to user home directory
                
            if not os.path.exists(path):
                return f"Path '{path}' does not exist"
            
            if not os.path.isdir(path):
                return f"'{path}' is not a directory"
            
            # Get all items in the directory
            try:
                items = os.listdir(path)
            except PermissionError:
                return f"Permission denied: Cannot access '{path}'"
            
            # Filter only folders
            folders = []
            for item in items:
                item_path = os.path.join(path, item)
                if os.path.isdir(item_path):
                    folders.append(item)
            
            # Sort folders alphabetically
            folders.sort()
            
            if folders:
                folder_list = "\n".join([f"• {folder}" for folder in folders])
                return f"Folders in {os.path.basename(path)}:\n{folder_list}"
            else:
                return f"No folders found in {os.path.basename(path)}"
                
        except Exception as e:
            return f"Error listing folders: {str(e)}"
            
    def check_folder_exists_os(self, folder_name, path=None):
        """Check if a folder exists using OS operations"""
        try:
            if path is None:
                path = os.path.expanduser("~")  # Default to user home directory
                
            folder_path = os.path.join(path, folder_name)
            
            if os.path.exists(folder_path):
                if os.path.isdir(folder_path):
                    # Get folder info
                    st = os.stat(folder_path)
                    created = datetime.datetime.fromtimestamp(st.st_ctime).strftime('%Y-%m-%d %H:%M:%S')
                    modified = datetime.datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                    
                    return f"Folder '{folder_name}' exists in {os.path.basename(path)}\nCreated: {created}\nModified: {modified}"
                else:
                    return f"'{folder_name}' exists in {os.path.basename(path)} but it's not a folder"
            else:
                return f"Folder '{folder_name}' does not exist in {os.path.basename(path)}"
                
        except Exception as e:
            return f"Error checking folder: {str(e)}"
            
    def get_folder_size_os(self, folder_name, path=None):
        """Get folder size using OS operations"""
        try:
            if path is None:
                path = os.path.expanduser("~")  # Default to user home directory
                
            folder_path = os.path.join(path, folder_name)
            
            if not os.path.exists(folder_path):
                return f"Folder '{folder_name}' not found in {os.path.basename(path)}"
            
            if not os.path.isdir(folder_path):
                return f"'{folder_name}' is not a folder in {os.path.basename(path)}"
            
            # Calculate folder size
            total_size = 0
            file_count = 0
            folder_count = 0
            
            for dirpath, dirnames, filenames in os.walk(folder_path):
                folder_count += len(dirnames)
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    try:
                        total_size += os.path.getsize(filepath)
                        file_count += 1
                    except:
                        pass
            
            # Format size
            def format_size(size_bytes):
                if size_bytes == 0:
                    return "0 B"
                size_names = ["B", "KB", "MB", "GB", "TB"]
                i = 0
                while size_bytes >= 1024 and i < len(size_names) - 1:
                    size_bytes /= 1024.0
                    i += 1
                return f"{size_bytes:.1f} {size_names[i]}"
            
            return f"Folder '{folder_name}' size: {format_size(total_size)}\nFiles: {file_count}\nSubfolders: {folder_count}"
            
        except Exception as e:
            return f"Error getting folder size: {str(e)}"
            
    def test_file_operations(self):
        """Test function to debug file operations"""
        try:
            print("=== Testing File Operations ===")
            
            # Test 1: Check if we can access user home directory
            home_dir = os.path.expanduser("~")
            print(f"Home directory: {home_dir}")
            print(f"Home directory exists: {os.path.exists(home_dir)}")
            
            # Test 2: Check if we can access common folders
            common_folders = ["Desktop", "Documents", "Downloads", "Pictures", "Music", "Videos"]
            for folder in common_folders:
                folder_path = os.path.join(home_dir, folder)
                exists = os.path.exists(folder_path)
                print(f"{folder}: {folder_path} - Exists: {exists}")
            
            # Test 3: Test creating a simple folder
            test_folder = "test_folder_123"
            test_path = os.path.join(home_dir, test_folder)
            
            if os.path.exists(test_path):
                print(f"Test folder already exists: {test_path}")
            else:
                try:
                    os.makedirs(test_path, exist_ok=True)
                    print(f"Successfully created test folder: {test_path}")
                    
                    # Clean up
                    import shutil
                    shutil.rmtree(test_path)
                    print(f"Successfully cleaned up test folder: {test_path}")
                except Exception as e:
                    print(f"Error creating test folder: {e}")
            
            # Test 4: Test subprocess for explorer
            try:
                import subprocess
                result = subprocess.run(['explorer', home_dir], capture_output=True, text=True, timeout=5)
                print(f"Explorer command test: {result.returncode}")
            except Exception as e:
                print(f"Explorer command error: {e}")
            
            print("=== File Operations Test Complete ===")
            return "File operations test completed. Check console for details."
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Test function error: {error_details}")
            return f"Test function error: {str(e)}"
            
    def handle_voice_command(self, command):
        """Handle voice commands for file operations"""
        command_lower = command.lower()
        print(f"Handling voice command: {command}")  # Debug log
        
        # Basic file explorer commands
        if 'open file explorer' in command_lower or 'file explorer' in command_lower:
            return self.open_explorer()
        
        elif 'open desktop' in command_lower:
            return self.open_desktop()
            
        elif 'open documents' in command_lower:
            return self.open_documents()
            
        elif 'open downloads' in command_lower:
            return self.open_downloads()
            
        elif 'open pictures' in command_lower:
            return self.open_pictures()
            
        elif 'open music' in command_lower:
            return self.open_music()
            
        elif 'open videos' in command_lower:
            return self.open_videos()
            
        # Drive commands
        elif 'open drive' in command_lower:
            # Extract drive letter
            for letter in 'abcdefghijklmnopqrstuvwxyz':
                if f'open drive {letter}' in command_lower:
                    return self.open_drive(letter)
            return "Please specify a drive letter (A-Z)"
            
        # Create folder command
        elif 'create folder' in command_lower or 'new folder' in command_lower:
            # Extract folder name and location from command
            command_parts = command_lower.replace('create folder', '').replace('new folder', '').strip()
            
            # Check for location keywords
            locations = {
                'desktop': os.path.join(os.path.expanduser("~"), "Desktop"),
                'documents': os.path.join(os.path.expanduser("~"), "Documents"),
                'downloads': os.path.join(os.path.expanduser("~"), "Downloads"),
                'pictures': os.path.join(os.path.expanduser("~"), "Pictures"),
                'music': os.path.join(os.path.expanduser("~"), "Music"),
                'videos': os.path.join(os.path.expanduser("~"), "Videos")
            }
            
            folder_name = command_parts
            target_path = None
            
            # Check if location is specified in the command
            for location_name, location_path in locations.items():
                if location_name in command_parts:
                    folder_name = command_parts.replace(location_name, '').strip()
                    target_path = location_path
                    break
            
            if folder_name:
                return self.create_folder(folder_name, target_path)
            else:
                return "Please specify a folder name"
                
        # Delete folder command
        elif 'delete folder' in command_lower or 'remove folder' in command_lower:
            # Extract folder name and location from command
            command_parts = command_lower.replace('delete folder', '').replace('remove folder', '').strip()
            
            # Check for location keywords
            locations = {
                'desktop': os.path.join(os.path.expanduser("~"), "Desktop"),
                'documents': os.path.join(os.path.expanduser("~"), "Documents"),
                'downloads': os.path.join(os.path.expanduser("~"), "Downloads"),
                'pictures': os.path.join(os.path.expanduser("~"), "Pictures"),
                'music': os.path.join(os.path.expanduser("~"), "Music"),
                'videos': os.path.join(os.path.expanduser("~"), "Videos")
            }
            
            folder_name = command_parts
            target_path = None
            
            # Check if location is specified in the command
            for location_name, location_path in locations.items():
                if location_name in command_parts:
                    folder_name = command_parts.replace(location_name, '').strip()
                    target_path = location_path
                    break
            
            if folder_name:
                return self.delete_folder(folder_name, target_path)
            else:
                return "Please specify a folder name"
                
        # Copy folder command
        elif 'copy folder' in command_lower:
            # Extract folder name and destination from command
            command_parts = command_lower.replace('copy folder', '').strip()
            
            # Check for destination keywords
            destinations = {
                'desktop': os.path.join(os.path.expanduser("~"), "Desktop"),
                'documents': os.path.join(os.path.expanduser("~"), "Documents"),
                'downloads': os.path.join(os.path.expanduser("~"), "Downloads"),
                'pictures': os.path.join(os.path.expanduser("~"), "Pictures"),
                'music': os.path.join(os.path.expanduser("~"), "Music"),
                'videos': os.path.join(os.path.expanduser("~"), "Videos")
            }
            
            folder_name = command_parts
            target_destination = None
            
            # Check if destination is specified in the command
            for dest_name, dest_path in destinations.items():
                if dest_name in command_parts:
                    folder_name = command_parts.replace(dest_name, '').strip()
                    target_destination = dest_path
                    break
            
            if folder_name:
                return self.copy_folder(folder_name, target_destination)
            else:
                return "Please specify a folder name"
                
        # Move folder command
        elif 'move folder' in command_lower:
            # Extract folder name and destination from command
            command_parts = command_lower.replace('move folder', '').strip()
            
            # Check for destination keywords
            destinations = {
                'desktop': os.path.join(os.path.expanduser("~"), "Desktop"),
                'documents': os.path.join(os.path.expanduser("~"), "Documents"),
                'downloads': os.path.join(os.path.expanduser("~"), "Downloads"),
                'pictures': os.path.join(os.path.expanduser("~"), "Pictures"),
                'music': os.path.join(os.path.expanduser("~"), "Music"),
                'videos': os.path.join(os.path.expanduser("~"), "Videos")
            }
            
            folder_name = command_parts
            target_destination = None
            
            # Check if destination is specified in the command
            for dest_name, dest_path in destinations.items():
                if dest_name in command_parts:
                    folder_name = command_parts.replace(dest_name, '').strip()
                    target_destination = dest_path
                    break
            
            if folder_name:
                return self.move_folder(folder_name, target_destination)
            else:
                return "Please specify a folder name"
                
        # Rename folder command
        elif 'rename folder' in command_lower:
            # Extract old name, new name, and location from command
            command_parts = command_lower.replace('rename folder', '').strip()
            
            # Look for "to" keyword to separate old and new names
            if ' to ' in command_parts:
                parts = command_parts.split(' to ')
                if len(parts) == 2:
                    old_name = parts[0].strip()
                    new_name = parts[1].strip()
                    
                    # Check for location keywords
                    locations = {
                        'desktop': os.path.join(os.path.expanduser("~"), "Desktop"),
                        'documents': os.path.join(os.path.expanduser("~"), "Documents"),
                        'downloads': os.path.join(os.path.expanduser("~"), "Downloads"),
                        'pictures': os.path.join(os.path.expanduser("~"), "Pictures"),
                        'music': os.path.join(os.path.expanduser("~"), "Music"),
                        'videos': os.path.join(os.path.expanduser("~"), "Videos")
                    }
                    
                    target_location = None
                    
                    # Check if location is specified in the command
                    for location_name, location_path in locations.items():
                        if location_name in old_name or location_name in new_name:
                            old_name = old_name.replace(location_name, '').strip()
                            new_name = new_name.replace(location_name, '').strip()
                            target_location = location_path
                            break
                    
                    if old_name and new_name:
                        return self.rename_folder(old_name, new_name, target_location)
                    else:
                        return "Please specify both old and new folder names"
                else:
                    return "Please use format: rename folder [old name] to [new name]"
            else:
                return "Please use format: rename folder [old name] to [new name]"
                
        # Show folder properties command
        elif 'folder properties' in command_lower or 'show properties' in command_lower:
            # Extract folder name and location from command
            command_parts = command_lower.replace('folder properties', '').replace('show properties', '').strip()
            
            # Check for location keywords
            locations = {
                'desktop': os.path.join(os.path.expanduser("~"), "Desktop"),
                'documents': os.path.join(os.path.expanduser("~"), "Documents"),
                'downloads': os.path.join(os.path.expanduser("~"), "Downloads"),
                'pictures': os.path.join(os.path.expanduser("~"), "Pictures"),
                'music': os.path.join(os.path.expanduser("~"), "Music"),
                'videos': os.path.join(os.path.expanduser("~"), "Videos")
            }
            
            folder_name = command_parts
            target_location = None
            
            # Check if location is specified in the command
            for location_name, location_path in locations.items():
                if location_name in command_parts:
                    folder_name = command_parts.replace(location_name, '').strip()
                    target_location = location_path
                    break
            
            if folder_name:
                return self.show_folder_properties(folder_name, target_location)
            else:
                return "Please specify a folder name"
                
        # ========== OS-BASED VOICE COMMANDS ==========
        
        # Create folder using OS operations
        elif 'create folder os' in command_lower or 'make folder os' in command_lower:
            # Extract folder name and location from command
            command_parts = command_lower.replace('create folder os', '').replace('make folder os', '').strip()
            
            # Check for location keywords
            locations = {
                'desktop': os.path.join(os.path.expanduser("~"), "Desktop"),
                'documents': os.path.join(os.path.expanduser("~"), "Documents"),
                'downloads': os.path.join(os.path.expanduser("~"), "Downloads"),
                'pictures': os.path.join(os.path.expanduser("~"), "Pictures"),
                'music': os.path.join(os.path.expanduser("~"), "Music"),
                'videos': os.path.join(os.path.expanduser("~"), "Videos")
            }
            
            folder_name = command_parts
            target_path = None
            
            # Check if location is specified in the command
            for location_name, location_path in locations.items():
                if location_name in command_parts:
                    folder_name = command_parts.replace(location_name, '').strip()
                    target_path = location_path
                    break
            
            if folder_name:
                return self.create_folder_os(folder_name, target_path)
            else:
                return "Please specify a folder name"
                
        # Delete folder using OS operations
        elif 'delete folder os' in command_lower or 'remove folder os' in command_lower:
            # Extract folder name and location from command
            command_parts = command_lower.replace('delete folder os', '').replace('remove folder os', '').strip()
            
            # Check for location keywords
            locations = {
                'desktop': os.path.join(os.path.expanduser("~"), "Desktop"),
                'documents': os.path.join(os.path.expanduser("~"), "Documents"),
                'downloads': os.path.join(os.path.expanduser("~"), "Downloads"),
                'pictures': os.path.join(os.path.expanduser("~"), "Pictures"),
                'music': os.path.join(os.path.expanduser("~"), "Music"),
                'videos': os.path.join(os.path.expanduser("~"), "Videos")
            }
            
            folder_name = command_parts
            target_path = None
            
            # Check if location is specified in the command
            for location_name, location_path in locations.items():
                if location_name in command_parts:
                    folder_name = command_parts.replace(location_name, '').strip()
                    target_path = location_path
                    break
            
            if folder_name:
                return self.delete_folder_os(folder_name, target_path)
            else:
                return "Please specify a folder name"
                
        # List folders using OS operations
        elif 'list folders' in command_lower or 'show folders' in command_lower:
            # Extract location from command
            command_parts = command_lower.replace('list folders', '').replace('show folders', '').strip()
            
            # Check for location keywords
            locations = {
                'desktop': os.path.join(os.path.expanduser("~"), "Desktop"),
                'documents': os.path.join(os.path.expanduser("~"), "Documents"),
                'downloads': os.path.join(os.path.expanduser("~"), "Downloads"),
                'pictures': os.path.join(os.path.expanduser("~"), "Pictures"),
                'music': os.path.join(os.path.expanduser("~"), "Music"),
                'videos': os.path.join(os.path.expanduser("~"), "Videos")
            }
            
            target_path = None
            
            # Check if location is specified in the command
            for location_name, location_path in locations.items():
                if location_name in command_parts:
                    target_path = location_path
                    break
            
            return self.list_folders_os(target_path)
            
        # Check if folder exists using OS operations
        elif 'check folder' in command_lower or 'folder exists' in command_lower:
            # Extract folder name and location from command
            command_parts = command_lower.replace('check folder', '').replace('folder exists', '').strip()
            
            # Check for location keywords
            locations = {
                'desktop': os.path.join(os.path.expanduser("~"), "Desktop"),
                'documents': os.path.join(os.path.expanduser("~"), "Documents"),
                'downloads': os.path.join(os.path.expanduser("~"), "Downloads"),
                'pictures': os.path.join(os.path.expanduser("~"), "Pictures"),
                'music': os.path.join(os.path.expanduser("~"), "Music"),
                'videos': os.path.join(os.path.expanduser("~"), "Videos")
            }
            
            folder_name = command_parts
            target_path = None
            
            # Check if location is specified in the command
            for location_name, location_path in locations.items():
                if location_name in command_parts:
                    folder_name = command_parts.replace(location_name, '').strip()
                    target_path = location_path
                    break
            
            if folder_name:
                return self.check_folder_exists_os(folder_name, target_path)
            else:
                return "Please specify a folder name"
                
        # Get folder size using OS operations
        elif 'folder size' in command_lower or 'get size' in command_lower:
            # Extract folder name and location from command
            command_parts = command_lower.replace('folder size', '').replace('get size', '').strip()
            
            # Check for location keywords
            locations = {
                'desktop': os.path.join(os.path.expanduser("~"), "Desktop"),
                'documents': os.path.join(os.path.expanduser("~"), "Documents"),
                'downloads': os.path.join(os.path.expanduser("~"), "Downloads"),
                'pictures': os.path.join(os.path.expanduser("~"), "Pictures"),
                'music': os.path.join(os.path.expanduser("~"), "Music"),
                'videos': os.path.join(os.path.expanduser("~"), "Videos")
            }
            
            folder_name = command_parts
            target_path = None
            
            # Check if location is specified in the command
            for location_name, location_path in locations.items():
                if location_name in command_parts:
                    folder_name = command_parts.replace(location_name, '').strip()
                    target_path = location_path
                    break
            
            if folder_name:
                return self.get_folder_size_os(folder_name, target_path)
            else:
                return "Please specify a folder name"
                
        # Test file operations
        elif 'test file operations' in command_lower or 'debug file operations' in command_lower:
            return self.test_file_operations()
            
        # 3D Circle controls
        elif 'show 3d circle' in command_lower or 'display 3d circle' in command_lower:
            if hasattr(self.parent, 'circle_3d'):
                self.parent.circle_3d.animation_speed = 0.02
                return "3D circle displayed"
            else:
                return "3D circle not available"
                
        elif 'hide 3d circle' in command_lower or 'remove 3d circle' in command_lower:
            if hasattr(self.parent, 'circle_3d'):
                self.parent.circle_3d.animation_speed = 0
                return "3D circle hidden"
            else:
                return "3D circle not available"
                
        elif 'faster circle' in command_lower or 'speed up circle' in command_lower:
            if hasattr(self.parent, 'circle_3d'):
                self.parent.circle_3d.animation_speed += 0.01
                return f"3D circle speed increased to {self.parent.circle_3d.animation_speed:.2f}"
            else:
                return "3D circle not available"
                
        elif 'slower circle' in command_lower or 'slow down circle' in command_lower:
            if hasattr(self.parent, 'circle_3d'):
                self.parent.circle_3d.animation_speed = max(0, self.parent.circle_3d.animation_speed - 0.01)
                return f"3D circle speed decreased to {self.parent.circle_3d.animation_speed:.2f}"
            else:
                return "3D circle not available"
                
        # Open specific path
        elif 'open path' in command_lower:
            path = command_lower.replace('open path', '').strip()
            if path:
                return self.open_path(path)
            else:
                return "Please specify a path"
                
        # If no command matched, return a helpful message
        print(f"No matching command found for: {command}")  # Debug log
        return f"Unrecognized file command: {command}. Try 'open file explorer', 'create folder', etc."
        
    def open_file_explorer(self):
        """Open Windows File Explorer"""
        return self.open_explorer()
        
    def load_drives(self):
        """Load available drives in tree view"""
        self.tree.delete(*self.tree.get_children())
        
        # Get drives (Windows)
        import string
        drives = []
        for letter in string.ascii_uppercase:
            drive = f"{letter}:\\"
            if os.path.exists(drive):
                drives.append(drive)
        
        for drive in drives:
            try:
                drive_name = f"{drive} ({os.path.basename(drive)})"
                self.tree.insert('', 'end', drive, text=drive_name, values=(drive,), tags=('drive',))
                # Add a placeholder for subdirectories
                self.tree.insert(drive, 'end', f"{drive}_placeholder", text="Loading...", tags=('placeholder',))
            except:
                continue
                
    def load_subdirectories(self, parent_path):
        """Load subdirectories for a given path"""
        try:
            # Remove placeholder if it exists
            placeholder_id = f"{parent_path}_placeholder"
            if self.tree.exists(placeholder_id):
                self.tree.delete(placeholder_id)
            
            # Get subdirectories
            subdirs = []
            try:
                for item in os.listdir(parent_path):
                    item_path = os.path.join(parent_path, item)
                    if os.path.isdir(item_path):
                        subdirs.append((item, item_path))
            except PermissionError:
                return
                
            # Sort alphabetically
            subdirs.sort(key=lambda x: x[0].lower())
            
            # Add subdirectories to tree
            for name, path in subdirs:
                try:
                    item_id = f"{parent_path}_{name}"
                    self.tree.insert(parent_path, 'end', item_id, text=name, values=(path,), tags=('folder',))
                    # Add placeholder for subdirectories
                    self.tree.insert(item_id, 'end', f"{item_id}_placeholder", text="Loading...", tags=('placeholder',))
                except:
                    continue
                    
        except Exception as e:
            print(f"Error loading subdirectories for {parent_path}: {e}")
            
    def on_tree_select(self, event):
        """Handle tree selection"""
        selection = self.tree.selection()
        if selection:
            item = selection[0]
            path = self.tree.item(item, 'values')[0]
            
            # Load subdirectories if this is a drive or folder
            if self.tree.item(item, 'tags') and 'placeholder' in self.tree.item(item, 'tags'):
                parent_item = self.tree.parent(item)
                if parent_item:
                    parent_path = self.tree.item(parent_item, 'values')[0]
                    self.load_subdirectories(parent_path)
            
            self.load_directory(path)
            
    def on_tree_expand(self, event):
        """Handle tree expansion"""
        item = self.tree.focus()
        if item:
            path = self.tree.item(item, 'values')[0]
            # Load subdirectories when expanding
            self.load_subdirectories(path)
                
    def load_directory(self, path):
        """Load directory contents"""
        try:
            self.current_path = path
            self.path_var.set(path)
            
            # Clear file list
            self.file_listbox.delete(0, tk.END)
            
            # Get directory contents
            try:
                items = os.listdir(path)
            except PermissionError:
                self.status_var.set("Access denied")
                return
                
            # Separate folders and files
            folders = []
            files = []
            
            for item in items:
                item_path = os.path.join(path, item)
                try:
                    if os.path.isdir(item_path):
                        folders.append(item)
                    else:
                        files.append(item)
                except:
                    continue
            
            # Sort alphabetically
            folders.sort()
            files.sort()
            
            # Add folders first
            for folder in folders:
                self.file_listbox.insert(tk.END, f"📁 {folder}")
                
            # Add files
            for file in files:
                # Get file extension for icon
                ext = os.path.splitext(file)[1].lower()
                icon = self.get_file_icon(ext)
                self.file_listbox.insert(tk.END, f"{icon} {file}")
                
            self.status_var.set(f"{len(folders)} folders, {len(files)} files")
            
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
            
    def get_file_icon(self, ext):
        """Get appropriate icon for file type"""
        icons = {
            '.txt': '📄', '.doc': '📄', '.docx': '📄', '.pdf': '📄',
            '.jpg': '🖼️', '.jpeg': '🖼️', '.png': '🖼️', '.gif': '🖼️', '.bmp': '🖼️',
            '.mp3': '🎵', '.wav': '🎵', '.flac': '🎵', '.aac': '🎵',
            '.mp4': '🎬', '.avi': '🎬', '.mkv': '🎬', '.mov': '🎬',
            '.py': '🐍', '.js': '📜', '.html': '🌐', '.css': '🎨',
            '.zip': '📦', '.rar': '📦', '.7z': '📦', '.tar': '📦',
            '.exe': '⚙️', '.msi': '⚙️', '.bat': '⚙️', '.cmd': '⚙️'
        }
        return icons.get(ext, '📄')
        
    def on_tree_select(self, event):
        """Handle tree selection"""
        selection = self.tree.selection()
        if selection:
            item = selection[0]
            path = self.tree.item(item, 'values')[0]
            self.load_directory(path)
            
    def on_file_double_click(self, event):
        """Handle file double click"""
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            item = self.file_listbox.get(index)
            
            # Remove icon and get filename
            filename = item.split(' ', 1)[1] if ' ' in item else item
            file_path = os.path.join(self.current_path, filename)
            
            if os.path.isdir(file_path):
                self.load_directory(file_path)
            else:
                self.open_file(file_path)
                
    def on_file_select(self, event):
        """Handle file selection"""
        selection = self.file_listbox.curselection()
        if selection:
            self.selected_items = []
            for index in selection:
                item = self.file_listbox.get(index)
                filename = item.split(' ', 1)[1] if ' ' in item else item
                file_path = os.path.join(self.current_path, filename)
                self.selected_items.append(file_path)
            
            # Show preview for single selection
            if len(self.selected_items) == 1:
                self.show_file_preview(self.selected_items[0])
            else:
                self.clear_preview()
                
    def show_file_preview(self, file_path):
        """Show preview of selected file"""
        try:
            if not os.path.exists(file_path):
                self.clear_preview()
                return
                
            # Clear previous preview
            self.clear_preview()
            
            # Get file extension
            ext = os.path.splitext(file_path)[1].lower()
            
            # Handle different file types
            if ext in ['.txt', '.py', '.js', '.html', '.css', '.json', '.xml', '.md', '.log']:
                self.show_text_preview(file_path)
            elif ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp']:
                self.show_image_preview(file_path)
            else:
                self.show_file_info_preview(file_path)
                
        except Exception as e:
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, f"Error loading preview: {str(e)}")
            self.preview_text.config(state=tk.DISABLED)
            
    def show_text_preview(self, file_path):
        """Show text file preview"""
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read(2000)  # Read first 2000 characters
                
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, content)
            if len(content) >= 2000:
                self.preview_text.insert(tk.END, "\n\n... (truncated)")
            self.preview_text.config(state=tk.DISABLED)
            
        except Exception as e:
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, f"Error reading file: {str(e)}")
            self.preview_text.config(state=tk.DISABLED)
            
    def show_image_preview(self, file_path):
        """Show image preview"""
        try:
            # Load and resize image
            image = Image.open(file_path)
            # Resize to fit preview area
            image.thumbnail((200, 150), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(image)
            
            self.preview_image_label.config(image=photo, text="")
            self.preview_image_label.image = photo  # Keep a reference
            
        except Exception as e:
            self.preview_image_label.config(image="", text=f"Error loading image: {str(e)}")
            
    def show_file_info_preview(self, file_path):
        """Show file information preview"""
        try:
            filename = os.path.basename(file_path)
            size = os.path.getsize(file_path)
            size_str = self.format_size(size)
            modified = os.path.getmtime(file_path)
            modified_str = datetime.datetime.fromtimestamp(modified).strftime('%Y-%m-%d %H:%M:%S')
            
            info = f"File: {filename}\n"
            info += f"Size: {size_str}\n"
            info += f"Modified: {modified_str}\n"
            info += f"Type: {os.path.splitext(filename)[1].upper()} file"
            
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, info)
            self.preview_text.config(state=tk.DISABLED)
            
        except Exception as e:
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, f"Error getting file info: {str(e)}")
            self.preview_text.config(state=tk.DISABLED)
            
    def clear_preview(self):
        """Clear the preview panel"""
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete(1.0, tk.END)
        self.preview_text.config(state=tk.DISABLED)
        self.preview_image_label.config(image="", text="No preview available")
                
    def open_file(self, file_path):
        """Open file with default application"""
        try:
            os.startfile(file_path)
            self.status_var.set(f"Opened: {os.path.basename(file_path)}")
        except Exception as e:
            self.status_var.set(f"Error opening file: {str(e)}")
            
    def go_back(self):
        """Go back in history"""
        # Implementation for back functionality
        pass
        
    def go_forward(self):
        """Go forward in history"""
        # Implementation for forward functionality
        pass
        
    def go_up(self):
        """Go to parent directory"""
        parent = os.path.dirname(self.current_path)
        if parent != self.current_path:
            self.load_directory(parent)
            
    def navigate_to_path(self, event):
        """Navigate to entered path"""
        path = self.path_var.get()
        if os.path.exists(path):
            self.load_directory(path)
        else:
            self.status_var.set("Path does not exist")
            
    def create_new_folder(self):
        """Create new folder"""
        folder_name = simpledialog.askstring("New Folder", "Enter folder name:")
        if folder_name:
            try:
                new_path = os.path.join(self.current_path, folder_name)
                os.makedirs(new_path, exist_ok=True)
                self.load_directory(self.current_path)
                self.status_var.set(f"Created folder: {folder_name}")
            except Exception as e:
                self.status_var.set(f"Error creating folder: {str(e)}")
                
    def copy_selected(self):
        """Copy selected items to clipboard"""
        if self.selected_items:
            self.clipboard = self.selected_items.copy()
            self.clipboard_mode = 'copy'
            self.status_var.set(f"Copied {len(self.selected_items)} items")
            
    def cut_selected(self):
        """Cut selected items to clipboard"""
        if self.selected_items:
            self.clipboard = self.selected_items.copy()
            self.clipboard_mode = 'cut'
            self.status_var.set(f"Cut {len(self.selected_items)} items")
            
    def paste_items(self):
        """Paste items from clipboard"""
        if not self.clipboard:
            self.status_var.set("Clipboard is empty")
            return
            
        try:
            for item in self.clipboard:
                if os.path.exists(item):
                    filename = os.path.basename(item)
                    dest_path = os.path.join(self.current_path, filename)
                    
                    if self.clipboard_mode == 'copy':
                        if os.path.isdir(item):
                            self.copy_directory(item, dest_path)
                        else:
                            import shutil
                            shutil.copy2(item, dest_path)
                    elif self.clipboard_mode == 'cut':
                        import shutil
                        shutil.move(item, dest_path)
                        
            self.load_directory(self.current_path)
            self.status_var.set(f"Pasted {len(self.clipboard)} items")
            
            if self.clipboard_mode == 'cut':
                self.clipboard = []
                self.clipboard_mode = None
                
        except Exception as e:
            self.status_var.set(f"Error pasting items: {str(e)}")
            
    def copy_directory(self, src, dst):
        """Copy directory recursively"""
        import shutil
        if os.path.exists(dst):
            dst = dst + "_copy"
        shutil.copytree(src, dst)
        
    def delete_selected(self):
        """Delete selected items"""
        if not self.selected_items:
            self.status_var.set("No items selected")
            return
            
        result = messagebox.askyesno("Confirm Delete", 
                                   f"Are you sure you want to delete {len(self.selected_items)} items?")
        if result:
            try:
                for item in self.selected_items:
                    if os.path.isdir(item):
                        import shutil
                        shutil.rmtree(item)
                    else:
                        os.remove(item)
                        
                self.load_directory(self.current_path)
                self.selected_items = []
                self.status_var.set("Items deleted successfully")
                
            except Exception as e:
                self.status_var.set(f"Error deleting items: {str(e)}")
                
    def refresh_view(self):
        """Refresh the current view"""
        self.load_directory(self.current_path)
        
    def show_file_manager(self):
        """Show the file manager window"""
        self.create_file_manager_window()
        
    def handle_voice_command(self, command):
        """Handle voice commands for file operations"""
        command_lower = command.lower()
        
        # File operations
        if 'create folder' in command_lower or 'new folder' in command_lower:
            folder_name = command_lower.replace('create folder', '').replace('new folder', '').strip()
            if folder_name:
                self.create_folder_by_voice(folder_name)
            else:
                self.create_new_folder()
                
        elif 'delete file' in command_lower or 'delete folder' in command_lower:
            item_name = command_lower.replace('delete file', '').replace('delete folder', '').strip()
            if item_name:
                self.delete_item_by_name(item_name)
                
        elif 'copy file' in command_lower or 'copy folder' in command_lower:
            item_name = command_lower.replace('copy file', '').replace('copy folder', '').strip()
            if item_name:
                self.copy_item_by_name(item_name)
                
        elif 'move file' in command_lower or 'move folder' in command_lower:
            item_name = command_lower.replace('move file', '').replace('move folder', '').strip()
            if item_name:
                self.move_item_by_name(item_name)
                
        elif 'rename file' in command_lower or 'rename folder' in command_lower:
            # Extract old and new names
            parts = command_lower.split('to')
            if len(parts) == 2:
                old_name = parts[0].replace('rename file', '').replace('rename folder', '').strip()
                new_name = parts[1].strip()
                if old_name and new_name:
                    self.rename_item_by_voice(old_name, new_name)
                    
        elif 'search for' in command_lower:
            search_term = command_lower.replace('search for', '').strip()
            if search_term:
                self.search_by_voice(search_term)
                
    def create_folder_by_voice(self, folder_name):
        """Create folder using voice command"""
        try:
            new_path = os.path.join(self.current_path, folder_name)
            os.makedirs(new_path, exist_ok=True)
            self.load_directory(self.current_path)
            return f"Created folder: {folder_name}"
        except Exception as e:
            return f"Error creating folder: {str(e)}"
            
    def delete_item_by_name(self, item_name):
        """Delete item by name using voice command"""
        try:
            item_path = os.path.join(self.current_path, item_name)
            if os.path.exists(item_path):
                if os.path.isdir(item_path):
                    import shutil
                    shutil.rmtree(item_path)
                else:
                    os.remove(item_path)
                self.load_directory(self.current_path)
                return f"Deleted: {item_name}"
            else:
                return f"Item not found: {item_name}"
        except Exception as e:
            return f"Error deleting item: {str(e)}"
            
    def copy_item_by_name(self, item_name):
        """Copy item by name using voice command"""
        try:
            item_path = os.path.join(self.current_path, item_name)
            if os.path.exists(item_path):
                self.selected_items = [item_path]
                self.copy_selected()
                return f"Copied: {item_name}"
            else:
                return f"Item not found: {item_name}"
        except Exception as e:
            return f"Error copying item: {str(e)}"
            
    def move_item_by_name(self, item_name):
        """Move item by name using voice command"""
        try:
            item_path = os.path.join(self.current_path, item_name)
            if os.path.exists(item_path):
                self.selected_items = [item_path]
                self.cut_selected()
                return f"Cut: {item_name} (ready to paste)"
            else:
                return f"Item not found: {item_name}"
        except Exception as e:
            return f"Error moving item: {str(e)}"
            
    def rename_item_by_voice(self, old_name, new_name):
        """Rename item using voice command"""
        try:
            old_path = os.path.join(self.current_path, old_name)
            new_path = os.path.join(self.current_path, new_name)
            if os.path.exists(old_path):
                os.rename(old_path, new_path)
                self.load_directory(self.current_path)
                return f"Renamed '{old_name}' to '{new_name}'"
            else:
                return f"Item not found: {old_name}"
        except Exception as e:
            return f"Error renaming item: {str(e)}"
            
    def search_by_voice(self, search_term):
        """Search using voice command"""
        try:
            self.show_search_dialog()
            # Set the search term in the dialog
            if hasattr(self, 'file_manager_window') and self.file_manager_window:
                # Find the search entry and set the value
                for widget in self.file_manager_window.winfo_children():
                    if isinstance(widget, tk.Toplevel):
                        for child in widget.winfo_children():
                            if isinstance(child, tk.Frame):
                                for grandchild in child.winfo_children():
                                    if isinstance(grandchild, tk.Entry):
                                        grandchild.delete(0, tk.END)
                                        grandchild.insert(0, search_term)
                                        return f"Searching for: {search_term}"
            return f"Search dialog opened for: {search_term}"
        except Exception as e:
            return f"Error searching: {str(e)}"
        
    def show_properties(self):
        """Show properties of selected item"""
        if not self.selected_items:
            self.status_var.set("No item selected")
            return
            
        if len(self.selected_items) > 1:
            self.status_var.set("Please select only one item for properties")
            return
            
        item_path = self.selected_items[0]
        if not os.path.exists(item_path):
            self.status_var.set("Item no longer exists")
            return
            
        # Create properties window
        props_window = tk.Toplevel(self.file_manager_window)
        props_window.title("Properties")
        props_window.geometry("400x500")
        props_window.configure(bg='#1E1E1E')
        
        # Make window draggable
        props_window.overrideredirect(True)
        props_window.attributes('-topmost', True)
        
        # Title bar
        title_bar = tk.Frame(props_window, bg='#2C2C2C', height=30)
        title_bar.pack(fill=tk.X)
        title_bar.pack_propagate(False)
        
        title_label = tk.Label(title_bar, text="Properties", bg='#2C2C2C', fg='#00FF9D', 
                              font=('Arial', 10, 'bold'))
        title_label.pack(side=tk.LEFT, padx=10, pady=5)
        
        close_btn = tk.Button(title_bar, text="✕", bg='#2C2C2C', fg='#FF5555', 
                             font=('Arial', 10, 'bold'), bd=0, padx=10,
                             command=props_window.destroy)
        close_btn.pack(side=tk.RIGHT, padx=5, pady=5)
        
        # Content
        content_frame = tk.Frame(props_window, bg='#1E1E1E')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Get file info
        filename = os.path.basename(item_path)
        is_dir = os.path.isdir(item_path)
        
        # Name
        tk.Label(content_frame, text=f"Name: {filename}", bg='#1E1E1E', fg='#FFFFFF',
                font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=5)
        
        # Type
        item_type = "Folder" if is_dir else "File"
        tk.Label(content_frame, text=f"Type: {item_type}", bg='#1E1E1E', fg='#FFFFFF',
                font=('Arial', 10)).pack(anchor=tk.W, pady=5)
        
        # Location
        location = os.path.dirname(item_path)
        tk.Label(content_frame, text=f"Location: {location}", bg='#1E1E1E', fg='#FFFFFF',
                font=('Arial', 10)).pack(anchor=tk.W, pady=5)
        
        # Size
        try:
            if is_dir:
                size = self.get_folder_size(item_path)
                size_str = self.format_size(size)
            else:
                size = os.path.getsize(item_path)
                size_str = self.format_size(size)
            tk.Label(content_frame, text=f"Size: {size_str}", bg='#1E1E1E', fg='#FFFFFF',
                    font=('Arial', 10)).pack(anchor=tk.W, pady=5)
        except:
            tk.Label(content_frame, text="Size: Unknown", bg='#1E1E1E', fg='#FFFFFF',
                    font=('Arial', 10)).pack(anchor=tk.W, pady=5)
        
        # Created
        try:
            created = os.path.getctime(item_path)
            created_str = datetime.datetime.fromtimestamp(created).strftime('%Y-%m-%d %H:%M:%S')
            tk.Label(content_frame, text=f"Created: {created_str}", bg='#1E1E1E', fg='#FFFFFF',
                    font=('Arial', 10)).pack(anchor=tk.W, pady=5)
        except:
            tk.Label(content_frame, text="Created: Unknown", bg='#1E1E1E', fg='#FFFFFF',
                    font=('Arial', 10)).pack(anchor=tk.W, pady=5)
        
        # Modified
        try:
            modified = os.path.getmtime(item_path)
            modified_str = datetime.datetime.fromtimestamp(modified).strftime('%Y-%m-%d %H:%M:%S')
            tk.Label(content_frame, text=f"Modified: {modified_str}", bg='#1E1E1E', fg='#FFFFFF',
                    font=('Arial', 10)).pack(anchor=tk.W, pady=5)
        except:
            tk.Label(content_frame, text="Modified: Unknown", bg='#1E1E1E', fg='#FFFFFF',
                    font=('Arial', 10)).pack(anchor=tk.W, pady=5)
        
        # Attributes
        try:
            import stat
            st = os.stat(item_path)
            attributes = []
            if stat.S_IREAD & st.st_mode:
                attributes.append("Read-only")
            if stat.S_IWRITE & st.st_mode:
                attributes.append("Writable")
            if stat.S_IEXEC & st.st_mode:
                attributes.append("Executable")
            if stat.S_IFDIR & st.st_mode:
                attributes.append("Directory")
            if stat.S_IFREG & st.st_mode:
                attributes.append("Regular file")
                
            attrs_str = ", ".join(attributes) if attributes else "Normal"
            tk.Label(content_frame, text=f"Attributes: {attrs_str}", bg='#1E1E1E', fg='#FFFFFF',
                    font=('Arial', 10)).pack(anchor=tk.W, pady=5)
        except:
            tk.Label(content_frame, text="Attributes: Unknown", bg='#1E1E1E', fg='#FFFFFF',
                    font=('Arial', 10)).pack(anchor=tk.W, pady=5)
        
        # Center window
        props_window.update_idletasks()
        width = props_window.winfo_width()
        height = props_window.winfo_height()
        x = (props_window.winfo_screenwidth() // 2) - (width // 2)
        y = (props_window.winfo_screenheight() // 2) - (height // 2)
        props_window.geometry(f'{width}x{height}+{x}+{y}')
        
    def get_folder_size(self, folder_path):
        """Get total size of folder"""
        total_size = 0
        try:
            for dirpath, dirnames, filenames in os.walk(folder_path):
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    try:
                        total_size += os.path.getsize(filepath)
                    except:
                        continue
        except:
            pass
        return total_size
        
    def format_size(self, size_bytes):
        """Format size in human readable format"""
        if size_bytes == 0:
            return "0 B"
        size_names = ["B", "KB", "MB", "GB", "TB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
        return f"{size_bytes:.1f} {size_names[i]}"
        
    def show_search_dialog(self):
        """Show search dialog"""
        search_window = tk.Toplevel(self.file_manager_window)
        search_window.title("Search Files")
        search_window.geometry("500x400")
        search_window.configure(bg='#1E1E1E')
        
        # Make window draggable
        search_window.overrideredirect(True)
        search_window.attributes('-topmost', True)
        
        # Title bar
        title_bar = tk.Frame(search_window, bg='#2C2C2C', height=30)
        title_bar.pack(fill=tk.X)
        title_bar.pack_propagate(False)
        
        title_label = tk.Label(title_bar, text="Search Files", bg='#2C2C2C', fg='#00FF9D', 
                              font=('Arial', 10, 'bold'))
        title_label.pack(side=tk.LEFT, padx=10, pady=5)
        
        close_btn = tk.Button(title_bar, text="✕", bg='#2C2C2C', fg='#FF5555', 
                             font=('Arial', 10, 'bold'), bd=0, padx=10,
                             command=search_window.destroy)
        close_btn.pack(side=tk.RIGHT, padx=5, pady=5)
        
        # Content
        content_frame = tk.Frame(search_window, bg='#1E1E1E')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Search input
        search_label = tk.Label(content_frame, text="Search for:", bg='#1E1E1E', fg='#00FF9D',
                               font=('Arial', 10, 'bold'))
        search_label.pack(anchor=tk.W, pady=5)
        
        search_var = tk.StringVar()
        search_entry = tk.Entry(content_frame, textvariable=search_var, bg='#2C2C2C', fg='#FFFFFF',
                               font=('Arial', 10), bd=1, relief=tk.SUNKEN)
        search_entry.pack(fill=tk.X, pady=5)
        search_entry.focus()
        
        # Search options
        options_frame = tk.Frame(content_frame, bg='#1E1E1E')
        options_frame.pack(fill=tk.X, pady=10)
        
        # Search in subdirectories
        subdir_var = tk.BooleanVar(value=True)
        subdir_check = tk.Checkbutton(options_frame, text="Search in subdirectories", 
                                     variable=subdir_var, bg='#1E1E1E', fg='#FFFFFF',
                                     selectcolor='#2C2C2C', activebackground='#1E1E1E',
                                     activeforeground='#FFFFFF')
        subdir_check.pack(anchor=tk.W, pady=2)
        
        # Case sensitive
        case_var = tk.BooleanVar(value=False)
        case_check = tk.Checkbutton(options_frame, text="Case sensitive", 
                                   variable=case_var, bg='#1E1E1E', fg='#FFFFFF',
                                   selectcolor='#2C2C2C', activebackground='#1E1E1E',
                                   activeforeground='#FFFFFF')
        case_check.pack(anchor=tk.W, pady=2)
        
        # Search button
        search_btn = tk.Button(content_frame, text="🔍 Search", bg='#2C2C2C', fg='#00FF9D',
                              font=('Arial', 10, 'bold'), bd=1, relief=tk.RAISED,
                              command=lambda: self.perform_search(search_var.get(), subdir_var.get(), 
                                                                 case_var.get(), search_window))
        search_btn.pack(pady=10)
        
        # Results list
        results_label = tk.Label(content_frame, text="Results:", bg='#1E1E1E', fg='#00FF9D',
                                font=('Arial', 10, 'bold'))
        results_label.pack(anchor=tk.W, pady=5)
        
        results_frame = tk.Frame(content_frame, bg='#1E1E1E')
        results_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        results_listbox = tk.Listbox(results_frame, bg='#1E1E1E', fg='#FFFFFF',
                                    selectbackground='#404040', selectforeground='#00FF9D',
                                    font=('Arial', 9))
        results_scrollbar = tk.Scrollbar(results_frame, orient=tk.VERTICAL, command=results_listbox.yview)
        results_listbox.configure(yscrollcommand=results_scrollbar.set)
        
        results_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        results_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind double-click to open file
        results_listbox.bind('<Double-Button-1>', lambda e: self.open_search_result(results_listbox))
        
        # Center window
        search_window.update_idletasks()
        width = search_window.winfo_width()
        height = search_window.winfo_height()
        x = (search_window.winfo_screenwidth() // 2) - (width // 2)
        y = (search_window.winfo_screenheight() // 2) - (height // 2)
        search_window.geometry(f'{width}x{height}+{x}+{y}')
        
        # Store references
        search_window.results_listbox = results_listbox
        
    def perform_search(self, search_term, search_subdirs, case_sensitive, search_window):
        """Perform file search"""
        if not search_term.strip():
            return
            
        results_listbox = search_window.results_listbox
        results_listbox.delete(0, tk.END)
        
        search_term = search_term if case_sensitive else search_term.lower()
        found_files = []
        
        try:
            if search_subdirs:
                for root, dirs, files in os.walk(self.current_path):
                    for file in files:
                        if case_sensitive:
                            if search_term in file:
                                found_files.append(os.path.join(root, file))
                        else:
                            if search_term in file.lower():
                                found_files.append(os.path.join(root, file))
            else:
                for item in os.listdir(self.current_path):
                    item_path = os.path.join(self.current_path, item)
                    if os.path.isfile(item_path):
                        if case_sensitive:
                            if search_term in item:
                                found_files.append(item_path)
                        else:
                            if search_term in item.lower():
                                found_files.append(item_path)
                                
            # Display results
            for file_path in found_files[:100]:  # Limit to 100 results
                rel_path = os.path.relpath(file_path, self.current_path)
                results_listbox.insert(tk.END, rel_path)
                
            if len(found_files) > 100:
                results_listbox.insert(tk.END, f"... and {len(found_files) - 100} more files")
                
            search_window.title(f"Search Files - {len(found_files)} results")
            
        except Exception as e:
            results_listbox.insert(tk.END, f"Error during search: {str(e)}")
            
    def open_search_result(self, results_listbox):
        """Open selected search result"""
        selection = results_listbox.curselection()
        if selection:
            index = selection[0]
            rel_path = results_listbox.get(index)
            if not rel_path.startswith("..."):
                file_path = os.path.join(self.current_path, rel_path)
                if os.path.exists(file_path):
                    if os.path.isdir(file_path):
                        self.load_directory(file_path)
                    else:
                        self.open_file(file_path)
                        
    def show_context_menu(self, event):
        """Show context menu for right-click"""
        if not self.selected_items:
            return
            
        context_menu = tk.Menu(self.file_manager_window, tearoff=0, bg='#2C2C2C', fg='#FFFFFF',
                              activebackground='#404040', activeforeground='#00FF9D')
        
        # Open
        context_menu.add_command(label="Open", command=self.open_selected)
        
        # Open with
        context_menu.add_command(label="Open with...", command=self.open_with)
        
        context_menu.add_separator()
        
        # Cut, Copy, Paste
        context_menu.add_command(label="Cut", command=self.cut_selected)
        context_menu.add_command(label="Copy", command=self.copy_selected)
        context_menu.add_command(label="Paste", command=self.paste_items)
        
        context_menu.add_separator()
        
        # Rename, Delete
        context_menu.add_command(label="Rename", command=self.rename_selected)
        context_menu.add_command(label="Delete", command=self.delete_selected)
        
        context_menu.add_separator()
        
        # Properties
        context_menu.add_command(label="Properties", command=self.show_properties)
        
        # Show menu at cursor position
        context_menu.tk_popup(event.x_root, event.y_root)
        
    def open_selected(self):
        """Open selected item"""
        if len(self.selected_items) == 1:
            item_path = self.selected_items[0]
            if os.path.isdir(item_path):
                self.load_directory(item_path)
            else:
                self.open_file(item_path)
                
    def open_with(self):
        """Open with dialog (placeholder)"""
        if len(self.selected_items) == 1:
            self.status_var.set("Open with dialog not implemented yet")
            
    def rename_selected(self):
        """Rename selected item"""
        if len(self.selected_items) != 1:
            self.status_var.set("Please select exactly one item to rename")
            return
            
        item_path = self.selected_items[0]
        old_name = os.path.basename(item_path)
        parent_dir = os.path.dirname(item_path)
        
        new_name = simpledialog.askstring("Rename", f"Enter new name for '{old_name}':", initialvalue=old_name)
        if new_name and new_name != old_name:
            try:
                new_path = os.path.join(parent_dir, new_name)
                os.rename(item_path, new_path)
                self.load_directory(self.current_path)
                self.status_var.set(f"Renamed '{old_name}' to '{new_name}'")
            except Exception as e:
                self.status_var.set(f"Error renaming: {str(e)}")
                
    def select_all(self):
        """Select all items in current directory"""
        self.file_listbox.selection_set(0, tk.END)
        self.on_file_select(None)  # Update selected_items

# ---------- API Keys ----------
GEMINI_API_KEY = "YOUR-API-KEY"
WEATHER_API_KEY = "YOUR-API-KEY"
NEWS_API_KEY = "YOUR-API-KEY"

genai.configure(api_key=GEMINI_API_KEY)
gemini_model = genai.GenerativeModel("gemini-2.5-flash")

# ---------- Enhanced TTS Engine ----------
# ---------- Global Variables ----------
LISTENING = True
TTS_ENABLED = True
FULLSCREEN = False
AUTO_HIDE = False  # Enable auto-hide feature by default
_last_browser_spoken = ""

# Initialize TTS engine
def init_tts_engine():
    """Initialize TTS engine with advanced features"""
    try:
        # Try to use Windows SAPI voice for better quality
        import win32com.client
        engine = win32com.client.Dispatch("SAPI.SpVoice")
        print(f"Using voice: {engine.Voice.GetDescription()}")
        return engine
    except Exception as e:
        print(f"Falling back to pyttsx3: {e}")
        try:
            import pyttsx3
            engine = pyttsx3.init()
            # Set a slower speaking rate for better clarity
            engine.setProperty('rate', 175)
            return engine
        except Exception as e:
            print(f"Failed to initialize TTS: {e}")
            return None

# Initialize the TTS engine
try:
    engine = init_tts_engine()
    if engine is None:
        print("Warning: TTS engine could not be initialized. Voice output will be disabled.")
        TTS_ENABLED = False
except Exception as e:
    print(f"Error initializing TTS engine: {e}")
    TTS_ENABLED = False

# Speech functions
def speak(text: str):
    """Speak text using TTS"""
    global TTS_ENABLED
    
    if not TTS_ENABLED or engine is None:
        print(f"[TTS] {text}")
        return
        
    try:
        if hasattr(engine, 'Speak'):  # Windows SAPI
            engine.Speak(text)
        else:  # pyttsx3
            engine.say(text)
            engine.runAndWait()
    except Exception as e:
        print(f"Error in TTS: {e}")
        TTS_ENABLED = False  # Disable TTS on error

def start_listening():
    """Start listening for voice commands"""
    global LISTENING
    LISTENING = True
    print("Voice recognition started")

def stop_listening():
    """Stop listening for voice commands"""
    global LISTENING
    LISTENING = False
    print("Voice recognition stopped")

# ---------- Enhanced 3D Animation Class ----------
class Advanced3DAnimation:
    def __init__(self, canvas):
        self.canvas = canvas
        self.width = 400
        self.height = 600
        self.center_x = self.width // 2
        self.center_y = self.height // 2 - 50
        self.angle = 0
        self.scan_angle = 0
        self.pulse_radius = 0
        self.pulse_growing = True
        self.scan_line = None
        self.circles = []
        self.dots = []
        self.status_text = None
        self.command_text = None
        self.conversation_text = []
        self.conversation_bg = None
        self.particles = []
        self.wave_effect = 0
        self.neural_network = []
        self.voice_waves = []
        self.floating_particles = []
        self.energy_field = []
        self.init_3d_hud()
        self.animate()
        self.update_status("Desktop Bot Online", "#00FF9D")
        
    def init_3d_hud(self):
        """Initialize advanced 3D HUD with Google AI Studio-like effects"""
        # Clear any existing items
        self.canvas.delete('hud')
        
        # Create gradient background effect
        self.create_gradient_background()
        
        # Create neural network nodes
        self.create_neural_network()
        
        # Create energy field
        self.create_energy_field()
        
        # Create floating particles
        self.create_floating_particles()
        
        # Create concentric circles with 3D effect and glow
        colors = ['#00FF9D', '#00CC7F', '#009966', '#006633']
        for i, radius in enumerate([140, 110, 80, 50]):
            # Create glow effect for circles
            glow_circle = self.canvas.create_oval(
                self.center_x - radius - 5, self.center_y - radius - 5,
                self.center_x + radius + 5, self.center_y + radius + 5,
                outline=colors[i], width=1, tags='hud'
            )
            circle = self.canvas.create_oval(
                self.center_x - radius, self.center_y - radius,
                self.center_x + radius, self.center_y + radius,
                outline=colors[i], width=3, tags='hud'
            )
            self.circles.append((glow_circle, circle))
        
        # Create outer dots with enhanced glow effect
        dot_count = 32
        for i in range(dot_count):
            angle = (2 * math.pi * i) / dot_count
            # Create multiple glow layers
            glow1 = self.canvas.create_oval(
                self.center_x - 12 + 150 * math.cos(angle),
                self.center_y - 12 + 150 * math.sin(angle),
                self.center_x + 12 + 150 * math.cos(angle),
                self.center_y + 12 + 150 * math.sin(angle),
                fill='#00FF9D', stipple='gray25', tags='hud'
            )
            glow2 = self.canvas.create_oval(
                self.center_x - 8 + 150 * math.cos(angle),
                self.center_y - 8 + 150 * math.sin(angle),
                self.center_x + 8 + 150 * math.cos(angle),
                self.center_y + 8 + 150 * math.sin(angle),
                fill='#00FF9D', stipple='gray50', tags='hud'
            )
            dot = self.canvas.create_oval(
                self.center_x - 4 + 150 * math.cos(angle),
                self.center_y - 4 + 150 * math.sin(angle),
                self.center_x + 4 + 150 * math.cos(angle),
                self.center_y + 4 + 150 * math.sin(angle),
                fill='#00FF9D', tags='hud'
            )
            self.dots.append((glow1, glow2, dot))
        
        # Add JARVIS text with enhanced 3D effect and holographic appearance
        # Shadow layer
        self.canvas.create_text(
            self.center_x + 3, self.center_y - 17,
            text="J.A.R.V.I.S.",
            font=('Arial', 18, 'bold'),
            fill='#006633',
            tags='hud'
        )
        
        # Main text
        self.canvas.create_text(
            self.center_x, self.center_y - 20,
            text="Desktop Bot.",
            font=('Arial', 18, 'bold'),
            fill='#00FF9D',
            tags='hud'
        )
        
        # Holographic overlay
        self.canvas.create_text(
            self.center_x - 1, self.center_y - 21,
            text="",
            font=('Arial', 18, 'bold'),
            fill='#00FFFF',
            tags='hud'
        )
        
        # Create conversation area with enhanced glass effect
        self.conversation_bg = self.canvas.create_rectangle(
            10, self.center_y + 120, self.width - 10, self.height - 10,
            fill='#0A0A0A', outline='#1E1E1E', width=2, tags='hud'
        )
        
        # Create voice wave visualization
        self.create_voice_waves()
        
        # Create status indicator
        self.create_status_indicator()
        
        # Create 3D circle
        self.circle_3d = Circle3D(self)
    
    def create_gradient_background(self):
        """Create gradient background effect - now transparent to show the 4K background"""
        # Skip creating gradient lines since we now have a high-quality background image
        # Instead, create a semi-transparent overlay for better HUD visibility
        overlay = self.canvas.create_rectangle(
            0, 0, self.width, self.height,
            fill='black', stipple='gray25', tags='hud'
        )
        self.canvas.tag_lower(overlay)  # Keep it behind other HUD elements but above the background
    
    def create_neural_network(self):
        """Create neural network visualization"""
        nodes = []
        connections = []
        
        # Create nodes
        for i in range(8):
            angle = (2 * math.pi * i) / 8
            x = self.center_x + 100 * math.cos(angle)
            y = self.center_y + 100 * math.sin(angle)
            node = self.canvas.create_oval(
                x - 3, y - 3, x + 3, y + 3,
                fill='#00FF9D', tags='hud'
            )
            nodes.append((x, y, node))
        
        # Create connections
        for i in range(len(nodes)):
            for j in range(i + 1, len(nodes)):
                if random.random() < 0.3:  # 30% chance of connection
                    connection = self.canvas.create_line(
                        nodes[i][0], nodes[i][1],
                        nodes[j][0], nodes[j][1],
                        fill='#00FF9D', width=1, tags='hud'
                    )
                    connections.append(connection)
        
        self.neural_network = nodes + connections
    
    def create_energy_field(self):
        """Create energy field effect"""
        field_count = 12
        for i in range(field_count):
            angle = (2 * math.pi * i) / field_count
            x = self.center_x + 200 * math.cos(angle)
            y = self.center_y + 200 * math.sin(angle)
            field = self.canvas.create_oval(
                x - 2, y - 2, x + 2, y + 2,
                fill='#00FFFF', tags='hud'
            )
            self.energy_field.append((x, y, field, angle))
    
    def create_floating_particles(self):
        """Create floating particles for 3D effect"""
        particle_count = 15
        for i in range(particle_count):
            x = random.randint(0, self.width)
            y = random.randint(0, self.height)
            size = random.randint(1, 3)
            particle = self.canvas.create_oval(
                x - size, y - size, x + size, y + size,
                fill='#00FF9D', tags='hud'
            )
            self.floating_particles.append((x, y, particle, random.uniform(0, 2*math.pi)))
    
    def create_voice_waves(self):
        """Create voice wave visualization"""
        wave_count = 5
        for i in range(wave_count):
            wave = []
            for j in range(20):
                x = 20 + j * 18
                y = self.center_y + 80 + i * 15
                point = self.canvas.create_oval(
                    x - 1, y - 1, x + 1, y + 1,
                    fill='#00FF9D', tags='hud'
                )
                wave.append((x, y, point))
            self.voice_waves.append(wave)
    
    def create_status_indicator(self):
        """Create animated status indicator"""
        self.status_indicator = self.canvas.create_oval(
            self.center_x - 15, self.center_y + 60,
            self.center_x + 15, self.center_y + 90,
            outline='#00FF9D', width=2, tags='hud'
        )
        
        # Create pulsing center
        self.status_center = self.canvas.create_oval(
            self.center_x - 5, self.center_y + 75,
            self.center_x + 5, self.center_y + 85,
            fill='#00FF9D', tags='hud'
        )
    
    def update_status(self, text, color="#00FF9D"):
        """Update status text with animation"""
        if hasattr(self, 'status_text'):
            self.canvas.delete(self.status_text)
        self.status_text = self.canvas.create_text(
            self.center_x, self.center_y + 100,
            text=text,
            font=('Arial', 12, 'bold'),
            fill=color,
            tags='hud'
        )
    
    def add_to_conversation(self, speaker, text, color="#00FF9D"):
        """Add conversation text with enhanced visibility for 4K background"""
        if not hasattr(self, 'conversation_text'):
            self.conversation_text = []
        if not hasattr(self, 'conversation_history'):
            self.conversation_history = []
            
        # Add new message to history
        self.conversation_history.append({
            'speaker': speaker,
            'text': text,
            'color': color,
            'timestamp': time.time()
        })
        
        # Keep only the last 8 messages to avoid clutter
        if len(self.conversation_history) > 8:
            self.conversation_history = self.conversation_history[-8:]
        
        # Clear previous conversation display
        for item in self.conversation_text:
            self.canvas.delete(item)
        
        # Create semi-transparent background for conversation area
        conv_bg = self.canvas.create_rectangle(
            10, self.center_y + 120, self.width - 10, self.height - 10,
            fill='black', stipple='gray50', outline='#1E1E1E', width=2, tags='hud'
        )
        self.conversation_text.append(conv_bg)
        
        # Display all messages in conversation history
        start_y = self.center_y + 140
        line_height = 25
        max_width = 360
        
        for i, message in enumerate(self.conversation_history):
            # Calculate position for this line
            y_pos = start_y + (i * line_height)
            
            # Create the message text with proper formatting
            msg_text = f"{message['speaker']}: {message['text']}"
            
            # Create text shadow for better readability against 4K background
            shadow = self.canvas.create_text(
                22, y_pos + 2,
                text=msg_text,
                font=('Arial', 9, 'bold'),
                fill='black',
                anchor='nw',
                tags='hud',
                width=max_width
            )
            self.conversation_text.append(shadow)
            
            # Create the text item with enhanced visibility
            msg = self.canvas.create_text(
                20, y_pos,
                text=msg_text,
                font=('Arial', 9, 'bold'),
                fill=message['color'],
                anchor='nw',
                tags='hud',
                width=max_width
            )
            
            self.conversation_text.append(msg)
            
            # Add a stronger glow effect for the most recent message
            if i == len(self.conversation_history) - 1:
                # Create a subtle background glow for the latest message
                bbox = self.canvas.bbox(msg)
                if bbox:
                    glow_bg = self.canvas.create_rectangle(
                        bbox[0] - 5, bbox[1] - 2,
                        bbox[2] + 5, bbox[3] + 2,
                        fill='', outline=message['color'],
                        width=2, stipple='gray12',
                        tags='hud'
                    )
                    self.conversation_text.append(glow_bg)
        
        # Removed automatic clearing - messages will persist until manually cleared
        # If a chat history window is open, refresh its contents
        try:
            if 'refresh_chat_history' in globals():
                refresh_chat_history()
        except Exception:
            pass
    
    def clear_old_messages(self):
        """Clear old messages from conversation history"""
        if hasattr(self, 'conversation_history') and self.conversation_history:
            # Keep only the last 4 messages
            if len(self.conversation_history) > 4:
                self.conversation_history = self.conversation_history[-4:]
                # Redraw the conversation
                self.redraw_conversation()
    
    def redraw_conversation(self):
        """Redraw the entire conversation display"""
        if not hasattr(self, 'conversation_text'):
            return
            
        # Clear current display
        for item in self.conversation_text:
            self.canvas.delete(item)
        
        self.conversation_text = []
        
        if not hasattr(self, 'conversation_history'):
            return
            
        # Redraw all messages
        start_y = self.center_y + 140
        line_height = 25
        max_width = 360
        
        for i, message in enumerate(self.conversation_history):
            y_pos = start_y + (i * line_height)
            msg_text = f"{message['speaker']}: {message['text']}"
            
            msg = self.canvas.create_text(
                20, y_pos,
                text=msg_text,
                font=('Arial', 9),
                fill=message['color'],
                anchor='nw',
                tags='hud',
                width=max_width
            )
            
            self.conversation_text.append(msg)
            
            # Add glow effect for the most recent message
            if i == len(self.conversation_history) - 1:
                bbox = self.canvas.bbox(msg)
                if bbox:
                    glow_bg = self.canvas.create_rectangle(
                        bbox[0] - 5, bbox[1] - 2,
                        bbox[2] + 5, bbox[3] + 2,
                        fill='', outline=message['color'],
                        width=1, stipple='gray25',
                        tags='hud'
                    )
                    self.conversation_text.append(glow_bg)
    
    def clear_conversation(self):
        """Clear all conversation messages"""
        if hasattr(self, 'conversation_text'):
            for item in self.conversation_text:
                self.canvas.delete(item)
            self.conversation_text = []
        if hasattr(self, 'conversation_history'):
            self.conversation_history = []
    
    def animate(self):
        """Advanced animation loop with 3D effects"""
        # Rotate outer dots with enhanced glow effect
        self.angle = (self.angle + 1) % 360
        rad_angle = math.radians(self.angle)
        
        # Update dot positions with enhanced glow
        for i, (glow1, glow2, dot) in enumerate(self.dots):
            angle = (2 * math.pi * i / len(self.dots)) + rad_angle
            x = self.center_x + 150 * math.cos(angle)
            y = self.center_y + 150 * math.sin(angle)
            
            # Update glow positions
            self.canvas.coords(
                glow1,
                x - 12, y - 12, x + 12, y + 12
            )
            self.canvas.coords(
                glow2,
                x - 8, y - 8, x + 8, y + 8
            )
            
            # Update dot position
            self.canvas.coords(
                dot,
                x - 4, y - 4, x + 4, y + 4
            )
        
        # Pulsing effect for center circle
        if self.pulse_growing:
            self.pulse_radius += 0.8
            if self.pulse_radius > 25:
                self.pulse_growing = False
        else:
            self.pulse_radius -= 0.8
            if self.pulse_radius < 0:
                self.pulse_growing = True
        
        # Update scan line with enhanced wave effect
        self.scan_angle = (self.scan_angle + 3) % 360
        scan_rad = math.radians(self.scan_angle)
        
        if hasattr(self, 'scan_line'):
            self.canvas.delete(self.scan_line)
        
        # Create enhanced wave effect
        wave_offset = math.sin(math.radians(self.scan_angle * 2)) * 10
        self.scan_line = self.canvas.create_line(
            self.center_x,
            self.center_y,
            self.center_x + (120 + wave_offset) * math.cos(scan_rad),
            self.center_y + (120 + wave_offset) * math.sin(scan_rad),
            fill='#00FF9D',
            width=2,
            tags='hud'
        )
        
        # Animate voice waves
        self.animate_voice_waves()
        
        # Animate neural network
        self.animate_neural_network()
        
        # Animate energy field
        self.animate_energy_field()
        
        # Animate floating particles
        self.animate_floating_particles()
        
        # Animate status indicator
        self.animate_status_indicator()
        
        # Animate circles with glow effect
        self.animate_circles()
        
        # Schedule next frame
        self.canvas.after(25, self.animate)
    
    def animate_voice_waves(self):
        """Animate voice wave visualization"""
        for wave in self.voice_waves:
            for x, y, point in wave:
                # Create wave effect based on time
                wave_offset = math.sin(time.time() * 2 + x * 0.1) * 5
                self.canvas.coords(
                    point,
                    x - 1, y + wave_offset - 1,
                    x + 1, y + wave_offset + 1
                )
    
    def animate_neural_network(self):
        """Animate neural network connections"""
        for i, (x, y, node) in enumerate(self.neural_network[:8]):  # Only nodes
            # Pulsing effect for nodes
            pulse = math.sin(time.time() * 3 + i) * 2
            self.canvas.coords(
                node,
                x - 3 - pulse, y - 3 - pulse,
                x + 3 + pulse, y + 3 + pulse
            )
    
    def animate_energy_field(self):
        """Animate energy field particles"""
        for x, y, field, angle in self.energy_field:
            # Rotate energy field particles
            new_angle = angle + 0.02
            new_x = self.center_x + 200 * math.cos(new_angle)
            new_y = self.center_y + 200 * math.sin(new_angle)
            
            # Update position
            self.canvas.coords(
                field,
                new_x - 2, new_y - 2, new_x + 2, new_y + 2
            )
            
            # Update stored values
            self.energy_field[self.energy_field.index((x, y, field, angle))] = (new_x, new_y, field, new_angle)
    
    def animate_floating_particles(self):
        """Animate floating particles"""
        for i, (x, y, particle, angle) in enumerate(self.floating_particles):
            # Move particles in circular motion
            new_x = x + math.cos(angle) * 0.5
            new_y = y + math.sin(angle) * 0.5
            
            # Wrap around screen
            if new_x < 0:
                new_x = self.width
            elif new_x > self.width:
                new_x = 0
            if new_y < 0:
                new_y = self.height
            elif new_y > self.height:
                new_y = 0
            
            # Update position
            self.canvas.coords(
                particle,
                new_x - 1, new_y - 1, new_x + 1, new_y + 1
            )
            
            # Update stored values
            self.floating_particles[i] = (new_x, new_y, particle, angle + 0.01)

    def animate_status_indicator(self):
        """Animate status indicator with pulsing effect"""
        if hasattr(self, 'status_center'):
            # Pulsing effect for status center
            pulse = math.sin(time.time() * 4) * 3
            self.canvas.coords(
                self.status_center,
                self.center_x - 5 - pulse, self.center_y + 75 - pulse,
                self.center_x + 5 + pulse, self.center_y + 85 + pulse
            )
    
    def animate_circles(self):
        """Animate circles with glow effect"""
        for i, (glow_circle, circle) in enumerate(self.circles):
            # Subtle pulsing effect for circles
            pulse = math.sin(time.time() * 2 + i) * 2
            radius = [140, 110, 80, 50][i]
            
            # Update glow circle
            self.canvas.coords(
                glow_circle,
                self.center_x - radius - 5 - pulse, self.center_y - radius - 5 - pulse,
                self.center_x + radius + 5 + pulse, self.center_y + radius + 5 + pulse
            )
            
            # Update main circle
            self.canvas.coords(
                circle,
                self.center_x - radius - pulse, self.center_y - radius - pulse,
                self.center_x + radius + pulse, self.center_y + radius + pulse
            )

# ---------- AI Response Functions ----------
def ask_gemini(prompt: str) -> str:
    try:
        response = gemini_model.generate_content(
            prompt,
            generation_config=genai.types.GenerationConfig(
                max_output_tokens=120,
                temperature=0.3,
                top_p=0.8,
                top_k=20
            )
        )
        return response.text.strip()
    except Exception as e:
        return get_offline_response(prompt)

def get_offline_response(prompt: str) -> str:
    """Provide offline responses when API is unavailable"""
    prompt_lower = prompt.lower()
    
    offline_responses = {
        "hello": "Hello! I'm running in offline mode.",
        "hi": "Hi there! Currently using offline responses.",
        "how are you": "I'm doing well, though running offline right now.",
        "what is your name": "I'm your AI assistant, currently in offline mode.",
        "help": "I can help with basic tasks like time, date, opening websites, and system commands.",
        "what can you do": "I can check time/date, open websites, control volume, take screenshots, and more.",
        "thank you": "You're welcome!",
        "thanks": "Happy to help!",
        "good morning": "Good morning! Hope you have a great day!",
        "good evening": "Good evening! How can I assist you?",
        "good night": "Good night! Sleep well!",
        "weather": "Weather service requires internet. Try checking your local weather app.",
        "news": "News service is currently unavailable. Check your preferred news website.",
        "joke": "Why don't scientists trust atoms? Because they make up everything!",
        "tell me a joke": "What do you call a bear with no teeth? A gummy bear!",
        "funny": "Here's something funny: I tried to catch some fog, but I mist!",
        "who are you": "I'm your AI assistant, currently running in offline mode.",
        "what time": "Use the 'time' command to get the current time.",
        "what date": "Use the 'date' command to get today's date.",
    }
    
    # Check for exact matches
    for key, response in offline_responses.items():
        if key in prompt_lower:
            return response
    
    # # Generic responses
    # if "what" in prompt_lower or "how" in prompt_lower or "why" in prompt_lower:
    #     return "I'm currently in offline mode. I can help with basic system tasks."
    # elif "can you" in prompt_lower:
    #     return "I can help with system tasks. AI responses are limited due to daily quota."
    # elif "?" in prompt:
    #     return "I'm in offline mode right now. Try asking for time, date, or system commands."
    # else:
    #     return "API quota exceeded. I can still help with system commands."

# ---------- AI Response Functions ----------
def get_ai_response(query):
    """Generate a response using Google's Gemini AI with weather integration"""
    try:
        query_lower = query.lower()
        
        # Check for weather queries first
        if "weather" in query_lower:
            # Extract location if mentioned
            location = ""
            if "in " in query_lower:
                location = query_lower.split("in ", 1)[1].strip()
            
            if not location:
                # Default to current location if not specified
                try:
                    # Try to get approximate location from IP
                    location_data = requests.get('https://ipinfo.io/json').json()
                    location = location_data.get('city', 'New York')  # Default to New York if can't detect
                except:
                    location = "New York"
            
            return get_weather(location)
        
        # Use Gemini for all other queries
        try:
            # Configure Gemini with the API key
            genai.configure(api_key=GEMINI_API_KEY)
            
            # Create the model
            model = genai.GenerativeModel('gemini-2.5-flash')
            
            # Generate response
            response = model.generate_content(
                f"You are DESKTOP BOT, a helpful AI assistant. Answer concisely in 1-2 sentences. Question: {query}"
            )
            
            return response.text
            
        except Exception as e:
            print(f"Gemini API Error: {e}")
            # Fallback to basic responses if Gemini fails
            return "I'm having trouble connecting to the AI service. Please try again later."
            
    except Exception as e:
        print(f"Error in AI response: {e}")
        return "I encountered an error processing your request. Could you please rephrase?"

def get_weather(city):
    """Get weather information for a city"""
    try:
        if not WEATHER_API_KEY:
            return "Weather API key is not configured."
            
        # Get coordinates for the city
        geo_url = f"http://api.openweathermap.org/geo/1.0/direct?q={city}&limit=1&appid={WEATHER_API_KEY}"
        geo_response = requests.get(geo_url)
        geo_data = geo_response.json()
        
        if not geo_data:
            return f"Could not find weather information for {city}."
            
        lat = geo_data[0]['lat']
        lon = geo_data[0]['lon']
        
        # Get weather data
        weather_url = f"https://api.openweathermap.org/data/2.5/weather?lat={lat}&lon={lon}&appid={WEATHER_API_KEY}&units=metric"
        weather_response = requests.get(weather_url)
        weather_data = weather_response.json()
        
        if weather_data.get('cod') != 200:
            return f"Could not fetch weather data for {city}."
        
        # Format weather information
        temp = weather_data['main']['temp']
        feels_like = weather_data['main']['feels_like']
        humidity = weather_data['main']['humidity']
        description = weather_data['weather'][0]['description']
        wind_speed = weather_data['wind']['speed']
        
        return (
            f"Weather in {city}: {description}. "
            f"Temperature: {temp}°C (feels like {feels_like}°C). "
            f"Humidity: {humidity}%. Wind: {wind_speed} m/s."
        )
        
    except Exception as e:
        print(f"Weather API Error: {e}")
        return "I couldn't fetch the weather information right now. Please try again later."

# ---------- System Control Functions ----------
def shutdown_system():
    window_shake_effect()
    pwd = getpass.getpass("Shutdown password: ")
    if pwd == "1234":
        speak("Shutting down…")
        subprocess.call(["shutdown", "/s", "/t", "0"])
    else:
        speak("Wrong password.")

def restart_system():
    window_shake_effect()
    pwd = getpass.getpass("Restart password: ")
    if pwd == "5678":
        speak("Restarting…")
        subprocess.call(["shutdown", "/r", "/t", "0"])
    else:
        speak("Wrong password.")

# ---------- System Settings Functions ----------
def open_windows_settings():
    """Open Windows Settings"""
    os.system('start ms-settings:')
    return "Opening Windows Settings"

def open_display_settings():
    """Open Display Settings"""
    os.system('start ms-settings:display')
    return "Opening Display Settings"

def open_sound_settings():
    """Open Sound Settings"""
    os.system('start ms-settings:sound')
    return "Opening Sound Settings"

def open_network_settings():
    """Open Network & Internet Settings"""
    os.system('start ms-settings:network')
    return "Opening Network Settings"

def open_wifi_settings():
    """Open Wi-Fi Settings"""
    os.system('start ms-settings:network-wifi')
    return "Opening Wi-Fi Settings"

def open_bluetooth_settings():
    """Open Bluetooth Settings"""
    os.system('start ms-settings:bluetooth')
    return "Opening Bluetooth Settings"

def open_notification_settings():
    """Open Notification Settings"""
    os.system('start ms-settings:notifications')
    return "Opening Notification Settings"

def open_storage_settings():
    """Open Storage Settings"""
    os.system('start ms-settings:storagesense')
    return "Opening Storage Settings"

def open_battery_settings():
    """Open Battery Settings"""
    os.system('start ms-settings:batterysaver')
    return "Opening Battery Settings"

def open_privacy_settings():
    """Open Privacy Settings"""
    os.system('start ms-settings:privacy')
    return "Opening Privacy Settings"

def open_update_settings():
    """Open Windows Update Settings"""
    os.system('start ms-settings:windowsupdate')
    return "Opening Windows Update Settings"

def open_default_apps():
    """Open Default Apps Settings"""
    os.system('start ms-settings:defaultapps')
    return "Opening Default Apps Settings"

def open_control_panel():
    """Open Control Panel"""
    os.system('control')
    return "Opening Control Panel"

def open_task_manager():
    """Open Task Manager"""
    os.system('taskmgr')
    return "Opening Task Manager"

def open_system_properties():
    """Open System Properties"""
    os.system('sysdm.cpl')
    return "Opening System Properties"

def open_device_manager():
    """Open Device Manager"""
    os.system('devmgmt.msc')
    return "Opening Device Manager"

def open_disk_cleanup():
    """Open Disk Cleanup"""
    os.system('cleanmgr')
    return "Opening Disk Cleanup"

def open_system_restore():
    """Open System Restore"""
    os.system('rstrui')
    return "Opening System Restore"

def open_registry_editor():
    """Open Registry Editor"""
    os.system('regedit')
    return "Opening Registry Editor"

def open_event_viewer():
    """Open Event Viewer"""
    os.system('eventvwr')
    return "Opening Event Viewer"

# ---------- Utility Functions ----------
def volume_up():
    [pyautogui.press("volumeup") for _ in range(5)]
    speak("Volume up")

def volume_down():
    [pyautogui.press("volumedown") for _ in range(5)]
    speak("Volume down")

def brightness_up():
    os.system("powershell (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,100)")

def brightness_down():
    os.system("powershell (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,20)")

def screenshot():
    pyautogui.screenshot("screenshot.png")
    speak("Screenshot saved")

def lock_pc():
    ctypes.windll.user32.LockWorkStation()
    speak("Locked")

def new_desktop():
    pyautogui.hotkey("win", "ctrl", "d")
    speak("New desktop")

def close_desktop():
    pyautogui.hotkey("win", "ctrl", "f4")
    speak("Desktop closed")

def switch_desktop():
    pyautogui.hotkey("win", "ctrl", "right")
    speak("Switched desktop")

def system_info():
    info = f"CPU: {psutil.cpu_percent()}%  RAM: {psutil.virtual_memory().percent}%"
    speak(info)

def minimize_window():
    """Minimize the active window"""
    try:
        hwnd = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        speak("Window minimized")
    except Exception as e:
        print(f"Minimize error: {e}")
        speak("Could not minimize window")

def minimize_google():
    """Minimize Google Chrome specifically"""
    try:
        # Find Chrome windows with better detection
        def enum_windows_callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd).lower()
                # More flexible Chrome detection
                if ('chrome' in window_text or 
                    'google' in window_text or 
                    'chromium' in window_text or
                    'browser' in window_text):
                    # Additional check: make sure it's not just a partial match
                    if len(window_text) > 3:  # Avoid very short window titles
                        windows.append((hwnd, window_text))
            return True
        
        chrome_windows = []
        win32gui.EnumWindows(enum_windows_callback, chrome_windows)
        
        print(f"Found {len(chrome_windows)} potential Chrome windows:")
        for hwnd, title in chrome_windows:
            print(f"  - {title} (hwnd: {hwnd})")
        
        if chrome_windows:
            minimized_count = 0
            for hwnd, title in chrome_windows:
                try:
                    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                    minimized_count += 1
                    print(f"Successfully minimized: {title}")
                except Exception as e:
                    print(f"Failed to minimize {title}: {e}")
            
            if minimized_count > 0:
                speak(f"Minimized {minimized_count} Chrome windows")
            else:
                speak("Failed to minimize Chrome windows")
        else:
            print("No Chrome windows found")
            speak("Google Chrome not found")
    except Exception as e:
        print(f"Minimize Google error: {e}")
        speak("Could not minimize Google Chrome")

def minimize_browser():
    """Minimize any browser application"""
    try:
        browsers = ['chrome', 'firefox', 'edge', 'opera', 'brave', 'safari', 'chromium']
        minimized_count = 0
        found_windows = []
        
        def enum_windows_callback(hwnd, data):
            nonlocal minimized_count
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd).lower()
                for browser in browsers:
                    if browser in window_text and len(window_text) > 3:
                        found_windows.append((hwnd, window_text, browser))
                        break
            return True
        
        win32gui.EnumWindows(enum_windows_callback, None)
        
        print(f"Found {len(found_windows)} browser windows:")
        for hwnd, title, browser_type in found_windows:
            print(f"  - {title} ({browser_type})")
        
        for hwnd, title, browser_type in found_windows:
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                minimized_count += 1
                print(f"Successfully minimized: {title}")
            except Exception as e:
                print(f"Failed to minimize {title}: {e}")
        
        if minimized_count > 0:
            speak(f"Minimized {minimized_count} browser windows")
        else:
            speak("No browser windows found or failed to minimize")
    except Exception as e:
        print(f"Minimize browser error: {e}")
        speak("Could not minimize browser")

def minimize_application(app_name):
    """Minimize a specific application by name"""
    try:
        minimized_count = 0
        found_windows = []
        
        def enum_windows_callback(hwnd, data):
            nonlocal minimized_count
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd).lower()
                if app_name.lower() in window_text and len(window_text) > 3:
                    found_windows.append((hwnd, window_text))
            return True
        
        win32gui.EnumWindows(enum_windows_callback, None)
        
        print(f"Searching for '{app_name}' - Found {len(found_windows)} windows:")
        for hwnd, title in found_windows:
            print(f"  - {title}")
        
        for hwnd, title in found_windows:
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                minimized_count += 1
                print(f"Successfully minimized: {title}")
            except Exception as e:
                print(f"Failed to minimize {title}: {e}")
        
        if minimized_count > 0:
            speak(f"Minimized {app_name}")
        else:
            speak(f"{app_name} not found or failed to minimize")
    except Exception as e:
        print(f"Minimize {app_name} error: {e}")
        speak(f"Could not minimize {app_name}")

def close_application(app_name):
    """Close a specific application by name"""
    try:
        closed_count = 0
        found_windows = []
        
        def enum_windows_callback(hwnd, data):
            nonlocal closed_count
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd).lower()
                if app_name.lower() in window_text and len(window_text) > 3:
                    found_windows.append((hwnd, window_text))
            return True
        
        win32gui.EnumWindows(enum_windows_callback, None)
        
        print(f"Searching for '{app_name}' to close - Found {len(found_windows)} windows:")
        for hwnd, title in found_windows:
            print(f"  - {title}")
        
        for hwnd, title in found_windows:
            try:
                # Try to close gracefully first
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                closed_count += 1
                print(f"Sent close signal to: {title}")
            except Exception as e:
                print(f"Failed to close {title}: {e}")
                # Try force close if graceful close fails
                try:
                    import psutil
                    for proc in psutil.process_iter(['pid', 'name']):
                        if proc.info['name'] and app_name.lower() in proc.info['name'].lower():
                            proc.terminate()
                            print(f"Force terminated process: {proc.info['name']}")
                            closed_count += 1
                except Exception as force_e:
                    print(f"Force close failed for {title}: {force_e}")
        
        if closed_count > 0:
            speak(f"Closed {app_name}")
        else:
            speak(f"{app_name} not found or failed to close")
    except Exception as e:
        print(f"Close {app_name} error: {e}")
        speak(f"Could not close {app_name}")

def close_browser():
    """Close any browser application"""
    try:
        browsers = ['chrome', 'firefox', 'edge', 'opera', 'brave', 'safari', 'chromium']
        closed_count = 0
        found_windows = []
        
        def enum_windows_callback(hwnd, data):
            nonlocal closed_count
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd).lower()
                for browser in browsers:
                    if browser in window_text and len(window_text) > 3:
                        found_windows.append((hwnd, window_text, browser))
                        break
            return True
        
        win32gui.EnumWindows(enum_windows_callback, None)
        
        print(f"Found {len(found_windows)} browser windows to close:")
        for hwnd, title, browser_type in found_windows:
            print(f"  - {title} ({browser_type})")
        
        for hwnd, title, browser_type in found_windows:
            try:
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                closed_count += 1
                print(f"Sent close signal to: {title}")
            except Exception as e:
                print(f"Failed to close {title}: {e}")
        
        if closed_count > 0:
            speak(f"Closed {closed_count} browser windows")
        else:
            speak("No browser windows found or failed to close")
    except Exception as e:
        print(f"Close browser error: {e}")
        speak("Could not close browser")

def close_current_tab():
    """Close current browser tab"""
    try:
        # Simulate Ctrl+W to close current tab
        pyautogui.hotkey('ctrl', 'w')
        speak("Closed current tab")
    except Exception as e:
        print(f"Close tab error: {e}")
        speak("Could not close current tab")

def close_all_applications():
    """Close all running applications (use with caution)"""
    try:
        import psutil
        closed_count = 0
        
        # Get list of running processes
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                proc_name = proc.info['name']
                if proc_name and proc_name.lower() not in ['explorer.exe', 'svchost.exe', 'winlogon.exe', 'csrss.exe', 'wininit.exe']:
                    proc.terminate()
                    closed_count += 1
                    print(f"Closed: {proc_name}")
            except Exception as e:
                print(f"Failed to close {proc_name}: {e}")
        
        speak(f"Closed {closed_count} applications")
    except Exception as e:
        print(f"Close all applications error: {e}")
        speak("Could not close all applications")

def force_close_application(app_name):
    """Force close a specific application by terminating its process"""
    try:
        import psutil
        closed_count = 0
        
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                proc_name = proc.info['name']
                if proc_name and app_name.lower() in proc_name.lower():
                    proc.terminate()
                    closed_count += 1
                    print(f"Force closed: {proc_name}")
            except Exception as e:
                print(f"Failed to force close {proc_name}: {e}")
        
        if closed_count > 0:
            speak(f"Force closed {app_name}")
        else:
            speak(f"{app_name} not found or failed to force close")
    except Exception as e:
        print(f"Force close {app_name} error: {e}")
        speak(f"Could not force close {app_name}")

# ---------- Weather with API ----------
def get_weather(city):
    try:
        url = f"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={WEATHER_API_KEY}&units=metric"
        response = requests.get(url, timeout=5)
        data = response.json()
        
        if data["cod"] != 200:
            return "Weather service error: " + data.get("message", "Unknown error")
        
        temp = data["main"]["temp"]
        desc = data["weather"][0]["description"]
        feels_like = data["main"]["feels_like"]
        humidity = data["main"]["humidity"]
        
        return f"{city}: {temp}°C, {desc}\nFeels like: {feels_like}°C\nHumidity: {humidity}%"
    except Exception as e:
        return f"Weather service unavailable: {str(e)}"

# ---------- News with API ----------
def get_news():
    try:
        url = f"https://newsapi.org/v2/top-headlines?country=us&apiKey={NEWS_API_KEY}"
        response = requests.get(url, timeout=5)
        data = response.json()
        
        if data["status"] != "ok":
            return "News service error"
        
        headlines = [article["title"] for article in data["articles"][:3]]
        return ". ".join(headlines)
    except Exception as e:
        return f"News service unavailable: {str(e)}"

# ---------- Calculator ----------
def calculate(expr):
    try:
        allowed = {"builtins": None}
        result = eval(expr, allowed, {"math": math})
        return str(result)
    except:
        return "Cannot calculate"

# ---------- Alarm ----------
def set_alarm(time_str):
    global ALARM_THREAD
    try:
        alarm_time = datetime.datetime.strptime(time_str.strip(), "%I %p").time()
        def alarm_loop():
            while True:
                now = datetime.datetime.now().time()
                if now.hour == alarm_time.hour and now.minute == alarm_time.minute:
                    winsound.Beep(1000, 1000)
                    speak("Alarm! Wake up!")
                    break
                time.sleep(10)
        ALARM_THREAD = threading.Thread(target=alarm_loop, daemon=True)
        ALARM_THREAD.start()
        return f"Alarm set for {time_str.strip()}"
    except:
        return "Invalid time format. Say like '7 am'"

# ---------- Browser announcer ----------
def browser_announce_loop():
    global _last_browser_spoken
    browsers = ("chrome", "msedge", "firefox", "brave", "opera")
    while True:
        try:
            title = win32gui.GetWindowText(win32gui.GetForegroundWindow()).lower()
            for b in browsers:
                if b in title and title != _last_browser_spoken and LISTENING:
                    speak(b.capitalize() + " opened")
                    _last_browser_spoken = title
                    break
        except:
            pass
        time.sleep(1)

# ---------- Advanced Animation Functions ----------
class HUD3DAnimation:
    def __init__(self, canvas):
        self.canvas = canvas
        self.width = 400
        self.height = 500
        self.center_x = self.width // 2
        self.center_y = self.height // 2 - 50
        self.angle = 0
        self.scan_angle = 0
        self.pulse_radius = 0
        self.pulse_growing = True
        self.scan_line = None
        self.circles = []
        self.dots = []
        self.status_text = None
        self.command_text = None
        self.conversation_text = []
        self.conversation_bg = None
        self.init_hud()
        self.animate()
        self.update_status("DESKTOP BOT", "#00FF9D")
        
    def init_hud(self):
        # Clear any existing items
        self.canvas.delete('hud')
        
        # Create concentric circles
        colors = ['#00FF9D', '#00CC7F', '#009966']
        for i, radius in enumerate([120, 90, 60]):
            circle = self.canvas.create_oval(
                self.center_x - radius, self.center_y - radius,
                self.center_x + radius, self.center_y + radius,
                outline=colors[i], width=2, tags='hud'
            )
            self.circles.append(circle)
        
        # Create outer dots
        dot_count = 24
        for i in range(dot_count):
            angle = (2 * math.pi * i) / dot_count
            dot = self.canvas.create_oval(
                self.center_x - 3 + 130 * math.cos(angle),
                self.center_y - 3 + 130 * math.sin(angle),
                self.center_x + 3 + 130 * math.cos(angle),
                self.center_y + 3 + 130 * math.sin(angle),
                fill='#00FF9D', tags='hud'
            )
            self.dots.append(dot)
        
        # Add JARVIS text in the center
        self.canvas.create_text(
            self.center_x, self.center_y - 20,
            text="J.A.R.V.I.S.",
            font=('Arial', 16, 'bold'),
            fill='#00FF9D',
            tags='hud'
        )
        
        # Add conversation area background
        self.conversation_bg = self.canvas.create_rectangle(
            10, self.center_y + 100, self.width - 10, self.height - 10,
            fill='#0A0A0A', outline='#1E1E1E', width=1, tags='hud'
        )
    
    def update_status(self, text, color="#00FF9D"):
        if hasattr(self, 'status_text'):
            self.canvas.delete(self.status_text)
        self.status_text = self.canvas.create_text(
            self.center_x, self.center_y + 80,
            text=text,
            font=('Arial', 10, 'bold'),
            fill=color,
            tags='hud'
        )
    
    def add_to_conversation(self, speaker, text, color="#00FF9D"):
        if not hasattr(self, 'conversation_text'):
            self.conversation_text = []
            
        # Clear previous messages if any
        for item in self.conversation_text:
            self.canvas.delete(item)
        
        # Create new message
        msg = self.canvas.create_text(
            20, self.center_y + 120,
            text=f"{speaker}: {text}",
            font=('Arial', 8),
            fill=color,
            anchor='nw',
            tags='hud',
            width=360
        )
        
        # Store only the current message
        self.conversation_text = [msg]
        
        # Removed automatic clearing - messages will persist until manually cleared
    
    def clear_conversation(self):
        """Clear all conversation messages"""
        if hasattr(self, 'conversation_text'):
            for item in self.conversation_text:
                self.canvas.delete(item)
            self.conversation_text = []
    
    def animate(self):
        # Rotate outer dots
        self.angle = (self.angle + 0.5) % 360
        rad_angle = math.radians(self.angle)
        
        # Update dot positions
        for i, dot in enumerate(self.dots):
            angle = (2 * math.pi * i / len(self.dots)) + rad_angle
            self.canvas.coords(
                dot,
                self.center_x - 3 + 130 * math.cos(angle),
                self.center_y - 3 + 130 * math.sin(angle),
                self.center_x + 3 + 130 * math.cos(angle),
                self.center_y + 3 + 130 * math.sin(angle)
            )
        
        # Pulsing effect for center circle
        if self.pulse_growing:
            self.pulse_radius += 0.5
            if self.pulse_radius > 20:
                self.pulse_growing = False
        else:
            self.pulse_radius -= 0.5
            if self.pulse_radius < 0:
                self.pulse_growing = True
        
        # Update scan line
        self.scan_angle = (self.scan_angle + 2) % 360
        scan_rad = math.radians(self.scan_angle)
        
        if hasattr(self, 'scan_line'):
            self.canvas.delete(self.scan_line)
        
        self.scan_line = self.canvas.create_line(
            self.center_x,
            self.center_y,
            self.center_x + 120 * math.cos(scan_rad),
            self.center_y + 120 * math.sin(scan_rad),
            fill='#00FF9D',
            width=1,
            tags='hud'
        )
        
        # Schedule next frame
        self.canvas.after(30, self.animate)

# ---------- Web Browsing Functions ----------
def open_website(url):
    """Open a website in the default browser"""
    try:
        # Clean up the URL
        url = url.strip().lower()
        if not url:
            return "Please specify a website to open"
            
        # Common domain mappings
        domain_mapping = {
            'youtube': 'youtube.com',
            'facebook': 'facebook.com',
            'twitter': 'twitter.com',
            'instagram': 'instagram.com',
            'linkedin': 'linkedin.com',
            'github': 'github.com',
            'google': 'google.com',
            'chrome': 'google.com/chrome'
        }
        
        # Check if it's an exact mapped domain or a simple variant (avoid replacing multi-word phrases like "google drive")
        for domain, mapped_url in domain_mapping.items():
            # Accept exact matches like "google", "google.com", or "www.google"
            if url == domain or url == f"{domain}.com" or url == f"www.{domain}":
                url = mapped_url
                break
                
        # Remove common prefixes and protocols
        url = url.replace('https://', '').replace('http://', '').replace('www.', '')
        
        # Add .com if no domain extension is present
        if '.' not in url and url not in domain_mapping.values():
            url += '.com'
            
        # Construct the full URL
        full_url = f'https://{url}'
        
        # Open the website
        webbrowser.open(full_url)
        
        return f"Opening {url}"
        
    except Exception as e:
        error_msg = f"Failed to open website: {str(e)}"
        print(error_msg)
        return error_msg

# ---------- Voice Command Functions ----------
def listen_for_commands(hud):
    """Listen for voice commands and process them"""
    recognizer = sr.Recognizer()
    
    while True:
        if not LISTENING:
            time.sleep(0.5)
            continue
            
        try:
            with sr.Microphone() as source:
                recognizer.adjust_for_ambient_noise(source)
                hud.update_status("Listening...", "#55FF55")
                audio = recognizer.listen(source, timeout=5, phrase_time_limit=5)
                
            try:
                query = recognizer.recognize_google(audio)
                hud.add_to_conversation("You", query, "#4FC3F7")
                process_command(query, hud)
                
            except sr.UnknownValueError:
                hud.update_status("Could not understand audio", "#FF5555")
                hud.add_to_conversation("DESKTOP BOT", "I didn't catch that. Could you please repeat?", "#FF5555")
                
            except sr.RequestError as e:
                error_msg = f"Speech service error: {e}"
                hud.update_status("Speech service error", "#FF5555")
                hud.add_to_conversation("DESKTOP BOT", error_msg, "#FF5555")
                print(error_msg)
                
        except sr.WaitTimeoutError:
            continue
            
        except Exception as e:
            error_msg = f"Error in voice recognition: {e}"
            print(error_msg)
            hud.add_to_conversation("System", error_msg, "#FF5555")
            time.sleep(1)

# ---------- Main ----------
def process_command(query, hud):
    """Process voice commands with AI assistance"""
    global LISTENING
    
    if not query or not query.strip():
        return
        
    query = query.strip()
    print(f"Processing command: {query}")  # Debug log
    
    # Show user's command in the HUD
    hud.add_to_conversation("You", query, "#4FC3F7")
    
    # Check for search engine commands
    search_engines = {
        'youtube': 'https://www.youtube.com/results?search_query=',
        'google': 'https://www.google.com/search?q=',
        'chatgpt': 'https://chat.openai.com/?q=',
        'perplexity': 'https://www.perplexity.ai/search/'
    }
    
    # Check if the query starts with a search engine name
    query_lower = query.lower()
    for engine, base_url in search_engines.items():
        if query_lower.startswith(engine + ' '):
            search_term = query[len(engine):].strip()
            if search_term:
                search_url = base_url + search_term.replace(' ', '+')
                webbrowser.open(search_url)
                    
                response = f"Searching {engine.capitalize()} for: {search_term}"
                hud.add_to_conversation("JARVIS", response, "#00FF9D")
                speak(response)
                return
    
    # Check for music playback commands
    if any(cmd in query_lower for cmd in ['play ', 'play the song ', 'play music', 'play song']):
        # Extract song name from command
        song_name = query_lower
        for prefix in ['play ', 'play the song ', 'play music ', 'play song ']:
            if song_name.startswith(prefix):
                song_name = song_name[len(prefix):].strip()
                break
                
        if song_name:
            # Search and play on YouTube
            search_query = f"https://www.youtube.com/results?search_query={song_name.replace(' ', '+')}"
            webbrowser.open(search_query)
                
            response = f"Playing {song_name} on YouTube"
            hud.add_to_conversation("JARVIS", response, "#00FF9D")
            speak(response)
            return
    
    # Check for system settings commands
    system_commands = {
        # Windows Settings
        'open windows settings': open_windows_settings,
        'windows settings': open_windows_settings,
        'settings': open_windows_settings,
        
        # Sound Settings
        'open sound settings': open_sound_settings,
        'sound settings': open_sound_settings,
        'audio settings': open_sound_settings,
        'volume settings': open_sound_settings,
        
        # Network Settings
        'open network settings': open_network_settings,
        'network settings': open_network_settings,
        'internet settings': open_network_settings,
        
        # WiFi Settings
        'open wifi settings': open_wifi_settings,
        'wifi settings': open_wifi_settings,
        'wireless settings': open_wifi_settings,
        
        # Bluetooth Settings
        'open bluetooth settings': open_bluetooth_settings,
        'bluetooth settings': open_bluetooth_settings,
        'bluetooth': open_bluetooth_settings,
        
        # Notification Settings
        'open notification settings': open_notification_settings,
        'notification settings': open_notification_settings,
        'notifications': open_notification_settings,
        
        # Storage Settings
        'open storage settings': open_storage_settings,
        'storage settings': open_storage_settings,
        'disk settings': open_storage_settings,
        
        # Battery Settings
        'open battery settings': open_battery_settings,
        'battery settings': open_battery_settings,
        'power settings': open_battery_settings,
        
        # Privacy Settings
        'open privacy settings': open_privacy_settings,
        'privacy settings': open_privacy_settings,
        'privacy': open_privacy_settings,
        
        # Update Settings
        'open update settings': open_update_settings,
        'update settings': open_update_settings,
        'windows update': open_update_settings,
        
        # Default Apps
        'open default apps': open_default_apps,
        'default apps': open_default_apps,
        'default applications': open_default_apps,
        
        # Control Panel
        'open control panel': open_control_panel,
        'control panel': open_control_panel,
        'control': open_control_panel,
        
        # Task Manager
        'open task manager': open_task_manager,
        'task manager': open_task_manager,
        'tasks': open_task_manager,
        
        # System Properties
        'open system properties': open_system_properties,
        'system properties': open_system_properties,
        'system info': open_system_properties,
        
        # Device Manager
        'open device manager': open_device_manager,
        'device manager': open_device_manager,
        'devices': open_device_manager,
        
        # Disk Cleanup
        'open disk cleanup': open_disk_cleanup,
        'disk cleanup': open_disk_cleanup,
        'clean disk': open_disk_cleanup,
        
        # System Restore
        'open system restore': open_system_restore,
        'system restore': open_system_restore,
        'restore system': open_system_restore,
        
        # Registry Editor
        'open registry editor': open_registry_editor,
        'registry editor': open_registry_editor,
        'registry': open_registry_editor,
        
        # Event Viewer
        'open event viewer': open_event_viewer,
        'event viewer': open_event_viewer,
        'events': open_event_viewer,
        
        # Volume Control
        'volume up': volume_up,
        'increase volume': volume_up,
        'turn up volume': volume_up,
        'volume down': volume_down,
        'decrease volume': volume_down,
        'turn down volume': volume_down,
        
        # Brightness Control
        'brightness up': brightness_up,
        'increase brightness': brightness_up,
        'turn up brightness': brightness_up,
        'brightness down': brightness_down,
        'decrease brightness': brightness_down,
        'turn down brightness': brightness_down,
        
        # Screenshot
        'take screenshot': screenshot,
        'screenshot': screenshot,
        'capture screen': screenshot,
        
        # PC Control
        'lock pc': lock_pc,
        'lock computer': lock_pc,
        'lock': lock_pc,
        
        # Virtual Desktop
        'new desktop': new_desktop,
        'create desktop': new_desktop,
        'add desktop': new_desktop,
        'close desktop': close_desktop,
        'remove desktop': close_desktop,
        'switch desktop': switch_desktop,
        'change desktop': switch_desktop,
        
        # System Information
        'system info': system_info,
        'system information': system_info,
        'computer info': system_info,
        
        # Window Control
        'minimize window': minimize_window,
        'minimize': minimize_window,
        'hide window': minimize_window,
        'minimize google': minimize_google,
        'minimize chrome': minimize_google,
        'minimize browser': minimize_browser,
        'minimize browsers': minimize_browser,
        'minimize application': minimize_application,
        'minimize app': minimize_application,
        'minimize website': minimize_window,
        'minimize websites': minimize_window,
        
        # Close Commands
        'close window': lambda: close_application('window'),
        'close google': lambda: close_application('chrome'),
        'close chrome': lambda: close_application('chrome'),
        'close browser': close_browser,
        'close browsers': close_browser,
        'close application': lambda: close_application('application'),
        'close app': lambda: close_application('app'),
        'close website': close_browser,
        'close websites': close_browser,
        'close tab': close_current_tab,
        'close current tab': close_current_tab,
        'close all applications': close_all_applications,
        'close everything': close_all_applications,
        'force close': lambda: force_close_application('application'),
        'kill application': lambda: force_close_application('application'),
        'terminate application': lambda: force_close_application('application'),
        
        # File Explorer
        'open file explorer': lambda: file_manager.open_explorer(),
        'file explorer': lambda: file_manager.open_explorer(),
        'files': lambda: file_manager.open_explorer(),
        'explorer': lambda: file_manager.open_explorer(),
        'file explorer': lambda: file_manager.open_explorer(),
        'open files': lambda: file_manager.open_explorer(),
        'show files': lambda: file_manager.open_explorer(),
        'browse files': lambda: file_manager.open_explorer(),
        'file browser': lambda: file_manager.open_explorer(),
        'open desktop': lambda: file_manager.open_desktop(),
        'open documents': lambda: file_manager.open_documents(),
        'open downloads': lambda: file_manager.open_downloads(),
        'open pictures': lambda: file_manager.open_pictures(),
        'open music': lambda: file_manager.open_music(),
        'open videos': lambda: file_manager.open_videos()
    }
    
    # Check for system commands
    for command, function in system_commands.items():
        if command in query_lower:
            try:
                function()
                response = f"Executed: {command}"
                hud.add_to_conversation("DESKTOP BOT", response, "#00FF9D")
                speak(response)
                return
            except Exception as e:
                error_msg = f"Error executing {command}: {str(e)}"
                hud.add_to_conversation("DESKTOP BOT", error_msg, "#FF5555")
                speak(f"Error executing {command}")
                return
    
    # Check for dynamic minimize commands (e.g., "minimize notepad", "minimize word")
    if 'minimize' in query_lower:
        # Extract application name after "minimize"
        parts = query_lower.split('minimize')
        if len(parts) > 1:
            app_name = parts[1].strip()
            if app_name and app_name not in ['window', 'google', 'chrome', 'browser', 'browsers', 'application', 'app', 'website', 'websites']:
                try:
                    minimize_application(app_name)
                    response = f"Minimized {app_name}"
                    hud.add_to_conversation("DESKTOP BOT", response, "#00FF9D")
                    speak(response)
                    return
                except Exception as e:
                    error_msg = f"Error minimizing {app_name}: {str(e)}"
                    hud.add_to_conversation("DESKTOP BOT", error_msg, "#FF5555")
                    speak(f"Error minimizing {app_name}")
                    return
    
    # Check for dynamic close commands (e.g., "close notepad", "close word", "close chrome")
    if 'close' in query_lower:
        # Extract application name after "close"
        parts = query_lower.split('close')
        if len(parts) > 1:
            app_name = parts[1].strip()
            if app_name and app_name not in ['window', 'google', 'chrome', 'browser', 'browsers', 'application', 'app', 'website', 'websites', 'tab', 'current tab', 'all applications', 'everything']:
                try:
                    close_application(app_name)
                    response = f"Closed {app_name}"
                    hud.add_to_conversation("DESKTOP BOT", response, "#00FF9D")
                    speak(response)
                    return
                except Exception as e:
                    error_msg = f"Error closing {app_name}: {str(e)}"
                    hud.add_to_conversation("DESKTOP BOT", error_msg, "#FF5555")
                    speak(f"Error closing {app_name}")
                    return
    
    # Check for force close commands (e.g., "force close notepad", "kill word", "terminate chrome")
    if any(phrase in query_lower for phrase in ['force close', 'kill', 'terminate']):
        # Extract application name after force close keywords
        app_name = None
        for phrase in ['force close', 'kill', 'terminate']:
            if phrase in query_lower:
                parts = query_lower.split(phrase)
                if len(parts) > 1:
                    app_name = parts[1].strip()
                    break
        
        if app_name and app_name not in ['window', 'google', 'chrome', 'browser', 'browsers', 'application', 'app', 'website', 'websites', 'tab', 'current tab', 'all applications', 'everything']:
            try:
                force_close_application(app_name)
                response = f"Force closed {app_name}"
                hud.add_to_conversation("DESKTOP BOT", response, "#00FF9D")
                speak(response)
                return
            except Exception as e:
                error_msg = f"Error force closing {app_name}: {str(e)}"
                hud.add_to_conversation("DESKTOP BOT", error_msg, "#FF5555")
                speak(f"Error force closing {app_name}")
                return
    
    # Check for file explorer commands
    file_commands = [
        'open file explorer', 'file explorer', 'files', 'explorer', 'open files', 
        'show files', 'browse files', 'file browser', 'open desktop', 'open documents',
        'open downloads', 'open pictures', 'open music', 'open videos', 'open drive',
        'create folder', 'new folder', 'delete folder', 'remove folder', 'copy folder',
        'move folder', 'rename folder', 'folder properties', 'show properties', 'open path',
        'create folder os', 'make folder os', 'delete folder os', 'remove folder os',
        'list folders', 'show folders', 'check folder', 'folder exists', 'folder size', 'get size',
        'test file operations', 'debug file operations',
        'show 3d circle', 'display 3d circle', 'hide 3d circle', 'remove 3d circle',
        'faster circle', 'speed up circle', 'slower circle', 'slow down circle'
    ]
    
    if any(cmd in query_lower for cmd in file_commands):
        try:
            print(f"Processing file command: {query}")  # Debug log
            result = file_manager.handle_voice_command(query)
            if result:
                hud.add_to_conversation("JARVIS", result, "#00FF9D")
                speak(result)
                return
            else:
                # If no result returned, it might be an unrecognized command
                hud.add_to_conversation("DESKTOP BOT", "File operation completed but no result returned", "#FFAA00")
                speak("File operation completed")
                return
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"File operation error details: {error_details}")  # Debug log
            
            error_msg = f"File operation error: {str(e)}"
            hud.add_to_conversation("DESKTOP BOT", error_msg, "#FF5555")
            speak("File operation error occurred")
            return
    
    # Check for weather commands
    if any(cmd in query_lower for cmd in ['weather', 'temperature', 'forecast']):
        # Extract city name from command
        city = query_lower
        for prefix in ['weather in ', 'temperature in ', 'forecast in ', 'weather for ', 'temperature for ', 'forecast for ']:
            if city.startswith(prefix):
                city = city[len(prefix):].strip()
                break
        if city:
            try:
                weather_info = get_weather(city)
                hud.add_to_conversation("DESKTOP BOT", weather_info, "#00FF9D")
                speak(weather_info)
                return
            except Exception as e:
                error_msg = f"Error getting weather for {city}: {str(e)}"
                hud.add_to_conversation("DESKTOP BOT", error_msg, "#FF5555")
                speak(f"Error getting weather for {city}")
                return
    
    # Check for news commands
    if any(cmd in query_lower for cmd in ['news', 'headlines', 'latest news']):
        try:
            news_info = get_news()
            hud.add_to_conversation("DESKTOP BOT", news_info, "#00FF9D")
            speak(news_info)
            return
        except Exception as e:
            error_msg = f"Error getting news: {str(e)}"
            hud.add_to_conversation("DESKTOP BOT", error_msg, "#FF5555")
            speak("Error getting news")
            return
    
    # Check for exit command - single meaning: close Jarvis only
    if any(cmd in query_lower for cmd in ['exit', 'quit', 'close', 'stop', 'shutdown']):
        response = "Shutting down Desktop Bot. Goodbye!"
        hud.add_to_conversation("Desktop Bot", response, "#00FF9D")
        speak(response)
        # Small delay to let the message be spoken
        root.after(1000, close_app)
        return
        
    # Alternative exit commands that might be used
    if any(phrase in query_lower for phrase in ['exit window', 'close window', 'exit the window', 'close the window', 'exit application', 'close application', 'exit the application', 'close the application']):
        response = "Closing the application. Goodbye!"
        hud.add_to_conversation("DESKTOP BOT", response, "#00FF9D")
        speak(response)
        # Small delay to let the message be spoken
        root.after(1000, close_app)
        return
    
    # Check for calculation commands
    if any(cmd in query_lower for cmd in ['calculate', 'math', 'compute']):
        # Extract mathematical expression
        expr = query_lower
        for prefix in ['calculate ', 'math ', 'compute ']:
            if expr.startswith(prefix):
                expr = expr[len(prefix):].strip()
                break
        if expr:
            try:
                result = calculate(expr)
                hud.add_to_conversation("DESKTOP BOT", f"Result: {result}", "#00FF9D")
                speak(f"The result is {result}")
                return
            except Exception as e:
                error_msg = f"Error calculating {expr}: {str(e)}"
                hud.add_to_conversation("DESKTOP BOT", error_msg, "#FF5555")
                speak(f"Error calculating {expr}")
                return
    
    # Check for web browsing commands
    if any(cmd in query_lower for cmd in ['open ', 'go to ', 'navigate to ']):
        # Extract URL from command
        url = query.lower()
        for prefix in ['open ', 'go to ', 'navigate to ']:
            if url.startswith(prefix):
                url = url[len(prefix):].strip()
                break
                
        # Process the URL
        result = open_website(url)
        hud.add_to_conversation("DESKTOP BOT", result, "#00FF9D")
        speak(result)
        return
        
    # Get AI response for the query
    try:
        hud.update_status("Thinking...", "#FFA500")
        response = get_ai_response(query)
        
        if response:
            hud.add_to_conversation("DESKTOP BOT", response, "#00FF9D")
            speak(response)
    except Exception as e:
        error_msg = f"Error processing your request: {str(e)}"
        print(error_msg)
        hud.add_to_conversation("DESKTOP BOT", "I encountered an error processing your request", "#FF5555")
    finally:
        hud.update_status("Ready", "#00FF9D")

# Import the background manager
from background_manager import BackgroundManager

# Main window setup with enhanced 4K background
root = tk.Tk()
root.title(" Desktop bot AI Assistant")
root.overrideredirect(True)
root.attributes('-transparentcolor', 'black')
root.attributes('-topmost', True)
root.attributes('-alpha', 0.95)

# Make window draggable
def move_window(event):
    x = root.winfo_pointerx() - offset_x
    y = root.winfo_pointery() - offset_y
    root.geometry(f'+{x}+{y}')

def on_click(event):
    global offset_x, offset_y
    offset_x = event.x
    offset_y = event.y

# Center the window on screen
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = 400
window_height = 600
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2
root.geometry(f'{window_width}x{window_height}+{x}+{y}')

# Create canvas with transparent background for the 4K image
canvas = tk.Canvas(root, width=window_width, height=window_height, 
                  bg='black', highlightthickness=0, bd=0)
canvas.pack()

# Initialize and load the high-quality background
background_manager = BackgroundManager()
background_manager.load_background_image(canvas, window_width, window_height)

# Create right frame for additional controls
right_frame = tk.Frame(root, bg='black')
right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10)

# Make the window draggable
canvas.bind('<Button-1>', on_click)
canvas.bind('<B1-Motion>', move_window)

# Close on right-click
def close_app(event=None):
    print("Shutting down Bot...")
    root.quit()
canvas.bind('<Button-3>', close_app)
root.protocol("WM_DELETE_WINDOW", close_app)

# Handle window resizing to update background
def on_resize(event):
    global window_width, window_height
    # Only update if size actually changed
    if event.width != window_width or event.height != window_height:
        window_width, window_height = event.width, event.height
        # Update canvas size
        canvas.config(width=window_width, height=window_height)
        # Update background image
        if 'background_manager' in globals():
            background_manager.update_background(canvas, window_width, window_height)
        # Refresh HUD if it exists
        if 'hud' in globals():
            hud.width = window_width
            hud.height = window_height
            hud.center_x = window_width // 2
            hud.center_y = window_height // 2 - 50
            hud.init_3d_hud()

# Bind resize event
root.bind("<Configure>", on_resize)

# Start the enhanced 3D HUD and voice recognition
hud = Advanced3DAnimation(canvas)

# Initial greeting
initial_greeting = "Desktop Bot online with enhanced 3D interface."
hud.add_to_conversation("Desktop Bot", initial_greeting, "#00FF9D")
speak(initial_greeting)

# Start voice recognition in a separate thread
voice_thread = threading.Thread(target=listen_for_commands, args=(hud,), daemon=True)
voice_thread.start()

# Add 3D Interface Launch Button
def launch_advanced_3d_interface():
    """Launch the advanced 3D interface"""
    try:
        from launch_3d_interface import launch_3d_interface
        # Update the animation status
        if 'animation' in globals() and animation:
            animation.update_status("Launching 3D Interface", "#00FFFF")
            animation.add_to_conversation("Desktop Bot", "Initializing advanced 3D interface...")
        
        # Launch the 3D interface
        thread = launch_3d_interface(root)
        
        # Update status after launch
        if 'animation' in globals() and animation:
            animation.update_status("3D Interface Active", "#00FF9D")
    except Exception as e:
        print(f"Error launching 3D interface: {e}")
        if 'animation' in globals() and animation:
            animation.update_status("3D Interface Error", "#FF0000")
            animation.add_to_conversation("Desktop Bot", f"Error launching 3D interface: {str(e)}")

# Create enhanced 3D interface launch button with glowing effect
def create_3d_interface_button(parent):
    # Create a frame with semi-transparent background
    button_frame = tk.Frame(parent, bg='#1E1E1E')
    button_frame.pack(side=tk.BOTTOM, pady=10)
    
    # Create a glowing effect around the button
    def pulse_button():
        # Alternate between two colors for pulsing effect
        current_color = launch_button.cget("bg")
        if current_color == "#2C2C2C":
            launch_button.config(bg="#3C3C3C", fg="#00FFFF")
        else:
            launch_button.config(bg="#2C2C2C", fg="#00FF9D")
        # Schedule the next pulse
        button_frame.after(1000, pulse_button)
    
    # Create the button with enhanced visual style
    launch_button = tk.Button(
        button_frame,
        text="Launch Advanced 4K 3D Interface",
        font=("Arial", 10, "bold"),
        bg="#2C2C2C",
        fg="#00FF9D",
        activebackground="#3C3C3C",
        activeforeground="#00FFFF",
        relief=tk.RAISED,
        bd=2,
        padx=8,
        pady=5,
        command=launch_advanced_3d_interface
    )
    launch_button.pack(padx=5, pady=5)
    
    # Start the pulsing effect
    pulse_button()
    
    return launch_button

# Add the 3D interface button to the UI
interface_3d_button = create_3d_interface_button(right_frame)

# Auto-hide functionality
def check_window_focus():
    """Check if JARVIS window has focus and hide/show accordingly"""
    global AUTO_HIDE
    if not AUTO_HIDE:
        # If auto-hide is disabled, schedule next check and return
        root.after(1000, check_window_focus)
        return

    try:
        # Get the currently active window
        active_window = win32gui.GetForegroundWindow()
        jarvis_window = root.winfo_id()

        # If active window is invalid or is Jarvis itself, keep visible
        if not active_window or active_window == jarvis_window:
            root.deiconify()
            root.attributes('-topmost', True)
            root.attributes('-topmost', False)
            root.after(1000, check_window_focus)
            return

        # If the active window is minimized (iconic), don't hide Jarvis
        try:
            if win32gui.IsIconic(active_window):
                root.deiconify()
                root.attributes('-topmost', True)
                root.attributes('-topmost', False)
                root.after(1000, check_window_focus)
                return
        except Exception:
            pass

        # If the active window is the desktop or taskbar, don't hide Jarvis
        try:
            cls_name = win32gui.GetClassName(active_window)
            if cls_name in ('Progman', 'WorkerW', 'Shell_TrayWnd'):
                root.deiconify()
                root.attributes('-topmost', True)
                root.attributes('-topmost', False)
                root.after(1000, check_window_focus)
                return
        except Exception:
            pass

        # Otherwise another normal window has focus - hide Jarvis
        root.withdraw()

    except Exception:
        # On any error, just continue and reschedule
        pass

    # Schedule the next check
    root.after(1000, check_window_focus)

# Toggle auto-hide feature
def toggle_auto_hide():
    """Toggle the auto-hide feature"""
    global AUTO_HIDE
    AUTO_HIDE = not AUTO_HIDE
    
    if AUTO_HIDE:
        hud.update_status("Auto-hide enabled", "#00FF9D")
        hud.add_to_conversation("Desktop Bot", "Auto-hide mode enabled. I'll hide when other applications are in focus.")
        speak("Auto-hide mode enabled")
    else:
        hud.update_status("Auto-hide disabled", "#FFAA00")
        hud.add_to_conversation("Desktop Bot", "Auto-hide mode disabled. I'll stay visible at all times.")
        speak("Auto-hide mode disabled")
        root.deiconify()  # Make sure window is visible

# Initialize file manager
file_manager = SystemFileExplorer()

# Add file explorer button
file_explorer_button = tk.Button(
    right_frame,
    text="📁 File Explorer",
    font=("Arial", 10, "bold"),
    bg="#2C2C2C",
    fg="#00FF9D",
    activebackground="#3C3C3C",
    activeforeground="#00FFFF",
    relief=tk.RAISED,
    bd=2,
    padx=8,
    pady=5,
    command=lambda: file_manager.open_explorer()
)
file_explorer_button.pack(side=tk.TOP, padx=5, pady=5)

# Add auto-hide toggle button
auto_hide_button = tk.Button(
    right_frame,
    text="Toggle Auto-Hide",
    font=("Arial", 10, "bold"),
    bg="#2C2C2C",
    fg="#00FF9D",
    activebackground="#3C3C3C",
    activeforeground="#00FFFF",
    relief=tk.RAISED,
    bd=2,
    padx=8,
    pady=5,
    command=toggle_auto_hide
)
auto_hide_button.pack(side=tk.TOP, padx=5, pady=10)

# Start the auto-hide check
root.after(1000, check_window_focus)

# ---------------- Chat history window (persistent, scrollable) -----------------
chat_history_window = None
chat_history_text = None

def open_chat_history():
    """Open a persistent chat history window showing full conversation history."""
    global chat_history_window, chat_history_text, hud
    try:
        if chat_history_window and tk.Toplevel.winfo_exists(chat_history_window):
            chat_history_window.lift()
            return

        chat_history_window = tk.Toplevel(root)
        chat_history_window.title("Chat History")
        chat_history_window.geometry("480x600")
        chat_history_window.configure(bg='#111111')

        # Scrolled text widget for history
        chat_history_text = scrolledtext.ScrolledText(chat_history_window, wrap=tk.WORD, bg='#0F0F0F', fg='#EAEAEA')
        chat_history_text.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        chat_history_text.config(state=tk.NORMAL)

        # Buttons frame
        btn_frame = tk.Frame(chat_history_window, bg='#111111')
        btn_frame.pack(fill=tk.X, padx=8, pady=(0,8))

        def save_history():
            try:
                if not hasattr(hud, 'conversation_history') or not hud.conversation_history:
                    messagebox.showinfo("Save History", "No chat history available to save.")
                    return
                file_path = filedialog.asksaveasfilename(defaultextension='.json', filetypes=[('JSON', '*.json'), ('Text', '*.txt')])
                if not file_path:
                    return
                # Save as JSON for structure
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(hud.conversation_history, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("Save History", f"Chat history saved to {file_path}")
            except Exception as e:
                messagebox.showerror("Save Error", str(e))

        def clear_history():
            try:
                if messagebox.askyesno("Clear History", "Clear the in-memory chat history? This cannot be undone."):
                    if hasattr(hud, 'conversation_history'):
                        hud.conversation_history = []
                    refresh_chat_history()
            except Exception as e:
                messagebox.showerror("Error", str(e))

        save_btn = tk.Button(btn_frame, text="Save", command=save_history, bg='#2C2C2C', fg='#00FF9D')
        save_btn.pack(side=tk.LEFT, padx=(0,8))
        clear_btn = tk.Button(btn_frame, text="Clear", command=clear_history, bg='#2C2C2C', fg='#FF5555')
        clear_btn.pack(side=tk.LEFT)

        # Populate initial contents
        refresh_chat_history()

        # Make sure the window closes cleanly
        def on_close():
            global chat_history_window, chat_history_text
            try:
                chat_history_window.destroy()
            except:
                pass
            chat_history_window = None
            chat_history_text = None

        chat_history_window.protocol("WM_DELETE_WINDOW", on_close)
    except Exception as e:
        print(f"Error opening chat history window: {e}")

def refresh_chat_history():
    """Refresh the chat history window contents from hud.conversation_history."""
    global chat_history_window, chat_history_text, hud
    try:
        if not chat_history_window or not chat_history_text:
            return
        chat_history_text.config(state=tk.NORMAL)
        chat_history_text.delete('1.0', tk.END)
        if hasattr(hud, 'conversation_history') and hud.conversation_history:
            for msg in hud.conversation_history:
                ts = datetime.datetime.fromtimestamp(msg.get('timestamp', time.time())).strftime('%Y-%m-%d %H:%M:%S')
                speaker = msg.get('speaker', 'Unknown')
                text = msg.get('text', '')
                chat_history_text.insert(tk.END, f"[{ts}] {speaker}: {text}\n\n")
        else:
            chat_history_text.insert(tk.END, "No chat history yet.")
        chat_history_text.see(tk.END)
        chat_history_text.config(state=tk.DISABLED)
    except Exception as e:
        print(f"Error refreshing chat history: {e}")

# Add Chat History button to the right frame
chat_history_button = tk.Button(
    right_frame,
    text="💬 Chat History",
    font=("Arial", 10, "bold"),
    bg="#2C2C2C",
    fg="#00FF9D",
    activebackground="#3C3C3C",
    activeforeground="#00FFFF",
    relief=tk.RAISED,
    bd=2,
    padx=8,
    pady=5,
    command=open_chat_history
)
chat_history_button.pack(side=tk.TOP, padx=5, pady=5)

# Start the main loop
if __name__ == "__main__":
    root.mainloop()
