#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Speaking Test Application
A program to display images and record mouse click responses for student assessments.

Double-click this file to run the application (no batch file needed).
"""

import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
import os
import re
from datetime import datetime
from pathlib import Path
import configparser
import shutil


class SpeakingTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Speaking Test Application")
        self.root.geometry("800x600")
        
        # Set up window close handler
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Get the script's directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Data storage
        self.class_number = ""
        self.class_folder = ""
        self.student_id = ""
        self.current_image_index = 1
        self.image_files = []
        self.slides_folder = os.path.join(script_dir, "Slides")
        self.selected_slides_subfolder = None  # Track which subfolder is selected
        self.records_folder = os.path.join(script_dir, "Records")
        self.config_file = os.path.join(script_dir, "mouse_config.ini")
        
        # Mouse button mapping (will be loaded from config or calibrated)
        # Format: {tkinter_button_event: position_number}
        self.mouse_button_map = {}
        
        # Calibration state
        self.calibration_step = 0
        self.calibration_mapping = {}
        
        # State management
        self.current_screen = "class"
        
        # Ensure required folders exist
        self.ensure_required_folders()
        
        # Check for and handle duplicates BEFORE starting
        self.check_and_handle_duplicates()
        
        # Load or create mouse button configuration
        if self.load_mouse_config():
            # Config loaded successfully, start with class entry
            self.bind_escape()
            self.show_class_entry()
        else:
            # Need to calibrate
            self.bind_escape()
            self.start_calibration()
    
    def check_and_handle_duplicates(self):
        """Check for duplicate test files and move them to Duplicates folder"""
        try:
            duplicates_moved = self.find_and_move_duplicates()
            
            # Only show message if duplicates were found
            if duplicates_moved > 0:
                messagebox.showinfo(
                    "Duplicates Found",
                    f"Found and moved {duplicates_moved} duplicate test file(s)\n"
                    f"to Records/Duplicates folder.\n\n"
                    f"Review them to check for:\n"
                    f"• Students who took test twice\n"
                    f"• Accidental duplicate student numbers"
                )
        except Exception as e:
            # Don't stop the app if duplicate checking fails
            print(f"Error checking duplicates: {e}")
    
    def find_and_move_duplicates(self):
        """Find duplicate test files and move older ones to Duplicates folder. Returns count moved."""
        from collections import defaultdict
        
        duplicates_folder = os.path.join(self.records_folder, "Duplicates")
        Path(duplicates_folder).mkdir(parents=True, exist_ok=True)
        
        total_moved = 0
        
        # Process each class folder
        for item in os.listdir(self.records_folder):
            item_path = os.path.join(self.records_folder, item)
            
            # Skip Duplicates folder itself and non-directories
            if item == "Duplicates" or not os.path.isdir(item_path):
                continue
            
            # Group files by (class, normalized_student_id)
            student_files = defaultdict(list)
            
            for filename in os.listdir(item_path):
                filepath = os.path.join(item_path, filename)
                
                # Skip non-files and processed files
                if not os.path.isfile(filepath) or '_PROCESSED.txt' in filename:
                    continue
                
                # Parse filename: SpeakingTest_CLASS_STUDENTID_TIMESTAMP.txt
                if filename.startswith('SpeakingTest_') and filename.endswith('.txt'):
                    parts = filename.replace('.txt', '').split('_')
                    if len(parts) >= 4:
                        class_number = parts[1]
                        student_id = parts[2]
                        timestamp = parts[3]
                        
                        # Normalize student ID (1 == 01)
                        normalized_id = str(int(student_id)) if student_id.isdigit() else student_id
                        
                        # Parse timestamp for sorting
                        try:
                            ts_parts = timestamp.split('.')
                            if len(ts_parts) == 4:
                                year, month, day, time = ts_parts
                                if len(year) == 2:
                                    year = '20' + year
                                hour = time[:2] if len(time) >= 2 else '00'
                                minute = time[2:4] if len(time) >= 4 else '00'
                                dt = datetime(int(year), int(month), int(day), int(hour), int(minute))
                            else:
                                dt = datetime(1900, 1, 1)
                        except:
                            dt = datetime(1900, 1, 1)
                        
                        key = (class_number, normalized_id)
                        student_files[key].append({
                            'filepath': filepath,
                            'filename': filename,
                            'datetime': dt
                        })
            
            # Find and move duplicates
            for key, files in student_files.items():
                if len(files) > 1:
                    # Sort by datetime (oldest first)
                    files.sort(key=lambda x: x['datetime'])
                    
                    # Move all but the newest
                    for file_info in files[:-1]:  # All except last (newest)
                        dest_path = os.path.join(duplicates_folder, file_info['filename'])
                        
                        # Handle name conflicts
                        if os.path.exists(dest_path):
                            base, ext = os.path.splitext(file_info['filename'])
                            counter = 1
                            while os.path.exists(dest_path):
                                dest_path = os.path.join(duplicates_folder, f"{base}_dup{counter}{ext}")
                                counter += 1
                        
                        # Move the file
                        try:
                            shutil.move(file_info['filepath'], dest_path)
                            total_moved += 1
                        except:
                            pass
        
        return total_moved
    
    def on_closing(self):
        """Handle window close event with confirmation"""
        if self.current_screen == "image":
            # During a test - confirm exit
            result = messagebox.askyesno(
                "Exit Speaking Test?",
                "A test is currently in progress.\n\n"
                "Are you sure you want to exit?\n\n"
                "Note: The current test will be saved as incomplete.",
                icon='warning'
            )
            if result:
                self.root.destroy()
        else:
            # Not during a test - just confirm
            result = messagebox.askyesno(
                "Exit Application?",
                "Are you sure you want to exit the Speaking Test Application?",
                icon='question'
            )
            if result:
                self.root.destroy()
    
    def ensure_required_folders(self):
        """Create required folders if they don't exist"""
        # Use Path.mkdir with parents=True and exist_ok=True (same as working RaffleSystem)
        Path(self.records_folder).mkdir(parents=True, exist_ok=True)
        Path(self.slides_folder).mkdir(parents=True, exist_ok=True)
    
    def bind_escape(self):
        """Bind ESC key globally"""
        self.root.bind("<Escape>", self.handle_escape)
    
    def load_mouse_config(self):
        """Load mouse button configuration from INI file. Returns True if successful."""
        if not os.path.exists(self.config_file):
            return False
        
        try:
            config = configparser.ConfigParser()
            config.read(self.config_file)
            
            # Check if the config has the required section and keys
            if 'MouseButtons' not in config:
                return False
            
            # Load all 5 button mappings
            required_positions = ['1', '2', '3', '4', '5']
            for pos in required_positions:
                if pos not in config['MouseButtons']:
                    return False
                
                # Parse the button event (e.g., "Button-1")
                button_event = config['MouseButtons'][pos]
                self.mouse_button_map[button_event] = int(pos)
            
            # Load point values if they exist
            self.point_values = {}
            if 'PointValues' in config:
                for pos in required_positions:
                    if pos in config['PointValues']:
                        try:
                            self.point_values[int(pos)] = int(config['PointValues'][pos])
                        except ValueError:
                            # Invalid point value, use default
                            self.point_values[int(pos)] = 0
                    else:
                        # No point value specified, use default
                        self.point_values[int(pos)] = 0
            else:
                # No PointValues section, use defaults
                for pos in required_positions:
                    self.point_values[int(pos)] = 0
            
            # Load point names if they exist
            self.point_names = {}
            if 'PointNames' in config:
                for pos in required_positions:
                    if pos in config['PointNames']:
                        self.point_names[int(pos)] = config['PointNames'][pos]
                    else:
                        self.point_names[int(pos)] = ""
            else:
                # No PointNames section, use empty strings
                for pos in required_positions:
                    self.point_names[int(pos)] = ""
            
            # Load timer settings
            if 'TimerSettings' in config:
                self.timer_seconds = config['TimerSettings'].getint('seconds', 0)
                self.timer_start_text_color = config['TimerSettings'].get('start_text_color', '#000000')
                self.timer_start_bg_color = config['TimerSettings'].get('start_bg_color', '#D3D3D3')
                self.timer_end_text_color = config['TimerSettings'].get('end_text_color', '#00FF00')
                self.timer_end_bg_color = config['TimerSettings'].get('end_bg_color', '#D3D3D3')
            else:
                # Default timer settings
                self.timer_seconds = 0
                self.timer_start_text_color = '#000000'  # Black
                self.timer_start_bg_color = '#D3D3D3'    # Light grey
                self.timer_end_text_color = '#00FF00'    # Green
                self.timer_end_bg_color = '#D3D3D3'      # Light grey
            
            # Load roster setting
            if 'GeneralSettings' in config:
                self.use_roster = config['GeneralSettings'].getboolean('use_roster', False)
                self.reverse_count = config['GeneralSettings'].getboolean('reverse_count', False)
            else:
                self.use_roster = False
                self.reverse_count = False
            
            return True
            
        except Exception as e:
            # If there's any error, backup the corrupted file and recalibrate
            if os.path.exists(self.config_file):
                backup_name = f"{self.config_file}.backup"
                shutil.copy(self.config_file, backup_name)
            return False
    
    def save_mouse_config(self):
        """Save mouse button configuration to INI file"""
        config = configparser.ConfigParser()
        config['MouseButtons'] = {}
        
        # Save the mapping (position: button_event)
        for button_event, position in self.mouse_button_map.items():
            config['MouseButtons'][str(position)] = button_event
        
        # Save point values
        config['PointValues'] = {}
        if hasattr(self, 'point_values'):
            for position, points in self.point_values.items():
                config['PointValues'][str(position)] = str(points)
        
        # Save point names
        config['PointNames'] = {}
        if hasattr(self, 'point_names'):
            for position, name in self.point_names.items():
                config['PointNames'][str(position)] = name
        
        # Save timer settings
        config['TimerSettings'] = {}
        if hasattr(self, 'timer_seconds'):
            config['TimerSettings']['seconds'] = str(self.timer_seconds)
            config['TimerSettings']['start_text_color'] = self.timer_start_text_color
            config['TimerSettings']['start_bg_color'] = self.timer_start_bg_color
            config['TimerSettings']['end_text_color'] = self.timer_end_text_color
            config['TimerSettings']['end_bg_color'] = self.timer_end_bg_color
        
        # Save general settings
        config['GeneralSettings'] = {}
        if hasattr(self, 'use_roster'):
            config['GeneralSettings']['use_roster'] = str(self.use_roster)
        if hasattr(self, 'reverse_count'):
            config['GeneralSettings']['reverse_count'] = str(self.reverse_count)
        
        with open(self.config_file, 'w') as f:
            config.write(f)
    
    def start_calibration(self):
        """Start the mouse button calibration process"""
        self.clear_screen()
        self.current_screen = "calibration"
        self.calibration_step = 1
        self.calibration_mapping = {}
        
        # Create frame
        frame = tk.Frame(self.root, bg="white")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title = tk.Label(frame, text="Mouse Button Calibration", 
                        font=("Arial", 28, "bold"), bg="white")
        title.pack(pady=40)
        
        # Instructions
        instructions = tk.Label(frame, 
                              text="We need to calibrate your mouse buttons.\n\n"
                                   "You will be asked to click each mouse button\n"
                                   "that you want to use for recording answers.\n\n"
                                   "This only needs to be done once.",
                              font=("Arial", 14), bg="white", justify="center")
        instructions.pack(pady=20)
        
        # Current step label
        self.calibration_label = tk.Label(frame, 
                                         text="Click POSITION 1 mouse button\n(e.g., Left Click for full credit)",
                                         font=("Arial", 18, "bold"), bg="white", fg="blue")
        self.calibration_label.pack(pady=40)
        
        # Bind all mouse buttons to calibration handler
        # tkinter only supports Button-1 through Button-5
        for i in range(1, 6):  # Only buttons 1-5
            frame.bind(f"<Button-{i}>", self.handle_calibration_click)
        
        frame.focus_set()
    
    def handle_calibration_click(self, event):
        """Handle mouse clicks during calibration"""
        button_event = f"Button-{event.num}"
        
        # Check if this button was already assigned
        if button_event in self.calibration_mapping.values():
            messagebox.showwarning("Duplicate Button", 
                                 f"You already assigned this button to another position.\n"
                                 f"Please use a different mouse button.")
            return
        
        # Record the mapping
        self.calibration_mapping[self.calibration_step] = button_event
        
        # Move to next step
        self.calibration_step += 1
        
        if self.calibration_step <= 5:
            # Update label for next button
            position_descriptions = {
                2: "(e.g., Right Click for completely wrong)",
                3: "(e.g., Middle Click for mispronunciation)",
                4: "(custom category)",
                5: "(custom category)"
            }
            desc = position_descriptions.get(self.calibration_step, "")
            self.calibration_label.config(
                text=f"Click POSITION {self.calibration_step} mouse button\n{desc}"
            )
        else:
            # Calibration complete
            self.complete_calibration()
    
    def complete_calibration(self):
        """Complete the calibration and save the configuration"""
        # Transfer calibration mapping to mouse button map
        self.mouse_button_map = {}
        for position, button_event in self.calibration_mapping.items():
            self.mouse_button_map[button_event] = position
        
        # Now ask for point values
        self.show_point_value_configuration()
    
    def show_point_value_configuration(self):
        """Show screen to configure point values and names for each position"""
        self.clear_screen()
        self.current_screen = "point_config"
        
        # Create scrollable frame
        main_frame = tk.Frame(self.root, bg="white")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=40)
        
        # Title
        title = tk.Label(main_frame, text="Configure Point Values and Names", 
                        font=("Arial", 24, "bold"), bg="white")
        title.pack(pady=20)
        
        # Instructions
        instructions = tk.Label(main_frame, 
                              text="Assign point values and names to each mouse button.\n"
                                   "The button numbers match YOUR calibration.",
                              font=("Arial", 12), bg="white", justify="center")
        instructions.pack(pady=10)
        
        # Create entry fields for each position
        self.point_entries = {}
        self.name_entries = {}
        entries_frame = tk.Frame(main_frame, bg="white")
        entries_frame.pack(pady=20)
        
        for position in range(1, 6):
            row_frame = tk.Frame(entries_frame, bg="white")
            row_frame.pack(pady=8)
            
            # Label showing which mouse button this is
            label = tk.Label(row_frame, text=f"Mouse Button {position}:", 
                           font=("Arial", 14, "bold"), bg="white", width=18, anchor="e")
            label.pack(side=tk.LEFT, padx=5)
            
            # Point value entry - pre-fill with existing value
            point_entry = tk.Entry(row_frame, font=("Arial", 14), width=8)
            existing_points = self.point_values.get(position, 0) if hasattr(self, 'point_values') else 0
            point_entry.insert(0, str(existing_points))
            point_entry.pack(side=tk.LEFT, padx=5)
            
            points_label = tk.Label(row_frame, text="points", 
                                   font=("Arial", 14), bg="white", width=6, anchor="w")
            points_label.pack(side=tk.LEFT, padx=5)
            
            # Name entry - pre-fill with existing name
            name_entry = tk.Entry(row_frame, font=("Arial", 14), width=20)
            existing_name = self.point_names.get(position, "") if hasattr(self, 'point_names') else ""
            name_entry.insert(0, existing_name)
            name_entry.pack(side=tk.LEFT, padx=5)
            
            name_label = tk.Label(row_frame, text="name", 
                                 font=("Arial", 14), bg="white", width=5, anchor="w")
            name_label.pack(side=tk.LEFT, padx=5)
            
            self.point_entries[position] = point_entry
            self.name_entries[position] = name_entry
        
        # Example text
        example = tk.Label(main_frame, 
                         text='Example: Button 1 = 5 points, "Correct" | Button 2 = 0 points, "Incorrect"',
                         font=("Arial", 10), bg="white", fg="gray", justify="center")
        example.pack(pady=15)
        
        # Save button
        btn = tk.Button(main_frame, text="Save Configuration", font=("Arial", 16), 
                       command=self.save_point_values_and_names, width=20, bg="lightblue")
        btn.pack(pady=20)
    
    def save_point_values_and_names(self):
        """Save the point values and names and complete setup"""
        self.point_values = {}
        self.point_names = {}
        
        # Get values from entry fields
        for position in range(1, 6):
            # Get point value
            try:
                value = int(self.point_entries[position].get())
                self.point_values[position] = value
            except ValueError:
                # Invalid input, default to 0
                self.point_values[position] = 0
            
            # Get name
            name = self.name_entries[position].get().strip()
            self.point_names[position] = name if name else ""
        
        # Save everything to INI file
        self.save_mouse_config()
        
        # Show success message
        messagebox.showinfo("Configuration Complete", 
                          "Mouse buttons, point values, and names configured successfully!\n\n"
                          "Your settings have been saved.")
        
        # Start the application
        self.show_class_entry()
    
    def show_timer_settings(self):
        """Show timer configuration screen"""
        self.clear_screen()
        self.current_screen = "timer_settings"
        
        # Create main frame
        main_frame = tk.Frame(self.root, bg="white")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=40)
        
        # Title
        title = tk.Label(main_frame, text="Timer Settings", 
                        font=("Arial", 24, "bold"), bg="white")
        title.pack(pady=20)
        
        # Timer seconds setting
        seconds_frame = tk.Frame(main_frame, bg="white")
        seconds_frame.pack(pady=15)
        
        seconds_label = tk.Label(seconds_frame, text="Countdown Seconds:", 
                                font=("Arial", 14), bg="white", width=20, anchor="e")
        seconds_label.pack(side=tk.LEFT, padx=5)
        
        self.timer_seconds_entry = tk.Entry(seconds_frame, font=("Arial", 14), width=10)
        current_seconds = self.timer_seconds if hasattr(self, 'timer_seconds') else 0
        self.timer_seconds_entry.insert(0, str(current_seconds))
        self.timer_seconds_entry.pack(side=tk.LEFT, padx=5)
        
        # Validate that it's a number
        def validate_number(event):
            value = self.timer_seconds_entry.get()
            if value and not value.isdigit():
                # Remove non-digit characters
                self.timer_seconds_entry.delete(0, tk.END)
                self.timer_seconds_entry.insert(0, ''.join(filter(str.isdigit, value)))
        
        self.timer_seconds_entry.bind("<KeyRelease>", validate_number)
        
        info_label = tk.Label(seconds_frame, text="(0 = no timer)", 
                             font=("Arial", 10), bg="white", fg="gray")
        info_label.pack(side=tk.LEFT, padx=5)
        
        # Color settings
        colors_frame = tk.Frame(main_frame, bg="white")
        colors_frame.pack(pady=20)
        
        # Start colors (before reaching zero)
        start_label = tk.Label(colors_frame, text="Start Colors (before zero):", 
                              font=("Arial", 12, "bold"), bg="white")
        start_label.grid(row=0, column=0, columnspan=3, pady=5, sticky="w")
        
        tk.Label(colors_frame, text="Text Color:", font=("Arial", 11), bg="white").grid(row=1, column=0, sticky="e", padx=5)
        self.timer_start_text_entry = tk.Entry(colors_frame, font=("Arial", 11), width=10)
        start_text = self.timer_start_text_color if hasattr(self, 'timer_start_text_color') else '#000000'
        self.timer_start_text_entry.insert(0, start_text)
        self.timer_start_text_entry.grid(row=1, column=1, padx=5)
        tk.Label(colors_frame, text="(e.g., #000000 = black)", font=("Arial", 9), bg="white", fg="gray").grid(row=1, column=2, sticky="w")
        
        tk.Label(colors_frame, text="Background:", font=("Arial", 11), bg="white").grid(row=2, column=0, sticky="e", padx=5)
        self.timer_start_bg_entry = tk.Entry(colors_frame, font=("Arial", 11), width=10)
        start_bg = self.timer_start_bg_color if hasattr(self, 'timer_start_bg_color') else '#D3D3D3'
        self.timer_start_bg_entry.insert(0, start_bg)
        self.timer_start_bg_entry.grid(row=2, column=1, padx=5)
        tk.Label(colors_frame, text="(e.g., #D3D3D3 = light grey)", font=("Arial", 9), bg="white", fg="gray").grid(row=2, column=2, sticky="w")
        
        # End colors (at zero)
        end_label = tk.Label(colors_frame, text="End Colors (at zero):", 
                            font=("Arial", 12, "bold"), bg="white")
        end_label.grid(row=3, column=0, columnspan=3, pady=(15,5), sticky="w")
        
        tk.Label(colors_frame, text="Text Color:", font=("Arial", 11), bg="white").grid(row=4, column=0, sticky="e", padx=5)
        self.timer_end_text_entry = tk.Entry(colors_frame, font=("Arial", 11), width=10)
        end_text = self.timer_end_text_color if hasattr(self, 'timer_end_text_color') else '#00FF00'
        self.timer_end_text_entry.insert(0, end_text)
        self.timer_end_text_entry.grid(row=4, column=1, padx=5)
        tk.Label(colors_frame, text="(e.g., #00FF00 = green)", font=("Arial", 9), bg="white", fg="gray").grid(row=4, column=2, sticky="w")
        
        tk.Label(colors_frame, text="Background:", font=("Arial", 11), bg="white").grid(row=5, column=0, sticky="e", padx=5)
        self.timer_end_bg_entry = tk.Entry(colors_frame, font=("Arial", 11), width=10)
        end_bg = self.timer_end_bg_color if hasattr(self, 'timer_end_bg_color') else '#D3D3D3'
        self.timer_end_bg_entry.insert(0, end_bg)
        self.timer_end_bg_entry.grid(row=5, column=1, padx=5)
        tk.Label(colors_frame, text="(e.g., #D3D3D3 = light grey)", font=("Arial", 9), bg="white", fg="gray").grid(row=5, column=2, sticky="w")
        
        # Save and Cancel buttons
        button_frame = tk.Frame(main_frame, bg="white")
        button_frame.pack(pady=20)
        
        save_btn = tk.Button(button_frame, text="Save", font=("Arial", 14), 
                            command=self.save_timer_settings, width=12, bg="lightblue")
        save_btn.pack(side=tk.LEFT, padx=10)
        
        cancel_btn = tk.Button(button_frame, text="Cancel", font=("Arial", 14), 
                              command=self.show_class_entry, width=12)
        cancel_btn.pack(side=tk.LEFT, padx=10)
    
    def save_timer_settings(self):
        """Save timer settings to config"""
        # Get seconds value
        try:
            self.timer_seconds = int(self.timer_seconds_entry.get())
        except ValueError:
            self.timer_seconds = 0
        
        # Get color values
        self.timer_start_text_color = self.timer_start_text_entry.get().strip() or '#000000'
        self.timer_start_bg_color = self.timer_start_bg_entry.get().strip() or '#D3D3D3'
        self.timer_end_text_color = self.timer_end_text_entry.get().strip() or '#00FF00'
        self.timer_end_bg_color = self.timer_end_bg_entry.get().strip() or '#D3D3D3'
        
        # Save to config file
        self.save_mouse_config()
        
        messagebox.showinfo("Timer Settings Saved", 
                          f"Timer set to {self.timer_seconds} seconds\n\n"
                          "Settings have been saved.")
        
        self.show_class_entry()
    
    def find_special_image(self, prefix):
        """Find a special image file with given prefix (Start_, PreScore_, PostScore_) in root directory"""
        supported_formats = ['.png', '.jpg', '.jpeg', '.bmp', '.gif']
        
        # Look in the current directory (where the script is)
        for file in os.listdir('.'):
            if os.path.isfile(file):
                filename_lower = file.lower()
                if filename_lower.startswith(prefix.lower()):
                    _, ext = os.path.splitext(file)
                    if ext.lower() in supported_formats:
                        return file
        return None
    
    def show_special_image_screen(self, image_path, next_action):
        """Display a special image (Start, PreScore, PostScore) and wait for click"""
        self.clear_screen()
        self.current_screen = "special_image"
        
        # Store the image path and next action for resize handling
        self.special_image_path = image_path
        self.special_image_next_action = next_action
        
        # Create canvas
        self.special_canvas = tk.Canvas(self.root, bg="black", highlightthickness=0)
        self.special_canvas.pack(fill=tk.BOTH, expand=True)
        
        # Bind window resize event
        self.root.bind("<Configure>", self.on_special_image_resize)
        
        # Display the image
        self.display_special_image()
        
        # Bind click to continue
        self.special_canvas.bind("<Button-1>", lambda e: next_action())
        self.root.bind("<Return>", lambda e: next_action())
        
        self.special_canvas.focus_set()
    
    def on_special_image_resize(self, event):
        """Handle window resize for special images"""
        if self.current_screen == "special_image" and hasattr(self, 'special_canvas'):
            try:
                self.special_canvas.winfo_exists()
                # Add a small delay to avoid excessive redraws during resize
                if hasattr(self, '_special_resize_timer'):
                    self.root.after_cancel(self._special_resize_timer)
                self._special_resize_timer = self.root.after(100, self.display_special_image)
            except tk.TclError:
                pass
    
    def display_special_image(self):
        """Load and display the special image, scaled to window"""
        if not hasattr(self, 'special_canvas') or not hasattr(self, 'special_image_path'):
            return
        
        try:
            # Check if canvas still exists
            self.special_canvas.winfo_exists()
        except tk.TclError:
            return
        
        try:
            # Load image
            image = Image.open(self.special_image_path)
            
            # Update window to get accurate dimensions
            self.root.update_idletasks()
            
            # Get current canvas size
            canvas_width = self.special_canvas.winfo_width()
            canvas_height = self.special_canvas.winfo_height()
            
            # Fallback to window size if canvas size is too small (not yet rendered)
            if canvas_width < 100:
                canvas_width = self.root.winfo_width()
            if canvas_height < 100:
                canvas_height = self.root.winfo_height()
            
            # Get image dimensions
            img_width, img_height = image.size
            
            # Calculate scaling factor to fit within canvas (allow upscaling)
            if img_width > 0 and img_height > 0 and canvas_width > 0 and canvas_height > 0:
                width_ratio = canvas_width / img_width
                height_ratio = canvas_height / img_height
                scale = min(width_ratio, height_ratio)  # Removed the 1.0 cap to allow upscaling
                
                new_width = int(img_width * scale)
                new_height = int(img_height * scale)
                
                # Resize image
                image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # Convert to PhotoImage
            self.special_photo = ImageTk.PhotoImage(image)
            
            # Clear canvas and display image
            self.special_canvas.delete("all")
            self.special_canvas.create_image(canvas_width // 2, canvas_height // 2, 
                                           image=self.special_photo, anchor="center")
            
        except Exception as e:
            print(f"Error displaying special image: {e}")
    
    def lookup_student_in_roster(self, student_id):
        """Look up student name in roster. Returns name or None if not found."""
        if not hasattr(self, 'roster_file') or not os.path.exists(self.roster_file):
            return None
        
        # Normalize student ID (remove leading zeros for comparison)
        normalized_id = str(int(student_id)) if student_id.isdigit() else student_id
        
        try:
            with open(self.roster_file, 'r') as f:
                for line in f:
                    line = line.strip()
                    if line:
                        parts = line.split(' ', 1)
                        if len(parts) >= 1:
                            roster_id = parts[0]
                            # Normalize roster ID for comparison
                            normalized_roster_id = str(int(roster_id)) if roster_id.isdigit() else roster_id
                            
                            if normalized_id == normalized_roster_id:
                                # Found the student
                                return parts[1] if len(parts) > 1 else ""
        except Exception as e:
            print(f"Error reading roster: {e}")
        
        return None
    
    def update_roster(self, student_id, student_name):
        """Update or add student to roster, keeping numerical order"""
        if not hasattr(self, 'roster_file'):
            return
        
        # Read existing roster
        roster = []
        if os.path.exists(self.roster_file):
            try:
                with open(self.roster_file, 'r') as f:
                    for line in f:
                        line = line.strip()
                        if line:
                            roster.append(line)
            except Exception as e:
                print(f"Error reading roster: {e}")
        
        # Normalize student ID (pad with leading zero if single digit)
        if student_id.isdigit():
            formatted_id = student_id.zfill(2)
        else:
            formatted_id = student_id
        
        # Check if student already exists
        normalized_id = str(int(student_id)) if student_id.isdigit() else student_id
        found = False
        updated_roster = []
        
        for line in roster:
            parts = line.split(' ', 1)
            if len(parts) >= 1:
                roster_id = parts[0]
                normalized_roster_id = str(int(roster_id)) if roster_id.isdigit() else roster_id
                
                if normalized_id == normalized_roster_id:
                    # Update existing entry
                    updated_roster.append(f"{formatted_id} {student_name}")
                    found = True
                else:
                    updated_roster.append(line)
        
        # If not found, add new entry
        if not found:
            updated_roster.append(f"{formatted_id} {student_name}")
        
        # Sort roster numerically
        def get_sort_key(line):
            parts = line.split(' ', 1)
            if parts[0].isdigit():
                return int(parts[0])
            else:
                return float('inf')  # Non-numeric IDs go to the end
        
        updated_roster.sort(key=get_sort_key)
        
        # Write back to file
        try:
            with open(self.roster_file, 'w') as f:
                for line in updated_roster:
                    f.write(line + '\n')
        except Exception as e:
            print(f"Error writing roster: {e}")
    
    def sanitize_folder_name(self, name):
        """Remove invalid characters from folder name"""
        # Remove invisible characters and non-printable characters
        # Keep only alphanumeric, spaces, hyphens, and underscores
        sanitized = re.sub(r'[^\w\s-]', '', name)
        # Replace multiple spaces with single space
        sanitized = re.sub(r'\s+', ' ', sanitized)
        # Strip leading/trailing whitespace
        sanitized = sanitized.strip()
        # Replace spaces with underscores for folder name
        sanitized = sanitized.replace(' ', '_')
        return sanitized
    
    def show_class_entry(self):
        """Display the class number entry screen"""
        self.clear_screen()
        self.current_screen = "class"
        
        # Create frame
        frame = tk.Frame(self.root)
        frame.place(relx=0.5, rely=0.5, anchor="center")
        
        # Label
        label = tk.Label(frame, text="Class Number", font=("Arial", 24))
        label.pack(pady=20)
        
        # Entry field
        self.class_entry = tk.Entry(frame, font=("Arial", 18), width=20)
        self.class_entry.pack(pady=10)
        self.class_entry.focus_set()
        
        # Bind Enter and Tab
        self.class_entry.bind("<Return>", lambda e: self.submit_class())
        self.class_entry.bind("<Tab>", lambda e: self.submit_class())
        
        # Next button
        btn = tk.Button(frame, text="Next", font=("Arial", 16), 
                       command=self.submit_class, width=10)
        btn.pack(pady=10)
        
        # Checkbox frame for both checkboxes
        checkbox_frame = tk.Frame(frame)
        checkbox_frame.pack(pady=5)
        
        # Use Roster checkbox
        self.use_roster_var = tk.BooleanVar()
        self.use_roster_var.set(self.use_roster if hasattr(self, 'use_roster') else False)
        
        roster_check = tk.Checkbutton(checkbox_frame, text="Use Roster (track student names)", 
                                     variable=self.use_roster_var, font=("Arial", 11))
        roster_check.grid(row=0, column=0, padx=5, pady=2, sticky="w")
        
        # Bind checkbox change to save settings
        def on_roster_change():
            self.use_roster = self.use_roster_var.get()
            self.save_mouse_config()
        
        roster_check.config(command=on_roster_change)
        
        # Reverse Count checkbox
        self.reverse_count_var = tk.BooleanVar()
        self.reverse_count_var.set(self.reverse_count if hasattr(self, 'reverse_count') else False)
        
        reverse_check = tk.Checkbutton(checkbox_frame, text="Reverse Count (countdown student numbers)", 
                                      variable=self.reverse_count_var, font=("Arial", 11))
        reverse_check.grid(row=1, column=0, padx=5, pady=2, sticky="w")
        
        # Bind checkbox change to save settings
        def on_reverse_change():
            self.reverse_count = self.reverse_count_var.get()
            self.save_mouse_config()
        
        reverse_check.config(command=on_reverse_change)
        
        # Recalibrate button
        recal_btn = tk.Button(frame, text="Recalibrate Mouse", font=("Arial", 12), 
                             command=self.start_calibration, width=20, fg="blue")
        recal_btn.pack(pady=5)
        
        # Timer settings button
        timer_btn = tk.Button(frame, text="Timer Settings", font=("Arial", 12), 
                             command=self.show_timer_settings, width=20, fg="green")
        timer_btn.pack(pady=5)
    
    def submit_class(self):
        """Process class number and move to student ID entry"""
        raw_class_number = self.class_entry.get().strip()
        if not raw_class_number:
            return
        
        # Sanitize the class number for folder name
        sanitized_class = self.sanitize_folder_name(raw_class_number)
        
        if not sanitized_class:
            messagebox.showerror("Invalid Class Number", 
                               "Class number contains only invalid characters. Please enter a valid class number.")
            return
        
        # Store both versions
        self.class_number = raw_class_number
        self.class_folder = os.path.join(self.records_folder, sanitized_class)
        
        # Check if folder already exists
        if os.path.exists(self.class_folder):
            result = messagebox.askyesno(
                "Class Folder Exists",
                f"A folder for class '{sanitized_class}' already exists.\n\n"
                f"Do you want to use this existing folder?\n\n"
                f"Yes: Use existing folder\n"
                f"No: Go back and enter a different class number"
            )
            if not result:
                # User wants to cancel and rename
                self.class_entry.delete(0, tk.END)
                self.class_entry.focus_set()
                return
        else:
            # Create the class folder
            try:
                os.makedirs(self.class_folder)
            except Exception as e:
                messagebox.showerror("Error", f"Could not create class folder: {str(e)}")
                return
        
        # Create class summary filename (date only, no time)
        now = datetime.now()
        date_str = now.strftime("%y.%m.%d")
        self.class_summary_file = os.path.join(
            self.class_folder, 
            f"{self.class_number}_SpeakingTest.{date_str}.txt"
        )
        
        # Check if class summary file already exists for today
        if not os.path.exists(self.class_summary_file):
            # Create new file with header
            try:
                with open(self.class_summary_file, 'w') as f:
                    f.write(f"Class {self.class_number} - Speaking Test Summary\n")
                    f.write(f"Date: {now.strftime('%Y-%m-%d')}\n")
                    f.write("=" * 80 + "\n\n")
            except Exception as e:
                print(f"Error creating class summary file: {e}")
        
        # Set roster file path
        self.roster_file = os.path.join(
            self.class_folder,
            f"{self.class_number}_Roster.txt"
        )
        
        # Create roster file if Use Roster is enabled and file doesn't exist
        if hasattr(self, 'use_roster') and self.use_roster:
            if not os.path.exists(self.roster_file):
                try:
                    with open(self.roster_file, 'w') as f:
                        f.write("")  # Create empty roster
                except Exception as e:
                    print(f"Error creating roster file: {e}")
        
        # Check if there are subfolders in Slides directory
        self.check_for_slide_subfolders()
    
    def check_for_slide_subfolders(self):
        """Check if there are subfolders in the Slides directory and show selection if needed"""
        subfolders = []
        
        # Look for subdirectories in Slides folder
        if os.path.exists(self.slides_folder):
            for item in os.listdir(self.slides_folder):
                item_path = os.path.join(self.slides_folder, item)
                if os.path.isdir(item_path):
                    subfolders.append(item)
        
        if len(subfolders) > 0:
            # Multiple subfolders found, show selection screen
            self.show_slides_folder_selection(subfolders)
        else:
            # No subfolders, use the main Slides folder
            self.selected_slides_subfolder = None
            self.show_student_entry()
    
    def show_slides_folder_selection(self, subfolders):
        """Display a selection screen for choosing which slides subfolder to use"""
        self.clear_screen()
        self.current_screen = "folder_selection"
        
        # Create frame
        frame = tk.Frame(self.root)
        frame.place(relx=0.5, rely=0.5, anchor="center")
        
        # Title
        title = tk.Label(frame, text="Select Slides Folder", font=("Arial", 24, "bold"))
        title.pack(pady=20)
        
        # Instructions
        instructions = tk.Label(frame, 
                              text="Multiple slide folders found.\nPlease select which one to use for this test:",
                              font=("Arial", 14), justify="center")
        instructions.pack(pady=10)
        
        # Create a frame for the listbox and scrollbar
        list_frame = tk.Frame(frame)
        list_frame.pack(pady=20)
        
        # Scrollbar
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Listbox
        self.folder_listbox = tk.Listbox(list_frame, font=("Arial", 14), 
                                        width=40, height=min(15, len(subfolders)),
                                        yscrollcommand=scrollbar.set)
        self.folder_listbox.pack(side=tk.LEFT, fill=tk.BOTH)
        scrollbar.config(command=self.folder_listbox.yview)
        
        # Add "Use Main Slides Folder" option
        self.folder_listbox.insert(tk.END, "[ Use Main Slides Folder ]")
        
        # Add subfolders to listbox
        for folder in sorted(subfolders):
            self.folder_listbox.insert(tk.END, folder)
        
        # Select first item by default
        self.folder_listbox.select_set(0)
        self.folder_listbox.focus_set()
        
        # Bind double-click and Enter key
        self.folder_listbox.bind("<Double-Button-1>", lambda e: self.select_slides_folder())
        self.folder_listbox.bind("<Return>", lambda e: self.select_slides_folder())
        
        # Select button
        btn = tk.Button(frame, text="Select", font=("Arial", 16), 
                       command=self.select_slides_folder, width=15)
        btn.pack(pady=10)
    
    def select_slides_folder(self):
        """Process the selected slides folder and continue"""
        selection = self.folder_listbox.curselection()
        if not selection:
            return
        
        selected_index = selection[0]
        selected_text = self.folder_listbox.get(selected_index)
        
        if selected_index == 0:  # "Use Main Slides Folder"
            self.selected_slides_subfolder = None
        else:
            self.selected_slides_subfolder = selected_text
        
        # Proceed to student entry
        self.show_student_entry()
    
    def show_student_entry(self):
        """Display the student ID entry screen"""
        self.clear_screen()
        self.current_screen = "student"
        
        # Reset for new student - start at image 0 (first image)
        self.current_image_index = 0
        # Reset last clicked button
        self.last_clicked_button = None
        
        # Create frame
        frame = tk.Frame(self.root)
        frame.place(relx=0.5, rely=0.5, anchor="center")
        
        # Label
        label = tk.Label(frame, text="Student ID", font=("Arial", 24))
        label.pack(pady=20)
        
        # Entry field
        self.student_entry = tk.Entry(frame, font=("Arial", 18), width=20)
        self.student_entry.pack(pady=10)
        
        # Auto-increment logic: if last student ID was a number, increment or decrement it
        if hasattr(self, 'student_id') and self.student_id:
            try:
                # Try to convert last student ID to integer
                last_number = int(self.student_id)
                
                # Check if reverse count is enabled
                if hasattr(self, 'reverse_count') and self.reverse_count:
                    # Countdown - subtract 1, but don't go below 1
                    next_number = max(1, last_number - 1)
                else:
                    # Normal - add 1
                    next_number = last_number + 1
                
                # Pre-fill with incremented/decremented number
                self.student_entry.insert(0, str(next_number))
                # Select all text so typing immediately replaces it
                self.student_entry.select_range(0, tk.END)
                self.student_entry.icursor(tk.END)
            except ValueError:
                # Last student ID wasn't a number, leave field empty
                pass
        
        self.student_entry.focus_set()
        
        # Bind Enter and Tab
        self.student_entry.bind("<Return>", lambda e: self.submit_student())
        self.student_entry.bind("<Tab>", lambda e: self.submit_student())
        
        # Button
        btn = tk.Button(frame, text="Start Test", font=("Arial", 16), 
                       command=self.submit_student, width=10)
        btn.pack(pady=10)
    
    def submit_student(self):
        """Process student ID and start the image test"""
        self.student_id = self.student_entry.get().strip()
        if self.student_id:
            # Check if a file already exists for this student
            if self.check_existing_student_file():
                # File exists - ask for confirmation
                result = messagebox.askyesno(
                    "Student Already Tested",
                    f"A test file already exists for Student {self.student_id}.\n\n"
                    "This could mean:\n"
                    "• Student is retaking the test\n"
                    "• Wrong student number entered\n"
                    "• Previous test was incomplete\n\n"
                    "Do you want to continue?\n\n"
                    "Note: A new test file will be created (old file will NOT be overwritten).",
                    icon='warning'
                )
                if not result:
                    # User chose not to continue - stay on student entry screen
                    return
            
            # Proceed with test
            # Check if Use Roster is enabled
            if hasattr(self, 'use_roster') and self.use_roster:
                # Look up student in roster
                student_name = self.lookup_student_in_roster(self.student_id)
                self.show_student_name_screen(student_name if student_name else "")
            else:
                # No roster, proceed directly to test
                self.student_name = ""
                self.start_test()
    
    def check_existing_student_file(self):
        """Check if a test file already exists for the current student ID"""
        if not hasattr(self, 'class_folder') or not hasattr(self, 'student_id'):
            return False
        
        # Look for any file matching the pattern: SpeakingTest_CLASS_STUDENTID_*.txt
        # Need to check various formats of student ID (1, 01, etc.)
        pattern_ids = [self.student_id]
        
        # If numeric, also check zero-padded version
        if self.student_id.isdigit():
            padded = self.student_id.zfill(2)
            if padded != self.student_id:
                pattern_ids.append(padded)
            # Also check unpadded version
            unpadded = str(int(self.student_id))
            if unpadded != self.student_id:
                pattern_ids.append(unpadded)
        
        # Check if class folder exists
        if not os.path.exists(self.class_folder):
            return False
        
        # Look for matching files
        for filename in os.listdir(self.class_folder):
            if filename.startswith('SpeakingTest_') and filename.endswith('.txt'):
                # Parse filename: SpeakingTest_CLASS_STUDENTID_TIMESTAMP.txt
                parts = filename.replace('.txt', '').split('_')
                if len(parts) >= 3:
                    file_student_id = parts[2]
                    if file_student_id in pattern_ids:
                        return True
        
        return False
    
    def show_student_name_screen(self, current_name):
        """Show screen to confirm or enter student name"""
        self.clear_screen()
        self.current_screen = "student_name"
        
        # Create frame
        frame = tk.Frame(self.root)
        frame.place(relx=0.5, rely=0.5, anchor="center")
        
        # Title label
        title = tk.Label(frame, text="Student Information", font=("Arial", 24, "bold"))
        title.pack(pady=20)
        
        # Student Number section
        number_frame = tk.Frame(frame)
        number_frame.pack(pady=10)
        
        number_label = tk.Label(number_frame, text="Student Number:", font=("Arial", 14))
        number_label.pack(side=tk.LEFT, padx=5)
        
        self.student_number_entry = tk.Entry(number_frame, font=("Arial", 14), width=10)
        self.student_number_entry.insert(0, self.student_id)
        # Do NOT select/highlight the text - just leave cursor at end
        self.student_number_entry.pack(side=tk.LEFT, padx=5)
        
        # Student Name section
        name_frame = tk.Frame(frame)
        name_frame.pack(pady=10)
        
        name_label = tk.Label(name_frame, text="Student Name:", font=("Arial", 14))
        name_label.pack(side=tk.LEFT, padx=5)
        
        self.name_entry = tk.Entry(name_frame, font=("Arial", 14), width=30)
        self.name_entry.insert(0, current_name)
        if current_name:
            # Select all text in NAME field so typing replaces it
            self.name_entry.select_range(0, tk.END)
            self.name_entry.icursor(tk.END)
        self.name_entry.pack(side=tk.LEFT, padx=5)
        
        # Set focus to name field
        self.name_entry.focus_set()
        
        # Bind Enter on name field - explicitly prevent Tab from doing anything
        self.name_entry.bind("<Return>", lambda e: self.submit_student_name(current_name))
        self.name_entry.bind("<Tab>", lambda e: "break")  # Prevent Tab from propagating
        
        # Also bind Enter on number field
        self.student_number_entry.bind("<Return>", lambda e: self.submit_student_name(current_name))
        
        # Submit button
        btn = tk.Button(frame, text="Submit", font=("Arial", 16), 
                       command=lambda: self.submit_student_name(current_name), width=10)
        btn.pack(pady=20)
    
    def submit_student_name(self, original_name):
        """Process student name and update roster if changed"""
        # Store the original student ID before potentially changing it
        original_id = self.student_id
        
        # Get the student number from the entry field (may have been changed)
        entered_number = self.student_number_entry.get().strip()
        
        # Get the student name
        entered_name = self.name_entry.get().strip()
        
        # Use the entered number as the official student_id
        if entered_number:
            self.student_id = entered_number
        
        # Update roster - always update if we got here, because either name or number may have changed
        self.update_roster(self.student_id, entered_name)
        
        # Store student name
        self.student_name = entered_name
        
        # Proceed to test
        self.start_test()
    
    def start_test(self):
        """Start the actual test (called after student ID and optionally name)"""
        self.load_images()
        self.create_output_file()
        # Ensure we start at the first image
        self.current_image_index = 0
        # Reset last clicked button
        self.last_clicked_button = None
        
        # Check for Start image
        start_image = self.find_special_image("Start_")
        if start_image:
            # Store the fact that we need to show image screen after Start
            self.show_special_image_screen(start_image, self.begin_test_after_start)
        else:
            self.begin_test_after_start()
    
    def begin_test_after_start(self):
        """Begin the test after Start screen (or directly if no Start screen)"""
        # Make absolutely sure we're at the first image
        self.current_image_index = 0
        self.show_image_screen()
    
    def load_images(self):
        """Load all images from the slides folder or selected subfolder in alphabetical order"""
        self.image_files = {}
        self.unscored_images = set()  # Track which images should not be scored
        
        # Determine which folder to load images from
        if self.selected_slides_subfolder:
            images_folder = os.path.join(self.slides_folder, self.selected_slides_subfolder)
        else:
            images_folder = self.slides_folder
        
        # Supported image formats
        supported_formats = ['.png', '.jpg', '.jpeg', '.bmp', '.gif']
        
        # Check if there are category subfolders
        category_folders = []
        if os.path.exists(images_folder):
            for item in os.listdir(images_folder):
                item_path = os.path.join(images_folder, item)
                if os.path.isdir(item_path):
                    category_folders.append((item, item_path))
        
        current_index = 0
        
        if category_folders:
            # We have category subfolders - load from each in alphabetical order
            category_folders.sort(key=lambda x: x[0].lower())
            
            for category_name, category_path in category_folders:
                # Collect all image files in this category
                category_images = []
                for file in os.listdir(category_path):
                    file_path = os.path.join(category_path, file)
                    if os.path.isfile(file_path):
                        name, ext = os.path.splitext(file)
                        if ext.lower() in supported_formats:
                            category_images.append((file, file_path))
                
                # Sort files alphabetically
                category_images.sort(key=lambda x: x[0].lower())
                
                # Add to main image list
                for i, (filename, filepath) in enumerate(category_images):
                    self.image_files[current_index] = filepath
                    
                    # Mark first image in each category as unscored
                    if i == 0:
                        self.unscored_images.add(current_index)
                    
                    current_index += 1
        else:
            # No category folders - load images directly from the folder
            image_file_list = []
            if os.path.exists(images_folder):
                for file in os.listdir(images_folder):
                    file_path = os.path.join(images_folder, file)
                    if os.path.isfile(file_path):
                        name, ext = os.path.splitext(file)
                        if ext.lower() in supported_formats:
                            image_file_list.append((file, file_path))
            
            # Sort files alphabetically (case-insensitive)
            image_file_list.sort(key=lambda x: x[0].lower())
            
            # Assign sequential numbers starting from 0
            for index, (filename, filepath) in enumerate(image_file_list):
                self.image_files[index] = filepath
    
    def create_output_file(self):
        """Create the output text file with timestamp in the class folder"""
        now = datetime.now()
        timestamp = now.strftime("%Y.%m.%d.%H%M")
        filename = f"SpeakingTest_{self.class_number}_{self.student_id}_{timestamp}.txt"
        
        # Save in the class folder
        self.output_file = os.path.join(self.class_folder, filename)
        
        # Create empty file
        with open(self.output_file, 'w') as f:
            f.write("")  # Empty file for now
    
    def show_image_screen(self):
        """Display the current image with mouse controls"""
        self.clear_screen()
        self.current_screen = "image"
        
        # Create main canvas
        self.canvas = tk.Canvas(self.root, bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # Bind window resize event to redisplay image
        self.root.bind("<Configure>", self.on_window_resize)
        
        # Bind mouse events based on calibrated mapping
        for button_event, position in self.mouse_button_map.items():
            # Extract button number from "Button-X" format
            button_num = button_event.split('-')[1]
            self.canvas.bind(f"<{button_event}>", 
                           lambda e, pos=position: self.handle_click(pos))
        
        # Bind scroll wheel for Windows
        self.canvas.bind("<MouseWheel>", self.handle_scroll)
        
        # Bind keyboard keys
        self.root.bind("<BackSpace>", self.handle_backspace)
        self.root.bind("<Up>", lambda e: self.go_previous())  # Up arrow - go back
        self.root.bind("<Down>", lambda e: self.go_next())  # Down arrow - go forward
        
        # Bind number keys 1-5 to function like mouse positions 1-5
        self.root.bind("1", lambda e: self.handle_click(1))
        self.root.bind("2", lambda e: self.handle_click(2))
        self.root.bind("3", lambda e: self.handle_click(3))
        self.root.bind("4", lambda e: self.handle_click(4))
        self.root.bind("5", lambda e: self.handle_click(5))
        
        # Bind Tab to jump to student ID input
        self.root.bind("<Tab>", lambda e: self.show_student_entry())
        # Bind Shift+Tab to jump to class input
        self.root.bind("<Shift-Tab>", lambda e: self.show_class_entry())
        
        self.canvas.focus_set()
        
        # Initialize elapsed time clock
        self.test_start_time = None
        self.elapsed_seconds = 0
        self.start_elapsed_clock()
        
        # Initialize timer if enabled
        if hasattr(self, 'timer_seconds') and self.timer_seconds > 0:
            self.current_timer_value = self.timer_seconds
            self.start_countdown_timer()
        else:
            self.current_timer_value = None
        
        # Display the current image
        self.display_current_image()
    
    def start_elapsed_clock(self):
        """Start the elapsed time clock"""
        if self.test_start_time is None:
            self.test_start_time = datetime.now()
        
        if self.current_screen == "image":
            # Calculate elapsed time
            elapsed = datetime.now() - self.test_start_time
            self.elapsed_seconds = int(elapsed.total_seconds())
            
            # Schedule next update
            self.elapsed_clock_id = self.root.after(1000, self.start_elapsed_clock)
            
            # Redisplay to update clock
            self.display_current_image()
    
    def start_countdown_timer(self):
        """Start the countdown timer"""
        if self.current_screen == "image" and self.current_timer_value is not None:
            if self.current_timer_value > 0:
                self.current_timer_value -= 1
                # Schedule next countdown
                self.timer_id = self.root.after(1000, self.start_countdown_timer)
                # Redisplay to update timer
                self.display_current_image()
            elif self.current_timer_value == 0:
                # Timer just reached zero - beep!
                try:
                    import winsound
                    winsound.Beep(1000, 200)  # 1000 Hz for 200ms
                except:
                    # If winsound not available (non-Windows), try system beep
                    try:
                        print('\a')  # ASCII bell character
                    except:
                        pass
                # Don't count down below zero, just stay at 0
                # Redisplay to update timer color
                self.display_current_image()
            # If it reaches 0, it just stops (stays at 0 and turns green)
    
    def on_window_resize(self, event):
        """Handle window resize events - redisplay the current image"""
        # Only redisplay if we're on the image screen and the canvas still exists
        if self.current_screen == "image" and hasattr(self, 'canvas'):
            try:
                # Check if canvas still exists and is valid
                self.canvas.winfo_exists()
                # Add a small delay to avoid excessive redraws during resize
                if hasattr(self, '_resize_timer'):
                    self.root.after_cancel(self._resize_timer)
                self._resize_timer = self.root.after(100, self.display_current_image)
            except tk.TclError:
                # Canvas was destroyed, ignore
                pass
    
    def display_current_image(self):
        """Load and display the current image"""
        # Safety check: make sure we're still on the image screen
        if self.current_screen != "image" or not hasattr(self, 'canvas'):
            return
        
        try:
            # Check if canvas still exists
            self.canvas.winfo_exists()
        except tk.TclError:
            # Canvas was destroyed, stop here
            return
        
        # Check if current image index is valid
        max_index = max(self.image_files.keys(), default=-1)
        
        # If current index is beyond the max, loop to the last image
        if self.current_image_index > max_index:
            if max_index >= 0:
                self.current_image_index = max_index
            else:
                messagebox.showerror("Error", "No images found!")
                return
        
        # If current index is below 0, set to first image
        if self.current_image_index < 0:
            self.current_image_index = 0
        
        if self.current_image_index not in self.image_files:
            messagebox.showerror("Error", f"Image {self.current_image_index} not found!")
            return
        
        # Load image
        image_path = self.image_files[self.current_image_index]
        image = Image.open(image_path)
        
        # Resize to fit window while maintaining aspect ratio
        canvas_width = self.root.winfo_width()
        canvas_height = self.root.winfo_height()
        
        # Get image dimensions
        img_width, img_height = image.size
        
        # Calculate scaling factor
        width_ratio = canvas_width / img_width
        height_ratio = canvas_height / img_height
        scale = min(width_ratio, height_ratio)  # Allow upscaling
        
        new_width = int(img_width * scale)
        new_height = int(img_height * scale)
        
        image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        
        # Convert to PhotoImage
        self.photo = ImageTk.PhotoImage(image)
        
        # Clear canvas and display image
        self.canvas.delete("all")
        self.canvas.create_image(canvas_width // 2, canvas_height // 2, 
                                image=self.photo, anchor="center")
        
        # Display last clicked button number in lower left corner
        if hasattr(self, 'last_clicked_button') and self.last_clicked_button is not None:
            # Create background rectangle
            padding = 5
            text_x = 10 + padding
            text_y = canvas_height - 10 - padding
            
            # Create text (will be used to measure size)
            text = f"{self.last_clicked_button}"
            
            # Estimate text size (approximate)
            text_width = len(text) * 8  # Rough estimate for 12pt font
            text_height = 16
            
            # Draw light grey background rectangle
            self.canvas.create_rectangle(
                text_x - padding,
                text_y - text_height - padding,
                text_x + text_width + padding,
                text_y + padding,
                fill="lightgrey",
                outline=""
            )
            
            # Draw black text on top
            self.canvas.create_text(
                text_x,
                text_y - text_height // 2,
                text=text,
                font=("Arial", 12),
                fill="black",
                anchor="w"
            )
        
        # Display point value in lower right corner
        if hasattr(self, 'last_clicked_button') and self.last_clicked_button is not None:
            if hasattr(self, 'point_values') and self.last_clicked_button in self.point_values:
                points = self.point_values[self.last_clicked_button]
                
                # Create background rectangle
                padding = 5
                text = f"{points}"  # Just the number, no "pts"
                
                # Estimate text size
                text_width = len(text) * 8
                text_height = 16
                
                text_x = canvas_width - 10 - text_width - padding
                text_y = canvas_height - 10 - padding
                
                # Draw light grey background rectangle
                self.canvas.create_rectangle(
                    text_x - padding,
                    text_y - text_height - padding,
                    text_x + text_width + padding * 2,
                    text_y + padding,
                    fill="lightgrey",
                    outline=""
                )
                
                # Draw black text on top
                self.canvas.create_text(
                    text_x + padding,
                    text_y - text_height // 2,
                    text=text,
                    font=("Arial", 12),
                    fill="black",
                    anchor="w"
                )
        
        # Display elapsed time clock in upper left corner
        if hasattr(self, 'elapsed_seconds'):
            # Format as MM:SS
            minutes = self.elapsed_seconds // 60
            seconds = self.elapsed_seconds % 60
            clock_text = f"{minutes:02d}:{seconds:02d}"
            
            padding = 5
            text_width = len(clock_text) * 15  # Estimate for 24pt font
            text_height = 28
            
            text_x = 10 + padding
            text_y = 10 + padding
            
            # Draw light grey background rectangle
            self.canvas.create_rectangle(
                text_x - padding,
                text_y - padding,
                text_x + text_width + padding,
                text_y + text_height + padding,
                fill="lightgrey",
                outline=""
            )
            
            # Draw black text
            self.canvas.create_text(
                text_x,
                text_y + text_height // 2,
                text=clock_text,
                font=("Arial", 24),
                fill="black",
                anchor="w"
            )
        
        # Display "UNSCORED" indicator for category intro slides (center top)
        if hasattr(self, 'unscored_images') and self.current_image_index in self.unscored_images:
            unscored_text = "UNSCORED"
            padding = 8
            text_width = len(unscored_text) * 18  # Estimate for 28pt bold font
            text_height = 32
            
            # Center horizontally
            text_x = canvas_width // 2
            text_y = 15 + padding
            
            # Draw yellow background rectangle
            self.canvas.create_rectangle(
                text_x - text_width // 2 - padding,
                text_y - padding,
                text_x + text_width // 2 + padding,
                text_y + text_height - padding,
                fill="yellow",
                outline="black",
                width=2
            )
            
            # Draw bold black text
            self.canvas.create_text(
                text_x,
                text_y + text_height // 2 - padding,
                text=unscored_text,
                font=("Arial", 28, "bold"),
                fill="black",
                anchor="center"
            )
        
        # Display countdown timer in upper right corner
        if hasattr(self, 'current_timer_value') and self.current_timer_value is not None:
            padding = 5
            timer_text = str(self.current_timer_value)
            
            # Estimate text size for size 24 font
            text_width = len(timer_text) * 15  # Rough estimate for 24pt font
            text_height = 28
            
            text_x = canvas_width - 10 - text_width - padding
            text_y = 10 + padding
            
            # Determine colors based on timer value
            if self.current_timer_value > 0:
                # Use start colors
                text_color = self.timer_start_text_color if hasattr(self, 'timer_start_text_color') else '#000000'
                bg_color = self.timer_start_bg_color if hasattr(self, 'timer_start_bg_color') else '#D3D3D3'
            else:
                # Use end colors (at zero)
                text_color = self.timer_end_text_color if hasattr(self, 'timer_end_text_color') else '#00FF00'
                bg_color = self.timer_end_bg_color if hasattr(self, 'timer_end_bg_color') else '#D3D3D3'
            
            # Draw background rectangle
            self.canvas.create_rectangle(
                text_x - padding,
                text_y - padding,
                text_x + text_width + padding * 2,
                text_y + text_height + padding,
                fill=bg_color,
                outline=""
            )
            
            # Draw timer text
            self.canvas.create_text(
                text_x + padding,
                text_y + text_height // 2,
                text=timer_text,
                font=("Arial", 24),
                fill=text_color,
                anchor="w"
            )
    
    def handle_click(self, button):
        """Handle mouse button clicks"""
        # Check if this is the last image
        max_image_index = max(self.image_files.keys(), default=0)
        is_last_image = (self.current_image_index == max_image_index)
        
        # Only record if this image is NOT marked as unscored
        if not (hasattr(self, 'unscored_images') and self.current_image_index in self.unscored_images):
            # Record the click for current image
            self.save_record(self.current_image_index, button)
        
        # Store the last clicked button number for display (even if unscored)
        self.last_clicked_button = button
        
        # Move to next image or show score summary
        if is_last_image:
            # Cancel timers
            if hasattr(self, 'timer_id'):
                self.root.after_cancel(self.timer_id)
            if hasattr(self, 'elapsed_clock_id'):
                self.root.after_cancel(self.elapsed_clock_id)
            
            # Last image - check for PreScore image, then show score summary
            prescore_image = self.find_special_image("PreScore_")
            if prescore_image:
                self.show_special_image_screen(prescore_image, self.show_score_summary)
            else:
                self.show_score_summary()
        else:
            self.current_image_index += 1
            
            # Reset timer for next image
            if hasattr(self, 'timer_seconds') and self.timer_seconds > 0:
                # Cancel previous timer
                if hasattr(self, 'timer_id'):
                    self.root.after_cancel(self.timer_id)
                # Reset timer value
                self.current_timer_value = self.timer_seconds
                # Start countdown again
                self.start_countdown_timer()
            
            self.display_current_image()
    
    def calculate_score(self):
        """Calculate the student's score from the output file"""
        # Read the output file
        try:
            with open(self.output_file, 'r') as f:
                lines = f.readlines()
        except:
            return 0, 0, 0  # score, max_score, percentage
        
        # Parse lines and get the latest entry for each question
        question_scores = {}  # question_num: points
        
        for line in lines:
            # Parse format: "Question 00: 5 = 1 0 0 0 0 filename"
            if line.startswith("Question"):
                parts = line.split(":")
                if len(parts) >= 2:
                    # Get question number
                    q_num_str = parts[0].replace("Question", "").strip()
                    try:
                        q_num = int(q_num_str)
                    except:
                        continue
                    
                    # Get point value (between : and =)
                    rest = parts[1].strip()
                    if "=" in rest:
                        point_str = rest.split("=")[0].strip()
                        try:
                            points = int(point_str)
                            # Store this score (overwrites previous if exists)
                            question_scores[q_num] = points
                        except:
                            continue
        
        # Calculate total score and max possible score
        num_questions = len(question_scores)
        if num_questions == 0:
            return 0, 0, 0
        
        # Get max point value (highest configured point value)
        max_point_value = max(self.point_values.values()) if hasattr(self, 'point_values') and self.point_values else 5
        
        total_score = sum(question_scores.values())
        max_score = num_questions * max_point_value
        
        # Calculate percentage
        if max_score > 0:
            percentage = round((total_score / max_score) * 100)
        else:
            percentage = 0
        
        return total_score, max_score, percentage
    
    def show_score_summary(self):
        """Show the score summary screen"""
        self.clear_screen()
        self.current_screen = "score_summary"
        
        # Calculate score
        total_score, max_score, percentage = self.calculate_score()
        
        # Add header to the output file
        self.add_header_to_output_file(total_score, max_score, percentage)
        
        # Add entry to class summary file
        self.add_to_class_summary(total_score, max_score, percentage)
        
        # Create frame with black background
        frame = tk.Frame(self.root, bg="black")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Center container
        center_frame = tk.Frame(frame, bg="black")
        center_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        # Percentage (larger font, white text)
        percentage_label = tk.Label(center_frame, 
                                   text=f"{percentage}%",
                                   font=("Arial", 72, "bold"),
                                   bg="black",
                                   fg="white")
        percentage_label.pack()
        
        # Score fraction (smaller font, white text, centered below)
        score_label = tk.Label(center_frame,
                              text=f"{total_score}/{max_score}",
                              font=("Arial", 48),
                              bg="black",
                              fg="white")
        score_label.pack()
        
        # Continue button
        continue_btn = tk.Button(frame,
                                text="Next Student",
                                font=("Arial", 16),
                                command=self.continue_after_score,
                                width=15)
        continue_btn.pack(side=tk.BOTTOM, pady=30)
        
        # Also bind Enter key to continue
        self.root.bind("<Return>", lambda e: self.continue_after_score())
        
        # Bind mouse click anywhere to continue
        frame.bind("<Button-1>", lambda e: self.continue_after_score())
        percentage_label.bind("<Button-1>", lambda e: self.continue_after_score())
        score_label.bind("<Button-1>", lambda e: self.continue_after_score())
    
    def continue_after_score(self):
        """Check for PostScore image, then continue to student entry"""
        postscore_image = self.find_special_image("PostScore_")
        if postscore_image:
            self.show_special_image_screen(postscore_image, self.show_student_entry)
        else:
            self.show_student_entry()
    
    def add_to_class_summary(self, total_score, max_score, percentage):
        """Add student's score to the class summary file"""
        try:
            if hasattr(self, 'class_summary_file'):
                now = datetime.now()
                timestamp = now.strftime("%y.%m.%d.%H%M")
                
                # Format student ID (pad to 2 digits if numeric)
                if self.student_id.isdigit():
                    student_display = self.student_id.zfill(2)
                else:
                    student_display = self.student_id.ljust(2)
                
                # Format student name if available
                if hasattr(self, 'student_name') and self.student_name:
                    name_display = self.student_name.ljust(20)[:20]  # Limit to 20 chars, pad if shorter
                else:
                    name_display = "".ljust(20)
                
                # Format score (right-aligned)
                score_display = f"{total_score}/{max_score}".rjust(10)
                
                # Format percentage (right-aligned)
                percent_display = f"{percentage}%".rjust(5)
                
                # Create line with aligned columns
                line = f"{timestamp}:   Student {student_display} {name_display}:  {score_display} = {percent_display}\n"
                
                with open(self.class_summary_file, 'a') as f:
                    f.write(line)
        except Exception as e:
            print(f"Error adding to class summary: {e}")
    
    def add_header_to_output_file(self, total_score, max_score, percentage):
        """Add header with summary information and point scale to the output file"""
        try:
            # Read existing content
            with open(self.output_file, 'r') as f:
                existing_content = f.read()
            
            # Count total questions (unique question numbers)
            question_nums = set()
            for line in existing_content.split('\n'):
                if line.startswith("Question"):
                    parts = line.split(":")
                    if len(parts) >= 1:
                        q_num_str = parts[0].replace("Question", "").strip()
                        try:
                            q_num = int(q_num_str)
                            question_nums.add(q_num)
                        except:
                            pass
            
            total_questions = len(question_nums)
            
            # Get the filename for the header
            filename = os.path.basename(self.output_file)
            
            # Create header
            header = f"{filename}\n"
            
            # Add student info if available
            if hasattr(self, 'student_id'):
                # Format student ID (pad if numeric)
                if self.student_id.isdigit():
                    student_display = self.student_id.zfill(2)
                else:
                    student_display = self.student_id
                
                if hasattr(self, 'student_name') and self.student_name:
                    header += f"Student: {student_display} {self.student_name}\n"
                else:
                    header += f"Student: {student_display}\n"
            
            header += "=" * 67 + "\n"
            header += f"Total Questions: {total_questions}   Max Score: {max_score}   Score: {total_score}   Percentage: {percentage}%\n"
            header += "=" * 67 + "\n"
            
            # Add point scale (sorted by point value, descending)
            if hasattr(self, 'point_values') and hasattr(self, 'point_names'):
                # Create list of (points, name, position) tuples
                point_scale = []
                for position in range(1, 6):
                    points = self.point_values.get(position, 0)
                    name = self.point_names.get(position, "")
                    
                    # Include if: points != 0 OR name is not empty
                    if points != 0 or name:
                        point_scale.append((points, name, position))
                
                # Sort by points (descending), then by position if tied
                point_scale.sort(key=lambda x: (-x[0], x[2]))
                
                # Add to header
                for points, name, position in point_scale:
                    header += f"{points} = {name}\n"
            
            header += "\n"
            
            # Write header + existing content
            with open(self.output_file, 'w') as f:
                f.write(header + existing_content)
                
        except Exception as e:
            # If there's an error, just continue without adding header
            print(f"Error adding header: {e}")
    
    def handle_scroll(self, event):
        """Handle mouse wheel scrolling on Windows"""
        if event.delta > 0:
            # Scroll up - go to previous image
            self.go_previous()
        else:
            # Scroll down - go to next image
            self.go_next()
    
    def go_previous(self):
        """Navigate to previous image"""
        if self.current_image_index > 0:
            self.current_image_index -= 1
            self.display_current_image()
        # If already at first image (0), do nothing
    
    def go_next(self):
        """Navigate to next image"""
        max_index = max(self.image_files.keys(), default=0)
        if self.current_image_index < max_index:
            self.current_image_index += 1
            self.display_current_image()
        # If already at last image, do nothing
    
    def save_record(self, image_num, button):
        """Save the click record to the output file"""
        # Create the record line
        clicks = [0, 0, 0, 0, 0]
        clicks[button - 1] = 1
        
        # Get the point value for this button
        if hasattr(self, 'point_values') and button in self.point_values:
            points = self.point_values[button]
        else:
            points = 0
        
        # Get the filename without extension
        if image_num in self.image_files:
            filepath = self.image_files[image_num]
            filename = os.path.basename(filepath)
            filename_no_ext = os.path.splitext(filename)[0]
        else:
            filename_no_ext = "Unknown"
        
        line = f"Question {image_num:02d}: {points} = {' '.join(map(str, clicks))} {filename_no_ext}\n"
        
        # Append to file
        with open(self.output_file, 'a') as f:
            f.write(line)
    
    def handle_backspace(self, event):
        """Handle backspace key - go back and delete last recorded line"""
        # Only allow backspace if we're past the first image (index 0)
        if self.current_image_index > 0:
            # Delete the last line from the output file
            self.delete_last_line()
            
            # Clear the last clicked button display
            self.last_clicked_button = None
            
            # Go back to previous image
            self.go_previous()
    
    def delete_last_line(self):
        """Delete the last line from the output file"""
        try:
            # Read all lines from the file
            with open(self.output_file, 'r') as f:
                lines = f.readlines()
            
            # If there are lines to delete, remove the last one
            if lines:
                lines = lines[:-1]
                
                # Write back all lines except the last one
                with open(self.output_file, 'w') as f:
                    f.writelines(lines)
        except Exception as e:
            # If file doesn't exist yet or other error, just continue
            pass
    
    def handle_escape(self, event):
        """Handle ESC key press"""
        if self.current_screen == "image":
            result = messagebox.askyesno("Exit", "Do you want to exit the test?")
            if result:
                self.show_class_entry()
    
    def clear_screen(self):
        """Clear all widgets from the screen"""
        # Cancel any pending resize timer
        if hasattr(self, '_resize_timer'):
            try:
                self.root.after_cancel(self._resize_timer)
            except:
                pass
            delattr(self, '_resize_timer')
        
        # Cancel any pending special image resize timer
        if hasattr(self, '_special_resize_timer'):
            try:
                self.root.after_cancel(self._special_resize_timer)
            except:
                pass
            delattr(self, '_special_resize_timer')
        
        # Cancel any running countdown timer
        if hasattr(self, 'timer_id'):
            try:
                self.root.after_cancel(self.timer_id)
            except:
                pass
            delattr(self, 'timer_id')
        
        # Cancel any running elapsed clock
        if hasattr(self, 'elapsed_clock_id'):
            try:
                self.root.after_cancel(self.elapsed_clock_id)
            except:
                pass
            delattr(self, 'elapsed_clock_id')
        
        # Clean up special image references
        if hasattr(self, 'special_image_path'):
            delattr(self, 'special_image_path')
        if hasattr(self, 'special_image_next_action'):
            delattr(self, 'special_image_next_action')
        if hasattr(self, 'special_canvas'):
            delattr(self, 'special_canvas')
        if hasattr(self, 'special_photo'):
            delattr(self, 'special_photo')
        
        # Unbind all common event bindings to prevent errors and conflicts
        try:
            self.root.unbind("<Configure>")
            self.root.unbind("<Return>")
            self.root.unbind("<Tab>")
            self.root.unbind("<Shift-Tab>")
            self.root.unbind("<BackSpace>")
            self.root.unbind("<Up>")
            self.root.unbind("<Down>")
            self.root.unbind("<Key>")
            self.root.unbind("1")
            self.root.unbind("2")
            self.root.unbind("3")
            self.root.unbind("4")
            self.root.unbind("5")
        except:
            pass
        
        for widget in self.root.winfo_children():
            widget.destroy()


def main():
    """Main entry point with error handling"""
    try:
        root = tk.Tk()
        app = SpeakingTestApp(root)
        root.mainloop()
    except Exception as e:
        # If there's an error, show it in a message box
        import traceback
        error_msg = f"An error occurred:\n\n{str(e)}\n\n"
        error_msg += "Full error details:\n" + traceback.format_exc()
        
        # Try to show in messagebox
        try:
            root = tk.Tk()
            root.withdraw()  # Hide the main window
            messagebox.showerror("Speaking Test Error", error_msg)
        except:
            # If messagebox fails, print to console and wait
            print("=" * 80)
            print("ERROR IN SPEAKING TEST APPLICATION")
            print("=" * 80)
            print(error_msg)
            print("=" * 80)
            input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
