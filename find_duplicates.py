#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Find Duplicate Speaking Tests
==============================
Finds duplicate test files based on class number and student ID.
Moves older duplicates to Records/Duplicates folder for review.

A duplicate is when the same student (class + student ID) has multiple test files.
The OLDEST file is moved to Duplicates, keeping the most recent test.

Usage: python find_duplicates.py

This helps you:
- Find students who took the test twice
- Detect accidental duplicate student numbers
- Review which tests to keep
"""

import os
import re
import shutil
from datetime import datetime
from pathlib import Path
from collections import defaultdict


def parse_filename(filename):
    """
    Parse test filename to extract class number, student ID, and timestamp.
    Format: SpeakingTest_CLASS_STUDENTID_TIMESTAMP.txt
    Returns: (class_number, student_id, timestamp_str, full_filename)
    Returns None if filename doesn't match pattern
    """
    if not filename.startswith('SpeakingTest_') or not filename.endswith('.txt'):
        return None
    
    # Remove .txt extension
    name = filename.replace('.txt', '')
    
    # Split by underscores
    parts = name.split('_')
    
    if len(parts) >= 4 and parts[0] == 'SpeakingTest':
        class_number = parts[1]
        student_id = parts[2]
        timestamp = parts[3]  # Format: YYYY.MM.DD.HHMM or YY.MM.DD.HHMM
        
        return (class_number, student_id, timestamp, filename)
    
    return None


def normalize_student_id(student_id):
    """
    Normalize student ID for comparison.
    '1' and '01' should be considered the same student.
    """
    # If it's numeric, remove leading zeros for comparison
    if student_id.isdigit():
        return str(int(student_id))
    return student_id


def get_timestamp_for_sorting(timestamp_str):
    """
    Convert timestamp string to datetime for sorting.
    Handles both YYYY.MM.DD.HHMM and YY.MM.DD.HHMM formats.
    Returns datetime object or current time if parsing fails.
    """
    try:
        parts = timestamp_str.split('.')
        if len(parts) == 4:
            year, month, day, time = parts
            
            # Determine if year is YY or YYYY
            if len(year) == 2:
                year = '20' + year  # Assume 20xx
            
            # Extract hour and minute from time (HHMM)
            if len(time) >= 4:
                hour = time[:2]
                minute = time[2:4]
            else:
                hour = '00'
                minute = '00'
            
            # Create datetime
            dt = datetime(int(year), int(month), int(day), int(hour), int(minute))
            return dt
    except:
        pass
    
    # If parsing fails, return a very old date so it gets moved as duplicate
    return datetime(1900, 1, 1)


def find_duplicates_in_folder(class_folder):
    """
    Find duplicate test files in a single class folder.
    Returns: list of tuples (filepath, class_num, student_id, timestamp, is_duplicate)
    """
    # Dictionary to group files by (class_number, normalized_student_id)
    student_files = defaultdict(list)
    
    # Collect all test files
    for filename in os.listdir(class_folder):
        filepath = os.path.join(class_folder, filename)
        
        # Skip directories
        if not os.path.isfile(filepath):
            continue
        
        # Skip processed files
        if '_PROCESSED.txt' in filename:
            continue
        
        # Parse filename
        parsed = parse_filename(filename)
        if not parsed:
            continue
        
        class_number, student_id, timestamp, _ = parsed
        
        # Normalize student ID for comparison
        normalized_id = normalize_student_id(student_id)
        
        # Group by (class, student)
        key = (class_number, normalized_id)
        student_files[key].append({
            'filepath': filepath,
            'filename': filename,
            'class': class_number,
            'student_id': student_id,
            'timestamp': timestamp,
            'datetime': get_timestamp_for_sorting(timestamp)
        })
    
    # Find duplicates
    duplicates = []
    
    for key, files in student_files.items():
        if len(files) > 1:
            # Multiple files for same student - sort by datetime (oldest first)
            files.sort(key=lambda x: x['datetime'])
            
            # All but the newest are duplicates
            for i, file_info in enumerate(files):
                is_duplicate = (i < len(files) - 1)  # All except last (newest) are duplicates
                duplicates.append((
                    file_info['filepath'],
                    file_info['filename'],
                    file_info['class'],
                    file_info['student_id'],
                    file_info['timestamp'],
                    is_duplicate
                ))
    
    return duplicates


def move_duplicates(duplicates, duplicates_folder):
    """
    Move duplicate files to the Duplicates folder.
    Returns: count of files moved
    """
    moved_count = 0
    
    for filepath, filename, class_num, student_id, timestamp, is_duplicate in duplicates:
        if is_duplicate:
            # Create destination path
            dest_path = os.path.join(duplicates_folder, filename)
            
            # If file already exists in Duplicates, add a number
            if os.path.exists(dest_path):
                base, ext = os.path.splitext(filename)
                counter = 1
                while os.path.exists(dest_path):
                    dest_path = os.path.join(duplicates_folder, f"{base}_dup{counter}{ext}")
                    counter += 1
            
            # Move the file
            try:
                shutil.move(filepath, dest_path)
                print(f"  ✓ Moved: {filename}")
                print(f"    → Class {class_num}, Student {student_id}, {timestamp}")
                moved_count += 1
            except Exception as e:
                print(f"  ✗ Error moving {filename}: {e}")
    
    return moved_count


def main():
    """Main function to find and move duplicates"""
    print("=" * 80)
    print("Find Duplicate Speaking Tests")
    print("=" * 80)
    
    # Find Records directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    records_dir = os.path.join(script_dir, "Records")
    
    if not os.path.exists(records_dir):
        print(f"\n❌ Error: 'Records' directory not found!")
        print("   Make sure to run this script from the same directory as speaking_test.py")
        input("\nPress Enter to exit...")
        return
    
    # Create Duplicates folder
    duplicates_folder = os.path.join(records_dir, "Duplicates")
    Path(duplicates_folder).mkdir(parents=True, exist_ok=True)
    
    print(f"\nScanning for duplicates in: {records_dir}")
    print(f"Duplicates will be moved to: {duplicates_folder}\n")
    
    total_duplicates_found = 0
    total_files_moved = 0
    classes_processed = 0
    
    # Process each class folder
    for item in sorted(os.listdir(records_dir)):
        item_path = os.path.join(records_dir, item)
        
        # Skip the Duplicates folder itself
        if item == "Duplicates":
            continue
        
        # Only process directories
        if not os.path.isdir(item_path):
            continue
        
        classes_processed += 1
        print(f"Processing class: {item}")
        
        # Find duplicates in this folder
        duplicates = find_duplicates_in_folder(item_path)
        
        if duplicates:
            # Count how many are actually duplicates (not the kept newest one)
            dup_count = sum(1 for _, _, _, _, _, is_dup in duplicates if is_dup)
            
            if dup_count > 0:
                total_duplicates_found += dup_count
                
                # Show what was found
                # Group by student to show clearly
                student_groups = defaultdict(list)
                for filepath, filename, class_num, student_id, timestamp, is_dup in duplicates:
                    student_groups[student_id].append((filename, timestamp, is_dup))
                
                for student_id, files in student_groups.items():
                    print(f"\n  Student {student_id}: {len(files)} tests found")
                    for filename, timestamp, is_dup in sorted(files, key=lambda x: x[1]):
                        status = "→ MOVING (older)" if is_dup else "✓ KEEPING (newest)"
                        print(f"    {status}: {timestamp}")
                
                # Move the duplicates
                moved = move_duplicates(duplicates, duplicates_folder)
                total_files_moved += moved
                print()
        else:
            print("  No duplicates found\n")
    
    # Summary
    print("=" * 80)
    print("SUMMARY")
    print("=" * 80)
    print(f"Classes processed:    {classes_processed}")
    print(f"Duplicates found:     {total_duplicates_found}")
    print(f"Files moved:          {total_files_moved}")
    
    if total_files_moved > 0:
        print(f"\nDuplicates moved to: {duplicates_folder}")
        print("\nReview the Duplicates folder to:")
        print("  • Check for students who took test twice")
        print("  • Find accidental duplicate student numbers")
        print("  • Delete or restore files as needed")
    else:
        print("\n✓ No duplicates found - all students have unique tests!")
    
    print("=" * 80)
    
    # Wait for user
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print("\n" + "=" * 80)
        print("ERROR OCCURRED")
        print("=" * 80)
        print(f"\n{str(e)}\n")
        print("Full traceback:")
        print(traceback.format_exc())
        print("=" * 80)
        input("\nPress Enter to exit...")
