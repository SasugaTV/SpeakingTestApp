#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Process and Compile Speaking Tests
===================================
This script does a complete processing pass on all test records:

1. Finds incomplete (unfinished) test files
2. Completes them with headers and summaries
3. Renames originals as *_Unprocessed.txt
4. Moves unprocessed files to Duplicates folder
5. Compiles class summary files with all student scores

Usage: python process_and_compile.py

This ensures all records are properly formatted and class summaries are complete.
"""

import os
import re
from datetime import datetime
from pathlib import Path
from collections import defaultdict


def parse_incomplete_test(filepath):
    """
    Parse a test file and check if it has a header.
    Returns: (questions_data, has_header)
    questions_data = [(question_num, points, binary_string, filename), ...]
    """
    questions = []
    has_header = False
    
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
    except:
        return questions, True  # Can't read, skip it
    
    # Check if file has header (look for "=" line or "Total Questions:")
    if '=' * 20 in content or 'Total Questions:' in content or 'Student:' in content:
        has_header = True
        return questions, has_header
    
    # Parse question lines
    # Format: Question 00: 5 = 1 0 0 0 0 intro
    for line in content.split('\n'):
        line = line.strip()
        if line.startswith('Question'):
            match = re.match(r'Question\s+(\d+):\s+(\d+)\s+=\s+([\d\s]+)\s+(.+)', line)
            if match:
                q_num = int(match.group(1))
                points = int(match.group(2))
                binary = match.group(3).strip()
                filename = match.group(4).strip()
                questions.append((q_num, points, binary, filename))
    
    return questions, has_header


def calculate_score(questions):
    """
    Calculate score from question data.
    Takes only the LAST entry for each question number.
    Returns: (total_score, max_score, percentage, total_questions)
    """
    # Group by question number, keep only last entry for each
    question_dict = {}
    for q_num, points, binary, filename in questions:
        question_dict[q_num] = points
    
    if not question_dict:
        return 0, 0, 0, 0
    
    # Find max point value from the questions
    all_points = [p for p in question_dict.values()]
    max_point_value = max(all_points) if all_points else 5
    
    total_questions = len(question_dict)
    total_score = sum(question_dict.values())
    max_score = total_questions * max_point_value
    
    if max_score > 0:
        percentage = round((total_score / max_score) * 100)
    else:
        percentage = 0
    
    return total_score, max_score, percentage, total_questions


def get_point_scale_from_directory(class_folder):
    """
    Get point scale from any complete test file in the directory.
    Returns dict: {position: (points, 'name'), ...}
    """
    point_scale = {}
    
    # Look for any file with a header
    for file in os.listdir(class_folder):
        if file.endswith('.txt') and file.startswith('SpeakingTest_'):
            filepath = os.path.join(class_folder, file)
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Check if it has point scale
                lines = content.split('\n')
                in_scale_section = False
                for line in lines:
                    if '=' * 20 in line and in_scale_section:
                        break
                    if '=' * 20 in line:
                        in_scale_section = True
                        continue
                    
                    if in_scale_section and '=' in line and 'Total Questions' not in line:
                        # Parse: "5 = Correct"
                        match = re.match(r'(\d+)\s*=\s*(.+)', line.strip())
                        if match:
                            points = int(match.group(1))
                            name = match.group(2).strip()
                            if not any(p == points for p in [v[0] for v in point_scale.values()]):
                                position = len(point_scale) + 1
                                point_scale[position] = (points, name)
                
                if point_scale:
                    return point_scale
            except:
                continue
    
    # Default if none found
    return {1: (5, 'Correct'), 2: (0, 'Incorrect'), 3: (3, 'Fluent'), 4: (2, 'Partial'), 5: (1, 'Attempted')}


def extract_info_from_filename(filename):
    """
    Extract class, student ID, and timestamp from filename.
    Format: SpeakingTest_CLASS_STUDENTID_TIMESTAMP.txt
    Returns: (class_number, student_id, timestamp_str)
    """
    name = filename.replace('.txt', '')
    parts = name.split('_')
    
    if len(parts) >= 4 and parts[0] == 'SpeakingTest':
        return parts[1], parts[2], parts[3]
    
    return None, None, None


def generate_complete_report(original_file, questions, class_number, student_id, 
                            timestamp_str, total_score, max_score, percentage, 
                            total_questions, point_scale):
    """
    Generate a complete report file with header.
    Returns path to new file.
    """
    # Read original content
    with open(original_file, 'r', encoding='utf-8') as f:
        original_content = f.read()
    
    # Build header
    original_filename = os.path.basename(original_file)
    header = f"{original_filename}\n"
    
    # Add student info
    if student_id:
        student_display = student_id.zfill(2) if student_id.isdigit() else student_id
        header += f"Student: {student_display}\n"
    
    header += "=" * 67 + "\n"
    header += f"Total Questions: {total_questions}   Max Score: {max_score}   Score: {total_score}   Percentage: {percentage}%\n"
    header += "=" * 67 + "\n"
    
    # Add point scale
    sorted_scale = sorted(point_scale.items(), key=lambda x: (-x[1][0], x[0]))
    for position, (points, name) in sorted_scale:
        if points != 0 or name:
            header += f"{points} = {name}\n"
    
    header += "\n"
    
    # Write complete file (same filename, original will be renamed)
    with open(original_file, 'w', encoding='utf-8') as f:
        f.write(header + original_content)
    
    return original_file


def compile_class_summary(class_folder, class_number):
    """
    Compile or update class summary file for a class.
    Returns: (summary_file_path, entries_added)
    """
    # Collect all test files and their scores
    test_entries = []
    
    for filename in os.listdir(class_folder):
        if not filename.startswith('SpeakingTest_') or not filename.endswith('.txt'):
            continue
        
        if '_Unprocessed.txt' in filename or '_PROCESSED.txt' in filename:
            continue
        
        filepath = os.path.join(class_folder, filename)
        
        # Parse file to get score
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Extract info from header
            student_id = None
            student_name = ""
            total_score = 0
            max_score = 0
            percentage = 0
            
            # Parse header
            for line in content.split('\n'):
                if line.startswith('Student:'):
                    # Format: "Student: 01" or "Student: 01 Name"
                    parts = line.replace('Student:', '').strip().split(' ', 1)
                    student_id = parts[0]
                    student_name = parts[1] if len(parts) > 1 else ""
                elif 'Total Questions:' in line and 'Score:' in line:
                    # Parse: Total Questions: 10   Max Score: 50   Score: 41   Percentage: 82%
                    score_match = re.search(r'Score:\s*(\d+)', line)
                    max_match = re.search(r'Max Score:\s*(\d+)', line)
                    pct_match = re.search(r'Percentage:\s*(\d+)%', line)
                    
                    if score_match:
                        total_score = int(score_match.group(1))
                    if max_match:
                        max_score = int(max_match.group(1))
                    if pct_match:
                        percentage = int(pct_match.group(1))
            
            # Get timestamp from filename
            _, _, timestamp_str = extract_info_from_filename(filename)
            
            if student_id and timestamp_str:
                test_entries.append({
                    'student_id': student_id,
                    'student_name': student_name,
                    'total_score': total_score,
                    'max_score': max_score,
                    'percentage': percentage,
                    'timestamp': timestamp_str
                })
        except:
            continue
    
    if not test_entries:
        return None, 0
    
    # Determine date from first entry
    if test_entries:
        first_ts = test_entries[0]['timestamp']
        try:
            parts = first_ts.split('.')
            if len(parts) >= 3:
                yy, mm, dd = parts[0], parts[1], parts[2]
                if len(yy) == 4:
                    yy = yy[-2:]
                date_str = f"{yy}.{mm}.{dd}"
            else:
                date_str = datetime.now().strftime("%y.%m.%d")
        except:
            date_str = datetime.now().strftime("%y.%m.%d")
    else:
        date_str = datetime.now().strftime("%y.%m.%d")
    
    # Create summary file
    summary_file = os.path.join(class_folder, f"{class_number}_SpeakingTest.{date_str}.txt")
    
    # Write summary
    with open(summary_file, 'w', encoding='utf-8') as f:
        f.write(f"Class {class_number} - Speaking Test Summary\n")
        f.write(f"Date: 20{date_str.replace('.', '-')}\n")
        f.write("=" * 80 + "\n\n")
        
        # Write entries
        for entry in test_entries:
            student_display = entry['student_id'].zfill(2) if entry['student_id'].isdigit() else entry['student_id'].ljust(2)
            name_display = entry['student_name'].ljust(20)[:20] if entry['student_name'] else "".ljust(20)
            score_display = f"{entry['total_score']}/{entry['max_score']}".rjust(10)
            percent_display = f"{entry['percentage']}%".rjust(5)
            
            # Extract time from timestamp
            ts_parts = entry['timestamp'].split('.')
            time_part = ts_parts[-1][:4] if len(ts_parts) >= 4 else "0000"
            
            line = f"{date_str}.{time_part}:   Student {student_display} {name_display}:  {score_display} = {percent_display}\n"
            f.write(line)
    
    return summary_file, len(test_entries)


def main():
    """Main processing function"""
    print("=" * 80)
    print("Process and Compile Speaking Tests")
    print("=" * 80)
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    records_dir = os.path.join(script_dir, "Records")
    
    if not os.path.exists(records_dir):
        print(f"\n❌ Error: 'Records' directory not found!")
        input("\nPress Enter to exit...")
        return
    
    # Create Duplicates folder
    duplicates_folder = os.path.join(records_dir, "Duplicates")
    Path(duplicates_folder).mkdir(parents=True, exist_ok=True)
    
    print(f"\nProcessing records in: {records_dir}\n")
    
    total_incomplete = 0
    total_processed = 0
    total_summaries = 0
    
    # Process each class folder
    for item in sorted(os.listdir(records_dir)):
        item_path = os.path.join(records_dir, item)
        
        if item == "Duplicates" or not os.path.isdir(item_path):
            continue
        
        print(f"Processing class: {item}")
        
        # Get point scale for this class
        point_scale = get_point_scale_from_directory(item_path)
        
        incomplete_in_class = 0
        
        # Process each test file
        for filename in sorted(os.listdir(item_path)):
            if not filename.startswith('SpeakingTest_') or not filename.endswith('.txt'):
                continue
            
            if '_Unprocessed.txt' in filename or '_PROCESSED.txt' in filename:
                continue
            
            filepath = os.path.join(item_path, filename)
            
            # Check if incomplete
            questions, has_header = parse_incomplete_test(filepath)
            
            if not has_header and questions:
                # This is an incomplete test - process it
                class_number, student_id, timestamp_str = extract_info_from_filename(filename)
                
                if not class_number or not student_id:
                    continue
                
                # Calculate score
                total_score, max_score, percentage, total_questions = calculate_score(questions)
                
                # Rename original to *_Unprocessed.txt
                unprocessed_name = filename.replace('.txt', '_Unprocessed.txt')
                unprocessed_path = os.path.join(item_path, unprocessed_name)
                os.rename(filepath, unprocessed_path)
                
                # Generate complete report (recreate with original filename)
                generate_complete_report(
                    filepath, questions, class_number, student_id, timestamp_str,
                    total_score, max_score, percentage, total_questions, point_scale
                )
                
                # Move unprocessed to Duplicates
                dest_path = os.path.join(duplicates_folder, unprocessed_name)
                if os.path.exists(dest_path):
                    base, ext = os.path.splitext(unprocessed_name)
                    counter = 1
                    while os.path.exists(dest_path):
                        dest_path = os.path.join(duplicates_folder, f"{base}_dup{counter}{ext}")
                        counter += 1
                
                import shutil
                shutil.move(unprocessed_path, dest_path)
                
                print(f"  ✓ Completed: Student {student_id} - {total_score}/{max_score} = {percentage}%")
                incomplete_in_class += 1
                total_processed += 1
        
        if incomplete_in_class > 0:
            total_incomplete += incomplete_in_class
            print(f"  → Processed {incomplete_in_class} incomplete test(s)")
        
        # Compile class summary
        summary_file, entries = compile_class_summary(item_path, item)
        
        if summary_file:
            print(f"  ✓ Compiled summary: {os.path.basename(summary_file)} ({entries} students)")
            total_summaries += 1
        
        print()
    
    # Summary
    print("=" * 80)
    print("SUMMARY")
    print("=" * 80)
    print(f"Incomplete tests found:    {total_incomplete}")
    print(f"Tests completed:           {total_processed}")
    print(f"Class summaries created:   {total_summaries}")
    
    if total_processed > 0:
        print(f"\nUnprocessed originals moved to: {duplicates_folder}")
    
    print("=" * 80)
    
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
