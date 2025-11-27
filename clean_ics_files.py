#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
clean_ics_files.py - Clean and split ICS files from Outlook export

This script:
1. Splits multi-calendar ICS exports into individual files
2. Removes problematic X- headers that cause parsing issues
3. PRESERVES X-MICROSOFT-CDO-ALLDAYEVENT (needed for all-day detection)
4. Ensures proper VCALENDAR structure
"""

import os
import re
import sys
import argparse

# Headers to KEEP (important for proper event handling)
HEADERS_TO_KEEP = {
    "X-MICROSOFT-CDO-ALLDAYEVENT",  # Critical for all-day detection
    "X-MICROSOFT-CDO-BUSYSTATUS",   # Useful for free/busy info
}

# Headers to REMOVE (cause parsing issues or are useless)
HEADERS_TO_REMOVE_PATTERNS = [
    r"^X-MICROSOFT-EXCHANGE-",      # Exchange-specific IDs
    r"^X-MICROSOFT-DISALLOW-",      # Not needed
    r"^X-MICROSOFT-DONOTFORWARD",   # Not needed
    r"^X-MS-OLK-",                   # Outlook-specific
]

def should_remove_line(line: str) -> bool:
    """Check if a line should be removed."""
    line_upper = line.upper().strip()
    
    # Check if it's an X- header we want to keep
    for keep in HEADERS_TO_KEEP:
        if line_upper.startswith(keep):
            return False
    
    # Check if it matches removal patterns
    for pattern in HEADERS_TO_REMOVE_PATTERNS:
        if re.match(pattern, line_upper):
            return True
    
    return False

def clean_ics_content(content: str) -> str:
    """Clean ICS content, preserving important headers."""
    lines = content.split('\n')
    cleaned_lines = []
    
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # Check if this line should be removed
        if should_remove_line(line):
            # Skip this line and any continuation lines
            i += 1
            while i < len(lines) and lines[i].startswith((' ', '\t')):
                i += 1
            continue
        
        cleaned_lines.append(line)
        i += 1
    
    return '\n'.join(cleaned_lines)

def split_and_clean(source_path: str, outbox_dir: str) -> int:
    """Split and clean ICS file, return count of cleaned files."""
    if not os.path.exists(source_path):
        print(f"[WARN] Source file not found: {source_path}")
        return 0
    
    with open(source_path, "r", errors="ignore") as f:
        content = f.read()
    
    # Split by VCALENDAR boundaries
    # This handles files with multiple VCALENDAR blocks
    blocks = re.split(r'(?=BEGIN:VCALENDAR)', content)
    blocks = [b.strip() for b in blocks if b.strip() and 'BEGIN:VCALENDAR' in b]
    
    if not blocks:
        print(f"[WARN] No VCALENDAR blocks found in {source_path}")
        return 0
    
    # Remove old cleaned files
    for old_file in os.listdir(outbox_dir):
        if old_file.startswith("clean_") and old_file.endswith(".ics"):
            os.remove(os.path.join(outbox_dir, old_file))
    
    cleaned_count = 0
    all_day_count = 0
    
    for i, block in enumerate(blocks, 1):
        # Clean the block
        cleaned = clean_ics_content(block)
        
        # Ensure it ends properly
        if not cleaned.rstrip().endswith("END:VCALENDAR"):
            cleaned = cleaned.rstrip() + "\nEND:VCALENDAR"
        
        # Count all-day events in this block
        if "X-MICROSOFT-CDO-ALLDAYEVENT:TRUE" in cleaned.upper():
            all_day_count += cleaned.upper().count("X-MICROSOFT-CDO-ALLDAYEVENT:TRUE")
        
        # Write cleaned file
        target = os.path.join(outbox_dir, f"clean_{i}.ics")
        with open(target, "w") as out:
            out.write(cleaned)
        cleaned_count += 1
    
    print(f"[INFO] Split into {cleaned_count} cleaned ICS files")
    print(f"[INFO] Preserved {all_day_count} all-day event markers")
    
    return cleaned_count

def main():
    parser = argparse.ArgumentParser(description="Clean and split ICS files")
    parser.add_argument("--inbox", default=os.path.expanduser("~/calendarBridge/outbox"),
                       help="Directory containing ICS files")
    args = parser.parse_args()
    
    source = os.path.join(args.inbox, "outlook_full_export.ics")
    count = split_and_clean(source, args.inbox)
    
    if count == 0:
        print("[ERROR] No files processed")
        sys.exit(1)

if __name__ == "__main__":
    main()
