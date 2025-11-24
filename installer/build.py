#!/usr/bin/env python3

"""
Prepare Installer Resources with Deduplication
This script scans all year directories, identifies duplicate DLLs by checksum,
and creates a deduplicated resource directory for the installer
"""

import hashlib
import os
import shutil
from collections import defaultdict
from pathlib import Path

SOURCE_PATH = Path("../commands/bin")
OUTPUT_PATH = Path("./resources")


def calculate_md5(file_path):
    """Calculate MD5 hash of a file"""
    hash_md5 = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()


def format_size(size_bytes):
    """Format bytes as human-readable size"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.2f} TB"


def main():
    print("Revit Ballet Installer - Resource Preparation with Deduplication")
    print("=" * 65)
    print()

    # Clean output directory
    if OUTPUT_PATH.exists():
        shutil.rmtree(OUTPUT_PATH)
    OUTPUT_PATH.mkdir(parents=True)

    # Get all year directories
    year_dirs = sorted([d for d in SOURCE_PATH.iterdir()
                       if d.is_dir() and d.name.isdigit() and len(d.name) == 4])

    if not year_dirs:
        print(f"ERROR: No year directories found in {SOURCE_PATH}")
        return 1

    print(f"Found {len(year_dirs)} year directories: {', '.join(d.name for d in year_dirs)}")
    print()

    # Build database: filename -> [(year, path, hash, size)]
    print("Scanning files and calculating checksums...")
    file_database = defaultdict(list)

    for year_dir in year_dirs:
        year = year_dir.name
        dll_files = list(year_dir.glob("*.dll"))

        for dll_file in dll_files:
            filename = dll_file.name
            file_hash = calculate_md5(dll_file)
            file_size = dll_file.stat().st_size

            file_database[filename].append({
                'year': year,
                'path': dll_file,
                'hash': file_hash,
                'size': file_size
            })

    print()
    print("Analyzing files for deduplication...")
    print()

    # Analyze deduplication opportunities
    deduplication_plan = []
    total_original_size = 0
    total_deduplicated_size = 0
    saved_space = 0

    for filename in sorted(file_database.keys()):
        entries = file_database[filename]

        # Group by hash
        hash_groups = defaultdict(list)
        for entry in entries:
            hash_groups[entry['hash']].append(entry)

        for file_hash, group_entries in hash_groups.items():
            years = [entry['year'] for entry in group_entries]
            sample_entry = group_entries[0]
            file_size = sample_entry['size']

            # Calculate space usage
            total_original_size += file_size * len(years)

            # Determine resource name
            if len(years) == len(year_dirs):
                # File is identical across ALL years - use _common prefix
                resource_name = f"_common.{filename}"
                saved_space += file_size * (len(years) - 1)
                print(f"  {resource_name} -> shared by ALL years")
            elif len(years) == 1:
                # File is unique to this year
                resource_name = f"_{years[0]}.{filename}"
            else:
                # File is shared by multiple (but not all) years
                resource_name = f"_{years[0]}.{filename}"
                saved_space += file_size * (len(years) - 1)
                print(f"  {resource_name} -> shared by years: {', '.join(years)}")

            total_deduplicated_size += file_size

            deduplication_plan.append({
                'filename': filename,
                'resource_name': resource_name,
                'hash': file_hash,
                'years': years,
                'source_path': sample_entry['path'],
                'size': file_size
            })

    # Display summary
    print()
    print("Deduplication Analysis:")
    print(f"  Total original size:      {format_size(total_original_size)}")
    print(f"  Deduplicated size:        {format_size(total_deduplicated_size)}")
    saved_pct = (saved_space / total_original_size * 100) if total_original_size > 0 else 0
    print(f"  Space saved:              {format_size(saved_space)} ({saved_pct:.1f}%)")
    print()

    # Copy deduplicated files
    print(f"Copying deduplicated files to {OUTPUT_PATH}...")

    for item in deduplication_plan:
        dest_path = OUTPUT_PATH / item['resource_name']
        shutil.copy2(item['source_path'], dest_path)

    # Generate mapping file
    mapping_path = OUTPUT_PATH / "file-mapping.txt"
    with open(mapping_path, 'w', encoding='utf-8') as f:
        f.write("# Revit Ballet Installer - File Mapping\n")
        f.write("# Format: ResourceName|TargetFileName|Years (comma-separated)\n")
        f.write("# This file is used by the installer to know which files to extract for each year\n")
        f.write("\n")

        for item in sorted(deduplication_plan, key=lambda x: x['resource_name']):
            years_str = ','.join(item['years'])
            f.write(f"{item['resource_name']}|{item['filename']}|{years_str}\n")

    print()
    print("Deduplication complete!")
    print(f"  Resources directory: {OUTPUT_PATH}")
    print(f"  Mapping file: {mapping_path}")
    print(f"  Total files: {len(deduplication_plan)}")
    print()
    print("Next step: Build the installer with 'dotnet build -c Release installer.csproj'")

    return 0


if __name__ == "__main__":
    exit(main())
