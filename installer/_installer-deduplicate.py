#!/usr/bin/env python3

"""
Prepare Installer Resources with Deduplication
This script scans all year directories, identifies duplicate DLLs by checksum,
and creates a deduplicated resource directory for the installer
"""

import hashlib
import os
import shutil
import sys
from collections import defaultdict
from pathlib import Path

SOURCE_PATH = Path("../commands/bin")
OUTPUT_PATH = Path("./bin/resources")


def print_error(*args, **kwargs):
    """Print to both stdout and stderr to ensure visibility in MSBuild"""
    print(*args, **kwargs)
    print(*args, **kwargs, file=sys.stderr)


def calculate_md5(file_path):
    """Calculate MD5 hash of a file - optimized for 9p filesystem"""
    hash_md5 = hashlib.md5()
    with open(file_path, "rb") as f:
        # Read entire file at once to minimize syscalls (files are small ~1-5MB)
        hash_md5.update(f.read())
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

    # Validate that the source path exists
    if not SOURCE_PATH.exists():
        print_error()
        print_error("=" * 65)
        print_error("ERROR: Commands build directory does not exist!")
        print_error("=" * 65)
        print_error()
        print_error(f"Expected location: {SOURCE_PATH.absolute()}")
        print_error()
        print_error("You need to build the commands project first.")
        print_error("Run these commands from the repository root:")
        print_error()
        print_error("  # Build all Revit versions:")
        print_error("  for year in 2017 2018 2019 2020 2021 2022 2023 2024 2025 2026; do")
        print_error("    dotnet build commands/revit-ballet.csproj -p:RevitYear=$year")
        print_error("  done")
        print_error()
        print_error("After building commands, run the installer build again.")
        print_error("=" * 65)
        return 1

    # Clean output directory
    if OUTPUT_PATH.exists():
        shutil.rmtree(OUTPUT_PATH)
    OUTPUT_PATH.mkdir(parents=True)

    # Define expected years for Revit Ballet installer
    EXPECTED_YEARS = ['2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025', '2026']

    # Get all year directories (exclude temp)
    year_dirs = sorted([d for d in SOURCE_PATH.iterdir()
                       if d.is_dir() and d.name.isdigit() and len(d.name) == 4])

    if not year_dirs:
        print_error()
        print_error("=" * 65)
        print_error("ERROR: No year directories found in commands build output!")
        print_error("=" * 65)
        print_error()
        print_error(f"Searched in: {SOURCE_PATH.absolute()}")
        print_error()
        print_error("You need to build the commands project first.")
        print_error("Run these commands from the repository root:")
        print_error()
        print_error("  # Build all Revit versions:")
        print_error("  for year in 2017 2018 2019 2020 2021 2022 2023 2024 2025 2026; do")
        print_error("    dotnet build commands/revit-ballet.csproj -p:RevitYear=$year")
        print_error("  done")
        print_error()
        print_error("After building commands, run the installer build again.")
        print_error("=" * 65)
        return 1

    found_years = set(d.name for d in year_dirs)
    print(f"Found {len(year_dirs)} year directories: {', '.join(d.name for d in year_dirs)}")
    print()

    # Check for missing year directories
    missing_year_dirs = sorted(set(EXPECTED_YEARS) - found_years)
    if missing_year_dirs:
        print_error()
        print_error("=" * 65)
        print_error("ERROR: Missing build directories for some Revit years!")
        print_error("=" * 65)
        print_error()
        print_error(f"Expected years: {', '.join(EXPECTED_YEARS)}")
        print_error(f"Found years:    {', '.join(sorted(found_years))}")
        print_error(f"Missing years:  {', '.join(missing_year_dirs)}")
        print_error()
        print_error("You need to build the commands project for all Revit versions.")
        print_error("Run these commands from the repository root:")
        print_error()
        for year in missing_year_dirs:
            print_error(f"  dotnet build commands/revit-ballet.csproj -p:RevitYear={year}")
        print_error()
        print_error("Or build all years at once:")
        print_error("  for year in 2017 2018 2019 2020 2021 2022 2023 2024 2025 2026; do")
        print_error("    dotnet build commands/revit-ballet.csproj -p:RevitYear=$year")
        print_error("  done")
        print_error()
        print_error("After rebuilding, run this script again.")
        print_error("=" * 65)
        return 1

    # Validate that all year directories have the required revit-ballet.dll
    print("Validating year directories...")
    missing_years = []
    incomplete_years = []

    for year_dir in year_dirs:
        year = year_dir.name
        dll_files = list(year_dir.glob("*.dll"))

        if not dll_files:
            missing_years.append(year)
        elif not any(f.name == "revit-ballet.dll" for f in dll_files):
            incomplete_years.append(year)

    if missing_years or incomplete_years:
        print_error()
        print_error("=" * 65)
        print_error("ERROR: Command DLLs are missing or incomplete!")
        print_error("=" * 65)

        if missing_years:
            print_error(f"\nYears with NO DLL files: {', '.join(missing_years)}")

        if incomplete_years:
            print_error(f"\nYears missing revit-ballet.dll: {', '.join(incomplete_years)}")

        print_error("\nYou need to build the commands project for all Revit versions first.")
        print_error("Run these commands from the repository root:")
        print_error()

        if missing_years or incomplete_years:
            all_problem_years = sorted(set(missing_years + incomplete_years))
            for year in all_problem_years:
                print_error(f"  dotnet build commands/revit-ballet.csproj -p:RevitYear={year}")

        print_error()
        print_error("Or build all years at once:")
        print_error("  for year in 2017 2018 2019 2020 2021 2022 2023 2024 2025 2026; do")
        print_error("    dotnet build commands/revit-ballet.csproj -p:RevitYear=$year")
        print_error("  done")
        print_error()
        print_error("After rebuilding, run this script again.")
        print_error("=" * 65)
        return 1

    print("All year directories validated successfully.")
    print()

    # Build database: filename -> [(year, path, hash, size)]
    print("Scanning files and calculating checksums...")
    file_database = defaultdict(list)

    for year_dir in year_dirs:
        year = year_dir.name
        dll_files = list(year_dir.glob("*.dll"))

        for dll_file in dll_files:
            filename = dll_file.name
            file_size = dll_file.stat().st_size

            # Always calculate MD5 to ensure correctness
            # Files with same name and size but different contents must be detected
            file_hash = calculate_md5(dll_file)

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
    shared_file_count = 0

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
                shared_file_count += 1
            elif len(years) == 1:
                # File is unique to this year
                resource_name = f"_{years[0]}.{filename}"
            else:
                # File is shared by multiple (but not all) years
                resource_name = f"_{years[0]}.{filename}"
                saved_space += file_size * (len(years) - 1)
                print(f"  {resource_name} -> shared by years: {', '.join(years)}")
                shared_file_count += 1

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
    print(f"Found {shared_file_count} shared files across multiple years")
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
        # Use shutil.copy() instead of copy2() - metadata not needed, saves syscalls
        shutil.copy(item['source_path'], dest_path)

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
    print("Next step: Build the installer with 'dotnet build installer.csproj'")

    return 0


if __name__ == "__main__":
    exit(main())
