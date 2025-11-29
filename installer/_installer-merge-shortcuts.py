#!/usr/bin/env python3
"""
Merge KeyboardShortcuts-custom.xml into default shortcuts for each Revit version.
Creates temporary merged files in keyboard-shortcuts-merged/ directory without modifying originals.
"""

import os
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
import shutil

def normalize_command_id(cmd_id):
    """Normalize CommandId for case-insensitive comparison"""
    return cmd_id.lower() if cmd_id else ""

def parse_shortcuts(shortcuts_attr):
    """Parse the Shortcuts attribute into individual shortcut keys"""
    if not shortcuts_attr:
        return []
    # Split by # to get individual shortcuts
    return [s.strip() for s in shortcuts_attr.split('#') if s.strip()]

def main():
    script_dir = Path(__file__).parent
    shortcuts_dir = script_dir / "KeyboardShortcuts"
    merged_dir = script_dir / "bin" / "keyboard-shortcuts-merged"
    custom_file = shortcuts_dir / "KeyboardShortcuts-custom.xml"
    addin_file = script_dir / "revit-ballet.addin"

    # Check if files exist
    if not custom_file.exists():
        print(f"Error: {custom_file} not found")
        return 1

    if not addin_file.exists():
        print(f"Error: {addin_file} not found")
        return 1

    # Create or clean merged directory
    if merged_dir.exists():
        shutil.rmtree(merged_dir)
    merged_dir.mkdir(parents=True)

    # Parse addin file to get UUIDs
    addin_tree = ET.parse(addin_file)
    addin_root = addin_tree.getroot()
    addin_uuids = set()
    for addin_id in addin_root.findall(".//AddInId"):
        if addin_id.text:
            addin_uuids.add(normalize_command_id(addin_id.text.strip()))

    # Parse custom shortcuts file
    with open(custom_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # Parse the XML
    if not content.strip().startswith('<?xml'):
        custom_xml = '<?xml version="1.0" encoding="utf-8"?>\n<Shortcuts>\n' + content + '\n</Shortcuts>'
    else:
        lines = content.split('\n')
        # Keep only the root and shortcut items
        if '<Shortcuts>' not in content:
            filtered_lines = [line for line in lines if not line.strip().startswith('<?xml')]
            custom_xml = '<?xml version="1.0" encoding="utf-8"?>\n<Shortcuts>\n' + '\n'.join(filtered_lines) + '\n</Shortcuts>'
        else:
            custom_xml = content

    custom_root = ET.fromstring(custom_xml)
    custom_shortcuts = list(custom_root.findall('ShortcutItem'))

    print(f"Merging {len(custom_shortcuts)} custom shortcuts into base files...")

    # Process each year
    years = [2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026]

    for year in years:
        base_file = shortcuts_dir / f"KeyboardShortcuts-{year}.xml"

        if not base_file.exists():
            print(f"Warning: {base_file} not found, skipping")
            continue

        # Parse base file
        tree = ET.parse(base_file)
        root = tree.getroot()

        # Create a map of existing shortcuts by normalized CommandId
        existing_shortcuts = {}
        for shortcut in root.findall('ShortcutItem'):
            cmd_id = shortcut.get('CommandId')
            if cmd_id:
                existing_shortcuts[normalize_command_id(cmd_id)] = shortcut

        # Build a map of shortcuts from custom file
        custom_shortcut_map = {}  # shortcut_key -> CommandId
        for custom_shortcut in custom_shortcuts:
            cmd_id = custom_shortcut.get('CommandId')
            shortcuts_attr = custom_shortcut.get('Shortcuts')
            if cmd_id and shortcuts_attr:
                shortcut_keys = parse_shortcuts(shortcuts_attr)
                for key in shortcut_keys:
                    custom_shortcut_map[key] = cmd_id

        # Filter base shortcuts to remove conflicts with custom shortcuts
        filtered_shortcuts = []
        conflict_count = 0

        for shortcut in root.findall('ShortcutItem'):
            cmd_id = shortcut.get('CommandId')
            shortcuts_attr = shortcut.get('Shortcuts')

            if not cmd_id or not shortcuts_attr:
                filtered_shortcuts.append(shortcut)
                continue

            # Check each shortcut key for conflicts
            shortcut_keys = parse_shortcuts(shortcuts_attr)
            non_conflicting_keys = []

            for key in shortcut_keys:
                if key in custom_shortcut_map:
                    # Conflict found - custom takes precedence
                    if custom_shortcut_map[key] != normalize_command_id(cmd_id):
                        conflict_count += 1
                        print(f"    Conflict: '{key}' used in base for '{cmd_id}' but overridden by custom for '{custom_shortcut_map[key]}'")
                    # Skip this key (whether same command or different)
                else:
                    non_conflicting_keys.append(key)

            # Only keep this shortcut if it has non-conflicting keys
            if non_conflicting_keys:
                # Update the Shortcuts attribute to only include non-conflicting keys
                shortcut.set('Shortcuts', '#'.join(non_conflicting_keys))
                filtered_shortcuts.append(shortcut)

        # Clear root and add filtered shortcuts
        for shortcut in root.findall('ShortcutItem'):
            root.remove(shortcut)

        for shortcut in filtered_shortcuts:
            root.append(shortcut)

        # Merge custom shortcuts
        added_count = 0
        updated_count = 0

        for custom_shortcut in custom_shortcuts:
            cmd_id = custom_shortcut.get('CommandId')
            if not cmd_id:
                continue

            normalized_id = normalize_command_id(cmd_id)

            # Check if this command already exists in base
            if normalized_id in existing_shortcuts:
                # This means we're updating/replacing
                updated_count += 1
            else:
                added_count += 1

            # Add the custom shortcut
            new_shortcut = ET.SubElement(root, 'ShortcutItem')
            for attr_name, attr_value in custom_shortcut.attrib.items():
                new_shortcut.set(attr_name, attr_value)

        # Write to merged directory (not original)
        merged_file = merged_dir / f"KeyboardShortcuts-{year}.xml"

        # Format output with proper indentation
        try:
            ET.indent(tree, space="  ")
        except AttributeError:
            # Fallback for older Python versions
            pass

        tree.write(merged_file, encoding='utf-8', xml_declaration=True)

        conflict_msg = f", {conflict_count} conflicts resolved" if conflict_count > 0 else ""
        print(f"  Revit {year}: +{added_count} new, ~{updated_count} updated{conflict_msg} -> {merged_file.name}")

    print(f"\nMerged files written to: {merged_dir}")
    return 0

if __name__ == '__main__':
    sys.exit(main())
