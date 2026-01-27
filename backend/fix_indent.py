#!/usr/bin/env python3
"""Fix indentation of lines 7456-8019 by adding 4 spaces"""

with open('app.py', 'r') as f:
    lines = f.readlines()

print(f"Total lines: {len(lines)}")

# Fix lines 7455-8018 (0-indexed) by adding 4 spaces
fixed_count = 0
for i in range(7455, min(8019, len(lines))):
    line = lines[i]
    if line.strip():  # Not empty
        lines[i] = '    ' + line
        fixed_count += 1

print(f"Fixed {fixed_count} lines")

with open('app.py', 'w') as f:
    f.writelines(lines)

print("Indentation fixed and saved")
