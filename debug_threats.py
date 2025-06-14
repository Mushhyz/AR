"""Debug script to inspect threats.csv file for loading issues."""

import csv
from pathlib import Path


def debug_threats_csv():
    """Debug the threats.csv file to identify loading issues."""
    threats_file = Path("config/threats.csv")

    print(f"🔍 Debugging {threats_file}")
    print(f"File exists: {threats_file.exists()}")

    if not threats_file.exists():
        print("❌ threats.csv not found in config directory")
        return

    # Check raw file content
    print("\n📄 Raw file inspection:")
    with open(threats_file, "rb") as f:
        raw_content = f.read(200)
        print(f"First 200 bytes: {raw_content}")

    # Check text content with BOM handling
    print("\n📝 Text content inspection:")
    with open(threats_file, "r", encoding="utf-8-sig") as f:
        lines = f.readlines()[:5]
        for i, line in enumerate(lines):
            print(f"Line {i}: {repr(line)}")

    # Check CSV parsing
    print("\n🔍 CSV inspection:")
    try:
        with open(threats_file, "r", newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            print(f"Headers: {reader.fieldnames}")
            print(f"Header types: {[(h, type(h)) for h in reader.fieldnames]}")

            # Check for problematic headers
            if reader.fieldnames:
                for h in reader.fieldnames:
                    if h is None:
                        print("❌ Found None in headers!")
                    elif not isinstance(h, str):
                        print(f"❌ Non-string header: {h} (type: {type(h)})")
                    elif h.strip() != h:
                        print(f"⚠️  Header with whitespace: '{h}'")

            # Check first few rows
            for i, row in enumerate(reader):
                print(f"\nRow {i}:")
                print(f"  Keys: {list(row.keys())}")
                print(f"  Key types: {[(k, type(k)) for k in row.keys()]}")
                print(f"  Values: {dict(row)}")

                # Check for problematic keys
                for k in row.keys():
                    if k is None:
                        print(f"  ❌ None key found in row {i}")
                    elif not isinstance(k, str):
                        print(f"  ❌ Non-string key in row {i}: {k} (type: {type(k)})")

                if i >= 2:  # Only check first 3 rows
                    break

    except Exception as e:
        print(f"❌ CSV parsing error: {e}")
        print(f"Error type: {type(e)}")


if __name__ == "__main__":
    debug_threats_csv()
