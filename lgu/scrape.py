import os
import urllib.request

SHEET_URL = "https://docs.google.com/spreadsheets/d/1B2N5SjZEBP-29tUioFHaSk-6R_1YFGcYIkSPjuApxZw/gviz/tq?tqx=out:csv&gid=1213553854"

req = urllib.request.Request(SHEET_URL, headers={"User-Agent": "Mozilla/5.0"})
print("Downloading 2025-2028 officials dataset...")

with urllib.request.urlopen(req) as response:
    content = response.read()

output_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "officials_2025_2028.csv")
os.makedirs(os.path.dirname(output_path), exist_ok=True)
with open(output_path, "wb") as f:
    f.write(content)

print(f"Done! Saved {len(content):,} bytes to {output_path}")
