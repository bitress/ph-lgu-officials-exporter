import urllib.request

SHEET_URL = "https://docs.google.com/spreadsheets/d/1B2N5SjZEBP-29tUioFHaSk-6R_1YFGcYIkSPjuApxZw/gviz/tq?tqx=out:csv&gid=1213553854"

req = urllib.request.Request(SHEET_URL, headers={"User-Agent": "Mozilla/5.0"})
print("Downloading 2025-2028 officials dataset...")

with urllib.request.urlopen(req) as response:
    content = response.read()

output_path = "officials_2025_2028.csv"
with open(output_path, "wb") as f:
    f.write(content)

print(f"Done! Saved {len(content):,} bytes to {output_path}")
