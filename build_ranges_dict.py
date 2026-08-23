"""
Build ranges dictionary for batch_update_cells MCP tool
"""

import json

with open("all_chunks.json", encoding="utf-8") as f:
    chunks = json.load(f)

ranges_dict = {}
for ch in chunks:
    ranges_dict[ch["range"]] = ch["data"]

print(f"Total ranges in batch update: {len(ranges_dict)}")
for k, v in ranges_dict.items():
    print(f"  Range: {k:<15} -> {len(v)} rows")

with open("ranges_dict.json", "w", encoding="utf-8") as f:
    json.dump(ranges_dict, f)
