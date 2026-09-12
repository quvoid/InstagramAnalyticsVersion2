"""
Save chunk JSON files for MCP updates
"""

import json

with open("all_chunks.json", encoding="utf-8") as f:
    chunks = json.load(f)

for i, ch in enumerate(chunks, 1):
    with open(f"chunk_{i}.json", "w", encoding="utf-8") as f_out:
        json.dump(ch["data"], f_out)
    print(f"Chunk {i}: Range {ch['range']} ({len(ch['data'])} rows)")
