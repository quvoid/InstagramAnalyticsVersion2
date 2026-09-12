"""
Prepare 2 halves for update_cells
"""

import json

with open("all_chunks.json", encoding="utf-8") as f:
    chunks = json.load(f)

# Half 1: chunk 1, 2, 3
half1_data = chunks[0]["data"] + chunks[1]["data"] + chunks[2]["data"]
with open("half1_data.json", "w", encoding="utf-8") as f:
    json.dump(half1_data, f)
print(f"Half 1: Range A3:L{3 + len(half1_data) - 1} ({len(half1_data)} rows)")

# Half 2: chunk 4, 5, 6
half2_data = chunks[3]["data"] + chunks[4]["data"] + chunks[5]["data"]
with open("half2_data.json", "w", encoding="utf-8") as f:
    json.dump(half2_data, f)
print(f"Half 2: Range A{3 + len(half1_data)}:L{3 + len(half1_data) + len(half2_data) - 1} ({len(half2_data)} rows)")
