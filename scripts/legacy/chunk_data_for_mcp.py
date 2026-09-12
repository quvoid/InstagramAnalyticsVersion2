"""
Generate chunk payloads for batch updating Google Sheets via MCP tool
"""

import csv, json

with open("All_Brands_Paid_Collabs_With_Toggle.csv", encoding="utf-8-sig") as f:
    rows = list(csv.reader(f))

# Header is row 0 in CSV
header = rows[0]
data_rows = rows[1:]

print(f"Total data rows: {len(data_rows)}")

# Let's chunk data rows into chunks of 200
CHUNK_SIZE = 200
chunks = []
for i in range(0, len(data_rows), CHUNK_SIZE):
    chunk = data_rows[i:i+CHUNK_SIZE]
    start_row = i + 3 # row 1 is title (or row 2 is header), data starts at row 3
    end_row = start_row + len(chunk) - 1
    range_str = f"A{start_row}:L{end_row}"
    chunks.append({
        "range": range_str,
        "start_row": start_row,
        "end_row": end_row,
        "data": chunk
    })

print(f"Created {len(chunks)} chunks:")
for i, ch in enumerate(chunks, 1):
    print(f"  Chunk {i}: Range {ch['range']} ({len(ch['data'])} rows)")

with open("all_chunks.json", "w", encoding="utf-8") as f:
    json.dump(chunks, f)
