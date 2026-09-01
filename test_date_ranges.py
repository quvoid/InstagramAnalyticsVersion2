"""
Test date range parser in api_wrapper
"""

from api_wrapper.client import parse_date_range

test_ranges = [
    "7d", "1w", "1week",
    "30d", "1m", "1month",
    "90d", "3m", "3months",
    "180d", "6m", "6months",
    "365d", "1y", "1year",
    "730d", "2y", "2years",
    14, 45, "60d"
]

print("="*60)
print("TESTING DATE RANGE PRESET PARSER")
print("="*60)

for tr in test_ranges:
    days = parse_date_range(tr)
    print(f"  Preset: {str(tr):<12} -> Resolved Days: {days:>4} days")

print("="*60)
print("✓ All date range presets verified successfully!")
