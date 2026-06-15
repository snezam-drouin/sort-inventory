#!/usr/bin/env python3
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import re
from urllib.parse import urlparse

# Read CSV file (traffic stats)
csv_file = 'All URLs - traffic stats-Free form 1.csv'
df_traffic = pd.read_csv(csv_file, skiprows=6)

# Read Excel file (content inventory)
excel_file = 'cta_content_inventory_output (1).xlsx'
df_inventory = pd.read_excel(excel_file, sheet_name='Inventory')

print("CSV columns:", df_traffic.columns.tolist())
print("CSV shape:", df_traffic.shape)
print("\nExcel columns:", df_inventory.columns.tolist())
print("Excel shape:", df_inventory.shape)

# Clean up CSV column names (remove extra spaces)
df_traffic.columns = df_traffic.columns.str.strip()
print("\nCleaned CSV columns:", df_traffic.columns.tolist())

# Function to extract the path from URL for matching
def extract_path(url):
    """Extract path from URL for matching"""
    if pd.isna(url):
        return None
    try:
        parsed = urlparse(str(url))
        path = parsed.path
        # Remove /eng/ or /fra/ prefix for matching
        path = re.sub(r'^/(eng|fra)/', '/', path)
        return path.lower()
    except:
        return None

# Add path columns for matching
df_traffic['url_path'] = df_traffic['Page location'].apply(extract_path)
df_inventory['url_path_en'] = df_inventory['English URL'].apply(extract_path)
df_inventory['url_path_fr'] = df_inventory['French URL'].apply(extract_path)

print("\n\nMatching URLs...")

# Create a dictionary for fast lookup of traffic data
traffic_dict = {}
for traffic_idx, traffic_row in df_traffic.iterrows():
    traffic_path = traffic_row['url_path']
    if traffic_path:
        traffic_dict[traffic_path] = (traffic_idx, traffic_row)

# Create merged dataset
matched_rows = []
matched_traffic_indices = set()

for idx, inv_row in df_inventory.iterrows():
    path_en = inv_row['url_path_en']
    path_fr = inv_row['url_path_fr']
    
    # Find matching traffic data
    traffic_match = None
    traffic_idx_match = None
    
    # Try English URL first
    if path_en in traffic_dict:
        traffic_idx_match, traffic_match = traffic_dict[path_en]
    # Then try French URL
    elif path_fr in traffic_dict:
        traffic_idx_match, traffic_match = traffic_dict[path_fr]
    
    if traffic_idx_match is not None:
        matched_traffic_indices.add(traffic_idx_match)
    
    # Add matched row to results
    merged_row = {
        'Nid': inv_row['Nid'],
        'Title': inv_row['Title'],
        'English URL': inv_row['English URL'],
        'French URL': inv_row['French URL'],
        'Creation Date': inv_row['Created Date'],
        'Updated Date': inv_row['Updated Date'],
        'Primary Section': inv_row['Primary Section'],
        'Secondary Section': inv_row['Secondary Section'],
        'Section Confidence': inv_row['Section Confidence'],
        'Migration Status': inv_row['Migration Status'],
        'Issue Flagged': inv_row['Issue Flag'],
        'Total Users': traffic_match['Total users'] if traffic_match is not None else None,
        'Views': traffic_match['Views'] if traffic_match is not None else None,
        'Sessions': traffic_match['Sessions'] if traffic_match is not None else None,
        'Content Type': inv_row['Content Type'],
        'Purple Excel Marker': inv_row['Purple Excel Marker'],
    }
    
    matched_rows.append(merged_row)

df_merged = pd.DataFrame(matched_rows)

# Find unmatched traffic URLs
unmatched_traffic_rows = []
for traffic_idx, traffic_row in df_traffic.iterrows():
    if traffic_idx not in matched_traffic_indices:
        unmatched_traffic_rows.append(traffic_row.to_dict())

df_unmatched = pd.DataFrame(unmatched_traffic_rows)

print(f"\nMatched rows: {len(df_merged)}")
print(f"Unmatched traffic rows: {len(df_unmatched)}")

# Export matched data to Excel with formatting
output_file_matched = 'merged_inventory_with_traffic.xlsx'
output_file_unmatched = 'unmatched_traffic_urls.csv'

# Write matched data
with pd.ExcelWriter(output_file_matched, engine='openpyxl') as writer:
    df_merged.to_excel(writer, sheet_name='Matched Content', index=False)
    
    # Add color formatting based on "Purple Excel Marker"
    workbook = writer.book
    worksheet = writer.sheets['Matched Content']
    
    # Color mapping (you can adjust these)
    colors = {
        'No': 'FFFFFF',      # White
        'Yes': 'C7B8F0',     # Light purple
    }
    
    for idx, row in enumerate(df_merged.iterrows(), 2):
        color_marker = row[1]['Purple Excel Marker']
        if color_marker in colors:
            fill = PatternFill(start_color=colors[color_marker], end_color=colors[color_marker], fill_type='solid')
            for cell in worksheet[idx]:
                cell.fill = fill

print(f"\nMatched data exported to: {output_file_matched}")

# Write unmatched traffic URLs to CSV
df_unmatched.to_csv(output_file_unmatched, index=False)
print(f"Unmatched URLs exported to: {output_file_unmatched}")

# Print summary
print("\n=== SUMMARY ===")
print(f"Total inventory items: {len(df_inventory)}")
print(f"Total traffic URLs: {len(df_traffic)}")
print(f"Matched: {len(df_merged)}")
print(f"Unmatched URLs (to crawl later): {len(df_unmatched)}")

print("\n=== First 10 Matched Rows ===")
print(df_merged.head(10))

print("\n=== First 10 Unmatched URLs ===")
print(df_unmatched[['Page location', 'Page title', 'Total users', 'Views', 'Sessions']].head(10))
