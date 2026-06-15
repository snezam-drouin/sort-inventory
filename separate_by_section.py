#!/usr/bin/env python3
import pandas as pd
from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill

# Read the merged file
merged_file = 'merged_inventory_with_traffic.xlsx'
df = pd.read_excel(merged_file)

print("=== SEPARATING BY PRIMARY SECTION ===\n")

# Create output directory
output_dir = Path('separated_by_section')
output_dir.mkdir(exist_ok=True)

# Get unique sections sorted alphabetically
sections = sorted(df['Primary Section'].unique())

# Separate and export each section
for section in sections:
    section_df = df[df['Primary Section'] == section].copy()
    
    # Create filename from section name (sanitized for file system)
    safe_section_name = section.replace('/', '_').replace(' ', '_')
    output_file = output_dir / f"{safe_section_name}.xlsx"
    
    # Sanitize sheet name (Excel doesn't allow /, \, *, ?, [, ])
    safe_sheet_name = section.replace('/', '-').replace('\\', '-').replace('*', '-').replace('?', '-').replace('[', '-').replace(']', '-')[:31]
    
    # Write to Excel with formatting
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        section_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
        
        # Add color formatting based on "Purple Excel Marker"
        workbook = writer.book
        worksheet = writer.sheets[safe_sheet_name]
        
        # Color mapping
        colors = {
            'No': 'FFFFFF',      # White
            'Yes': 'C7B8F0',     # Light purple
        }
        
        for idx, row in enumerate(section_df.iterrows(), 2):
            color_marker = row[1]['Purple Excel Marker']
            if color_marker in colors:
                fill = PatternFill(start_color=colors[color_marker], end_color=colors[color_marker], fill_type='solid')
                for cell in worksheet[idx]:
                    cell.fill = fill
        
        # Adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column = list(column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
    
    print(f"✓ {section}")
    print(f"  File: {output_file}")
    print(f"  Items: {len(section_df)}")
    print(f"  Total Views: {section_df['Views'].sum():,.0f}")
    print(f"  Total Users: {section_df['Total Users'].sum():,.0f}\n")

print(f"\n=== SUMMARY ===")
print(f"Separated files created in: {output_dir}")
print(f"Total sections: {len(sections)}")

# Also create a summary file
summary_data = []
for section in sections:
    section_df = df[df['Primary Section'] == section]
    summary_data.append({
        'Primary Section': section,
        'Item Count': len(section_df),
        'Total Views': section_df['Views'].sum(),
        'Total Users': section_df['Total Users'].sum(),
        'Total Sessions': section_df['Sessions'].sum(),
    })

df_summary = pd.DataFrame(summary_data)
df_summary.to_excel(output_dir / 'SECTION_SUMMARY.xlsx', index=False)
print(f"\nSummary file created: separated_by_section/SECTION_SUMMARY.xlsx")
print("\n" + df_summary.to_string(index=False))
