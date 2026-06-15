# Data Merge & Sort Summary

## Overview
Successfully merged traffic statistics from Google Analytics with the CTA content inventory, matching 1,623 URLs and creating organized output files.

## Process Completed

### 1. **Data Merge** ✓
- **Source Files:**
  - `All URLs - traffic stats-Free form 1.csv` - 2,307 traffic records (GA4 data: 2025-06-07 to 2026-06-07)
  - `cta_content_inventory_output (1).xlsx` - 1,623 inventory records

- **Matching Logic:**
  - URLs matched by extracting the path from both files (ignoring `/eng/` and `/fra/` prefixes)
  - First attempt match with English URL, then French URL
  - All 1,623 inventory items matched with traffic data

### 2. **Output Files Created**

#### Main Merged File
- **`merged_inventory_with_traffic.xlsx`** (214 KB)
  - Contains all 1,623 matched items in table format
  - Columns included: Nid, Title, English URL, French URL, Creation Date, Updated Date, Primary Section, Secondary Section, Section Confidence, Migration Status, Issue Flagged, Total Users, Views, Sessions, Content Type, Purple Excel Marker
  - Color formatting: Light purple background for items with Purple Excel Marker = "Yes"

#### Unmatched URLs (for future crawling)
- **`unmatched_traffic_urls.csv`** (232 KB)
  - 874 traffic URLs that did not match any inventory items
  - These URLs will need to be crawled/investigated later
  - Top traffic URLs in this file (by views):
    1. Flight Delays and Cancellations: A Guide - 171,543 views
    2. Air travel complaints resolution process - 101,068 views
    3. Air Passenger Protection Regulations Highlights - 67,616 views

### 3. **Separated by Primary Section** ✓

6 separate Excel files created in `separated_by_section/` folder:

| Section | Items | Total Views | Total Users |
|---------|-------|------------|------------|
| Accessible Transportation | 247 | 66,609 | 48,293 |
| Complaint and Dispute Resolution | 250 | 61,082 | 46,518 |
| Compliance Monitoring and Enforcement | 711 | 39,674 | 31,741 |
| Consultations | 52 | 15,456 | 11,550 |
| National Transportation System | 362 | 70,263 | 51,197 |
| About / Contact / Corporate | 1 | 42 | 28 |

**Organization Details:**
- Items with Primary Section = "About / Contact / Corporate" are sorted by their **Secondary Section** instead
- Secondary sections from About/Contact/Corporate are merged with their corresponding primary sections:
  - 54 items → Accessible Transportation
  - 102 items → Complaint and Dispute Resolution
  - 348 items → Compliance Monitoring and Enforcement
  - 9 items → Consultations
  - 138 items → National Transportation System
- 1 item with missing secondary section kept in About / Contact / Corporate file

Each section file includes:
- All matching columns from the inventory
- Color formatting based on Purple Excel Marker
- Adjusted column widths for readability
- Total: 1,623 items

Also includes **`SECTION_SUMMARY.xlsx`** - Summary statistics for all sections.

## Statistics

| Metric | Count |
|--------|-------|
| Total Inventory Items | 1,623 |
| Total Traffic URLs | 2,307 |
| **Successfully Matched** | **1,623 (100%)** |
| Unmatched URLs (need crawling) | 874 |
| Primary Sections (in inventory) | 6 |
| Export Files | 6 |

## Data Quality Notes

- All inventory items have been matched with traffic data
- Items categorized as "About / Contact / Corporate" are exported by their secondary section instead of as a single group
- Some URLs appear in traffic data but not in the inventory (874 URLs) - these should be evaluated for inclusion in future
- All colors from `cta_content_inventory_output` preserved in merged files
- Tables are properly formatted and ready for analysis/reporting

## Scripts Used

- **`merge_data.py`** - Merges traffic stats with content inventory
- **`separate_by_section.py`** - Separates merged data by primary section (already run; can be re-run if needed)

## Next Steps

1. Review the `unmatched_traffic_urls.csv` file for URLs that need to be crawled
2. Use `separate_by_section.py` anytime to regenerate section-specific files if the merged file is updated
3. The data is now ready for further analysis, reporting, or migration planning
