#!/usr/bin/env python3
import pandas as pd
import json
import glob
import os
import re

def extract_county_name(filename):
    """Extract county name from filename"""
    match = re.search(r'縣表3-(.+?)-全國性公民投票', filename)
    return match.group(1) if match else None

def process_excel_file(file_path):
    """Process a single Excel file and return structured data"""
    county_name = extract_county_name(os.path.basename(file_path))
    
    # Read Excel file without headers
    df = pd.read_excel(file_path, header=None)
    
    # Find the data start row (after headers)
    data_start_row = None
    for i, row in df.iterrows():
        if pd.notna(row[0]) and ('總' in str(row[0]) and '計' in str(row[0])):
            data_start_row = i
            break
    
    if data_start_row is None:
        return None
    
    # Extract data rows
    data_rows = []
    current_district = None
    current_village = None
    
    for i in range(data_start_row, len(df)):
        row = df.iloc[i]
        
        # Skip empty rows
        if pd.isna(row[0]) and pd.isna(row[1]) and pd.isna(row[2]):
            continue
            
        district = row[0] if pd.notna(row[0]) else current_district
        village = row[1] if pd.notna(row[1]) else current_village
        polling_station = row[2] if pd.notna(row[2]) else None
        
        # Update current district and village
        if pd.notna(row[0]):
            current_district = district
        if pd.notna(row[1]):
            current_village = village
            
        # Extract voting data
        agree_votes = row[3] if pd.notna(row[3]) else 0
        disagree_votes = row[4] if pd.notna(row[4]) else 0
        valid_votes = row[5] if pd.notna(row[5]) else 0
        invalid_votes = row[6] if pd.notna(row[6]) else 0
        total_votes = row[7] if pd.notna(row[7]) else 0
        unused_ballots = row[8] if pd.notna(row[8]) else 0
        issued_ballots = row[9] if pd.notna(row[9]) else 0
        remaining_ballots = row[10] if pd.notna(row[10]) else 0
        eligible_voters = row[11] if pd.notna(row[11]) else 0
        turnout_rate = row[12] if pd.notna(row[12]) else 0
        
        data_row = {
            'county': county_name,
            'district': district,
            'village': village,
            'polling_station': polling_station,
            'votes': {
                'agree': int(agree_votes) if pd.notna(agree_votes) and str(agree_votes).replace('.', '').isdigit() else 0,
                'disagree': int(disagree_votes) if pd.notna(disagree_votes) and str(disagree_votes).replace('.', '').isdigit() else 0,
                'valid': int(valid_votes) if pd.notna(valid_votes) and str(valid_votes).replace('.', '').isdigit() else 0,
                'invalid': int(invalid_votes) if pd.notna(invalid_votes) and str(invalid_votes).replace('.', '').isdigit() else 0,
                'total': int(total_votes) if pd.notna(total_votes) and str(total_votes).replace('.', '').isdigit() else 0
            },
            'ballots': {
                'unused': int(unused_ballots) if pd.notna(unused_ballots) and str(unused_ballots).replace('.', '').isdigit() else 0,
                'issued': int(issued_ballots) if pd.notna(issued_ballots) and str(issued_ballots).replace('.', '').isdigit() else 0,
                'remaining': int(remaining_ballots) if pd.notna(remaining_ballots) and str(remaining_ballots).replace('.', '').isdigit() else 0
            },
            'eligible_voters': int(eligible_voters) if pd.notna(eligible_voters) and str(eligible_voters).replace('.', '').isdigit() else 0,
            'turnout_rate': float(turnout_rate) if pd.notna(turnout_rate) and str(turnout_rate).replace('.', '').replace('%', '').isdigit() else 0.0
        }
        
        data_rows.append(data_row)
    
    return data_rows

def main():
    """Main function to process all Excel files and generate JSON"""
    all_data = []
    
    # Process all Excel files in raw directory
    excel_files = glob.glob('raw/*.xlsx')
    
    for file_path in excel_files:
        print(f"Processing: {file_path}")
        data = process_excel_file(file_path)
        if data:
            all_data.extend(data)
    
    # Save to JSON file
    output_file = 'referendum_data.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)
    
    print(f"Generated {output_file} with {len(all_data)} records")
    
    # Also create a summary by county
    county_summary = {}
    for record in all_data:
        county = record['county']
        if county not in county_summary:
            county_summary[county] = {
                'total_agree': 0,
                'total_disagree': 0,
                'total_valid': 0,
                'total_invalid': 0,
                'total_votes': 0,
                'total_eligible_voters': 0,
                'polling_stations': 0
            }
        
        county_summary[county]['total_agree'] += record['votes']['agree']
        county_summary[county]['total_disagree'] += record['votes']['disagree']
        county_summary[county]['total_valid'] += record['votes']['valid']
        county_summary[county]['total_invalid'] += record['votes']['invalid']
        county_summary[county]['total_votes'] += record['votes']['total']
        county_summary[county]['total_eligible_voters'] += record['eligible_voters']
        county_summary[county]['polling_stations'] += 1
    
    # Calculate turnout rates for summary
    for county in county_summary:
        if county_summary[county]['total_eligible_voters'] > 0:
            county_summary[county]['turnout_rate'] = (
                county_summary[county]['total_votes'] / 
                county_summary[county]['total_eligible_voters'] * 100
            )
        else:
            county_summary[county]['turnout_rate'] = 0.0
    
    # Save county summary
    summary_file = 'referendum_summary.json'
    with open(summary_file, 'w', encoding='utf-8') as f:
        json.dump(county_summary, f, ensure_ascii=False, indent=2)
    
    print(f"Generated {summary_file} with summary by county")

if __name__ == "__main__":
    main()