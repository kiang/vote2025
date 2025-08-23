#!/usr/bin/env python3
import pandas as pd
import glob
import os
import re

def extract_county_name(filename):
    """Extract county name from filename"""
    match = re.search(r'縣表3-(.+?)-全國性公民投票', filename)
    return match.group(1) if match else None

def process_excel_file_totals(file_path):
    """Process a single Excel file and return vote totals"""
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
        print(f"Warning: Could not find data start row in {file_path}")
        return None
    
    county_totals = {
        'county': county_name,
        'agree': 0,
        'disagree': 0,
        'valid': 0,
        'invalid': 0,
        'total': 0,
        'eligible_voters': 0,
        'polling_stations': 0
    }
    
    # Process each row to sum up totals
    for i in range(data_start_row, len(df)):
        row = df.iloc[i]
        
        # Skip empty rows or rows without polling station
        if pd.isna(row[2]):  # No polling station ID
            continue
            
        # Extract voting data
        agree_votes = row[3] if pd.notna(row[3]) else 0
        disagree_votes = row[4] if pd.notna(row[4]) else 0
        valid_votes = row[5] if pd.notna(row[5]) else 0
        invalid_votes = row[6] if pd.notna(row[6]) else 0
        total_votes = row[7] if pd.notna(row[7]) else 0
        eligible_voters = row[11] if pd.notna(row[11]) else 0
        
        # Convert to integers safely
        try:
            agree = int(float(str(agree_votes))) if pd.notna(agree_votes) else 0
            disagree = int(float(str(disagree_votes))) if pd.notna(disagree_votes) else 0
            valid = int(float(str(valid_votes))) if pd.notna(valid_votes) else 0
            invalid = int(float(str(invalid_votes))) if pd.notna(invalid_votes) else 0
            total = int(float(str(total_votes))) if pd.notna(total_votes) else 0
            eligible = int(float(str(eligible_voters))) if pd.notna(eligible_voters) else 0
        except (ValueError, TypeError):
            continue
        
        county_totals['agree'] += agree
        county_totals['disagree'] += disagree
        county_totals['valid'] += valid
        county_totals['invalid'] += invalid
        county_totals['total'] += total
        county_totals['eligible_voters'] += eligible
        county_totals['polling_stations'] += 1
    
    return county_totals

def main():
    """Main function to verify totals from raw Excel files"""
    
    print("=== VERIFYING TOTALS FROM RAW EXCEL FILES ===")
    print()
    
    excel_files = glob.glob('raw/*.xlsx')
    all_county_totals = []
    grand_totals = {
        'agree': 0,
        'disagree': 0,
        'valid': 0,
        'invalid': 0,
        'total': 0,
        'eligible_voters': 0,
        'polling_stations': 0
    }
    
    for file_path in sorted(excel_files):
        print(f"Processing: {os.path.basename(file_path)}")
        totals = process_excel_file_totals(file_path)
        
        if totals:
            all_county_totals.append(totals)
            
            # Add to grand totals
            grand_totals['agree'] += totals['agree']
            grand_totals['disagree'] += totals['disagree']
            grand_totals['valid'] += totals['valid']
            grand_totals['invalid'] += totals['invalid']
            grand_totals['total'] += totals['total']
            grand_totals['eligible_voters'] += totals['eligible_voters']
            grand_totals['polling_stations'] += totals['polling_stations']
            
            print(f"  {totals['county']}: agree={totals['agree']:,}, disagree={totals['disagree']:,}, total={totals['total']:,}")
    
    print()
    print("=== GRAND TOTALS FROM RAW FILES ===")
    print(f"同意票 (Agree):     {grand_totals['agree']:,}")
    print(f"不同意票 (Disagree): {grand_totals['disagree']:,}")
    print(f"有效票 (Valid):     {grand_totals['valid']:,}")
    print(f"無效票 (Invalid):   {grand_totals['invalid']:,}")
    print(f"總投票數 (Total):    {grand_totals['total']:,}")
    print(f"投票權人數:         {grand_totals['eligible_voters']:,}")
    print(f"投票所數量:         {grand_totals['polling_stations']:,}")
    
    agree_rate = (grand_totals['agree'] / grand_totals['valid'] * 100) if grand_totals['valid'] > 0 else 0
    disagree_rate = (grand_totals['disagree'] / grand_totals['valid'] * 100) if grand_totals['valid'] > 0 else 0
    turnout_rate = (grand_totals['total'] / grand_totals['eligible_voters'] * 100) if grand_totals['eligible_voters'] > 0 else 0
    
    print()
    print("=== RATES ===")
    print(f"同意票比例:   {agree_rate:.2f}%")
    print(f"不同意票比例: {disagree_rate:.2f}%")
    print(f"投票率:       {turnout_rate:.2f}%")
    
    # Save county breakdown
    with open('county_totals_verification.json', 'w', encoding='utf-8') as f:
        import json
        json.dump({
            'counties': all_county_totals,
            'grand_totals': grand_totals
        }, f, ensure_ascii=False, indent=2)
    
    print(f"\nSaved verification data to county_totals_verification.json")
    
    return grand_totals

if __name__ == "__main__":
    main()