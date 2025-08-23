#!/usr/bin/env python3
import pandas as pd
import json
import glob
import os
import re

def clean_field(field):
    """Clean and trim field names for matching"""
    if not field:
        return ""
    cleaned = str(field).strip()
    cleaned = re.sub(r'^\s+', '', cleaned)
    return cleaned

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

def load_manual_mappings():
    """Load manual VILLCODE mappings"""
    if not os.path.exists('manual_villcode_mapping.json'):
        return {}, {}
        
    with open('manual_villcode_mapping.json', 'r', encoding='utf-8') as f:
        manual_data = json.load(f)
    
    mappings = {}
    multi_village_mappings = {}
    
    for entry in manual_data:
        if not entry['suggested_villcode']:
            continue
            
        county = entry['county']
        district = entry['district'] 
        village = entry['village']
        villcode = entry['suggested_villcode']
        
        key = f"{county}|{district}|{village}"
        
        # Check if this is a multi-village case (contains 、)
        if '、' in village:
            # Split village names and villcodes
            village_names = [v.strip() for v in village.split('、')]
            villcodes = [v.strip() for v in villcode.split(',')]
            
            if len(village_names) == len(villcodes):
                multi_village_mappings[key] = {
                    'villcodes': villcodes,
                    'villages': village_names
                }
            else:
                print(f"Warning: Mismatch in village/villcode count for {key}")
                print(f"  Villages: {village_names}")
                print(f"  VILLCODEs: {villcodes}")
        else:
            mappings[key] = villcode
    
    return mappings, multi_village_mappings

def match_with_geo_data(referendum_data, geo_data, manual_mappings, multi_village_mappings):
    """Match referendum data with geo data and manual mappings"""
    
    # Build lookup dictionary from geo data
    geo_lookup = {}
    for feature in geo_data['features']:
        props = feature['properties']
        county = props['COUNTYNAME']
        town = props['TOWNNAME']
        village = props['VILLNAME']
        villcode = props['VILLCODE']
        
        key = f"{county}|{town}|{village}"
        geo_lookup[key] = villcode
    
    print(f"Built geo lookup with {len(geo_lookup)} entries")
    print(f"Manual mappings: {len(manual_mappings)} single, {len(multi_village_mappings)} multi-village")
    
    # Process referendum data
    matched_data = {}
    unmatched_records = []
    
    for record in referendum_data:
        # Skip total/summary records
        if not record['village'] or not record['polling_station']:
            continue
            
        county = clean_field(record['county'])
        district = clean_field(record['district'])
        village = clean_field(record['village'])
        
        lookup_key = f"{county}|{district}|{village}"
        villcode = None
        
        # Try manual mappings first
        if lookup_key in manual_mappings:
            villcode = manual_mappings[lookup_key]
        elif lookup_key in multi_village_mappings:
            # Handle multi-village case - distribute data across all villcodes
            multi_data = multi_village_mappings[lookup_key]
            villcodes = multi_data['villcodes']
            
            # Distribute the polling station data equally across all villcodes
            for vc in villcodes:
                if vc not in matched_data:
                    matched_data[vc] = {
                        'villcode': vc,
                        'county': county,
                        'district': district,
                        'village': village,
                        'polling_stations': [],
                        'total_votes': {
                            'agree': 0,
                            'disagree': 0,
                            'valid': 0,
                            'invalid': 0,
                            'total': 0
                        },
                        'total_ballots': {
                            'unused': 0,
                            'issued': 0,
                            'remaining': 0
                        },
                        'total_eligible_voters': 0,
                        'station_count': 0
                    }
                
                # Add the full polling station data to each villcode
                matched_data[vc]['polling_stations'].append({
                    'station_id': record['polling_station'],
                    'votes': record['votes'],
                    'ballots': record['ballots'],
                    'eligible_voters': record['eligible_voters'],
                    'turnout_rate': record['turnout_rate'],
                    'shared_station': True,
                    'shared_with_villages': len(villcodes)
                })
                
                # Add full totals to each villcode (will be the same for all)
                matched_data[vc]['total_votes']['agree'] += record['votes']['agree']
                matched_data[vc]['total_votes']['disagree'] += record['votes']['disagree']
                matched_data[vc]['total_votes']['valid'] += record['votes']['valid']
                matched_data[vc]['total_votes']['invalid'] += record['votes']['invalid']
                matched_data[vc]['total_votes']['total'] += record['votes']['total']
                
                matched_data[vc]['total_ballots']['unused'] += record['ballots']['unused']
                matched_data[vc]['total_ballots']['issued'] += record['ballots']['issued']
                matched_data[vc]['total_ballots']['remaining'] += record['ballots']['remaining']
                
                matched_data[vc]['total_eligible_voters'] += record['eligible_voters']
                matched_data[vc]['station_count'] += 1
            
            continue  # Skip the normal processing for this record
        else:
            # Try geo lookup
            villcode = geo_lookup.get(lookup_key)
            
            if not villcode:
                # Try alternative matching strategies
                district_clean = re.sub(r'^(市|縣|區|鄉|鎮)', '', district)
                lookup_key_alt = f"{county}|{district_clean}|{village}"
                villcode = geo_lookup.get(lookup_key_alt)
            
            if not villcode:
                # Try village variants
                village_variants = [
                    village,
                    village + '里' if not village.endswith('里') else village,
                    village + '村' if not village.endswith('村') else village,
                    village.replace('里', '').replace('村', '') + '里',
                    village.replace('里', '').replace('村', '') + '村'
                ]
                
                for variant in village_variants:
                    lookup_key_var = f"{county}|{district}|{variant}"
                    villcode = geo_lookup.get(lookup_key_var)
                    if villcode:
                        break
        
        if villcode:
            # Aggregate data by VILLCODE
            if villcode not in matched_data:
                matched_data[villcode] = {
                    'villcode': villcode,
                    'county': county,
                    'district': district,
                    'village': village,
                    'polling_stations': [],
                    'total_votes': {
                        'agree': 0,
                        'disagree': 0,
                        'valid': 0,
                        'invalid': 0,
                        'total': 0
                    },
                    'total_ballots': {
                        'unused': 0,
                        'issued': 0,
                        'remaining': 0
                    },
                    'total_eligible_voters': 0,
                    'station_count': 0
                }
            
            # Add polling station data
            matched_data[villcode]['polling_stations'].append({
                'station_id': record['polling_station'],
                'votes': record['votes'],
                'ballots': record['ballots'],
                'eligible_voters': record['eligible_voters'],
                'turnout_rate': record['turnout_rate']
            })
            
            # Aggregate totals
            matched_data[villcode]['total_votes']['agree'] += record['votes']['agree']
            matched_data[villcode]['total_votes']['disagree'] += record['votes']['disagree']
            matched_data[villcode]['total_votes']['valid'] += record['votes']['valid']
            matched_data[villcode]['total_votes']['invalid'] += record['votes']['invalid']
            matched_data[villcode]['total_votes']['total'] += record['votes']['total']
            
            matched_data[villcode]['total_ballots']['unused'] += record['ballots']['unused']
            matched_data[villcode]['total_ballots']['issued'] += record['ballots']['issued']
            matched_data[villcode]['total_ballots']['remaining'] += record['ballots']['remaining']
            
            matched_data[villcode]['total_eligible_voters'] += record['eligible_voters']
            matched_data[villcode]['station_count'] += 1
            
        else:
            unmatched_records.append({
                'county': county,
                'district': district,
                'village': village,
                'lookup_key': lookup_key
            })
    
    # Calculate turnout rates for aggregated data
    for villcode in matched_data:
        data = matched_data[villcode]
        if data['total_eligible_voters'] > 0:
            data['turnout_rate'] = (data['total_votes']['total'] / data['total_eligible_voters']) * 100
        else:
            data['turnout_rate'] = 0.0
    
    return matched_data, unmatched_records

def main():
    """Main function to process Excel files and generate cunli-based JSON"""
    
    # Process all Excel files
    print("Processing Excel files...")
    all_data = []
    excel_files = glob.glob('raw/*.xlsx')
    
    for file_path in excel_files:
        print(f"Processing: {file_path}")
        data = process_excel_file(file_path)
        if data:
            all_data.extend(data)
    
    print(f"Processed {len(all_data)} total records from {len(excel_files)} files")
    
    # Load geo data
    print("Loading geo data...")
    geo_file = '/home/kiang/public_html/taiwan_basecode/cunli/s_geo/20250620.json'
    with open(geo_file, 'r', encoding='utf-8') as f:
        geo_data = json.load(f)
    
    # Load manual mappings if available
    print("Loading manual mappings...")
    manual_mappings, multi_village_mappings = load_manual_mappings()
    
    # Match and aggregate data
    print("Matching and aggregating data...")
    matched_data, unmatched_records = match_with_geo_data(
        all_data, geo_data, manual_mappings, multi_village_mappings
    )
    
    # Convert to list format for JSON output
    cunli_data = list(matched_data.values())
    
    # Save final cunli-based data
    output_file = 'referendum_cunli_data.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(cunli_data, f, ensure_ascii=False, indent=2)
    
    print(f"Generated {output_file} with {len(cunli_data)} cunli records")
    print(f"Match rate: {len(matched_data)}/{len(matched_data) + len(set(r['lookup_key'] for r in unmatched_records))} villages")
    
    # Save unmatched records if any
    if unmatched_records:
        # Get unique unmatched villages
        unique_unmatched = {}
        for record in unmatched_records:
            key = record['lookup_key']
            if key not in unique_unmatched:
                unique_unmatched[key] = {
                    'county': record['county'],
                    'district': record['district'],
                    'village': record['village'],
                    'suggested_villcode': '',
                    'notes': 'Needs manual VILLCODE mapping'
                }
        
        unmatched_file = 'unmatched_for_mapping.json'
        with open(unmatched_file, 'w', encoding='utf-8') as f:
            json.dump(list(unique_unmatched.values()), f, ensure_ascii=False, indent=2)
        
        print(f"Saved {len(unique_unmatched)} unique unmatched villages to {unmatched_file}")
        print("Please add VILLCODEs to manual_villcode_mapping.json and re-run the script")

if __name__ == "__main__":
    main()