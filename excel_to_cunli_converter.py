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
            # Handle multi-village case - distribute data proportionally
            multi_data = multi_village_mappings[lookup_key]
            villcodes = multi_data['villcodes']
            num_villages = len(villcodes)
            
            # Distribute the polling station data proportionally across villcodes
            for i, vc in enumerate(villcodes):
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
                
                # Add the polling station data to each villcode
                matched_data[vc]['polling_stations'].append({
                    'station_id': record['polling_station'],
                    'votes': record['votes'],
                    'ballots': record['ballots'],
                    'eligible_voters': record['eligible_voters'],
                    'turnout_rate': record['turnout_rate'],
                    'shared_station': True,
                    'shared_with_villages': num_villages
                })
                
                # Add proportional totals (divide by number of villages to avoid double counting)
                vote_share = 1.0 / num_villages
                matched_data[vc]['total_votes']['agree'] += int(record['votes']['agree'] * vote_share)
                matched_data[vc]['total_votes']['disagree'] += int(record['votes']['disagree'] * vote_share)
                matched_data[vc]['total_votes']['valid'] += int(record['votes']['valid'] * vote_share)
                matched_data[vc]['total_votes']['invalid'] += int(record['votes']['invalid'] * vote_share)
                matched_data[vc]['total_votes']['total'] += int(record['votes']['total'] * vote_share)
                
                matched_data[vc]['total_ballots']['unused'] += int(record['ballots']['unused'] * vote_share)
                matched_data[vc]['total_ballots']['issued'] += int(record['ballots']['issued'] * vote_share)
                matched_data[vc]['total_ballots']['remaining'] += int(record['ballots']['remaining'] * vote_share)
                
                matched_data[vc]['total_eligible_voters'] += int(record['eligible_voters'] * vote_share)
                matched_data[vc]['station_count'] += vote_share
            
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

def verify_raw_totals(excel_files):
    """Verify totals directly from raw Excel files"""
    print("=== VERIFYING TOTALS FROM RAW FILES ===")
    
    grand_totals = {
        'agree': 0, 'disagree': 0, 'valid': 0, 'invalid': 0, 'total': 0,
        'eligible_voters': 0, 'station_count': 0, 'village_count': 0
    }
    
    unique_villages = set()
    
    for file_path in excel_files:
        county_name = extract_county_name(os.path.basename(file_path))
        df = pd.read_excel(file_path, header=None)
        
        # Find data start row
        data_start_row = None
        for i, row in df.iterrows():
            if pd.notna(row[0]) and ('總' in str(row[0]) and '計' in str(row[0])):
                data_start_row = i
                break
        
        if data_start_row is None:
            continue
        
        county_totals = {'agree': 0, 'disagree': 0, 'valid': 0, 'invalid': 0, 'total': 0, 'eligible_voters': 0, 'stations': 0}
        
        for i in range(data_start_row, len(df)):
            row = df.iloc[i]
            if pd.isna(row[2]):  # No polling station
                continue
            
            # Count unique villages
            if pd.notna(row[1]):  # Village name exists
                village_key = f"{county_name}|{row[0]}|{row[1]}"
                unique_villages.add(village_key)
                
            try:
                agree = int(float(str(row[3]))) if pd.notna(row[3]) else 0
                disagree = int(float(str(row[4]))) if pd.notna(row[4]) else 0
                valid = int(float(str(row[5]))) if pd.notna(row[5]) else 0
                invalid = int(float(str(row[6]))) if pd.notna(row[6]) else 0
                total = int(float(str(row[7]))) if pd.notna(row[7]) else 0
                eligible = int(float(str(row[11]))) if pd.notna(row[11]) else 0
                
                county_totals['agree'] += agree
                county_totals['disagree'] += disagree
                county_totals['valid'] += valid
                county_totals['invalid'] += invalid
                county_totals['total'] += total
                county_totals['eligible_voters'] += eligible
                county_totals['stations'] += 1
            except (ValueError, TypeError):
                continue
        
        grand_totals['agree'] += county_totals['agree']
        grand_totals['disagree'] += county_totals['disagree']
        grand_totals['valid'] += county_totals['valid']
        grand_totals['invalid'] += county_totals['invalid']
        grand_totals['total'] += county_totals['total']
        grand_totals['eligible_voters'] += county_totals['eligible_voters']
        grand_totals['station_count'] += county_totals['stations']
        
        print(f"{county_name}: agree={county_totals['agree']:,}, disagree={county_totals['disagree']:,}, stations={county_totals['stations']}")
    
    grand_totals['village_count'] = len(unique_villages)
    turnout_rate = (grand_totals['total'] / grand_totals['eligible_voters'] * 100) if grand_totals['eligible_voters'] > 0 else 0
    
    print(f"\nRAW FILE TOTALS:")
    print(f"村里數量:           {grand_totals['village_count']:,}")
    print(f"投票所數量:         {grand_totals['station_count']:,}")
    print(f"同意票 (Agree):     {grand_totals['agree']:,}")
    print(f"不同意票 (Disagree): {grand_totals['disagree']:,}")
    print(f"有效票:            {grand_totals['valid']:,}")
    print(f"總投票數:          {grand_totals['total']:,}")
    print(f"投票權人數:         {grand_totals['eligible_voters']:,}")
    print(f"投票率:            {turnout_rate:.2f}%")
    
    return grand_totals

def main():
    """Main function to process Excel files and generate cunli-based JSON"""
    
    # Get Excel files
    excel_files = glob.glob('raw/*.xlsx')
    
    # First verify totals from raw files
    expected_totals = verify_raw_totals(excel_files)
    
    # Process all Excel files
    print("\nProcessing Excel files for detailed data...")
    all_data = []
    
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
    
    # Check if all data is accounted for by comparing totals
    print("Verifying data completeness...")
    original_totals = {
        'agree': sum(r['votes']['agree'] for r in all_data if r['polling_station']),
        'disagree': sum(r['votes']['disagree'] for r in all_data if r['polling_station']),
        'valid': sum(r['votes']['valid'] for r in all_data if r['polling_station']),
        'total': sum(r['votes']['total'] for r in all_data if r['polling_station'])
    }
    
    matched_totals = {
        'agree': sum(v['total_votes']['agree'] for v in matched_data.values()),
        'disagree': sum(v['total_votes']['disagree'] for v in matched_data.values()),
        'valid': sum(v['total_votes']['valid'] for v in matched_data.values()),
        'total': sum(v['total_votes']['total'] for v in matched_data.values())
    }
    
    print(f"Original totals: agree={original_totals['agree']:,}, disagree={original_totals['disagree']:,}")
    print(f"Matched totals:  agree={matched_totals['agree']:,}, disagree={matched_totals['disagree']:,}")
    
    # Find truly unmatched records and add them
    unmatched_grouped = {}
    matched_stations = set()
    
    # Collect all matched station IDs
    for villcode, data in matched_data.items():
        for station in data['polling_stations']:
            station_key = f"{data['county']}|{data['district']}|{data['village']}|{station['station_id']}"
            matched_stations.add(station_key)
    
    # Find unmatched polling stations
    for data_record in all_data:
        if not data_record['village'] or not data_record['polling_station']:
            continue
            
        county = clean_field(data_record['county'])
        district = clean_field(data_record['district'])
        village = clean_field(data_record['village'])
        station_key = f"{county}|{district}|{village}|{data_record['polling_station']}"
        
        if station_key not in matched_stations:
            lookup_key = f"{county}|{district}|{village}"
            
            if lookup_key not in unmatched_grouped:
                unmatched_grouped[lookup_key] = {
                    'villcode': None,
                    'county': county,
                    'district': district,
                    'village': village,
                    'polling_stations': [],
                    'total_votes': {'agree': 0, 'disagree': 0, 'valid': 0, 'invalid': 0, 'total': 0},
                    'total_ballots': {'unused': 0, 'issued': 0, 'remaining': 0},
                    'total_eligible_voters': 0,
                    'station_count': 0
                }
            
            # Add polling station data
            unmatched_grouped[lookup_key]['polling_stations'].append({
                'station_id': data_record['polling_station'],
                'votes': data_record['votes'],
                'ballots': data_record['ballots'],
                'eligible_voters': data_record['eligible_voters'],
                'turnout_rate': data_record['turnout_rate']
            })
            
            # Aggregate totals
            unmatched_grouped[lookup_key]['total_votes']['agree'] += data_record['votes']['agree']
            unmatched_grouped[lookup_key]['total_votes']['disagree'] += data_record['votes']['disagree']
            unmatched_grouped[lookup_key]['total_votes']['valid'] += data_record['votes']['valid']
            unmatched_grouped[lookup_key]['total_votes']['invalid'] += data_record['votes']['invalid']
            unmatched_grouped[lookup_key]['total_votes']['total'] += data_record['votes']['total']
            
            unmatched_grouped[lookup_key]['total_ballots']['unused'] += data_record['ballots']['unused']
            unmatched_grouped[lookup_key]['total_ballots']['issued'] += data_record['ballots']['issued']
            unmatched_grouped[lookup_key]['total_ballots']['remaining'] += data_record['ballots']['remaining']
            
            unmatched_grouped[lookup_key]['total_eligible_voters'] += data_record['eligible_voters']
            unmatched_grouped[lookup_key]['station_count'] += 1
    
    # Calculate turnout rates for unmatched data
    for key in unmatched_grouped:
        data = unmatched_grouped[key]
        if data['total_eligible_voters'] > 0:
            data['turnout_rate'] = (data['total_votes']['total'] / data['total_eligible_voters']) * 100
        else:
            data['turnout_rate'] = 0.0
    
    unmatched_totals = {
        'agree': sum(v['total_votes']['agree'] for v in unmatched_grouped.values()),
        'disagree': sum(v['total_votes']['disagree'] for v in unmatched_grouped.values())
    }
    
    print(f"Unmatched totals: agree={unmatched_totals['agree']:,}, disagree={unmatched_totals['disagree']:,}")
    
    final_totals = {
        'agree': matched_totals['agree'] + unmatched_totals['agree'],
        'disagree': matched_totals['disagree'] + unmatched_totals['disagree']
    }
    
    print(f"Final totals:    agree={final_totals['agree']:,}, disagree={final_totals['disagree']:,}")
    print(f"Expected totals: agree={expected_totals['agree']:,}, disagree={expected_totals['disagree']:,}")
    
    # Check if totals match expected values
    agree_diff = final_totals['agree'] - expected_totals['agree']
    disagree_diff = final_totals['disagree'] - expected_totals['disagree']
    
    if agree_diff != 0 or disagree_diff != 0:
        print(f"WARNING: Totals don't match! Differences: agree={agree_diff:,}, disagree={disagree_diff:,}")
    else:
        print("✓ Totals match expected values perfectly!")
    
    # Combine matched and unmatched data
    all_cunli_data = list(matched_data.values()) + list(unmatched_grouped.values())
    
    # Create final output with verified totals at the top
    turnout_rate = (expected_totals['total'] / expected_totals['eligible_voters'] * 100) if expected_totals['eligible_voters'] > 0 else 0
    
    output_data = {
        'verified_totals': {
            'agree': expected_totals['agree'],
            'disagree': expected_totals['disagree'],
            'agree_rate': (expected_totals['agree'] / (expected_totals['agree'] + expected_totals['disagree']) * 100) if (expected_totals['agree'] + expected_totals['disagree']) > 0 else 0,
            'disagree_rate': (expected_totals['disagree'] / (expected_totals['agree'] + expected_totals['disagree']) * 100) if (expected_totals['agree'] + expected_totals['disagree']) > 0 else 0,
            '村里數量': expected_totals['village_count'],
            '投票所數量': expected_totals['station_count'],
            '有效票': expected_totals['valid'],
            '總投票數': expected_totals['total'],
            '投票權人數': expected_totals['eligible_voters'],
            '投票率': round(turnout_rate, 2),
            'note': 'Verified totals from raw CEC Excel files'
        },
        'villages': all_cunli_data
    }
    
    # Save final cunli-based data
    output_file = 'docs/referendum_cunli_data.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)
    
    print(f"Generated {output_file} with {len(all_cunli_data)} total records")
    print(f"  - {len(matched_data)} records with VILLCODE mapping")
    print(f"  - {len(unmatched_grouped)} records without VILLCODE (included for completeness)")
    print(f"Match rate: {len(matched_data)}/{len(matched_data) + len(unmatched_grouped)} villages")
    
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