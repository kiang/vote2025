#!/usr/bin/env python3
import json
import re

def clean_field(field):
    """Clean and trim field names for matching"""
    if not field:
        return ""
    # Remove whitespace and common prefixes/suffixes
    cleaned = str(field).strip()
    # Remove leading spaces/indentation
    cleaned = re.sub(r'^\s+', '', cleaned)
    # Remove common suffixes like 區, 市, 縣, 鄉, 鎮, 里, 村
    return cleaned

def match_with_geo_data(referendum_data, geo_data):
    """Match referendum data with geo data to find VILLCODE"""
    
    # Build lookup dictionary from geo data
    geo_lookup = {}
    for feature in geo_data['features']:
        props = feature['properties']
        county = props['COUNTYNAME']
        town = props['TOWNNAME']
        village = props['VILLNAME']
        villcode = props['VILLCODE']
        
        # Create lookup key
        key = f"{county}|{town}|{village}"
        geo_lookup[key] = villcode
    
    print(f"Built geo lookup with {len(geo_lookup)} entries")
    
    # Process referendum data and match with geo data
    matched_data = {}
    unmatched_records = []
    
    for record in referendum_data:
        # Skip total/summary records (those without polling station)
        if not record['village'] or not record['polling_station']:
            continue
            
        county = clean_field(record['county'])
        district = clean_field(record['district'])
        village = clean_field(record['village'])
        
        # Try exact match first
        lookup_key = f"{county}|{district}|{village}"
        villcode = geo_lookup.get(lookup_key)
        
        if not villcode:
            # Try alternative matching strategies
            # Remove common prefixes from district
            district_clean = re.sub(r'^(市|縣|區|鄉|鎮)', '', district)
            lookup_key_alt = f"{county}|{district_clean}|{village}"
            villcode = geo_lookup.get(lookup_key_alt)
        
        if not villcode:
            # Try with different village suffixes
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
                'lookup_key': lookup_key,
                'original_record': record
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
    """Main function to process and match data"""
    
    # Load referendum data
    print("Loading referendum data...")
    with open('referendum_data.json', 'r', encoding='utf-8') as f:
        referendum_data = json.load(f)
    
    # Load geo data
    print("Loading geo data...")
    with open('/home/kiang/public_html/taiwan_basecode/cunli/s_geo/20250620.json', 'r', encoding='utf-8') as f:
        geo_data = json.load(f)
    
    # Match and aggregate data
    print("Matching and aggregating data...")
    matched_data, unmatched_records = match_with_geo_data(referendum_data, geo_data)
    
    # Convert to list format for JSON output
    cunli_data = list(matched_data.values())
    
    # Save cunli-based data
    output_file = 'referendum_cunli_data.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(cunli_data, f, ensure_ascii=False, indent=2)
    
    print(f"Generated {output_file} with {len(cunli_data)} cunli records")
    print(f"Matched {len(matched_data)} villages")
    print(f"Unmatched records: {len(unmatched_records)}")
    
    # Save unmatched records for debugging
    if unmatched_records:
        unmatched_file = 'unmatched_records.json'
        with open(unmatched_file, 'w', encoding='utf-8') as f:
            json.dump(unmatched_records, f, ensure_ascii=False, indent=2)
        print(f"Saved {len(unmatched_records)} unmatched records to {unmatched_file}")
        
        # Show sample unmatched records
        print("\nSample unmatched records:")
        for i, record in enumerate(unmatched_records[:5]):
            print(f"  {i+1}. {record['county']} | {record['district']} | {record['village']}")

if __name__ == "__main__":
    main()