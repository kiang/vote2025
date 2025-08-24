#!/usr/bin/env python3
"""
Compare 2025 referendum data with 2021 referendum data (case 17)
Match by villcode and output comparison results
"""

import json
import pandas as pd

def load_2025_data():
    """Load 2025 referendum data"""
    with open('docs/referendum_cunli_data.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data['villages']

def load_2021_data():
    """Load 2021 referendum data (case 17)"""
    with open('/home/kiang/public_html/referendums2021/docs/data.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # Create villcode lookup for case 17
    villcode_lookup = {}
    for villcode, village_data in data.items():
        if '17_agree' in village_data and '17_disagree' in village_data:
            villcode_lookup[villcode] = {
                'agree_2021': village_data['17_agree'],
                'disagree_2021': village_data['17_disagree']
            }
    
    return villcode_lookup

def compare_data():
    """Compare 2025 and 2021 referendum data"""
    print("Loading 2025 referendum data...")
    villages_2025 = load_2025_data()
    
    print("Loading 2021 referendum data...")
    data_2021 = load_2021_data()
    
    comparison_results = []
    matched_count = 0
    unmatched_count = 0
    
    for village in villages_2025:
        villcode = village.get('villcode')
        
        result = {
            'villcode': villcode,
            'county': village.get('county', ''),
            'district': village.get('district', ''),
            'village': village.get('village', ''),
            'agree_2025': village['total_votes']['agree'],
            'disagree_2025': village['total_votes']['disagree'],
            'agree_2021': None,
            'disagree_2021': None,
            'agree_diff': None,
            'disagree_diff': None,
            'agree_pct_change': None,
            'disagree_pct_change': None
        }
        
        # Try to find matching 2021 data
        if villcode and villcode in data_2021:
            matched_count += 1
            result['agree_2021'] = data_2021[villcode]['agree_2021']
            result['disagree_2021'] = data_2021[villcode]['disagree_2021']
            
            # Calculate differences
            result['agree_diff'] = result['agree_2025'] - result['agree_2021']
            result['disagree_diff'] = result['disagree_2025'] - result['disagree_2021']
            
            # Calculate percentage changes
            if result['agree_2021'] > 0:
                result['agree_pct_change'] = ((result['agree_2025'] - result['agree_2021']) / result['agree_2021']) * 100
            if result['disagree_2021'] > 0:
                result['disagree_pct_change'] = ((result['disagree_2025'] - result['disagree_2021']) / result['disagree_2021']) * 100
        else:
            unmatched_count += 1
        
        comparison_results.append(result)
    
    # Convert to DataFrame for easier analysis
    df = pd.DataFrame(comparison_results)
    
    # Save to CSV for detailed analysis
    df.to_csv('referendum_comparison_2025_vs_2021.csv', index=False, encoding='utf-8')
    
    # Print summary statistics
    print("\n=== COMPARISON SUMMARY ===")
    print(f"Total villages in 2025 data: {len(villages_2025)}")
    print(f"Villages matched with 2021 data: {matched_count}")
    print(f"Villages not found in 2021 data: {unmatched_count}")
    print(f"Match rate: {(matched_count / len(villages_2025)) * 100:.2f}%")
    
    # Calculate aggregate totals for matched villages
    matched_df = df[df['agree_2021'].notna()]
    if not matched_df.empty:
        total_agree_2025 = matched_df['agree_2025'].sum()
        total_disagree_2025 = matched_df['disagree_2025'].sum()
        total_agree_2021 = matched_df['agree_2021'].sum()
        total_disagree_2021 = matched_df['disagree_2021'].sum()
        
        print(f"\n=== AGGREGATE TOTALS (Matched Villages Only) ===")
        print(f"2025 - Agree: {total_agree_2025:,}, Disagree: {total_disagree_2025:,}")
        print(f"2021 - Agree: {int(total_agree_2021):,}, Disagree: {int(total_disagree_2021):,}")
        print(f"Change - Agree: {int(total_agree_2025 - total_agree_2021):,} ({((total_agree_2025 - total_agree_2021) / total_agree_2021) * 100:+.2f}%)")
        print(f"Change - Disagree: {int(total_disagree_2025 - total_disagree_2021):,} ({((total_disagree_2025 - total_disagree_2021) / total_disagree_2021) * 100:+.2f}%)")
        
        # Find villages with largest changes
        print(f"\n=== TOP 10 VILLAGES WITH LARGEST AGREE VOTE INCREASES ===")
        top_agree_increase = matched_df.nlargest(10, 'agree_diff')[['county', 'district', 'village', 'agree_2025', 'agree_2021', 'agree_diff', 'agree_pct_change']]
        for _, row in top_agree_increase.iterrows():
            print(f"{row['county']} {row['district']} {row['village']}: "
                  f"{int(row['agree_2025'])} vs {int(row['agree_2021'])} "
                  f"(+{int(row['agree_diff'])}, {row['agree_pct_change']:+.1f}%)")
        
        print(f"\n=== TOP 10 VILLAGES WITH LARGEST DISAGREE VOTE INCREASES ===")
        top_disagree_increase = matched_df.nlargest(10, 'disagree_diff')[['county', 'district', 'village', 'disagree_2025', 'disagree_2021', 'disagree_diff', 'disagree_pct_change']]
        for _, row in top_disagree_increase.iterrows():
            print(f"{row['county']} {row['district']} {row['village']}: "
                  f"{int(row['disagree_2025'])} vs {int(row['disagree_2021'])} "
                  f"(+{int(row['disagree_diff'])}, {row['disagree_pct_change']:+.1f}%)")
    
    print(f"\nResults saved to: referendum_comparison_2025_vs_2021.csv")

if __name__ == "__main__":
    compare_data()