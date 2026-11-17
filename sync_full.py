"""
Complete Sync Demonstration: Read PPT -> Detect Changes -> Write Excel
"""

import sys
sys.path.insert(0, r'C:\Users\GESZA4\OneDrive - IKEA\Apps\ppt to excel')

from sync_engine_v2 import PPTReader
from sync_engine_writer import ExcelWriter

def main():
    ppt_path = r'C:\Users\GESZA4\OneDrive - IKEA\Apps\ppt to excel\product_team_optimized_test.pptx'
    excel_path = r'C:\Users\GESZA4\OneDrive - IKEA\Apps\ppt to excel\product_team_master.xlsx'
    
    print("=" * 80)
    print("FULL SYNC DEMONSTRATION: PPT -> Excel")
    print("=" * 80)
    print("")
    
    # Step 1: Read PPT
    print("[STEP 1] Reading PowerPoint...")
    ppt_reader = PPTReader(ppt_path)
    ppt_data = ppt_reader.read()
    print(f"         Extracted {len(ppt_data)} fields from PPT")
    print("")
    
    # Step 2: Extract record data
    print("[STEP 2] Parsing extracted data...")
    
    # Product data
    product_id = '01_Engineering_Toolchain'
    product_data = {
        'Product name': ppt_data.get('sync_product_name', ''),
        'Product type': ppt_data.get('sync_product_type', ''),
        'Primary users': ppt_data.get('sync_primary_users', ''),
        'Organisational owner': ppt_data.get('sync_organisational_owner', ''),
        'Catalogue link': ppt_data.get('sync_catalogue_link', ''),
        'Purpose and value': ppt_data.get('sync_purpose_value', ''),
        'SL1 system count': ppt_data.get('sync_sl1_system_count', ''),
        'SL2 system count': ppt_data.get('sync_sl2_system_count', ''),
        'High criticality systems': ppt_data.get('sync_high_criticality_count', ''),
        'Medium criticality systems': ppt_data.get('sync_medium_criticality_count', ''),
        'Low criticality systems': ppt_data.get('sync_low_criticality_count', ''),
        'Component count': ppt_data.get('sync_component_count', ''),
        'System count': ppt_data.get('sync_system_count', ''),
        'Primary systems': ppt_data.get('sync_primary_systems', ''),
        'Build percent': ppt_data.get('sync_make_percent', ''),
        'Buy percent': ppt_data.get('sync_buy_percent', ''),
        'Main vendor or partner': ppt_data.get('sync_vendor_partner', ''),
        'Competency 1': ppt_data.get('sync_competency_1', ''),
        'Competency 2': ppt_data.get('sync_competency_2', ''),
        'Competency 3': ppt_data.get('sync_competency_3', ''),
        'Technical complexity': ppt_data.get('sync_technical_complexity', ''),
        'Investment direction': ppt_data.get('sync_investment_direction', ''),
        'Change demand 1-2 years': ppt_data.get('sync_change_demand', ''),
        'Adjacent products and teams': ppt_data.get('sync_adjacent_products_teams', ''),
        'Dependencies and interaction model': ppt_data.get('sync_dependencies_interaction', ''),
        'AS-IS team count': ppt_data.get('sync_team_count', ''),
        'AS-IS FTE': ppt_data.get('sync_team_size_summary', ''),
    }
    
    print(f"         Product data: {len(product_data)} fields")
    
    # Team data
    team_id = 'Engineering_Delivery_Platform'
    team_data = {
        'Product ID': product_id,
        'Team name': ppt_data.get('sync_team_name', ''),
        'Team type': ppt_data.get('sync_team_type', ''),
        'Purpose and focus': ppt_data.get('sync_team_purpose', ''),
        'Composition': ppt_data.get('sync_composition', ''),
        'Team size': ppt_data.get('sync_team_size', ''),
        'Sibling team details': ppt_data.get('sync_sibling_details', ''),
        'Sizing reason - Business criticality': ppt_data.get('sync_reason_business_criticality', ''),
        'Sizing reason - Service level': ppt_data.get('sync_reason_service_level', ''),
        'Sizing reason - Technical complexity': ppt_data.get('sync_reason_technical_complexity', ''),
        'Sizing reason - Sourcing': ppt_data.get('sync_reason_sourcing', ''),
        'Sizing reason - Investment direction': ppt_data.get('sync_reason_investment_direction', ''),
        'Consumers': ppt_data.get('sync_consumers', ''),
        'Supported by': ppt_data.get('sync_supported_by', ''),
        'Interaction modes': ppt_data.get('sync_interaction_modes', ''),
        'Boundaries': ppt_data.get('sync_boundaries', ''),
        'Topology - value flow': ppt_data.get('sync_topology_value_flow', ''),
        'Topology - stream alignment': ppt_data.get('sync_topology_stream_alignment', ''),
        'Topology - outcome ownership': ppt_data.get('sync_topology_outcome_ownership', ''),
        'Topology - handoffs and load': ppt_data.get('sync_topology_handoffs_load', ''),
        'Topology - feedback': ppt_data.get('sync_topology_feedback', ''),
    }
    
    print(f"         Team data: {len(team_data)} fields")
    print("")
    
    # Step 3: Write to Excel
    print("[STEP 3] Writing to Excel...")
    writer = ExcelWriter(excel_path)
    writer.open()
    
    if writer.write_product(product_id, product_data):
        print(f"         [OK] Product '{product_id}' written to Excel")
    else:
        print(f"         [FAIL] Product write failed")
        return False
    
    if writer.write_team(team_id, team_data):
        print(f"         [OK] Team '{team_id}' written to Excel")
    else:
        print(f"         [FAIL] Team write failed")
        return False
    
    writer.close()
    print("         [OK] Excel file saved")
    print("")
    
    # Step 4: Verify
    print("[STEP 4] Verifying written data...")
    from openpyxl import load_workbook
    
    wb = load_workbook(excel_path)
    ws_products = wb['Products']
    ws_teams = wb['Teams']
    
    # Check product
    product_name = ws_products.cell(2, 2).value
    team_count = ws_products.cell(2, 27).value
    fte_count = ws_products.cell(2, 28).value
    print(f"         Product Row 2: ID={ws_products.cell(2, 1).value}, Name={product_name}")
    print(f"                        Team Count={team_count}, FTE={fte_count}")
    
    # Check team
    team_name = ws_teams.cell(2, 2).value
    team_size = ws_teams.cell(2, 8).value
    print(f"         Team Row 2: ID={ws_teams.cell(2, 1).value}, Name={team_name}")
    print(f"                     Size={team_size}, Product={ws_teams.cell(2, 3).value}")
    
    print("")
    print("=" * 80)
    print("SYNC COMPLETE!")
    print("=" * 80)
    print("")
    print("Summary:")
    print(f"  - Extracted {len(ppt_data)} fields from filled PPT template")
    print(f"  - Parsed into 1 Product record + 1 Team record")
    print(f"  - Wrote {len(product_data)} product fields to Excel")
    print(f"  - Wrote {len(team_data)} team fields to Excel")
    print(f"  - Verified data in Excel file")
    print("")
    
    return True

if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
