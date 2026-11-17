"""
PPT-to-Excel Sync Engine - Core Module
Reads PPT answer fields and Excel data, detects changes, proposes updates
"""

import os
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass
from enum import Enum

# Field mappings: ppt_shape_name -> (excel_column, table, data_type)
PPT_FIELD_MAP = {
    # Product fields - Slide 1
    'sync_product_name': ('Product name', 'Products', 'text'),
    'sync_product_type': ('Product type', 'Products', 'text'),
    'sync_primary_users': ('Primary users', 'Products', 'text'),
    'sync_organisational_owner': ('Organisational owner', 'Products', 'text'),
    'sync_catalogue_link': ('Catalogue link', 'Products', 'text'),
    'sync_purpose_value': ('Purpose and value', 'Products', 'text'),
    
    'sync_sl1_system_count': ('SL1 system count', 'Products', 'integer'),
    'sync_sl2_system_count': ('SL2 system count', 'Products', 'integer'),
    'sync_high_criticality_count': ('High criticality systems', 'Products', 'integer'),
    'sync_medium_criticality_count': ('Medium criticality systems', 'Products', 'integer'),
    'sync_low_criticality_count': ('Low criticality systems', 'Products', 'integer'),
    
    'sync_component_count': ('Component count', 'Products', 'integer'),
    'sync_system_count': ('System count', 'Products', 'integer'),
    'sync_primary_systems': ('Primary systems', 'Products', 'text'),
    'sync_make_percent': ('Build percent', 'Products', 'text'),
    'sync_buy_percent': ('Buy percent', 'Products', 'text'),
    'sync_vendor_partner': ('Main vendor or partner', 'Products', 'text'),
    'sync_competency_1': ('Competency 1', 'Products', 'text'),
    'sync_competency_2': ('Competency 2', 'Products', 'text'),
    'sync_competency_3': ('Competency 3', 'Products', 'text'),
    
    'sync_technical_complexity': ('Technical complexity', 'Products', 'text'),
    'sync_investment_direction': ('Investment direction', 'Products', 'text'),
    'sync_change_demand': ('Change demand 1-2 years', 'Products', 'text'),
    
    'sync_adjacent_products_teams': ('Adjacent products and teams', 'Products', 'text'),
    'sync_dependencies_interaction': ('Dependencies and interaction model', 'Products', 'text'),
    
    'sync_team_count': ('AS-IS team count', 'Products', 'integer'),
    'sync_team_size_summary': ('AS-IS FTE', 'Products', 'integer'),
    
    # Team fields - Slide 2+
    'sync_team_name': ('Team name', 'Teams', 'text'),
    'sync_team_purpose': ('Purpose and focus', 'Teams', 'text'),
    'sync_team_type': ('Team type', 'Teams', 'text'),
    'sync_composition': ('Composition', 'Teams', 'text'),
    
    'sync_team_size': ('Team size', 'Teams', 'text'),
    'sync_sibling_details': ('Sibling team details', 'Teams', 'text'),
    
    'sync_reason_business_criticality': ('Sizing reason - Business criticality', 'Teams', 'boolean'),
    'sync_reason_service_level': ('Sizing reason - Service level', 'Teams', 'boolean'),
    'sync_reason_technical_complexity': ('Sizing reason - Technical complexity', 'Teams', 'boolean'),
    'sync_reason_sourcing': ('Sizing reason - Sourcing', 'Teams', 'boolean'),
    'sync_reason_investment_direction': ('Sizing reason - Investment direction', 'Teams', 'boolean'),
    
    'sync_consumers': ('Consumers', 'Teams', 'text'),
    'sync_supported_by': ('Supported by', 'Teams', 'text'),
    'sync_interaction_modes': ('Interaction modes', 'Teams', 'text'),
    'sync_boundaries': ('Boundaries', 'Teams', 'text'),
    
    'sync_topology_value_flow': ('Topology - value flow', 'Teams', 'boolean'),
    'sync_topology_stream_alignment': ('Topology - stream alignment', 'Teams', 'boolean'),
    'sync_topology_outcome_ownership': ('Topology - outcome ownership', 'Teams', 'boolean'),
    'sync_topology_handoffs_load': ('Topology - handoffs and load', 'Teams', 'boolean'),
    'sync_topology_feedback': ('Topology - feedback', 'Teams', 'boolean'),
}


class ChangeType(Enum):
    NEW = "new"
    MODIFIED = "modified"
    DELETED = "deleted"
    CONFLICT = "conflict"
    MATCHED = "matched"


@dataclass
class FieldValue:
    field_key: str
    ppt_value: Optional[str]
    excel_value: Optional[str]
    data_type: str
    
    @property
    def status(self) -> ChangeType:
        """Determine if this field changed between PPT and Excel"""
        ppt_val = str(self.ppt_value or "").strip()
        excel_val = str(self.excel_value or "").strip()
        
        if not ppt_val and not excel_val:
            return ChangeType.MATCHED
        if ppt_val == excel_val:
            return ChangeType.MATCHED
        if ppt_val and not excel_val:
            return ChangeType.NEW
        if not ppt_val and excel_val:
            return ChangeType.DELETED
        return ChangeType.MODIFIED


@dataclass
class Record:
    """Represents a Product or Team record"""
    record_id: str
    record_name: str
    record_type: str  # 'Product' or 'Team'
    fields: Dict[str, FieldValue]
    
    @property
    def has_changes(self) -> bool:
        return any(f.status != ChangeType.MATCHED for f in self.fields.values())
    
    @property
    def change_summary(self) -> Dict[str, int]:
        return {
            'new': len([f for f in self.fields.values() if f.status == ChangeType.NEW]),
            'modified': len([f for f in self.fields.values() if f.status == ChangeType.MODIFIED]),
            'deleted': len([f for f in self.fields.values() if f.status == ChangeType.DELETED]),
            'matched': len([f for f in self.fields.values() if f.status == ChangeType.MATCHED]),
        }


class PPTReader:
    """Extract data from PowerPoint template"""
    
    def __init__(self, ppt_path: str):
        self.ppt_path = ppt_path
        self.data = {}
        
    def read(self) -> Dict[str, Any]:
        """Extract all named answer fields from PPT"""
        try:
            from pptx import Presentation
            prs = Presentation(self.ppt_path)
            
            extracted = {}
            
            for slide_idx, slide in enumerate(prs.slides, 1):
                for shape in slide.shapes:
                    if hasattr(shape, 'name') and shape.name.startswith('sync_'):
                        if hasattr(shape, 'text'):
                            text_val = shape.text.strip()
                            # Normalize text: remove empty brackets, clean whitespace
                            text_val = text_val.replace('[ ]', '').replace('[', '').replace(']', '').strip()
                            extracted[shape.name] = text_val
            
            self.data = extracted
            return extracted
            
        except ImportError:
            print("Error: python-pptx not installed. Install with: pip install python-pptx")
            return {}
        except Exception as e:
            print(f"Error reading PPT: {e}")
            return {}


class ExcelReader:
    """Extract data from Excel workbook"""
    
    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.data = {'Products': {}, 'Teams': {}}
        self.ws_products = None
        self.ws_teams = None
        
    def read(self) -> Dict[str, Dict[str, Any]]:
        """Extract all data from Products and Teams tables"""
        try:
            from openpyxl import load_workbook
            wb = load_workbook(self.excel_path, data_only=True)
            
            # Read Products
            self.ws_products = wb['Products']
            if 'Products' in self.ws_products.tables:
                products = self._read_table(self.ws_products, 'Products')
                self.data['Products'] = products
            
            # Read Teams
            self.ws_teams = wb['Teams']
            if 'Teams' in self.ws_teams.tables:
                teams = self._read_table(self.ws_teams, 'Teams')
                self.data['Teams'] = teams
            
            return self.data
            
        except ImportError:
            print("Error: openpyxl not installed. Install with: pip install openpyxl")
            return {}
        except Exception as e:
            print(f"Error reading Excel: {e}")
            return {}
    
    def _read_table(self, sheet, table_name: str) -> Dict[str, Dict[str, Any]]:
        """Extract all rows from a table"""
        table = sheet.tables[table_name]
        rows = {}
        
        # Get headers
        headers = []
        for col_idx in range(1, 100):
            cell_val = sheet.cell(1, col_idx).value
            if not cell_val:
                break
            headers.append(str(cell_val))
        
        # Read data rows
        for row_idx in range(2, 1000):
            id_val = sheet.cell(row_idx, 1).value
            if not id_val:
                break
            
            row_data = {'_row': row_idx}
            for col_idx, header in enumerate(headers, 1):
                cell_val = sheet.cell(row_idx, col_idx).value
                row_data[header] = str(cell_val) if cell_val else ""
            
            rows[str(id_val)] = row_data
        
        return rows


class ChangeDetector:
    """Detect and report changes between PPT and Excel"""
    
    def __init__(self, ppt_data: Dict, excel_data: Dict):
        self.ppt_data = ppt_data
        self.excel_data = excel_data
        self.changes = []
        
    def detect(self) -> List[Record]:
        """Identify all changes between PPT and Excel"""
        records = []
        
        # Group PPT data by record type (Product first, then Teams)
        ppt_products = self._extract_ppt_products()
        ppt_teams = self._extract_ppt_teams()
        
        # Get Excel records
        excel_products = self.excel_data.get('Products', {})
        excel_teams = self.excel_data.get('Teams', {})
        
        # Compare Products
        all_product_ids = set(ppt_products.keys()) | set(excel_products.keys())
        for product_id in sorted(all_product_ids):
            ppt_rec = ppt_products.get(product_id, {})
            excel_rec = excel_products.get(product_id, {})
            record = self._compare_record(product_id, 'Product', ppt_rec, excel_rec)
            if record:
                records.append(record)
        
        # Compare Teams
        all_team_ids = set(ppt_teams.keys()) | set(excel_teams.keys())
        for team_id in sorted(all_team_ids):
            ppt_rec = ppt_teams.get(team_id, {})
            excel_rec = excel_teams.get(team_id, {})
            record = self._compare_record(team_id, 'Team', ppt_rec, excel_rec)
            if record:
                records.append(record)
        
        return records
    
    def _extract_ppt_products(self) -> Dict[str, Dict[str, str]]:
        """Extract product data from PPT"""
        products = {}
        
        # First product is assumed to be slide 1
        # Extract product name as ID or use first value
        product_name = self.ppt_data.get('sync_product_name', 'Product_1')
        product_id = product_name.replace(' ', '_').replace('.', '')  # Sanitize
        
        product_data = {}
        for ppt_field, (excel_col, table, dtype) in PPT_FIELD_MAP.items():
            if table == 'Products':
                product_data[excel_col] = self.ppt_data.get(ppt_field, '')
        
        if any(product_data.values()):  # Only add if has data
            products[product_id] = product_data
        
        return products
    
    def _extract_ppt_teams(self) -> Dict[str, Dict[str, str]]:
        """Extract team data from PPT"""
        teams = {}
        
        # Teams start from slide 2
        team_name = self.ppt_data.get('sync_team_name', 'Team_1')
        team_id = team_name.replace(' ', '_').replace('.', '')  # Sanitize
        
        team_data = {}
        for ppt_field, (excel_col, table, dtype) in PPT_FIELD_MAP.items():
            if table == 'Teams':
                team_data[excel_col] = self.ppt_data.get(ppt_field, '')
        
        # Add product link
        team_data['Product ID'] = self.ppt_data.get('sync_product_name', '').replace(' ', '_')
        
        if any(team_data.values()):  # Only add if has data
            teams[team_id] = team_data
        
        return teams
    
    def _compare_record(self, record_id: str, record_type: str, ppt_data: Dict, excel_data: Dict) -> Optional[Record]:
        """Compare a single record between PPT and Excel"""
        
        # Determine record name
        if record_type == 'Product':
            record_name = ppt_data.get('Product name', excel_data.get('Product name', 'Unknown'))
        else:
            record_name = ppt_data.get('Team name', excel_data.get('Team name', 'Unknown'))
        
        # Build field comparisons
        fields = {}
        all_columns = set(ppt_data.keys()) | set(excel_data.keys())
        
        for col_name in sorted(all_columns):
            # Skip internal columns
            if col_name.startswith('_'):
                continue
            
            ppt_val = ppt_data.get(col_name, '')
            excel_val = excel_data.get(col_name, '')
            
            # Determine data type
            dtype = 'text'
            for ppt_field, (excel_col, table, ftype) in PPT_FIELD_MAP.items():
                if excel_col == col_name:
                    dtype = ftype
                    break
            
            field_value = FieldValue(col_name, ppt_val, excel_val, dtype)
            fields[col_name] = field_value
        
        # Create record if has any fields
        if fields:
            return Record(record_id, record_name, record_type, fields)
        
        return None


class SyncReport:
    """Generate change proposal report"""
    
    def __init__(self, records: List[Record]):
        self.records = records
    
    def generate_summary(self) -> str:
        """Generate human-readable summary of changes"""
        output = []
        output.append("=" * 80)
        output.append("PPT <-> EXCEL SYNC ENGINE - CHANGE REPORT")
        output.append("=" * 80)
        output.append("")
        
        # Summary stats
        total_records = len(self.records)
        records_with_changes = len([r for r in self.records if r.has_changes])
        output.append(f"Records reviewed: {total_records} | With changes: {records_with_changes}")
        output.append("")
        
        # Detailed report
        for record in sorted(self.records, key=lambda r: (r.record_type, r.record_name)):
            if record.has_changes:
                output.append(f"[{record.record_type.upper()}] {record.record_name} (ID: {record.record_id})")
                summary = record.change_summary
                output.append(f"  Status: NEW={summary['new']}, MODIFIED={summary['modified']}, DELETED={summary['deleted']}, MATCHED={summary['matched']}")
                output.append("")
                
                # Field details
                for field_key in sorted(record.fields.keys()):
                    field = record.fields[field_key]
                    if field.status != ChangeType.MATCHED:
                        status_str = field.status.value.upper()
                        ppt_str = f"'{field.ppt_value}'" if field.ppt_value else "[empty]"
                        excel_str = f"'{field.excel_value}'" if field.excel_value else "[empty]"
                        output.append(f"    {status_str:10} | {field_key:40} | PPT: {ppt_str:30} | Excel: {excel_str:30}")
                
                output.append("")
        
        output.append("=" * 80)
        return "\n".join(output)
    
    def save_report(self, output_path: str):
        """Save report to file"""
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(self.generate_summary())
        print(f"Report saved to: {output_path}")


# Main execution
if __name__ == '__main__':
    ppt_path = r'C:\Users\GESZA4\OneDrive - IKEA\Apps\ppt to excel\product_team_optimized_test.pptx'
    excel_path = r'C:\Users\GESZA4\OneDrive - IKEA\Apps\ppt to excel\product_team_master.xlsx'
    
    print("=" * 80)
    print("SYNC ENGINE - INITIALIZATION")
    print("=" * 80)
    print("")
    
    # Read PPT
    print("[1/3] Reading PowerPoint template...")
    ppt_reader = PPTReader(ppt_path)
    ppt_data = ppt_reader.read()
    print(f"      [OK] Extracted {len(ppt_data)} fields from PPT")
    print("")
    
    # Read Excel
    print("[2/3] Reading Excel workbook...")
    excel_reader = ExcelReader(excel_path)
    excel_data = excel_reader.read()
    products_count = len(excel_data.get('Products', {}))
    teams_count = len(excel_data.get('Teams', {}))
    print(f"      [OK] Extracted {products_count} products, {teams_count} teams from Excel")
    print("")
    
    # Detect changes
    print("[3/3] Detecting changes...")
    detector = ChangeDetector(ppt_data, excel_data)
    records = detector.detect()
    print(f"      [OK] Analyzed {len(records)} records")
    print("")
    
    # Generate report
    print("=" * 80)
    report = SyncReport(records)
    print(report.generate_summary())
    
    # Save report
    report_path = r'C:\Users\GESZA4\OneDrive - IKEA\Apps\ppt to excel\SYNC_REPORT.txt'
    report.save_report(report_path)
