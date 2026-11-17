"""
PPT-to-Excel Sync Engine - Core Module
Reads PPT answer fields and Excel data, detects changes, proposes updates
"""

import os
import json
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any, Tuple, Optional
from dataclasses import dataclass, asdict
from enum import Enum

# Field mappings: internal key -> (ppt_shape_name, excel_column, table, data_type)
PPT_FIELD_MAP = {
    # Product fields
    'product_id': ('sync_product_id', 'Product ID', 'Products', 'text'),
    'product_name': ('sync_product_name', 'Product name', 'Products', 'text'),
    'product_type': ('sync_product_type', 'Product type', 'Products', 'choice'),
    'primary_users': ('sync_primary_users', 'Primary users', 'Products', 'text'),
    'organisational_owner': ('sync_org_owner', 'Organisational owner', 'Products', 'text'),
    'purpose_value': ('sync_purpose_value', 'Purpose and value', 'Products', 'text'),
    
    'sl1_count': ('sync_sl1_count', 'SL1 system count', 'Products', 'integer'),
    'sl2_count': ('sync_sl2_count', 'SL2 system count', 'Products', 'integer'),
    'high_crit_count': ('sync_high_crit', 'High criticality systems', 'Products', 'integer'),
    'med_crit_count': ('sync_med_crit', 'Medium criticality systems', 'Products', 'integer'),
    'low_crit_count': ('sync_low_crit', 'Low criticality systems', 'Products', 'integer'),
    
    'component_count': ('sync_component_count', 'Component count', 'Products', 'integer'),
    'system_count': ('sync_system_count', 'System count', 'Products', 'integer'),
    'primary_systems': ('sync_primary_systems', 'Primary systems', 'Products', 'text'),
    'make_percent': ('sync_make_percent', 'Build percent', 'Products', 'percentage'),
    'buy_percent': ('sync_buy_percent', 'Buy percent', 'Products', 'percentage'),
    'vendor': ('sync_vendor', 'Main vendor or partner', 'Products', 'text'),
    'competency_1': ('sync_competency_1', 'Competency 1', 'Products', 'text'),
    'competency_2': ('sync_competency_2', 'Competency 2', 'Products', 'text'),
    'competency_3': ('sync_competency_3', 'Competency 3', 'Products', 'text'),
    
    'technical_complexity': ('sync_tech_complexity', 'Technical complexity', 'Products', 'choice'),
    'investment_direction': ('sync_investment', 'Investment direction', 'Products', 'choice'),
    'change_demand': ('sync_change_demand', 'Change demand 1-2 years', 'Products', 'choice'),
    
    'adjacent_products': ('sync_adjacent', 'Adjacent products and teams', 'Products', 'text'),
    'dependencies': ('sync_dependencies', 'Dependencies and interaction model', 'Products', 'text'),
    
    'as_is_team_count': ('sync_as_is_team_count', 'AS-IS team count', 'Products', 'integer'),
    'as_is_fte': ('sync_as_is_fte', 'AS-IS FTE', 'Products', 'decimal'),
    
    # Team fields
    'team_id': ('sync_team_id', 'Team ID', 'Teams', 'text'),
    'team_name': ('sync_team_name', 'Team name', 'Teams', 'text'),
    'product_id_link': ('sync_product_link', 'Product ID', 'Teams', 'text'),
    'assignment_status': ('sync_assignment_status', 'Assignment status', 'Teams', 'choice'),
    
    'team_type': ('sync_team_type', 'Team type', 'Teams', 'choice'),
    'team_purpose': ('sync_team_purpose', 'Purpose and focus', 'Teams', 'text'),
    'composition': ('sync_composition', 'Composition', 'Teams', 'text'),
    
    'team_size': ('sync_team_size', 'Team size', 'Teams', 'lookup'),
    'sibling_team': ('sync_sibling_team', 'Sibling team?', 'Teams', 'choice'),
    
    'sizing_criticality': ('sync_sizing_criticality', 'Sizing reason - Business criticality', 'Teams', 'boolean'),
    'sizing_service_level': ('sync_sizing_sl', 'Sizing reason - Service level', 'Teams', 'boolean'),
    'sizing_complexity': ('sync_sizing_complexity', 'Sizing reason - Technical complexity', 'Teams', 'boolean'),
    'sizing_sourcing': ('sync_sizing_sourcing', 'Sizing reason - Sourcing', 'Teams', 'boolean'),
    'sizing_investment': ('sync_sizing_investment', 'Sizing reason - Investment direction', 'Teams', 'boolean'),
    
    'consumers': ('sync_consumers', 'Consumers', 'Teams', 'text'),
    'supported_by': ('sync_supported_by', 'Supported by', 'Teams', 'text'),
    'interaction_modes': ('sync_interaction', 'Interaction modes', 'Teams', 'text'),
    'boundaries': ('sync_boundaries', 'Boundaries', 'Teams', 'text'),
    
    'topology_value_flow': ('sync_topology_value', 'Topology - value flow', 'Teams', 'boolean'),
    'topology_stream_align': ('sync_topology_stream', 'Topology - stream alignment', 'Teams', 'boolean'),
    'topology_outcome': ('sync_topology_outcome', 'Topology - outcome ownership', 'Teams', 'boolean'),
    'topology_handoffs': ('sync_topology_handoffs', 'Topology - handoffs and load', 'Teams', 'boolean'),
    'topology_feedback': ('sync_topology_feedback', 'Topology - feedback', 'Teams', 'boolean'),
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
    record_id: str  # Product ID or Team ID
    record_name: str  # Product name or Team name
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
                            extracted[shape.name] = shape.text.strip()
            
            self.data = extracted
            return extracted
            
        except ImportError:
            # Fallback: use python-pptx if available, otherwise return empty
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
    
    def read(self) -> Dict[str, Dict[str, Any]]:
        """Extract all data from Products and Teams tables"""
        try:
            from openpyxl import load_workbook
            wb = load_workbook(self.excel_path)
            
            # Read Products
            products_sheet = wb['Products']
            products_table = products_sheet.tables['Products']
            
            products = {}
            for row in products_table.ref[1:]:  # Skip header
                if products_sheet[row].value:
                    product_id = products_sheet[f'A{row.row}'].value
                    if product_id:
                        products[product_id] = self._extract_row(products_sheet, row.row, 'Products')
            
            # Read Teams
            teams_sheet = wb['Teams']
            teams_table = teams_sheet.tables['Teams']
            
            teams = {}
            for row in teams_table.ref[1:]:  # Skip header
                if teams_sheet[row].value:
                    team_id = teams_sheet[f'A{row.row}'].value
                    if team_id:
                        teams[team_id] = self._extract_row(teams_sheet, row.row, 'Teams')
            
            self.data = {'Products': products, 'Teams': teams}
            return self.data
            
        except ImportError:
            print("Error: openpyxl not installed. Install with: pip install openpyxl")
            return {}
        except Exception as e:
            print(f"Error reading Excel: {e}")
            return {}
    
    def _extract_row(self, sheet, row: int, table_type: str) -> Dict[str, Any]:
        """Extract all columns from a row"""
        row_data = {}
        # Simplified: would need actual column mapping
        return row_data

class ChangeDetector:
    """Detect and report changes between PPT and Excel"""
    
    def __init__(self, ppt_data: Dict, excel_data: Dict):
        self.ppt_data = ppt_data
        self.excel_data = excel_data
        self.changes = []
    
    def detect(self) -> List[Record]:
        """Identify all changes between PPT and Excel"""
        records = []
        
        # Map PPT fields to records
        ppt_records = self._group_ppt_by_record()
        excel_records = self._group_excel_by_record()
        
        # Compare all records
        all_ids = set(ppt_records.keys()) | set(excel_records.keys())
        
        for record_id in all_ids:
            ppt_rec = ppt_records.get(record_id, {})
            excel_rec = excel_records.get(record_id, {})
            
            record = self._compare_record(record_id, ppt_rec, excel_rec)
            if record:
                records.append(record)
        
        return records
    
    def _group_ppt_by_record(self) -> Dict[str, Dict]:
        """Group PPT fields by record ID"""
        groups = {}
        for field_key, ppt_name in PPT_FIELD_MAP.items():
            shape_name = ppt_name[0]
            value = self.ppt_data.get(shape_name)
            if value:
                # Extract record ID from value or field (simplified)
                record_id = "TODO"  # Need smarter parsing
                if record_id not in groups:
                    groups[record_id] = {}
                groups[record_id][field_key] = value
        return groups
    
    def _group_excel_by_record(self) -> Dict[str, Dict]:
        """Group Excel data by record ID"""
        # Simplified
        return {}
    
    def _compare_record(self, record_id: str, ppt_data: Dict, excel_data: Dict) -> Optional[Record]:
        """Compare a single record"""
        # Simplified comparison logic
        return None

class SyncReport:
    """Generate change proposal report"""
    
    def __init__(self, records: List[Record]):
        self.records = records
    
    def generate_summary(self) -> str:
        """Generate human-readable summary of changes"""
        output = []
        output.append("=" * 70)
        output.append("PPT ↔ EXCEL SYNC REPORT")
        output.append("=" * 70)
        output.append("")
        
        for record in self.records:
            if record.has_changes:
                output.append(f"[{record.record_type}] {record.record_name} ({record.record_id})")
                output.append(f"  Changes: {record.change_summary}")
                output.append("")
                
                for field_key, field in record.fields.items():
                    if field.status != ChangeType.MATCHED:
                        output.append(f"    {field_key}: {field.status.value}")
                        output.append(f"      PPT:   {field.ppt_value}")
                        output.append(f"      Excel: {field.excel_value}")
                output.append("")
        
        output.append("=" * 70)
        return "\n".join(output)

# Main execution
if __name__ == '__main__':
    ppt_path = r'C:\Users\GESZA4\OneDrive - IKEA\Apps\ppt to excel\product_team_optimized_test.pptx'
    excel_path = r'C:\Users\GESZA4\OneDrive - IKEA\Apps\ppt to excel\product_team_master.xlsx'
    
    print("Sync Engine - Reading Files")
    print("=" * 70)
    
    # Read PPT
    print("Reading PPT...")
    ppt_reader = PPTReader(ppt_path)
    ppt_data = ppt_reader.read()
    print(f"Extracted {len(ppt_data)} fields from PPT")
    
    # Read Excel
    print("Reading Excel...")
    excel_reader = ExcelReader(excel_path)
    excel_data = excel_reader.read()
    print(f"Extracted {len(excel_data.get('Products', {}))} products, {len(excel_data.get('Teams', {}))} teams")
    
    # Detect changes
    print("Detecting changes...")
    detector = ChangeDetector(ppt_data, excel_data)
    records = detector.detect()
    
    # Generate report
    print("")
    report = SyncReport(records)
    print(report.generate_summary())
