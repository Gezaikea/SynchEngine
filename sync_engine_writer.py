"""
PPT-to-Excel Sync Engine - Writer Module
Implements Excel writer to persist synced data
"""

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from typing import Dict, List, Any, Optional
from datetime import datetime


class ExcelWriter:
    """Write changes to Excel workbook"""
    
    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.wb = None
        self.ws_products = None
        self.ws_teams = None
        
    def open(self):
        """Open the workbook for writing"""
        self.wb = load_workbook(self.excel_path)
        self.ws_products = self.wb['Products']
        self.ws_teams = self.wb['Teams']
    
    def close(self):
        """Save and close the workbook"""
        if self.wb:
            self.wb.save(self.excel_path)
            self.wb.close()
    
    def write_product(self, product_id: str, product_data: Dict[str, Any]) -> bool:
        """
        Write or update a product record
        
        Args:
            product_id: Unique product identifier
            product_data: Dictionary of {column_name: value}
        
        Returns:
            True if successful
        """
        try:
            # Ensure product_id is in the data
            product_data_with_id = {'Product ID': product_id, **product_data}
            
            # Find or create row
            row_num = self._find_or_create_row(self.ws_products, product_id, 'Products')
            
            # Get column headers
            headers = self._get_headers(self.ws_products)
            
            # Write data
            for col_idx, header in enumerate(headers, 1):
                if header in product_data_with_id:
                    cell = self.ws_products.cell(row=row_num, column=col_idx)
                    cell.value = product_data_with_id[header]
            
            return True
        except Exception as e:
            print(f"Error writing product {product_id}: {e}")
            return False
    
    def write_team(self, team_id: str, team_data: Dict[str, Any]) -> bool:
        """
        Write or update a team record
        
        Args:
            team_id: Unique team identifier
            team_data: Dictionary of {column_name: value}
        
        Returns:
            True if successful
        """
        try:
            # Ensure team_id is in the data
            team_data_with_id = {'Team ID': team_id, **team_data}
            
            # Find or create row
            row_num = self._find_or_create_row(self.ws_teams, team_id, 'Teams')
            
            # Get column headers
            headers = self._get_headers(self.ws_teams)
            
            # Write data
            for col_idx, header in enumerate(headers, 1):
                if header in team_data_with_id:
                    cell = self.ws_teams.cell(row=row_num, column=col_idx)
                    cell.value = team_data_with_id[header]
            
            return True
        except Exception as e:
            print(f"Error writing team {team_id}: {e}")
            return False
    
    def _find_or_create_row(self, sheet, record_id: str, table_name: str) -> int:
        """
        Find existing row or create new one
        
        Returns:
            Row number (1-indexed)
        """
        # Find existing row (column A is ID), search only first 1000 rows
        for row_idx in range(2, min(1000, sheet.max_row + 1)):
            cell_val = sheet.cell(row=row_idx, column=1).value
            if not cell_val:
                # Found first empty row, use this
                return row_idx
            if str(cell_val).strip() == str(record_id).strip():
                # Found existing record, update it
                return row_idx
        
        # If we get here, table is full, find first empty row after header
        for row_idx in range(2, sheet.max_row + 1):
            cell_val = sheet.cell(row=row_idx, column=1).value
            if not cell_val:
                return row_idx
        
        # Fallback: add at end
        return sheet.max_row + 1
    
    def _get_headers(self, sheet) -> List[str]:
        """Extract column headers from first row"""
        headers = []
        for col_idx in range(1, 100):
            cell_val = sheet.cell(row=1, column=col_idx).value
            if not cell_val:
                break
            headers.append(str(cell_val))
        return headers
    
    def write_batch(self, products: Dict[str, Dict] = None, teams: Dict[str, Dict] = None) -> Dict[str, int]:
        """
        Write multiple records in batch
        
        Args:
            products: Dictionary of {product_id: {column: value}}
            teams: Dictionary of {team_id: {column: value}}
        
        Returns:
            Status dictionary with counts
        """
        status = {'products_written': 0, 'products_failed': 0, 'teams_written': 0, 'teams_failed': 0}
        
        if products:
            for product_id, product_data in products.items():
                if self.write_product(product_id, product_data):
                    status['products_written'] += 1
                else:
                    status['products_failed'] += 1
        
        if teams:
            for team_id, team_data in teams.items():
                if self.write_team(team_id, team_data):
                    status['teams_written'] += 1
                else:
                    status['teams_failed'] += 1
        
        return status


class PPTWriter:
    """Write changes to PowerPoint template (future implementation)"""
    
    def __init__(self, ppt_path: str):
        self.ppt_path = ppt_path
        self.prs = None
    
    def open(self):
        """Open the presentation for writing"""
        try:
            from pptx import Presentation
            self.prs = Presentation(self.ppt_path)
        except ImportError:
            print("Error: python-pptx not installed")
    
    def close(self):
        """Save and close the presentation"""
        if self.prs:
            self.prs.save(self.ppt_path)
    
    def update_product_slide(self, product_id: str, product_data: Dict[str, Any]) -> bool:
        """Update product slide with data from Excel"""
        # TODO: Implement finding and updating product slide
        # This would require identifying which slide contains the product
        # Then updating all named shapes with values from product_data
        pass
    
    def update_team_slide(self, team_id: str, team_data: Dict[str, Any]) -> bool:
        """Update team slide with data from Excel"""
        # TODO: Implement finding and updating team slide
        pass
    
    def create_product_slide(self, product_data: Dict[str, Any]) -> bool:
        """Create a new product slide from template"""
        # TODO: Implement slide cloning and data population
        pass
    
    def create_team_slide(self, team_data: Dict[str, Any]) -> bool:
        """Create a new team slide from template"""
        # TODO: Implement slide cloning and data population
        pass


# Example usage
if __name__ == '__main__':
    # Example data from PPT
    product_data = {
        'Product name': '01. Engineering Toolchain',
        'Product type': 'Technology',
        'Primary users': 'Engineering and product/delivery teams',
        'Organisational owner': 'To confirm',
        'Purpose and value': 'Shared delivery, automation and quality capabilities...',
        'SL1 system count': '',
        'SL2 system count': '2',
        'High criticality systems': '',
        'Medium criticality systems': '',
        'Low criticality systems': '19',
        'Component count': '6',
        'System count': '19',
        'Primary systems': 'Github, SauceLab, Zephyr',
        'Build percent': 'Build 60',
        'Buy percent': 'Buy 40',
        'Main vendor or partner': '',
        'Competency 1': 'Source Code Engineering',
        'Competency 2': 'Quality Test automation',
        'Competency 3': 'CI/CD',
        'Technical complexity': 'H',
        'Investment direction': 'Invest',
        'Change demand 1-2 years': 'H',
        'Adjacent products and teams': 'Adjacent products / teams',
        'Dependencies and interaction model': 'Dependencies and interaction model',
        'AS-IS team count': '5',
        'AS-IS FTE': '50',
        'Catalogue link': 'To confirm',
    }
    
    team_data = {
        'Product ID': '01. Engineering Toolchain',
        'Team name': 'Engineering Delivery Platform',
        'Team type': 'Platform',
        'Purpose and focus': 'Engineering teams collaborate on source and reliably build, test, release, deploy and recover changes',
        'Composition': 'Product · engineering · reliability · developer experience / adoption',
        'Team size': 'XL',
        'Sibling team details': 'Number and sizes, if applicable',
        'Sizing reason - Business criticality': 'x',
        'Sizing reason - Service level': '',
        'Sizing reason - Technical complexity': 'x',
        'Sizing reason - Sourcing': '',
        'Sizing reason - Investment direction': 'x',
        'Consumers': 'Product engineering teams; Quality team; release consumers',
        'Supported by': '',
        'Interaction modes': 'X-as-a-Service',
        'Boundaries': 'Delivery platform and build evidence to product teams...',
        'Topology - value flow': 'x',
        'Topology - stream alignment': 'x',
        'Topology - outcome ownership': 'x',
        'Topology - handoffs and load': 'x',
        'Topology - feedback': 'x',
    }
    
    # Write to Excel
    excel_path = r'C:\Users\GESZA4\OneDrive - IKEA\Apps\ppt to excel\product_team_master.xlsx'
    
    writer = ExcelWriter(excel_path)
    writer.open()
    
    # Write product
    if writer.write_product('01_Engineering_Toolchain', product_data):
        print("[OK] Product written to Excel")
    
    # Write team
    if writer.write_team('Engineering_Delivery_Platform', team_data):
        print("[OK] Team written to Excel")
    
    writer.close()
    print("[OK] Excel saved")
