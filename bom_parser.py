"""BOM file parsing logic for Excel files."""
import openpyxl
import csv
from pathlib import Path
from typing import List, Dict, Any
import logging

logger = logging.getLogger(__name__)


class BOMParser:
    """Parses Excel BOM files and extracts component data."""
    
    # Common column name variations
    COLUMN_MAPPINGS = {
        'refdes': ['refdes', 'ref', 'reference', 'reference designator', 'designator'],
        'mpn': ['mpn', 'manufacturer part number', 'part number', 'part#', 'mfr part number'],
        'value': ['value', 'component value', 'val'],
        'package': ['package', 'footprint', 'case', 'case code', 'size'],
        'voltage': ['voltage', 'voltage rating', 'v rating', 'v'],
        'tolerance': ['tolerance', 'tol'],
        'power': ['power', 'power rating', 'wattage', 'w'],
        'description': ['description', 'desc', 'comment', 'notes'],
        'quantity': ['quantity', 'qty', 'qty per board'],
    }
    
    def __init__(self):
        """Initialize the BOM parser."""
        pass
    
    def normalize_column_name(self, col_name: str) -> str:
        """Normalize column name to standard format."""
        if not col_name:
            return ""
        col_lower = col_name.lower().strip()
        
        # Check each standard column name
        for standard, variations in self.COLUMN_MAPPINGS.items():
            if col_lower in variations:
                return standard
        
        # Return original if no match (preserve for additional columns)
        return col_lower.replace(' ', '_')
    
    def _prepare_headers(self, original_headers: List[str]) -> tuple[List[str], Dict[str, str]]:
        """
        Build normalized header list and mapping to original names.
        
        Ensures duplicate normalized names become unique (e.g., package, package_2).
        """
        headers: List[str] = []
        column_mapping: Dict[str, str] = {}
        name_counts: Dict[str, int] = {}
        
        for idx, raw_header in enumerate(original_headers):
            orig_header = raw_header.strip() if isinstance(raw_header, str) else ""
            if not orig_header:
                orig_header = f"col_{idx}"
            
            base_name = self.normalize_column_name(orig_header) or f"col_{idx}"
            count = name_counts.get(base_name, 0)
            
            if count == 0:
                final_name = base_name
            else:
                final_name = f"{base_name}_{count + 1}"
                # Ensure uniqueness in case prior duplicates created same suffix
                while final_name in column_mapping:
                    count += 1
                    final_name = f"{base_name}_{count + 1}"
            
            name_counts[base_name] = count + 1
            headers.append(final_name)
            column_mapping[final_name] = orig_header
        
        return headers, column_mapping
    
    def parse_csv(self, file_path: str) -> tuple[List[Dict[str, Any]], Dict[str, str]]:
        """
        Parse CSV BOM file and return list of component dictionaries.
        
        Args:
            file_path: Path to CSV file
            
        Returns:
            Tuple of (list of component dictionaries with normalized column names, 
                     mapping of normalized -> original column names)
        """
        file_path = Path(file_path)
        if not file_path.exists():
            logger.error(f"CSV file not found: {file_path}")
            raise FileNotFoundError(f"BOM file not found: {file_path}")
            
        try:
            logger.info(f"Parsing CSV file: {file_path}")
            with open(file_path, 'r', newline='', encoding='utf-8-sig') as f:
                # Use Sniffer to detect dialect (delimiter, etc.)
                try:
                    dialect = csv.Sniffer().sniff(f.read(1024))
                    f.seek(0)
                    logger.debug(f"Detected CSV dialect: delimiter='{dialect.delimiter}'")
                except csv.Error:
                    # Fallback to standard excel dialect
                    f.seek(0)
                    dialect = 'excel'
                    logger.warning("Could not detect CSV dialect, falling back to 'excel'")
                
                reader = csv.reader(f, dialect=dialect)
                
                # Read header
                try:
                    header_row = next(reader)
                except StopIteration:
                    logger.warning("Empty CSV file")
                    return [], {}
                    
                # Store original column names and create normalized mapping
                original_headers = [cell.strip() if cell else f"col_{i}" for i, cell in enumerate(header_row)]
                headers, column_mapping = self._prepare_headers(original_headers)
                
                logger.debug(f"Original headers: {original_headers}")
                logger.debug(f"Normalized headers: {headers}")
                if len(original_headers) != len(set(headers)):
                    logger.debug("Detected duplicate column names; generated unique normalized headers.")
                
                # Read data rows
                components = []
                for row in reader:
                    # Skip empty rows
                    if not any(row):
                        continue
                    
                    component = {}
                    for col_idx, normalized_header in enumerate(headers):
                        if col_idx < len(row):
                            value = row[col_idx]
                            component[normalized_header] = str(value).strip() if value is not None else ""
                    
                    # Only add if component has some data
                    if component and any(v for v in component.values() if v):
                        components.append(component)
                
                logger.info(f"Extracted {len(components)} components from CSV")
                return components, column_mapping
                
        except Exception as e:
            logger.error(f"Error parsing CSV file: {e}", exc_info=True)
            raise Exception(f"Error parsing CSV file: {e}")

    def parse(self, file_path: str) -> tuple[List[Dict[str, Any]], Dict[str, str]]:
        """
        Parse BOM file (Excel or CSV) and return list of component dictionaries.
        
        Args:
            file_path: Path to BOM file
            
        Returns:
            Tuple of (list of component dictionaries, 
                     mapping of normalized -> original column names)
        """
        file_path = Path(file_path)
        suffix = file_path.suffix.lower()
        
        if suffix in ['.csv']:
            return self.parse_csv(file_path)
        elif suffix in ['.xlsx', '.xls', '.xlsm']:
            return self.parse_excel(file_path)
        else:
            raise ValueError(f"Unsupported file format: {suffix}")

    def parse_excel(self, file_path: str) -> tuple[List[Dict[str, Any]], Dict[str, str]]:
        """
        Parse Excel BOM file and return list of component dictionaries.
        
        Args:
            file_path: Path to Excel file
            
        Returns:
            Tuple of (list of component dictionaries with normalized column names,
                     mapping of normalized -> original column names)
        """
        file_path = Path(file_path)
        if not file_path.exists():
            logger.error(f"Excel file not found: {file_path}")
            raise FileNotFoundError(f"BOM file not found: {file_path}")
        
        try:
            logger.info(f"Parsing Excel file: {file_path}")
            workbook = openpyxl.load_workbook(file_path, data_only=True)
            # Use first sheet by default
            sheet = workbook.active
            logger.debug(f"Using sheet: {sheet.title}")
            
            # Read header row and store original column names
            original_headers = []
            header_row = sheet[1]
            for cell in header_row:
                orig_name = cell.value.strip() if cell.value else f"col_{len(original_headers)}"
                original_headers.append(orig_name)
            
            # Create normalized headers and mapping (handles duplicates)
            headers, column_mapping = self._prepare_headers(original_headers)
            
            logger.debug(f"Original headers: {original_headers}")
            logger.debug(f"Normalized headers: {headers}")
            if len(original_headers) != len(set(headers)):
                logger.debug("Detected duplicate column names; generated unique normalized headers.")
            
            # Read data rows
            components = []
            for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=False), start=2):
                # Skip empty rows
                if not any(cell.value for cell in row):
                    continue
                
                component = {}
                for col_idx, header in enumerate(headers):
                    if col_idx < len(row):
                        value = row[col_idx].value
                        # Convert to string if not None, preserve None for missing values
                        component[header] = str(value).strip() if value is not None else ""
                
                # Only add if component has some data
                if component and any(v for v in component.values() if v):
                    components.append(component)
            
            logger.info(f"Extracted {len(components)} components from Excel")
            return components, column_mapping
            
        except Exception as e:
            logger.error(f"Error parsing Excel file: {e}", exc_info=True)
            raise Exception(f"Error parsing BOM file: {e}")
    
    def get_consolidated_parts(self, components: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Consolidate components by grouping similar parts.
        Returns unique parts with aggregated reference designators.
        
        Args:
            components: List of component dictionaries
            
        Returns:
            List of consolidated part dictionaries
        """
        # Group by key attributes (MPN, Value, Package combination)
        parts_dict = {}
        
        for comp in components:
            # Create a key from MPN, Value, and Package
            mpn = comp.get('mpn', '').strip()
            value = comp.get('value', '').strip()
            package = comp.get('package', '').strip()
            
            # Use MPN as primary key if available, otherwise use Value+Package
            if mpn:
                key = f"mpn:{mpn}"
            elif value and package:
                key = f"value:{value}|package:{package}"
            elif value:
                key = f"value:{value}"
            else:
                # Fallback: use all available attributes
                key = str(hash(str(sorted(comp.items()))))
            
            if key not in parts_dict:
                parts_dict[key] = {
                    'refdes_list': [],
                    'mpn': mpn,
                    'value': value,
                    'package': package,
                    'voltage': comp.get('voltage', ''),
                    'tolerance': comp.get('tolerance', ''),
                    'power': comp.get('power', ''),
                    'description': comp.get('description', ''),
                    'quantity': 0,
                }
                # Copy any additional columns
                for k, v in comp.items():
                    if k not in parts_dict[key] and k not in ['refdes']:
                        parts_dict[key][k] = v
            
            # Aggregate reference designators
            refdes = comp.get('refdes', '').strip()
            if refdes:
                if refdes not in parts_dict[key]['refdes_list']:
                    parts_dict[key]['refdes_list'].append(refdes)
            
            # Sum quantities
            try:
                qty = comp.get('quantity', 0)
                if isinstance(qty, str):
                    qty = int(qty) if qty.isdigit() else 0
                parts_dict[key]['quantity'] += qty if qty else 1
            except:
                parts_dict[key]['quantity'] += 1
        
        # Convert to list and format refdes
        consolidated = []
        for part in parts_dict.values():
            part['refdes'] = ', '.join(part['refdes_list'])
            consolidated.append(part)
        
        return consolidated

