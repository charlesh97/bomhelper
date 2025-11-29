"""
DEPRECATED: This module is no longer used by the main application.
It is kept for reference only. 
Functionality has been replaced by direct Gemini integration in the main app for keyword generation.
"""
"""Component spec normalization with rule-based parsing and Gemini AI fallback."""
import re
import json
from typing import Dict, Any, Optional
import google.generativeai as genai
import logging
from config import Config

logger = logging.getLogger(__name__)


class SpecParser:
    """Parses component descriptions into normalized spec dictionaries."""
    
    def __init__(self, config: Config):
        """Initialize the spec parser with API configuration."""
        self.config = config
        self.gemini_client = None
        
        if config.get_gemini_api_key():
            try:
                genai.configure(api_key=config.get_gemini_api_key())
                self.gemini_client = genai.GenerativeModel('gemini-pro')
                logger.info("Gemini client initialized successfully")
            except Exception as e:
                logger.error(f"Could not initialize Gemini client: {e}")
        else:
            logger.info("Gemini API key not configured")
    
    def normalize_units(self, value_str: str) -> str:
        """Normalize unit representations (kΩ → 1000, µF → uF, etc.)."""
        if not value_str:
            return ""
        
        # Resistor values
        value_str = re.sub(r'(\d+\.?\d*)\s*kΩ', r'\1k', value_str, flags=re.IGNORECASE)
        value_str = re.sub(r'(\d+\.?\d*)\s*KΩ', r'\1k', value_str, flags=re.IGNORECASE)
        value_str = re.sub(r'(\d+\.?\d*)\s*MΩ', r'\1M', value_str, flags=re.IGNORECASE)
        value_str = re.sub(r'(\d+\.?\d*)\s*Ω', r'\1', value_str, flags=re.IGNORECASE)
        value_str = re.sub(r'(\d+\.?\d*)\s*Ohm', r'\1', value_str, flags=re.IGNORECASE)
        value_str = re.sub(r'(\d+\.?\d*)\s*ohms', r'\1', value_str, flags=re.IGNORECASE)
        
        # Capacitor values
        value_str = re.sub(r'(\d+\.?\d*)\s*µF', r'\1uF', value_str, flags=re.IGNORECASE)
        value_str = re.sub(r'(\d+\.?\d*)\s*μF', r'\1uF', value_str, flags=re.IGNORECASE)
        value_str = re.sub(r'(\d+\.?\d*)\s*uF', r'\1uF', value_str, flags=re.IGNORECASE)
        value_str = re.sub(r'(\d+\.?\d*)\s*nF', r'\1nF', value_str, flags=re.IGNORECASE)
        value_str = re.sub(r'(\d+\.?\d*)\s*pF', r'\1pF', value_str, flags=re.IGNORECASE)
        
        return value_str.strip()
    
    def parse_resistor(self, component: Dict[str, Any]) -> Dict[str, Any]:
        """Parse resistor component specs."""
        spec = {
            'category': 'resistor',
            'value': self.normalize_units(component.get('value', '')),
            'package': component.get('package', '').strip(),
            'tolerance': component.get('tolerance', '5%').strip(),
            'power': component.get('power', '0.1W').strip(),
            'voltage': component.get('voltage', '').strip(),
        }
        
        # Extract value from description if not in value field
        if not spec['value']:
            desc = component.get('description', '')
            # Look for resistor value patterns
            match = re.search(r'(\d+\.?\d*)\s*(k|K|M|m)?\s*(Ω|Ohm|ohms|ohm)', desc, re.IGNORECASE)
            if match:
                spec['value'] = match.group(1) + (match.group(2) or '')
        
        # Extract package from description if not in package field
        if not spec['package']:
            desc = component.get('description', '')
            # Look for package codes (0603, 0805, 1206, etc.)
            match = re.search(r'\b(0402|0603|0805|1206|1210|2010|2512)\b', desc)
            if match:
                spec['package'] = match.group(1)
        
        return spec
    
    def parse_capacitor(self, component: Dict[str, Any]) -> Dict[str, Any]:
        """Parse capacitor component specs."""
        spec = {
            'category': 'capacitor',
            'value': self.normalize_units(component.get('value', '')),
            'package': component.get('package', '').strip(),
            'voltage': component.get('voltage', '').strip(),
            'tolerance': component.get('tolerance', '20%').strip(),
            'dielectric': component.get('dielectric', '').strip(),
        }
        
        # Extract value from description if not in value field
        if not spec['value']:
            desc = component.get('description', '')
            # Look for capacitor value patterns
            match = re.search(r'(\d+\.?\d*)\s*(uF|µF|μF|nF|pF)', desc, re.IGNORECASE)
            if match:
                spec['value'] = match.group(1) + match.group(2)
        
        # Extract voltage from description
        if not spec['voltage']:
            desc = component.get('description', '')
            match = re.search(r'(\d+)\s*V', desc, re.IGNORECASE)
            if match:
                spec['voltage'] = match.group(1) + 'V'
        
        # Extract dielectric type
        if not spec['dielectric']:
            desc = component.get('description', '')
            for dielectric in ['X5R', 'X7R', 'X8R', 'NP0', 'C0G', 'Y5V', 'Z5U']:
                if dielectric in desc.upper():
                    spec['dielectric'] = dielectric
                    break
        
        return spec
    
    def parse_inductor(self, component: Dict[str, Any]) -> Dict[str, Any]:
        """Parse inductor component specs."""
        spec = {
            'category': 'inductor',
            'value': component.get('value', '').strip(),
            'package': component.get('package', '').strip(),
            'current': component.get('current', '').strip(),
            'tolerance': component.get('tolerance', '20%').strip(),
        }
        
        # Normalize inductance values
        if spec['value']:
            spec['value'] = self.normalize_units(spec['value'])
        
        return spec
    
    def detect_category(self, component: Dict[str, Any]) -> str:
        """Detect component category from description or MPN."""
        desc = (component.get('description', '') + ' ' + component.get('mpn', '')).lower()
        refdes = (component.get('refdes', '') or '').upper()
        
        # Check reference designator prefix
        if refdes.startswith('R'):
            return 'resistor'
        elif refdes.startswith('C'):
            return 'capacitor'
        elif refdes.startswith('L'):
            return 'inductor'
        elif refdes.startswith('U') or refdes.startswith('IC'):
            return 'ic'
        elif refdes.startswith('J') or refdes.startswith('CONN'):
            return 'connector'
        
        # Check description keywords
        if any(word in desc for word in ['resistor', 'res', 'ohm', 'kohm', 'mohm']):
            return 'resistor'
        elif any(word in desc for word in ['capacitor', 'cap', 'uf', 'nf', 'pf']):
            return 'capacitor'
        elif any(word in desc for word in ['inductor', 'inductance', 'uh', 'nh']):
            return 'inductor'
        elif any(word in desc for word in ['connector', 'header', 'socket', 'jack']):
            return 'connector'
        elif any(word in desc for word in ['ic', 'integrated circuit', 'chip', 'microcontroller']):
            return 'ic'
        
        return 'unknown'
    
    def parse_rule_based(self, component: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Attempt rule-based parsing of component specs."""
        category = self.detect_category(component)
        
        if category == 'resistor':
            return self.parse_resistor(component)
        elif category == 'capacitor':
            return self.parse_capacitor(component)
        elif category == 'inductor':
            return self.parse_inductor(component)
        elif category == 'ic' or category == 'connector':
            # For ICs and connectors, return basic structure
            return {
                'category': category,
                'mpn': component.get('mpn', '').strip(),
                'package': component.get('package', '').strip(),
                'description': component.get('description', '').strip(),
            }
        else:
            return None
    
    def parse_with_gemini(self, component: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Use Gemini AI to generate a search keyword phrase."""
        if not self.gemini_client:
            return None
        
        try:
            logger.info("Calling Gemini API for keyword generation")
            
            # Build description from component data
            desc_parts = []
            if component.get('description'):
                desc_parts.append(f"Description: {component['description']}")
            if component.get('value'):
                desc_parts.append(f"Value: {component['value']}")
            if component.get('package'):
                desc_parts.append(f"Package: {component['package']}")
            if component.get('voltage'):
                desc_parts.append(f"Voltage: {component['voltage']}")
            if component.get('mpn'):
                desc_parts.append(f"MPN: {component['mpn']}")
            
            description = ' '.join(desc_parts) if desc_parts else str(component)
            
            prompt = f"""Create a concise search keyword phrase for this electronic component for Mouser Electronics.
Input: {description}

Requirements:
- Include key specs (value, units, package, tolerance if critical)
- Include part type (resistor, capacitor, etc.)
- Do NOT include generic words like "Description" or "Value"
- Output ONLY the search phrase (e.g., "10k ohm resistor 0603 1%")
- Keep it under 50 characters if possible."""

            response = self.gemini_client.generate_content(prompt)
            search_phrase = response.text.strip().strip('"').strip("'")
            
            logger.info(f"Gemini generated search phrase: {search_phrase}")
            return {'keyword': search_phrase}
            
        except Exception as e:
            logger.error(f"Gemini parsing failed: {e}")
            return None
    
    def parse(self, component: Dict[str, Any]) -> Dict[str, Any]:
        """
        Parse component using Gemini to get a search keyword.
        
        Args:
            component: Component dictionary from BOM
            
        Returns:
            Dictionary with 'keyword' key
        """
        # Try Gemini first as requested
        gemini_spec = self.parse_with_gemini(component)
        if gemini_spec:
            return gemini_spec
        
        # Fallback to simple concatenation if Gemini fails
        logger.warning("Gemini failed, falling back to simple concatenation")
        parts = []
        if component.get('value'): parts.append(component['value'])
        if component.get('package'): parts.append(component['package'])
        if component.get('description'): parts.append(component['description'])
        
        return {'keyword': ' '.join(parts)}
