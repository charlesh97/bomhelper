"""Mouser API client wrapper for part search."""
import requests
import time
from typing import Dict, Any, List, Optional
import logging
from config import Config

logger = logging.getLogger(__name__)


class MouserAPI:
    """Client for Mouser API searches."""
    
    BASE_URL = "https://api.mouser.com/api/v1"
    
    def __init__(self, config: Config):
        """Initialize Mouser API client."""
        self.config = config
        self.api_key = config.get_mouser_api_key()
        if not self.api_key:
            raise ValueError("Mouser API key not configured")
        
        self.headers = {
            'Content-Type': 'application/json',
        }
        self.last_request_time = 0
        self.min_request_interval = 0.5  # Minimum seconds between requests
    
    def _rate_limit(self):
        """Enforce rate limiting between API requests."""
        elapsed = time.time() - self.last_request_time
        if elapsed < self.min_request_interval:
            time.sleep(self.min_request_interval - elapsed)
        self.last_request_time = time.time()
    
    def _make_request(self, endpoint: str, payload: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Make API request with error handling."""
        self._rate_limit()
        
        url = f"{self.BASE_URL}/{endpoint}"
        logger.debug(f"Making API request to {url}")
        
        try:
            response = requests.post(
                url,
                json=payload,
                headers=self.headers,
                params={'apiKey': self.api_key},
                timeout=30
            )
            response.raise_for_status()
            data = response.json()
            
            # Log result count if available
            if 'SearchResults' in data and data['SearchResults'] is not None:
                count = data['SearchResults'].get('NumberOfResult', 'unknown')
                logger.info(f"API request successful. Found {count} results.")
            
            # Check if SearchResults exists and has Parts
            if 'SearchResults' not in data or data['SearchResults'] is None:
                logger.warning("API response missing SearchResults")
                return None
            
            if 'Parts' not in data['SearchResults']:
                logger.warning("API response missing Parts in SearchResults")
                return None
            
            return data['SearchResults']['Parts']

        except requests.exceptions.RequestException as e:
            logger.error(f"Mouser API error: {e}")
            if hasattr(e.response, 'text'):
                logger.error(f"Response: {e.response.text}")
            return None
    
    def search_by_mpn(self, mpn: str) -> List[Dict[str, Any]]:
        """
        Search for exact MPN match.
        
        Args:
            mpn: Manufacturer Part Number
            
        Returns:
            List of matching parts
        """
        logger.info(f"Searching by MPN: {mpn}")
        payload = {
            "SearchByPartRequest": {
                "mouserPartNumber": "",
                "partNumber": mpn,
                "partSearchOptions": "Exact"
            }
        }
        
        result = self._make_request("search/partnumber", payload)
        if not result:
            return []
        
        return self._normalize_results(result)
    
    def search_keyword(self, keyword: str, max_results: int = 50, search_options: str = "None") -> List[Dict[str, Any]]:
        """
        Search by keyword.
        
        Args:
            keyword: Search keyword
            max_results: Maximum number of results
            search_options: API filtering options (None, Rohs, InStock, RohsAndInStock)
            
        Returns:
            List of matching parts
        """
        logger.info(f"Searching by keyword: {keyword} (Options: {search_options})")
        payload = {
            "SearchByKeywordRequest": {
                "keyword": keyword,
                "records": max_results,
                "startingRecord": 0,
                "searchOptions": search_options,
                "searchWithYourSignUpLanguage": "en"
            }
        }
        
        result = self._make_request("search/keyword", payload)
        if not result:
            return []
        
        return self._normalize_results(result)
    
    def _normalize_results(self, api_response: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Normalize Mouser API response to standard format."""
        parts = []
        
        # Handle different response structures
        search_results = None
        if 'SearchResults' in api_response:
            search_results = api_response['SearchResults']
        elif 'Parts' in api_response:
            search_results = api_response['Parts']
        elif isinstance(api_response, list):
            search_results = api_response
        else:
            # Try to find parts array in nested structure
            for key in ['Parts', 'Results', 'SearchResults', 'data']:
                if key in api_response:
                    search_results = api_response[key]
                    break
        
        if not search_results:
            return parts
        
        # Handle both list and single item responses
        if not isinstance(search_results, list):
            search_results = [search_results]
        
        for part in search_results:
            normalized = {
                'mpn': part.get('ManufacturerPartNumber', '') or part.get('MfrPartNumber', ''),
                'manufacturer': part.get('Manufacturer', '') or part.get('Mfr', ''),
                'mouser_part_number': part.get('MouserPartNumber', '') or part.get('PartNumber', ''),
                'description': part.get('Description', '') or part.get('ProductDescription', ''),
                'data_sheet_url': part.get('DataSheetUrl', '') or part.get('DataSheet', ''),
                'product_url': part.get('ProductDetailUrl', '') or part.get('ProductUrl', ''),
                'image_url': part.get('ImagePath', '') or part.get('ImageUrl', ''),
                'lifecycle': part.get('LifecycleStatus', '') or part.get('Status', ''),
                'rohs_status': part.get('ROHSStatus', ''),
                'package': part.get('Package', '') or part.get('CaseCode', ''),
                'stock': 0,
                'price_breaks': [],
            }
            
            # Extract availability/stock information
            # Try multiple possible fields for stock information
            stock = 0
            
            # Check AvailabilityInStock field (V1 API format)
            availability_in_stock = part.get('AvailabilityInStock', '')
            if availability_in_stock:
                try:
                    stock = int(availability_in_stock)
                except (ValueError, TypeError):
                    pass
            
            # Check Availability field (can be dict or string)
            if stock == 0:
                availability = part.get('Availability', {})
                if isinstance(availability, dict):
                    stock = availability.get('OnHand', 0) or availability.get('Quantity', 0)
                    normalized['lead_time'] = availability.get('LeadTime', '')
                elif isinstance(availability, str):
                    # Try to parse stock from string
                    try:
                        stock = int(availability)
                    except (ValueError, TypeError):
                        pass
            
            normalized['stock'] = stock
            
            # Extract price breaks
            price_breaks = part.get('PriceBreaks', [])
            if isinstance(price_breaks, list):
                for pb in price_breaks:
                    if isinstance(pb, dict):
                        normalized['price_breaks'].append({
                            'quantity': pb.get('Quantity', 0),
                            'price': pb.get('Price', ''),
                            'currency': pb.get('Currency', 'USD'),
                        })
            
            # Only add if we have at least an MPN or Mouser part number
            if normalized['mpn'] or normalized['mouser_part_number']:
                parts.append(normalized)
        
        return parts
    
    def search(self, component: Dict[str, Any], spec: Dict[str, Any], 
               in_stock_only: bool = True, active_only: bool = True) -> List[Dict[str, Any]]:
        """
        Main search method that tries different search strategies.
        
        Args:
            component: Original component dictionary
            spec: Normalized spec dictionary
            in_stock_only: Filter to only in-stock parts
            active_only: Filter to only active/lifecycle parts
            
        Returns:
            List of matching parts
        """
        results = []
        search_opts = "InStock" if in_stock_only else "None"
        
        # Strategy 1: If we have an MPN, try exact match first
        mpn = component.get('mpn', '').strip() or spec.get('mpn', '').strip()
        if mpn:
            logger.info(f"Strategy 1: Exact MPN search for '{mpn}'")
            # Note: SearchByPartRequest usually doesn't support InStock filtering directly in V1
            results = self.search_by_mpn(mpn)
            if results:
                logger.info(f"Strategy 1 successful: found {len(results)} parts")
                return self._apply_filters(results, in_stock_only, active_only)
            else:
                logger.info("Strategy 1 failed: no parts found")
        
        # Strategy 2: Keyword search
        keyword = spec.get('keyword')
        if not keyword:
            logger.info("Strategy 2 failed: no keyword provided")
            return []
        
        if keyword:
            logger.info(f"Strategy 2: Keyword search for '{keyword}' with options={search_opts}")
            results = self.search_keyword(keyword, max_results=50, search_options=search_opts)
            logger.info(f"Strategy 2 completed: found {len(results)} parts")
        
        # We still apply filters because parametric/MPN search might not have filtered
        # and keyword search 'InStock' option is good but double checking doesn't hurt
        return self._apply_filters(results, in_stock_only, active_only)
    
    def _apply_filters(self, parts: List[Dict[str, Any]], 
                      in_stock_only: bool, active_only: bool) -> List[Dict[str, Any]]:
        """
        Apply stock and lifecycle filters.
        
        Args:
            parts: List of part dictionaries with format:
                - 'stock': integer (0 if not in stock)
                - 'lifecycle': string (e.g., 'New Product', 'New at Mouser', '', 'OBSOLETE', 'EOL')
            in_stock_only: If True, only return parts with stock > 0
            active_only: If True, filter out obsolete/end-of-life parts
            
        Returns:
            Filtered list of parts
        """
        filtered = parts
        
        if in_stock_only:
            before_count = len(filtered)
            filtered = [p for p in filtered if p.get('stock', 0) > 0]
            logger.debug(f"In-stock filter: {before_count} -> {len(filtered)} parts")
        
        if active_only:
            before_count = len(filtered)
            # Filter out obsolete/end-of-life parts
            # Keep: 'New Product', 'New at Mouser', empty string, and any other non-obsolete status
            lifecycle_upper = lambda p: (p.get('lifecycle') or '').upper()
            filtered = [p for p in filtered 
                       if lifecycle_upper(p) not in ['OBSOLETE', 'EOL', 'END OF LIFE', 'NOT RECOMMENDED FOR NEW DESIGNS', 'END OF LIFE (EOL)']]
            logger.debug(f"Active-only filter: {before_count} -> {len(filtered)} parts")
        
        return filtered

