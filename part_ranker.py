"""Advanced part ranking and scoring engine."""
from typing import Dict, Any, List, Optional
import logging

logger = logging.getLogger(__name__)


class RankingEngine:
    """Ranks parts based on configurable scoring criteria."""
    
    def __init__(self, stock_weight: float = 0.3, price_weight: float = 0.5, 
                 lifecycle_weight: float = 0.1, package_match_weight: float = 0.1):
        """
        Initialize ranking engine with weights.
        
        Args:
            stock_weight: Weight for stock availability (default 0.5 - highest priority)
            price_weight: Weight for price competitiveness (default 0.3)
            lifecycle_weight: Weight for lifecycle status (default 0.1)
            package_match_weight: Weight for package/footprint matching (default 0.1)
        """
        self.stock_weight = stock_weight
        self.price_weight = price_weight
        self.lifecycle_weight = lifecycle_weight
        self.package_match_weight = package_match_weight
        
        # Normalize weights to sum to 1.0
        total = stock_weight + price_weight + lifecycle_weight + package_match_weight
        if total > 0:
            self.stock_weight /= total
            self.price_weight /= total
            self.lifecycle_weight /= total
            self.package_match_weight /= total
    
    def normalize_package(self, package: str) -> str:
        """Normalize package string for comparison."""
        if not package:
            return ""
        # Remove common suffixes and normalize
        package = package.upper().strip()
        # Remove spaces, dashes, underscores
        package = package.replace(' ', '').replace('-', '').replace('_', '')
        return package
    
    def packages_match(self, package1: str, package2: str) -> bool:
        """Check if two package strings match (with normalization)."""
        if not package1 or not package2:
            return False
        
        norm1 = self.normalize_package(package1)
        norm2 = self.normalize_package(package2)
        
        # Exact match
        if norm1 == norm2:
            return True
        
        # Check if one contains the other (for cases like "0603" vs "0603-0805")
        if norm1 in norm2 or norm2 in norm1:
            return True
        
        return False
    
    def score_stock(self, part: Dict[str, Any]) -> float:
        """Score based on stock availability (0-100)."""
        stock = part.get('stock', 0)
        
        if stock <= 0:
            return 0.0
        elif stock >= 10000:
            return 100.0
        elif stock >= 1000:
            return 80.0
        elif stock >= 100:
            return 60.0
        elif stock >= 10:
            return 40.0
        else:
            return 20.0
    
    def score_price(self, part: Dict[str, Any]) -> float:
        """Score based on price competitiveness (0-100, lower price = higher score)."""
        price_breaks = part.get('price_breaks', [])
        
        if not price_breaks:
            return 50.0  # Neutral score if no price data
        
        # Use the lowest quantity price break (typically 1 unit)
        # Lower price = higher score
        min_price = None
        for pb in price_breaks:
            try:
                price_str = pb.get('price', '').replace('$', '').replace(',', '').strip()
                price = float(price_str)
                if min_price is None or price < min_price:
                    min_price = price
            except (ValueError, AttributeError):
                continue
        
        if min_price is None:
            return 50.0
        
        # Normalize price score (assuming typical range 0.01 to 100.00)
        # Lower prices get higher scores
        if min_price <= 0.01:
            return 100.0
        elif min_price <= 0.10:
            return 90.0
        elif min_price <= 1.00:
            return 70.0
        elif min_price <= 10.00:
            return 50.0
        elif min_price <= 50.00:
            return 30.0
        else:
            return 10.0
    
    def score_lifecycle(self, part: Dict[str, Any]) -> float:
        """Score based on lifecycle status (0-100)."""
        lifecycle = part.get('lifecycle', '').upper()
        
        if not lifecycle:
            return 50.0  # Neutral if unknown
        
        # Active parts get highest score
        if lifecycle in ['ACTIVE', 'LIFEBUY', 'NEW']:
            return 100.0
        elif lifecycle in ['LAST TIME BUY', 'NOT RECOMMENDED FOR NEW DESIGNS']:
            return 30.0
        elif lifecycle in ['OBSOLETE', 'EOL', 'END OF LIFE']:
            return 0.0
        else:
            return 50.0  # Unknown status
    
    def score_package_match(self, part: Dict[str, Any], target_package: str) -> float:
        """Score based on package/footprint matching (0-100)."""
        if not target_package:
            return 50.0  # Neutral if no target package specified
        
        part_package = part.get('package', '')
        if not part_package:
            return 0.0  # No package info = no match
        
        if self.packages_match(part_package, target_package):
            return 100.0
        else:
            return 0.0
    
    def calculate_score(self, part: Dict[str, Any], target_package: Optional[str] = None) -> float:
        """
        Calculate overall score for a part.
        
        Args:
            part: Part dictionary from Mouser API
            target_package: Target package/footprint to match against
            
        Returns:
            Overall score (0-100)
        """
        stock_score = self.score_stock(part)
        price_score = self.score_price(part)
        lifecycle_score = self.score_lifecycle(part)
        package_score = self.score_package_match(part, target_package)
        
        total_score = (
            self.stock_weight * stock_score +
            self.price_weight * price_score +
            self.lifecycle_weight * lifecycle_score +
            self.package_match_weight * package_score
        )
        
        # Log detailed scoring info for debugging
        logger.debug(
            f"Part {part.get('mpn')}: "
            f"Stock={part.get('stock')} (Score={stock_score:.1f}), "
            f"Price={part.get('price_breaks', [{}])[0].get('price', 'N/A')} (Score={price_score:.1f}), "
            f"Lifecycle={part.get('lifecycle')} (Score={lifecycle_score:.1f}), "
            f"Package={part.get('package')} vs {target_package} (Score={package_score:.1f}) -> "
            f"Total={total_score:.1f}"
        )
        
        return round(total_score, 2)
    
    def rank_parts(self, parts: List[Dict[str, Any]], 
                   target_package: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        Rank parts by score (highest first).
        
        Args:
            parts: List of part dictionaries
            target_package: Target package/footprint for matching
            
        Returns:
            Sorted list of parts with 'score' field added
        """
        # Calculate scores
        for part in parts:
            part['score'] = self.calculate_score(part, target_package)
        
        # Sort by score (descending), then by stock (descending) as tiebreaker
        sorted_parts = sorted(
            parts,
            key=lambda p: (p.get('score', 0), p.get('stock', 0)),
            reverse=True
        )
        
        return sorted_parts
    
    def get_top_parts(self, parts: List[Dict[str, Any]], 
                     target_package: Optional[str] = None,
                     limit: int = 10) -> List[Dict[str, Any]]:
        """
        Get top N ranked parts.
        
        Args:
            parts: List of part dictionaries
            target_package: Target package/footprint for matching
            limit: Maximum number of parts to return
            
        Returns:
            Top N ranked parts
        """
        ranked = self.rank_parts(parts, target_package)
        return ranked[:limit]

