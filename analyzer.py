import statistics
from typing import Optional
from scraper import Product


class PriceAnalyzer:
    def __init__(self, products: list[Product]):
        self.products = [p for p in products if p.price > 0]

    def _prices(self) -> list[float]:
        return [p.price for p in self.products]

    def summary(self) -> dict:
        if not self.products:
            return {}

        prices = self._prices()
        rated = [p for p in self.products if p.rating is not None]
        discounted = [p for p in self.products if p.discount_pct]
        free_shipping = [p for p in self.products if p.shipping_free]

        return {
            "total": len(self.products),
            "price_min": min(prices),
            "price_max": max(prices),
            "price_avg": round(statistics.mean(prices), 2),
            "price_median": round(statistics.median(prices), 2),
            "price_stdev": round(statistics.stdev(prices), 2) if len(prices) > 1 else 0,
            "free_shipping_count": len(free_shipping),
            "free_shipping_pct": round(len(free_shipping) / len(self.products) * 100, 1),
            "discounted_count": len(discounted),
            "avg_discount_pct": (
                round(statistics.mean(p.discount_pct for p in discounted), 1)
                if discounted else 0
            ),
            "rated_count": len(rated),
            "avg_rating": (
                round(statistics.mean(p.rating for p in rated), 2) if rated else None
            ),
        }

    def top_deals(self, n: int = 5) -> list[Product]:
        def score(p: Product) -> float:
            s = 0.0
            if p.discount_pct:
                s += p.discount_pct * 0.5
            if p.rating:
                s += p.rating * 10
            if p.shipping_free:
                s += 15
            if p.reviews_count and p.reviews_count > 100:
                s += 10
            return s

        return sorted(self.products, key=score, reverse=True)[:n]

    def cheapest(self, n: int = 5) -> list[Product]:
        return sorted(self.products, key=lambda p: p.price)[:n]

    def most_reviewed(self, n: int = 5) -> list[Product]:
        return sorted(
            [p for p in self.products if p.reviews_count],
            key=lambda p: p.reviews_count,
            reverse=True
        )[:n]

    def price_ranges(self) -> dict[str, int]:
        prices = self._prices()
        if not prices:
            return {}

        min_p, max_p = min(prices), max(prices)
        step = (max_p - min_p) / 4 if max_p > min_p else 1
        ranges = {}

        for i in range(4):
            lo = min_p + i * step
            hi = min_p + (i + 1) * step
            label = f"R${lo:,.0f} - R${hi:,.0f}"
            count = sum(1 for p in prices if lo <= p < hi)
            ranges[label] = count

        last_key = list(ranges.keys())[-1]
        ranges[last_key] += sum(1 for p in prices if p == max_p)
        return ranges
