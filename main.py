#!/usr/bin/env python3

import sys
import os
import argparse
import logging

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from scraper import MLScraper
from analyzer import PriceAnalyzer
from exporter import export


def setup_logging(verbose: bool = False):
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S",
    )


def print_results(analyzer: PriceAnalyzer, keyword: str):
    s = analyzer.summary()
    sep = "-" * 48

    print(f"\n{'=' * 48}")
    print(f"  {keyword.upper()}")
    print(f"{'=' * 48}")
    print(f"  Products found     : {s.get('total', 0)}")
    print(sep)
    print(f"  Min price          : R$ {s.get('price_min', 0):>10,.2f}")
    print(f"  Max price          : R$ {s.get('price_max', 0):>10,.2f}")
    print(f"  Average            : R$ {s.get('price_avg', 0):>10,.2f}")
    print(f"  Median             : R$ {s.get('price_median', 0):>10,.2f}")
    print(sep)
    print(f"  Free shipping      : {s.get('free_shipping_count', 0)} ({s.get('free_shipping_pct', 0):.1f}%)")
    print(f"  Discounted         : {s.get('discounted_count', 0)} (avg {s.get('avg_discount_pct', 0):.1f}%)")
    print(f"  Avg rating         : {s.get('avg_rating') or 'N/A'}")
    print(f"{'=' * 48}")

    deals = analyzer.top_deals(3)
    if deals:
        print("\n  Top 3 deals:")
        for i, p in enumerate(deals, 1):
            print(f"  {i}. {p.title[:55]}")
            parts = [f"R$ {p.price:,.2f}"]
            if p.discount_pct:
                parts.append(f"-{p.discount_pct:.0f}%")
            if p.rating:
                parts.append(f"rating {p.rating}")
            if p.shipping_free:
                parts.append("free shipping")
            print(f"     {' | '.join(parts)}")
    print()


def main():
    parser = argparse.ArgumentParser(
        description="Scrapes product prices from Mercado Livre and exports to Excel.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
examples:
  python main.py "notebook gamer"
  python main.py "iphone 15" --pages 2
  python main.py "monitor 27" --pages 3 --output ./results
  python main.py "teclado mecanico" --no-excel
        """
    )
    parser.add_argument("keyword", help="Search term")
    parser.add_argument("--pages", type=int, default=2, metavar="N",
                        help="Number of pages to scrape (default: 2)")
    parser.add_argument("--output", default="output", metavar="DIR",
                        help="Output directory for Excel files (default: output/)")
    parser.add_argument("--no-excel", action="store_true",
                        help="Print summary only, skip Excel export")
    parser.add_argument("--verbose", action="store_true")

    args = parser.parse_args()
    setup_logging(args.verbose)

    print(f"\nml-price-scraper")
    print(f"Keyword: '{args.keyword}' | Pages: {args.pages}\n")

    scraper = MLScraper()
    products = scraper.search(args.keyword, max_pages=args.pages)

    if not products:
        print("No products found. Try a different search term.")
        sys.exit(1)

    analyzer = PriceAnalyzer(products)
    print_results(analyzer, args.keyword)

    if not args.no_excel:
        output_file = export(products, analyzer, args.keyword, args.output)
        print(f"Saved: {output_file}\n")


if __name__ == "__main__":
    main()
