# ml-price-scraper

Scrapes product listings from Mercado Livre and exports the results to a formatted Excel report.

Collects price, discount, rating, shipping, seller, and sales data across multiple pages, then generates a spreadsheet with per-product details, statistical summary, and a ranked list of best deals.

## Setup

```bash
git clone https://github.com/Kauajr13/ml-price-scraper.git
cd ml-price-scraper
pip install -r requirements.txt
```

## Usage

```bash
python main.py "notebook gamer"
python main.py "iphone 15" --pages 2
python main.py "monitor 27" --pages 3 --output ./results
python main.py "teclado mecanico" --no-excel
python main.py "headset" --verbose
```

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `--pages N` | 2 | Number of result pages to scrape |
| `--output DIR` | output/ | Directory for Excel files |
| `--no-excel` | false | Print summary only, skip export |
| `--verbose` | false | Enable debug logging |

## Output

Running the script produces a terminal summary and an `.xlsx` file with three sheets:

**Produtos** — full product table with price formatting, conditional highlighting for discounts and free shipping, and column filters.

**Resumo** — statistical breakdown: min/max/avg/median price, standard deviation, free shipping and discount rates, average rating, and a bar chart showing price distribution.

**Top Deals** — top 10 products ranked by a composite score (discount weight + rating + free shipping bonus + review volume).

## Project structure

```
ml-price-scraper/
├── main.py
├── requirements.txt
├── output/
└── src/
    ├── scraper.py    # requests + BeautifulSoup, session-based with random delay
    ├── analyzer.py   # statistics and scoring
    └── exporter.py   # openpyxl report generation
```

## Stack

- Python 3.11+
- requests, beautifulsoup4, lxml
- openpyxl

## Notes

Use responsibly. Add delays between requests (already configured by default) and respect Mercado Livre's terms of service. This project is intended for price research and portfolio demonstration.

## Author

Kauã Jr — [github.com/Kauajr13](https://github.com/Kauajr13)
