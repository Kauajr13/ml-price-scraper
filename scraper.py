import requests
from bs4 import BeautifulSoup
import time
import random
import logging
from dataclasses import dataclass, field
from typing import Optional
from datetime import datetime

logger = logging.getLogger(__name__)


@dataclass
class Product:
    title: str
    price: float
    original_price: Optional[float]
    discount_pct: Optional[float]
    rating: Optional[float]
    reviews_count: Optional[int]
    sold_count: Optional[str]
    shipping_free: bool
    condition: str
    url: str
    seller: Optional[str]
    scraped_at: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"))


class MLScraper:
    BASE_URL = "https://lista.mercadolivre.com.br"
    HEADERS = {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/122.0.0.0 Safari/537.36"
        ),
        "Accept-Language": "pt-BR,pt;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }

    def __init__(self, delay_min: float = 1.5, delay_max: float = 3.5):
        self.delay_min = delay_min
        self.delay_max = delay_max
        self.session = requests.Session()
        self.session.headers.update(self.HEADERS)

    def _get_page(self, url: str) -> Optional[BeautifulSoup]:
        try:
            response = self.session.get(url, timeout=15)
            response.raise_for_status()
            return BeautifulSoup(response.text, "lxml")
        except requests.RequestException as e:
            logger.error(f"Failed to fetch {url}: {e}")
            return None

    def _parse_price(self, text: str) -> Optional[float]:
        if not text:
            return None
        try:
            return float(text.replace("R$", "").replace(".", "").replace(",", ".").strip())
        except ValueError:
            return None

    def _parse_product(self, item) -> Optional[Product]:
        try:
            title_el = item.select_one("h2.poly-box a, a.poly-component__title")
            title = title_el.get_text(strip=True) if title_el else None
            if not title:
                return None

            url_el = item.select_one("a.poly-component__title, h2.poly-box a")
            url = url_el["href"].split("?")[0] if url_el and url_el.get("href") else ""

            # preço atual fica dentro de poly-price__current
            price_el = item.select_one("div.poly-price__current span.andes-money-amount__fraction")
            if not price_el:
                price_el = item.select_one("span.andes-money-amount__fraction")
            price = self._parse_price(price_el.get_text() if price_el else "")
            if price is None:
                return None

            orig_el = item.select_one("s span.andes-money-amount__fraction")
            original_price = self._parse_price(orig_el.get_text() if orig_el else "")

            discount_el = item.select_one("span.andes-money-amount__discount")
            discount_pct = None
            if discount_el:
                try:
                    # texto pode ser "32% OFF no Pix" ou "-32%"
                    raw = discount_el.get_text(strip=True)
                    number = raw.split("%")[0].replace("-", "").strip()
                    discount_pct = float(number)
                except (ValueError, IndexError):
                    pass

            # rating fica em poly-component__review-compacted > span.poly-phrase-label
            rating_el = item.select_one("span.poly-component__review-compacted span.poly-phrase-label")
            rating = None
            if rating_el:
                try:
                    rating = float(rating_el.get_text(strip=True).replace(",", "."))
                except ValueError:
                    pass

            # reviews ficam como "(312)" dentro do review-compacted, após o rating
            reviews_count = None
            review_block = item.select_one("span.poly-component__review-compacted")
            if review_block:
                labels = review_block.select("span.poly-phrase-label")
                # segundo label costuma ser o total entre parênteses
                for label in labels[1:]:
                    txt = label.get_text(strip=True).strip("()")
                    try:
                        reviews_count = int(txt.replace(".", "").replace(",", ""))
                        break
                    except ValueError:
                        continue

            sold_el = item.select_one("span.poly-component__sold")
            sold_count = sold_el.get_text(strip=True) if sold_el else None

            # frete fica em div.poly-component__shipping > span (filho direto)
            shipping_free = False
            shipping_div = item.select_one("div.poly-component__shipping")
            if shipping_div:
                shipping_text = shipping_div.get_text(strip=True).lower()
                shipping_free = "grátis" in shipping_text or "gratis" in shipping_text

            cond_el = item.select_one("span.poly-component__condition")
            condition = cond_el.get_text(strip=True) if cond_el else "Novo"

            seller_el = item.select_one("span.poly-component__seller")
            seller = seller_el.get_text(strip=True) if seller_el else None

            return Product(
                title=title,
                price=price,
                original_price=original_price,
                discount_pct=discount_pct,
                rating=rating,
                reviews_count=reviews_count,
                sold_count=sold_count,
                shipping_free=shipping_free,
                condition=condition,
                url=url,
                seller=seller,
            )
        except Exception as e:
            logger.debug(f"Skipped item: {e}")
            return None

    def search(self, keyword: str, max_pages: int = 3) -> list[Product]:
        products = []
        slug = keyword.replace(" ", "-")
        base_url = f"{self.BASE_URL}/{slug}"

        logger.info(f"Searching: '{keyword}' ({max_pages} page(s))")

        for page in range(1, max_pages + 1):
            if page == 1:
                url = base_url
            else:
                url = f"{base_url}_Desde_{(page - 1) * 48 + 1}_NoIndex_True"

            logger.info(f"  Page {page}: {url}")
            soup = self._get_page(url)
            if soup is None:
                break

            items = soup.select("li.ui-search-layout__item, div.poly-card")
            if not items:
                logger.info(f"  No items on page {page}, stopping.")
                break

            page_products = [p for item in items if (p := self._parse_product(item))]
            logger.info(f"  {len(page_products)} products extracted")
            products.extend(page_products)

            if page < max_pages:
                delay = random.uniform(self.delay_min, self.delay_max)
                logger.info(f"  Waiting {delay:.1f}s...")
                time.sleep(delay)

        logger.info(f"Total: {len(products)} products collected")
        return products
