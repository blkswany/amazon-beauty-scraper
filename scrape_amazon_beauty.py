#!/usr/bin/env python3
"""
Amazon & Rakuten Beauty Best Sellers Scraper
아마존 6개국 + 미국 하위카테고리 + 라쿠텐 + Qoo10
"""

import asyncio
import json
import random
import sys
from datetime import datetime
from pathlib import Path

# Windows terminal encoding fix
if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

import pandas as pd
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout
from playwright_stealth import Stealth

# ─── URLs ─────────────────────────────────────────────────────────────────────
AMAZON_COUNTRIES = {
    "US": [
        "https://www.amazon.com/Best-Sellers-Beauty-Personal-Care/zgbs/beauty/",
        "https://www.amazon.com/Best-Sellers-Beauty-Personal-Care/zgbs/beauty/ref=zg_bs_pg_2_beauty?_encoding=UTF8&pg=2",
    ],
    "US_SkinCare": [
        "https://www.amazon.com/Best-Sellers-Beauty-Personal-Care-Skin-Care-Products/zgbs/beauty/11060451/ref=zg_bs_pg_1_beauty?_encoding=UTF8&pg=1",
        "https://www.amazon.com/Best-Sellers-Beauty-Personal-Care-Skin-Care-Products/zgbs/beauty/11060451/ref=zg_bs_pg_2_beauty?_encoding=UTF8&pg=2",
    ],
    "US_SunCare": [
        "https://www.amazon.com/Best-Sellers-Beauty-Personal-Care-Sun-Skin-Care/zgbs/beauty/11062591/ref=zg_bs_pg_1_beauty?_encoding=UTF8&pg=1",
        "https://www.amazon.com/Best-Sellers-Beauty-Personal-Care-Sun-Skin-Care/zgbs/beauty/11062591/ref=zg_bs_pg_2_beauty?_encoding=UTF8&pg=2",
    ],
    "DE": [
        "https://www.amazon.de/-/en/gp/bestsellers/beauty/ref=zg_bs_pg_1_beauty?ie=UTF8&pg=1",
        "https://www.amazon.de/-/en/gp/bestsellers/beauty/ref=zg_bs_pg_2_beauty?ie=UTF8&pg=2",
    ],
    "FR": [
        "https://www.amazon.fr/-/en/gp/bestsellers/beauty/ref=zg_bs_pg_1_beauty?ie=UTF8&pg=1",
        "https://www.amazon.fr/-/en/gp/bestsellers/beauty/ref=zg_bs_pg_2_beauty?ie=UTF8&pg=2",
    ],
    "IT": [
        "https://www.amazon.it/-/en/gp/bestsellers/beauty/ref=zg_bs_pg_1_beauty?ie=UTF8&pg=1",
        "https://www.amazon.it/-/en/gp/bestsellers/beauty/ref=zg_bs_pg_2_beauty?ie=UTF8&pg=2",
    ],
    "UK": [
        "https://www.amazon.co.uk/Best-Sellers-Beauty/zgbs/beauty/",
        "https://www.amazon.co.uk/Best-Sellers-Beauty/zgbs/beauty/ref=zg_bs_pg_2_beauty?_encoding=UTF8&pg=2",
    ],
    "ES": [
        "https://www.amazon.es/-/en/gp/bestsellers/beauty/ref=zg_bs_pg_1_beauty?ie=UTF8&pg=1",
        "https://www.amazon.es/-/en/gp/bestsellers/beauty/ref=zg_bs_pg_2_beauty?ie=UTF8&pg=2",
    ],
}

RAKUTEN_COUNTRIES = {
    "Rakuten_JP": [
        "https://ranking.rakuten.co.jp/daily/100939/"
    ]
}

QOO10_COUNTRIES = {
    "Qoo10_JP": [
        "https://www.qoo10.jp/gmkt.inc/Bestsellers/?g=2"
    ]
}

# ─── JavaScript ───────────────────────────────────────────────────────────────
# 아마존용 추출 스크립트
JS_EXTRACT_AMAZON = """
() => {
    const items = [...document.querySelectorAll("div.zg-grid-general-faceout")];
    const rows = [];
    items.forEach((el, i) => {
        const titleEl = el.querySelector("div._cDEzb_p13n-sc-css-line-clamp-3_g3dy1")
                     || el.querySelector("div._cDEzb_p13n-sc-css-line-clamp-1_1tdkm")
                     || el.querySelector("span.p13n-sc-truncate-desktop-type2");
        const title = titleEl ? titleEl.innerText.trim() : "";

        const reviewLink = el.querySelector("a[aria-label*='stars']")
                        || el.querySelector("a[aria-label*='out of']");
        const ariaLabel = reviewLink ? reviewLink.getAttribute("aria-label") : "";
        const rating = (ariaLabel.match(/([0-9.]+) out of/) || [])[1] || "";

        const reviewsEl = el.querySelector("span.a-size-small[aria-hidden='true']");
        const reviews = reviewsEl ? reviewsEl.innerText.trim() : "";

        rows.push({ rank: i + 1, title, rating, reviews });
    });
    return JSON.stringify(rows);
}
"""

# 라쿠텐용 추출 스크립트
JS_EXTRACT_RAKUTEN = """
() => {
    const rows = [];
    const nameBoxes = [...document.querySelectorAll(".rnkRanking_itemName")];
    
    nameBoxes.forEach((nameBox, i) => {
        const title = nameBox.innerText.trim();
        if (!title) return;
        
        let parent = nameBox.parentElement;
        for(let j=0; j<5; j++) { if(parent && parent.parentElement) parent = parent.parentElement; }
        
        let reviews = "";
        if (parent) {
            const reviewLink = parent.querySelector('a[href*="review.rakuten.co.jp"]');
            if (reviewLink) {
                const match = reviewLink.innerText.match(/([0-9,]+)/);
                if (match) reviews = match[1];
            }
        }
        
        let rating = "";
        if (parent) {
            const onStars = parent.querySelectorAll('.rnkRanking_starON').length;
            const halfStars = parent.querySelectorAll('.rnkRanking_starHALF').length;
            if (onStars > 0 || halfStars > 0) {
                rating = (onStars + halfStars * 0.5).toFixed(1);
            }
        }
        
        rows.push({ rank: i + 1, title, rating, reviews });
    });
    return JSON.stringify(rows);
}
"""

# Qoo10용 추출 스크립트
JS_EXTRACT_QOO10 = r"""
() => {
    const rows = [];
    const items = [...document.querySelectorAll("div.item")];
    
    items.forEach((item, i) => {
        const titleEl = item.querySelector("a.tt");
        const title = titleEl ? titleEl.innerText.trim() : "";
        if (!title) return;
        
        const reviewCountEl = item.querySelector("span.review_total_count");
        let reviews = reviewCountEl ? reviewCountEl.innerText.replace(/[^0-9]/g, '') : "";
        
        const ratingEl = item.querySelector("div.review_rating_star");
        let rating = "";
        if (ratingEl) {
            const style = ratingEl.getAttribute("style");
            if (style) {
                const match = style.match(/width:\s*([0-9.]+)%/);
                if (match) {
                    // Qoo10 uses width percentage for rating. We could return the percentage or convert to 5-star format.
                    // Converting width to 5-star scale (e.g. 100% -> 5.0).
                    const percent = parseFloat(match[1]);
                    // Sometimes Qoo10 might use weird multipliers (like 1920% in your example?), 
                    // Let's just grab what we can, or just keep the raw value.
                    // Given the 1920% anomaly, maybe let's just return the raw style match for safety,
                    // or just standard percent/20. Let's return raw % for now or leave empty if weird.
                    rating = match[1] + "%"; 
                }
            }
        }
        
        rows.push({ rank: i + 1, title, rating, reviews });
    });
    return JSON.stringify(rows);
}
"""


# ─── Helpers (Amazon) ─────────────────────────────────────────────────────────
async def accept_cookies(page):
    selectors = [
        "#sp-cc-accept", "input#sp-cc-accept", "button#sp-cc-accept",
        "[data-cel-widget='sp-cc-accept']", "input[name='accept']",
        "#onetrust-accept-btn-handler", ".accept-cookies-button",
    ]
    for selector in selectors:
        try:
            btn = page.locator(selector)
            if await btn.is_visible(timeout=2000):
                await btn.click()
                await page.wait_for_timeout(1500)
                print(f"    ✓ 쿠키 동의 완료 ({selector})")
                return True
        except Exception:
            pass
    return False


async def scrape_amazon_page(page, url: str, rank_offset: int = 0, country: str = "", page_idx: int = 0) -> list:
    print(f"    GET {url}")
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=40000)
    except PlaywrightTimeout:
        print("    타임아웃, 재시도...")
        await page.goto(url, wait_until="domcontentloaded", timeout=40000)

    await accept_cookies(page)

    try:
        has_products = await page.locator("div.zg-grid-general-faceout").count()
        if has_products == 0:
            btn = page.locator("input[type='submit'], button[type='submit']").first
            if await btn.is_visible(timeout=3000):
                await btn.click()
                await page.wait_for_load_state("domcontentloaded")
                await page.wait_for_timeout(2000)
    except Exception:
        pass

    await page.wait_for_timeout(2000)

    try:
        await page.wait_for_selector("div.zg-grid-general-faceout", timeout=20000)
    except PlaywrightTimeout:
        screenshot_path = f"screenshot_{country}_p{page_idx + 1}.png"
        await page.screenshot(path=screenshot_path, full_page=False)
        print(f"    ⚠ 상품 목록 없음 — 스크린샷 저장됨: {screenshot_path}")
        return []

    for _ in range(50):
        count = await page.eval_on_selector_all("div.zg-grid-general-faceout", "els => els.length")
        if count >= 50:
            break
        await page.evaluate("window.scrollBy(0, window.innerHeight)")
        await page.wait_for_timeout(800)
        
    await page.evaluate("window.scrollTo(0, 0)")
    await page.wait_for_timeout(400)

    raw = await page.evaluate(JS_EXTRACT_AMAZON)
    rows = json.loads(raw)
    for row in rows:
        row["rank"] += rank_offset

    print(f"    → {len(rows)}개 추출")
    return rows


async def scrape_amazon_country(browser, country: str, urls: list) -> list:
    context = await browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        locale="en-US", viewport={"width": 1280, "height": 900}, java_script_enabled=True,
    )
    await context.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
    """)

    page = await context.new_page()
    all_rows = []
    rank_offset = 0

    try:
        for i, url in enumerate(urls):
            rows = await scrape_amazon_page(page, url, rank_offset, country=country, page_idx=i)
            all_rows.extend(rows)
            rank_offset += len(rows)
            if i < len(urls) - 1:
                delay = random.uniform(4, 7)
                await page.wait_for_timeout(int(delay * 1000))
    finally:
        await context.close()

    return all_rows


# ─── Helpers (Rakuten) ────────────────────────────────────────────────────────
async def scrape_rakuten_target(browser, country: str, urls: list) -> list:
    context = await browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        locale="ja-JP", viewport={"width": 1920, "height": 1080}, java_script_enabled=True,
        extra_http_headers={
            "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Sec-Ch-Ua": '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
            "Sec-Ch-Ua-Mobile": "?0",
            "Sec-Ch-Ua-Platform": '"Windows"',
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "none",
            "Sec-Fetch-User": "?1",
            "Upgrade-Insecure-Requests": "1"
        }
    )
    
    # 봇 탐지 우회
    await context.add_init_script("Object.defineProperty(navigator, 'webdriver', { get: () => undefined });")

    page = await context.new_page()
    await Stealth().apply_stealth_async(page)
    all_rows = []

    try:
        for url in urls:
            print(f"    GET {url}")
            await page.goto(url, wait_until="domcontentloaded", timeout=40000)
            
            # 상품 목록이 뜰 때까지 최대 15초 대기
            try:
                await page.wait_for_selector(".rnkRanking_itemName", timeout=15000)
            except:
                print(f"    ⚠ 라쿠텐 상품 목록 로딩 실패 (차단 가능성)")
                continue

            # 스크롤
            for _ in range(15):
                await page.evaluate("window.scrollBy(0, window.innerHeight)")
                await page.wait_for_timeout(600)
                
            raw = await page.evaluate(JS_EXTRACT_RAKUTEN)
            rows = json.loads(raw)
            print(f"    → {len(rows)}개 추출")
            all_rows.extend(rows)
            
    finally:
        await context.close()

    return all_rows


async def scrape_qoo10_target(browser, country: str, urls: list) -> list:
    context = await browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        locale="ja-JP", viewport={"width": 1920, "height": 1080}, java_script_enabled=True,
    )
    
    # 봇 탐지 우회
    await context.add_init_script("Object.defineProperty(navigator, 'webdriver', { get: () => undefined });")

    page = await context.new_page()
    await Stealth().apply_stealth_async(page)
    all_rows = []

    try:
        for url in urls:
            print(f"    GET {url}")
            await page.goto(url, wait_until="domcontentloaded", timeout=40000)
            
            # 상품 목록이 뜰 때까지 최대 15초 대기
            try:
                await page.wait_for_selector("div.item", timeout=15000)
            except:
                print(f"    ⚠ Qoo10 상품 목록 로딩 실패 (차단 가능성)")
                continue

            # 스크롤 (lazy load 방지)
            for _ in range(15):
                await page.evaluate("window.scrollBy(0, window.innerHeight)")
                await page.wait_for_timeout(600)
                
            raw = await page.evaluate(JS_EXTRACT_QOO10)
            rows = json.loads(raw)
            print(f"    → {len(rows)}개 추출")
            all_rows.extend(rows)
            
    finally:
        await context.close()

    return all_rows

# ─── Main ─────────────────────────────────────────────────────────────────────
async def main():
    today = datetime.now().strftime("%Y-%m-%d")
    output_path = Path(f"amazon_beauty_bestsellers_{today}.xlsx")

    print("=" * 55)
    print(" Amazon & Rakuten Beauty Best Sellers Scraper")
    print(f" 날짜: {today}")
    print(f" 저장 경로: {output_path}")
    print("=" * 55)

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-blink-features=AutomationControlled"],
        )

        country_data = {}
        
        # 1. 아마존 국가 및 카테고리 스크래핑
        for country, urls in AMAZON_COUNTRIES.items():
            print(f"\n[아마존 - {country}] 스크래핑 시작...")
            try:
                rows = await scrape_amazon_country(browser, country, urls)
                country_data[country] = rows
                print(f"[{country}] 완료 — 총 {len(rows)}개")
            except Exception as e:
                print(f"[{country}] 오류: {e}")
                country_data[country] = []
            
            await asyncio.sleep(random.uniform(3, 6))

        # 2. 라쿠텐 스크래핑
        for country, urls in RAKUTEN_COUNTRIES.items():
            print(f"\n[라쿠텐 - {country}] 스크래핑 시작...")
            try:
                rows = await scrape_rakuten_target(browser, country, urls)
                country_data[country] = rows
                print(f"[{country}] 완료 — 총 {len(rows)}개")
            except Exception as e:
                print(f"[{country}] 오류: {e}")
                country_data[country] = []

        # 3. Qoo10 스크래핑
        for country, urls in QOO10_COUNTRIES.items():
            print(f"\n[Qoo10 - {country}] 스크래핑 시작...")
            try:
                rows = await scrape_qoo10_target(browser, country, urls)
                country_data[country] = rows
                print(f"[{country}] 완료 — 총 {len(rows)}개")
            except Exception as e:
                print(f"[{country}] 오류: {e}")
                country_data[country] = []

        await browser.close()

    # ─── Excel 저장 ───────────────────────────────────────────────────────────
    print(f"\nExcel 파일 저장 중: {output_path}")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for country, rows in country_data.items():
            df = pd.DataFrame(rows if rows else [], columns=["rank", "title", "rating", "reviews"])
            df.to_excel(writer, sheet_name=country, index=False)
            print(f"  시트 '{country}': {len(df)}행")

    print(f"\n완료! → {output_path.resolve()}")


if __name__ == "__main__":
    asyncio.run(main())
