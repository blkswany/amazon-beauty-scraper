#!/usr/bin/env python3
"""
Amazon Beauty Best Sellers Scraper
6개국 (US/DE/FR/IT/UK/ES) x 2페이지 → Excel (국가별 시트)

Setup:
    pip install playwright pandas openpyxl
    playwright install chromium
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

# ─── URLs ─────────────────────────────────────────────────────────────────────
COUNTRIES = {
    "US": [
        "https://www.amazon.com/Best-Sellers-Beauty-Personal-Care/zgbs/beauty/",
        "https://www.amazon.com/Best-Sellers-Beauty-Personal-Care/zgbs/beauty/ref=zg_bs_pg_2_beauty?_encoding=UTF8&pg=2",
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

# ─── JavaScript ───────────────────────────────────────────────────────────────
JS_EXTRACT = """
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


# ─── Helpers ──────────────────────────────────────────────────────────────────
async def accept_cookies(page):
    """EU 사이트 쿠키 동의 팝업 처리 - 다양한 셀렉터 시도"""
    selectors = [
        "#sp-cc-accept",
        "input#sp-cc-accept",
        "button#sp-cc-accept",
        "[data-cel-widget='sp-cc-accept']",
        "input[name='accept']",
        # 일반적인 EU 쿠키 배너
        "#onetrust-accept-btn-handler",
        ".accept-cookies-button",
        "[id*='cookie'] button",
        "[class*='cookie'] button[class*='accept']",
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


async def scrape_page(page, url: str, rank_offset: int = 0, country: str = "", page_idx: int = 0) -> list:
    print(f"    GET {url}")
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=40000)
    except PlaywrightTimeout:
        print("    타임아웃, 재시도...")
        await page.goto(url, wait_until="domcontentloaded", timeout=40000)

    # 쿠키 팝업 처리
    await accept_cookies(page)

    # "Continue shopping" 인터셉트 페이지 처리 (UK/FR 등)
    try:
        btn = page.locator("input[value='Continue shopping'], button:has-text('Continue shopping')")
        if await btn.is_visible(timeout=3000):
            await btn.click()
            print("    ✓ Continue shopping 클릭")
            await page.wait_for_timeout(2000)
    except Exception:
        pass

    # 잠깐 대기 (팝업 처리 후 페이지 안정화)
    await page.wait_for_timeout(2000)

    # 상품 그리드 로딩 대기
    try:
        await page.wait_for_selector("div.zg-grid-general-faceout", timeout=20000)
    except PlaywrightTimeout:
        # 스크린샷 저장 (디버깅용)
        screenshot_path = f"screenshot_{country}_p{page_idx + 1}.png"
        await page.screenshot(path=screenshot_path, full_page=False)
        print(f"    ⚠ 상품 목록 없음 — 스크린샷 저장됨: {screenshot_path}")
        return []

    # 50개 로드될 때까지 스크롤
    for _ in range(50):
        count = await page.eval_on_selector_all(
            "div.zg-grid-general-faceout", "els => els.length"
        )
        print(f"    scroll... {count}개 로드됨")
        if count >= 50:
            break
        await page.evaluate("window.scrollBy(0, window.innerHeight)")
        await page.wait_for_timeout(800)
    await page.evaluate("window.scrollTo(0, 0)")
    await page.wait_for_timeout(400)

    raw = await page.evaluate(JS_EXTRACT)
    rows = json.loads(raw)
    for row in rows:
        row["rank"] += rank_offset

    print(f"    → {len(rows)}개 추출")
    return rows


async def scrape_country(browser, country: str, urls: list) -> list:
    context = await browser.new_context(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        locale="en-US",
        viewport={"width": 1280, "height": 900},
        java_script_enabled=True,
        extra_http_headers={
            "Accept-Language": "en-US,en;q=0.9",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        },
    )

    # headless 감지 우회
    await context.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
        Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3] });
        window.chrome = { runtime: {} };
    """)

    page = await context.new_page()
    all_rows = []
    rank_offset = 0

    try:
        for i, url in enumerate(urls):
            rows = await scrape_page(page, url, rank_offset, country=country, page_idx=i)
            all_rows.extend(rows)
            rank_offset += len(rows)
            if i < len(urls) - 1:
                delay = random.uniform(4, 7)
                print(f"    {delay:.1f}초 대기 중...")
                await page.wait_for_timeout(int(delay * 1000))
    finally:
        await context.close()

    return all_rows


# ─── Main ─────────────────────────────────────────────────────────────────────
async def main():
    today = datetime.now().strftime("%Y-%m-%d")
    output_path = Path(f"amazon_beauty_bestsellers_{today}.xlsx")

    print("=" * 55)
    print(" Amazon Beauty Best Sellers Scraper")
    print(f" 날짜: {today}")
    print(f" 저장 경로: {output_path}")
    print("=" * 55)

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-blink-features=AutomationControlled",
            ],
        )

        country_data = {}
        countries_list = list(COUNTRIES.keys())

        for country, urls in COUNTRIES.items():
            print(f"\n[{country}] 스크래핑 시작...")
            try:
                rows = await scrape_country(browser, country, urls)
                country_data[country] = rows
                print(f"[{country}] 완료 — 총 {len(rows)}개")
            except Exception as e:
                print(f"[{country}] 오류: {e}")
                country_data[country] = []

            if country != countries_list[-1]:
                delay = random.uniform(5, 10)
                print(f"\n다음 국가까지 {delay:.1f}초 대기...\n")
                await asyncio.sleep(delay)

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
