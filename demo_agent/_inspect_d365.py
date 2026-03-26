"""Inspect a live D365 page to discover real selectors."""
import asyncio
from playwright.async_api import async_playwright


async def inspect():
    p = await async_playwright().start()
    browser = await p.chromium.launch(headless=False, slow_mo=50)
    ctx = await browser.new_context(
        storage_state="auth_state.json",
        viewport={"width": 1920, "height": 1080},
        no_viewport=True,
        ignore_https_errors=True,
    )
    page = await ctx.new_page()
    page.set_default_timeout(30000)
    page.set_default_navigation_timeout(60000)
    await page.goto(
        "https://projectopscoreagentimplemented.crm.dynamics.com",
        wait_until="domcontentloaded",
    )
    await page.wait_for_timeout(8000)

    print("=" * 60)
    print("CURRENT URL:", page.url)
    print("PAGE TITLE:", await page.title())
    print("=" * 60)

    # --- Left nav / sitemap items ---
    print("\n--- SITEMAP / LEFT NAV ---")
    sitemap_els = await page.query_selector_all(
        '[data-id*="sitemap"], [data-lp-id*="sitemap"]'
    )
    print(f"Sitemap elements: {len(sitemap_els)}")
    for el in sitemap_els[:20]:
        tag = await el.evaluate("el => el.tagName")
        data_id = await el.get_attribute("data-id") or ""
        aria = await el.get_attribute("aria-label") or ""
        text = (await el.inner_text()).strip().replace("\n", " ")[:50]
        print(f"  {tag} data-id={data_id!r} aria={aria!r} text={text!r}")

    # --- Nav list items ---
    print("\n--- NAV LI ITEMS ---")
    nav_items = await page.query_selector_all('li[role="treeitem"], li[role="listitem"]')
    print(f"Nav li items: {len(nav_items)}")
    for item in nav_items[:20]:
        label = await item.get_attribute("aria-label") or ""
        data_id = await item.get_attribute("data-id") or ""
        text = (await item.inner_text()).strip().replace("\n", " ")[:50]
        print(f"  li aria-label={label!r} data-id={data_id!r} text={text!r}")

    # --- Buttons ---
    print("\n--- BUTTONS ---")
    buttons = await page.query_selector_all("button[aria-label]")
    print(f"Buttons: {len(buttons)}")
    for btn in buttons[:20]:
        label = await btn.get_attribute("aria-label") or ""
        data_id = await btn.get_attribute("data-id") or ""
        visible = await btn.is_visible()
        if visible:
            print(f"  button aria-label={label!r} data-id={data_id!r}")

    # --- Try navigating to time entries ---
    print("\n--- TRYING ENTITY LIST URL ---")
    url = "https://projectopscoreagentimplemented.crm.dynamics.com/main.aspx?etn=msdyn_timeentry&pagetype=entitylist"
    await page.goto(url, wait_until="domcontentloaded")
    await page.wait_for_timeout(5000)
    print("After nav URL:", page.url)
    print("Title:", await page.title())

    # Check for New button
    new_btns = await page.query_selector_all('button[aria-label*="New"], button[data-id*="new"]')
    print(f"'New' buttons found: {len(new_btns)}")
    for btn in new_btns:
        label = await btn.get_attribute("aria-label") or ""
        data_id = await btn.get_attribute("data-id") or ""
        visible = await btn.is_visible()
        print(f"  button aria-label={label!r} data-id={data_id!r} visible={visible}")

    await browser.close()
    await p.stop()


if __name__ == "__main__":
    asyncio.run(inspect())
