"""Quick diagnostic to inspect the chat panel state in the running browser."""
import asyncio
import json
from playwright.async_api import async_playwright

JS_DIAG = """() => {
    const panel = document.getElementById("demo-chat-panel");
    const input = document.getElementById("chat-input");
    const send = document.getElementById("chat-send");
    const inputArea = document.querySelector(".chat-input-area");
    const welcome = document.getElementById("chat-welcome");
    const toggle = document.getElementById("demo-chat-toggle");
    function cs(el, prop) { return el ? getComputedStyle(el)[prop] : null; }
    function rect(el) {
        if (!el) return null;
        const r = el.getBoundingClientRect();
        return {x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), h: Math.round(r.height)};
    }
    return {
        windowSize: [window.innerWidth, window.innerHeight],
        bodyMarginRight: cs(document.body, "marginRight"),
        panel: { exists: !!panel, rect: rect(panel), display: cs(panel,"display"), visibility: cs(panel,"visibility"), width: cs(panel,"width"), transform: cs(panel,"transform"), overflow: cs(panel,"overflow"), zIndex: cs(panel,"zIndex"), collapsed: panel ? panel.classList.contains("collapsed") : null },
        input: { exists: !!input, rect: rect(input), display: cs(input,"display"), visibility: cs(input,"visibility"), height: cs(input,"height"), disabled: input ? input.disabled : null },
        inputArea: { exists: !!inputArea, rect: rect(inputArea), display: cs(inputArea,"display"), visibility: cs(inputArea,"visibility") },
        send: { exists: !!send, rect: rect(send), display: cs(send,"display"), visibility: cs(send,"visibility") },
        welcome: { exists: !!welcome, rect: rect(welcome), display: cs(welcome,"display") },
        toggle: { exists: !!toggle, rect: rect(toggle), display: cs(toggle,"display") },
    };
}"""

async def main():
    p = await async_playwright().start()
    try:
        browser = await p.chromium.connect_over_cdp("http://127.0.0.1:9222")
    except Exception:
        print("ERROR: Cannot connect via CDP. The demo agent browser may not have CDP enabled.")
        print("We need to enable --remote-debugging-port=9222 in the browser launch.")
        await p.stop()
        return
    contexts = browser.contexts
    if not contexts:
        print("No browser contexts found")
        await p.stop()
        return
    pages = contexts[0].pages
    print(f"Found {len(pages)} page(s)")
    for i, pg in enumerate(pages):
        print(f"  [{i}] {pg.url[:100]}")
    page = pages[0]
    info = await page.evaluate(JS_DIAG)
    print(json.dumps(info, indent=2))
    await page.screenshot(path="chat_debug_screenshot.png", full_page=False)
    print("Screenshot saved: chat_debug_screenshot.png")
    await p.stop()

asyncio.run(main())
