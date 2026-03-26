"""Take a screenshot via CDP."""
import asyncio
import json
import base64
import websockets

async def main():
    import urllib.request
    resp = urllib.request.urlopen("http://127.0.0.1:9222/json")
    pages = json.loads(resp.read())
    ws_url = pages[0]["webSocketDebuggerUrl"]

    async with websockets.connect(ws_url, max_size=20_000_000) as ws:
        cmd = {"id": 1, "method": "Page.captureScreenshot", "params": {"format": "png"}}
        await ws.send(json.dumps(cmd))
        result = await ws.recv()
        data = json.loads(result)
        img_data = base64.b64decode(data["result"]["data"])
        with open("chat_panel_screenshot.png", "wb") as f:
            f.write(img_data)
        print(f"Screenshot saved: chat_panel_screenshot.png ({len(img_data)} bytes)")

asyncio.run(main())
