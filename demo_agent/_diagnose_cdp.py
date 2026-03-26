"""Quick CDP diagnostic — uses websocket to check Shadow DOM chat panel."""
import asyncio
import json
import websockets

async def main():
    # Get the page websocket URL
    import urllib.request
    resp = urllib.request.urlopen("http://127.0.0.1:9222/json")
    pages = json.loads(resp.read())
    ws_url = pages[0]["webSocketDebuggerUrl"]
    print(f"Connecting to {ws_url[:80]}...")

    async with websockets.connect(ws_url, max_size=10_000_000) as ws:
        # Execute JS to check the shadow DOM panel
        cmd = {
            "id": 1,
            "method": "Runtime.evaluate",
            "params": {
                "expression": """
(function() {
    var host = document.getElementById('demo-chat-panel-host');
    if (!host) return JSON.stringify({error: 'no host element found'});
    var sr = host.shadowRoot;
    if (!sr) return JSON.stringify({error: 'no shadow root'});
    var panel = sr.getElementById('demo-chat-panel');
    var input = sr.getElementById('chat-input');
    var send = sr.getElementById('chat-send');
    var welcome = sr.getElementById('chat-welcome');
    var toggle = document.getElementById('demo-chat-toggle');

    function cs(el, prop) { return el ? getComputedStyle(el)[prop] : null; }
    function rect(el) {
        if (!el) return null;
        var r = el.getBoundingClientRect();
        return {x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), h: Math.round(r.height)};
    }

    return JSON.stringify({
        windowSize: [window.innerWidth, window.innerHeight],
        bodyMarginRight: cs(document.body, 'marginRight'),
        host: { exists: true, rect: rect(host), display: cs(host, 'display'), position: cs(host, 'position'), width: cs(host, 'width'), zIndex: cs(host, 'zIndex'), collapsed: host.classList.contains('collapsed') },
        panel: { exists: !!panel, rect: rect(panel), display: cs(panel, 'display'), background: cs(panel, 'background')?.substring(0, 40) },
        input: { exists: !!input, rect: rect(input), display: cs(input, 'display'), visibility: cs(input, 'visibility'), height: cs(input, 'height'), disabled: input ? input.disabled : null },
        send: { exists: !!send, rect: rect(send), display: cs(send, 'display'), visibility: cs(send, 'visibility'), width: cs(send, 'width'), height: cs(send, 'height') },
        welcome: { exists: !!welcome, rect: rect(welcome), display: cs(welcome, 'display') },
        toggle: { exists: !!toggle, display: toggle ? toggle.style.display : null },
        api: typeof window.DemoCopilotChat !== 'undefined'
    });
})()
""",
                "returnByValue": True
            }
        }
        await ws.send(json.dumps(cmd))
        result = await ws.recv()
        data = json.loads(result)
        if "result" in data and "result" in data["result"]:
            val = data["result"]["result"].get("value", "")
            parsed = json.loads(val)
            print(json.dumps(parsed, indent=2))
        else:
            print("Unexpected result:", json.dumps(data, indent=2)[:500])

asyncio.run(main())
