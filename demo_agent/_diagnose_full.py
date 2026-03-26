"""Extended CDP diagnostic — check full layout + simulate smaller viewport."""
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
        # 1) Extended diagnostic at current size
        cmd1 = {
            "id": 1,
            "method": "Runtime.evaluate",
            "params": {
                "expression": """
(function() {
    var host = document.getElementById('demo-chat-panel-host');
    if (!host) return JSON.stringify({error: 'no host element'});
    var sr = host.shadowRoot;
    if (!sr) return JSON.stringify({error: 'no shadow root'});

    var panel = sr.getElementById('demo-chat-panel');
    var header = sr.querySelector('.chat-header');
    var welcome = sr.getElementById('chat-welcome');
    var msgs = sr.getElementById('chat-messages');
    var quickActs = sr.getElementById('chat-quick-actions');
    var inputArea = sr.querySelector('.chat-input-area');
    var inputWrap = sr.querySelector('.chat-input-wrap');
    var input = sr.getElementById('chat-input');
    var send = sr.getElementById('chat-send');

    function cs(el, props) {
        if (!el) return null;
        var s = getComputedStyle(el);
        var r = {};
        props.forEach(function(p) { r[p] = s[p]; });
        return r;
    }
    function rect(el) {
        if (!el) return null;
        var r = el.getBoundingClientRect();
        return {x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), h: Math.round(r.height)};
    }

    return JSON.stringify({
        screen: { width: screen.width, height: screen.height, availWidth: screen.availWidth, availHeight: screen.availHeight },
        window: { innerW: window.innerWidth, innerH: window.innerHeight, outerW: window.outerWidth, outerH: window.outerHeight, dpr: window.devicePixelRatio },
        host: { rect: rect(host), styles: cs(host, ['position','display','width','height','top','right','bottom','left','transform','zIndex','overflow']) },
        panel: { rect: rect(panel), styles: cs(panel, ['display','flexDirection','width','height','overflow','background']) },
        header: { rect: rect(header), styles: cs(header, ['display','flexShrink','padding','height']) },
        welcome: { rect: rect(welcome), styles: cs(welcome, ['display','flex','flexGrow','flexShrink','flexBasis','justifyContent','alignItems','padding','overflow','height']) },
        msgs: { rect: rect(msgs), styles: cs(msgs, ['display','flex','flexGrow','flexShrink','overflow','height']) },
        quickActs: { rect: rect(quickActs), styles: cs(quickActs, ['display']) },
        inputArea: { rect: rect(inputArea), styles: cs(inputArea, ['display','padding','background','flexShrink','height','visibility','overflow']) },
        inputWrap: { rect: rect(inputWrap), styles: cs(inputWrap, ['display','alignItems','background','border','height','overflow','visibility']) },
        input: { rect: rect(input), styles: cs(input, ['display','visibility','width','height','padding','fontSize','color','background','minHeight']), disabled: input ? input.disabled : null, placeholder: input ? input.placeholder : null },
        send: { rect: rect(send), styles: cs(send, ['display','visibility','width','height','background','color']) }
    });
})()
""",
                "returnByValue": True
            }
        }
        await ws.send(json.dumps(cmd1))
        result1 = await ws.recv()
        data1 = json.loads(result1)
        parsed = json.loads(data1["result"]["result"]["value"])
        print("=== FULL DIAGNOSTIC (current viewport) ===")
        print(json.dumps(parsed, indent=2))

        # 2) Try resizing viewport to 1366x768 (common laptop) and re-check
        resize_cmd = {
            "id": 2,
            "method": "Emulation.setDeviceMetricsOverride",
            "params": {
                "width": 1366,
                "height": 768,
                "deviceScaleFactor": 1,
                "mobile": False
            }
        }
        await ws.send(json.dumps(resize_cmd))
        await ws.recv()

        await asyncio.sleep(0.5)

        # Re-run diagnostic at smaller size
        cmd1["id"] = 3
        await ws.send(json.dumps(cmd1))
        result3 = await ws.recv()
        data3 = json.loads(result3)
        parsed3 = json.loads(data3["result"]["result"]["value"])
        print("\n=== DIAGNOSTIC AT 1366x768 ===")
        print(json.dumps(parsed3, indent=2))

        # Screenshot at small size
        ss_cmd = {"id": 4, "method": "Page.captureScreenshot", "params": {"format": "png"}}
        await ws.send(json.dumps(ss_cmd))
        ss_result = await ws.recv()
        ss_data = json.loads(ss_result)
        img = base64.b64decode(ss_data["result"]["data"])
        with open("chat_panel_small.png", "wb") as f:
            f.write(img)
        print(f"\nSmall viewport screenshot saved: chat_panel_small.png ({len(img)} bytes)")

        # Reset viewport
        clear_cmd = {"id": 5, "method": "Emulation.clearDeviceMetricsOverride", "params": {}}
        await ws.send(json.dumps(clear_cmd))
        await ws.recv()

asyncio.run(main())
