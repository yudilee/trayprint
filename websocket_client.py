"""
Optional WebSocket client for Print Hub's Reverb server.
Connects to Laravel Reverb (Pusher-compatible) over standard WS or Cloudflare Tunnel WSS.
"""

import json
import logging
import threading
import time
from urllib.parse import urlparse
from typing import Callable, Optional

log = logging.getLogger('trayprint.ws')

try:
    import websockets
    HAS_WEBSOCKETS = True
except ImportError:
    HAS_WEBSOCKETS = False
    log.info("'websockets' library not installed. WebSocket client disabled. Install with: pip install websockets")


class ReverbWebSocketClient:
    def __init__(
        self,
        host: str = '127.0.0.1',
        port: int = 8080,
        app_key: str = '',
        scheme: str = 'ws',
        agent_key: str = '',
        agent_id: str = '',
        hub_url: str = '',
        on_event: Optional[Callable] = None,
        reconnect_delay: float = 5.0,
        max_reconnect_attempts: int = 0,
    ):
        self.host = host
        self.port = port
        self.app_key = app_key
        self.agent_key = agent_key
        self.agent_id = str(agent_id) if agent_id else ''
        self.hub_url = hub_url
        self.on_event = on_event
        self.reconnect_delay = reconnect_delay
        self.max_reconnect_attempts = max_reconnect_attempts

        # Auto-detect host/scheme/port from hub_url if provided
        if hub_url:
            parsed = urlparse(hub_url)
            self.host = parsed.hostname or self.host
            if parsed.scheme == 'https':
                self.scheme = 'wss'
                self.port = parsed.port or 443
            else:
                self.scheme = 'ws'
                self.port = parsed.port or 80
        else:
            self.scheme = 'wss' if scheme in ('https', 'wss') else 'ws'

        if (self.scheme == 'wss' and self.port == 443) or (self.scheme == 'ws' and self.port == 80):
            self._ws_url = f"{self.scheme}://{self.host}/app/{self.app_key}"
        else:
            self._ws_url = f"{self.scheme}://{self.host}:{self.port}/app/{self.app_key}"

        self._thread: Optional[threading.Thread] = None
        self._stop_event = threading.Event()
        self._connected = False
        self._connect_count = 0

    @property
    def is_connected(self) -> bool:
        return self._connected

    @property
    def is_running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def start(self):
        if not HAS_WEBSOCKETS:
            log.warning("WebSocket client not started: 'websockets' library not available")
            return
        if self.is_running:
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run_loop, daemon=True, name='reverb-ws')
        self._thread.start()
        log.info("WebSocket client started (connecting to %s)", self._ws_url)

    def stop(self):
        self._stop_event.set()
        if self._thread:
            self._thread.join(timeout=3)
        self._connected = False
        log.info("WebSocket client stopped")

    def _run_loop(self):
        attempts = 0
        while not self._stop_event.is_set():
            if self.max_reconnect_attempts > 0 and attempts >= self.max_reconnect_attempts:
                log.warning("WebSocket client: max reconnect attempts reached, giving up")
                break
            try:
                attempts += 1
                import websockets.asyncio.client
                import asyncio

                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                try:
                    loop.run_until_complete(self._run_async())
                finally:
                    loop.close()
            except Exception as e:
                log.debug("WebSocket connection disconnected (%s), reconnecting in %.1fs...", e, self.reconnect_delay)
                self._connected = False

            if not self._stop_event.is_set():
                self._stop_event.wait(self.reconnect_delay)

    async def _run_async(self):
        import websockets.asyncio.client
        import asyncio

        async with websockets.asyncio.client.connect(
            self._ws_url,
            ping_interval=30,
            ping_timeout=10,
            close_timeout=5,
        ) as ws:
            self._connected = True
            log.info("WebSocket connected to %s", self._ws_url)

            # Subscribe to channels
            channels = ['admin.queue']
            if self.agent_id:
                channels.append(f"agent.{self.agent_id}")

            for channel in channels:
                subscribe_msg = json.dumps({
                    'event': 'pusher:subscribe',
                    'data': {
                        'channel': channel,
                    },
                })
                await ws.send(subscribe_msg)
                log.debug("Subscribed to channel: %s", channel)

            while not self._stop_event.is_set():
                try:
                    message = await asyncio.wait_for(ws.recv(), timeout=5.0)
                    if isinstance(message, bytes):
                        message = message.decode('utf-8')
                    data = json.loads(message)
                    event = data.get('event', '')
                    if event.startswith('pusher:'):
                        continue

                    log.info("WebSocket event received: %s on channel %s", event, data.get('channel', '?'))
                    if self.on_event:
                        self.on_event({
                            'event': event,
                            'channel': data.get('channel'),
                            'data': data.get('data', {}),
                        })
                except asyncio.TimeoutError:
                    continue
                except websockets.exceptions.ConnectionClosed:
                    log.debug("WebSocket connection closed during receive")
                    break


def start_websocket_client(
    config: dict,
    on_event: Optional[Callable] = None,
) -> Optional[ReverbWebSocketClient]:
    if not HAS_WEBSOCKETS:
        return None

    hub_url = config.get('hub_url', '')
    if not hub_url:
        return None

    reverb_host = config.get('reverb_host', '127.0.0.1')
    reverb_port = config.get('reverb_port', 8080)
    reverb_app_key = config.get('reverb_app_key', 'printhub-live-key')
    reverb_scheme = config.get('reverb_scheme', 'ws')
    agent_id = config.get('agent_id', '')

    client = ReverbWebSocketClient(
        host=reverb_host,
        port=reverb_port,
        app_key=reverb_app_key,
        scheme=reverb_scheme,
        agent_key=config.get('agent_key', ''),
        agent_id=agent_id,
        hub_url=hub_url,
        on_event=on_event,
    )
    client.start()
    return client
