"""
Optional WebSocket client for Print Hub's Reverb server.

Connects to the Laravel Reverb WebSocket server to receive real-time
push notifications about job status updates, agent status changes,
and queue updates. When an event is received, it triggers a callback
to refresh the queue immediately.

This is an optimization — if the WebSocket connection fails or the
required libraries are not available, the agent falls back to polling
(which is already implemented in server.py).

Reverb uses the Pusher protocol, so this client implements the
minimal Pusher WebSocket protocol needed to subscribe to channels
and receive events.
"""

import json
import logging
import threading
import time
from typing import Callable, Optional

log = logging.getLogger('trayprint.ws')

# Try to import websockets; if unavailable, WebSocket features are disabled
try:
    import websockets
    HAS_WEBSOCKETS = True
except ImportError:
    HAS_WEBSOCKETS = False
    log.info("'websockets' library not installed. WebSocket client disabled. Install with: pip install websockets")


class ReverbWebSocketClient:
    """
    Lightweight WebSocket client for Laravel Reverb (Pusher-compatible).

    Connects to the Reverb server, subscribes to channels, and invokes
    a callback when events are received.

    The Pusher WebSocket protocol is simple:
    1. Connect to ws://host:port/app/appKey
    2. Send a subscribe message: {"event":"pusher:subscribe","data":{"channel":"channelName"}}
    3. Receive events: {"event":"eventName","data":{...},"channel":"channelName"}
    """

    def __init__(
        self,
        host: str = '127.0.0.1',
        port: int = 8080,
        app_key: str = '',
        scheme: str = 'ws',
        agent_key: str = '',
        hub_url: str = '',
        on_event: Optional[Callable] = None,
        reconnect_delay: float = 5.0,
        max_reconnect_attempts: int = 0,  # 0 = unlimited
    ):
        self.host = host
        self.port = port
        self.app_key = app_key
        self.scheme = scheme
        self.agent_key = agent_key
        self.hub_url = hub_url
        self.on_event = on_event  # callback(event_data)
        self.reconnect_delay = reconnect_delay
        self.max_reconnect_attempts = max_reconnect_attempts

        self._ws_url = f"{scheme}://{host}:{port}/app/{app_key}"
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
        """Start the WebSocket client in a background thread."""
        if not HAS_WEBSOCKETS:
            log.warning("WebSocket client not started: 'websockets' library not available")
            return

        if self.is_running:
            log.warning("WebSocket client already running")
            return

        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run_loop, daemon=True, name='reverb-ws')
        self._thread.start()
        log.info("WebSocket client started (connecting to %s)", self._ws_url)

    def stop(self):
        """Signal the WebSocket client to stop."""
        self._stop_event.set()
        if self._thread:
            self._thread.join(timeout=3)
        self._connected = False
        log.info("WebSocket client stopped")

    def _run_loop(self):
        """Main loop with reconnection logic."""
        attempts = 0

        while not self._stop_event.is_set():
            if self.max_reconnect_attempts > 0 and attempts >= self.max_reconnect_attempts:
                log.warning("WebSocket client: max reconnect attempts (%d) reached, giving up", attempts)
                break

            try:
                self._connect_count += 1
                attempts += 1
                log.debug("WebSocket connection attempt %d to %s", attempts, self._ws_url)

                # Import here to ensure it's available
                import websockets.asyncio.client
                import asyncio

                # Since we're in a sync thread, run the async loop in a new event loop
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                try:
                    loop.run_until_complete(self._run_async())
                finally:
                    loop.close()

            except websockets.exceptions.ConnectionClosed:
                log.debug("WebSocket connection closed (normal)")
                self._connected = False

            except websockets.exceptions.WebSocketException as e:
                log.warning("WebSocket connection error: %s", e)
                self._connected = False

            except OSError as e:
                log.debug("WebSocket connection failed (server may be offline): %s", e)
                self._connected = False

            except Exception as e:
                log.warning("WebSocket unexpected error: %s", e)
                self._connected = False

            if not self._stop_event.is_set():
                log.debug("WebSocket reconnecting in %.1fs...", self.reconnect_delay)
                self._stop_event.wait(self.reconnect_delay)

    async def _run_async(self):
        """Async WebSocket connection handler."""
        import websockets.asyncio.client

        async with websockets.asyncio.client.connect(
            self._ws_url,
            ping_interval=30,
            ping_timeout=10,
            close_timeout=5,
        ) as ws:
            self._connected = True
            attempts_since_connected = 0
            log.info("WebSocket connected to %s", self._ws_url)

            # Subscribe to relevant channels
            # We subscribe to admin.queue for general queue updates
            # The agent will refresh its queue when it receives any event
            channels = [
                'admin.queue',
            ]

            for channel in channels:
                subscribe_msg = json.dumps({
                    'event': 'pusher:subscribe',
                    'data': {
                        'channel': channel,
                    },
                })
                await ws.send(subscribe_msg)
                log.debug("Subscribed to channel: %s", channel)

            # Listen for events
            while not self._stop_event.is_set():
                try:
                    message = await asyncio.wait_for(
                        ws.recv(),
                        timeout=5.0,
                    )

                    if isinstance(message, bytes):
                        message = message.decode('utf-8')

                    data = json.loads(message)
                    event = data.get('event', '')

                    # Skip Pusher internal events
                    if event.startswith('pusher:'):
                        continue

                    log.debug("WebSocket event received: %s on channel %s",
                              event, data.get('channel', '?'))

                    # Trigger callback (queue refresh)
                    if self.on_event:
                        try:
                            self.on_event({
                                'event': event,
                                'channel': data.get('channel'),
                                'data': data.get('data', {}),
                            })
                        except Exception as cb_err:
                            log.error("WebSocket event callback error: %s", cb_err)

                except asyncio.TimeoutError:
                    # Normal timeout to check stop flag
                    continue

                except websockets.exceptions.ConnectionClosed:
                    log.debug("WebSocket connection closed during receive")
                    break


def start_websocket_client(
    config: dict,
    on_event: Optional[Callable] = None,
) -> Optional[ReverbWebSocketClient]:
    """
    Start the WebSocket client based on config.

    Reads Reverb connection settings from the Print Hub configuration
    and starts a background WebSocket client if possible.

    Args:
        config: Configuration dict (from config.json)
        on_event: Callback invoked when events are received

    Returns:
        ReverbWebSocketClient instance, or None if not configured/available.
    """
    if not HAS_WEBSOCKETS:
        log.info("WebSocket client disabled: 'websockets' library not installed")
        return None

    # Check if we have hub connection configured
    hub_url = config.get('hub_url', '')
    if not hub_url:
        log.info("WebSocket client disabled: no hub_url configured")
        return None

    # Parse Reverb connection details from config or use defaults
    reverb_host = config.get('reverb_host', '127.0.0.1')
    reverb_port = config.get('reverb_port', 8080)
    reverb_app_key = config.get('reverb_app_key', '')
    reverb_scheme = config.get('reverb_scheme', 'ws')

    if not reverb_app_key:
        log.info("WebSocket client disabled: no reverb_app_key configured")
        return None

    client = ReverbWebSocketClient(
        host=reverb_host,
        port=reverb_port,
        app_key=reverb_app_key,
        scheme=reverb_scheme,
        agent_key=config.get('agent_key', ''),
        hub_url=hub_url,
        on_event=on_event,
    )

    client.start()
    return client
