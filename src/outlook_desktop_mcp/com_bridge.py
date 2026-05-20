"""
COM Threading Bridge
====================
Runs all Outlook COM calls on a dedicated STA (Single-Threaded Apartment)
thread so the async MCP event loop never touches COM objects directly.

Every COM function passed to bridge.call() receives (outlook, namespace, ...)
as its first two arguments — the live COM objects that only exist on the
COM thread.

The bridge serializes work onto a single STA thread. To avoid one slow
operation cascading into wedging the whole server, ``call`` supports a
per-call timeout and refuses to queue a new request while a previous one
is still in flight (the COM thread cannot be cancelled mid-COM-call).
"""
import asyncio
import logging
import os
import queue
import sys
import threading

logger = logging.getLogger("outlook_desktop_mcp.com_bridge")

DEFAULT_CALL_TIMEOUT = float(os.environ.get("OUTLOOK_MCP_COM_TIMEOUT", "60"))
DEFAULT_START_TIMEOUT = float(os.environ.get("OUTLOOK_MCP_COM_START_TIMEOUT", "15"))


class ComBridgeBusyError(RuntimeError):
    """Raised when a new COM request is submitted while another is in flight."""


class ComBridgeTimeoutError(TimeoutError):
    """Raised when a COM call exceeds its timeout."""


class OutlookBridge:
    """Manages a dedicated COM thread for Outlook operations."""

    def __init__(self):
        self._thread: threading.Thread | None = None
        self._request_queue: queue.Queue = queue.Queue()
        self._outlook = None
        self._namespace = None
        self._ready = threading.Event()
        self._shutdown = threading.Event()
        self._init_error: Exception | None = None
        self._in_flight_lock = threading.Lock()
        self._in_flight_label: str | None = None

    def start(self, timeout: float = DEFAULT_START_TIMEOUT):
        """Start the COM thread. Call once at server startup."""
        if self._thread is not None and self._thread.is_alive():
            logger.warning("OutlookBridge.start() called twice; ignoring.")
            return
        self._thread = threading.Thread(
            target=self._com_thread_main, daemon=True, name="outlook-com"
        )
        self._thread.start()
        if not self._ready.wait(timeout=timeout):
            if self._init_error:
                raise self._init_error
            raise RuntimeError(
                f"Outlook COM thread failed to initialize within {timeout}s. "
                "Is Outlook Desktop (Classic) running?"
            )

    def _com_thread_main(self):
        """Main loop for the COM thread."""
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            try:
                self._outlook = win32com.client.Dispatch("Outlook.Application")
                self._namespace = self._outlook.GetNamespace("MAPI")
                store_name = self._namespace.DefaultStore.DisplayName
                user_name = self._namespace.CurrentUser.Name
                logger.debug(
                    "COM thread ready. Store: %s, User: %s", store_name, user_name
                )
            except Exception as e:
                self._init_error = e
                logger.error("COM thread init failed: %s", e)
                self._ready.set()
                return
            self._ready.set()

            while not self._shutdown.is_set():
                try:
                    func, args, kwargs, result_event, result_holder = (
                        self._request_queue.get(timeout=0.5)
                    )
                except queue.Empty:
                    continue
                try:
                    result_holder["value"] = func(
                        self._outlook, self._namespace, *args, **kwargs
                    )
                except Exception as e:
                    result_holder["error"] = e
                finally:
                    result_event.set()

            self._drain_pending("COM bridge is shutting down")
        finally:
            pythoncom.CoUninitialize()

    def _drain_pending(self, message: str):
        """Reject any queued-but-not-started requests with a clear error."""
        while True:
            try:
                _, _, _, result_event, result_holder = (
                    self._request_queue.get_nowait()
                )
            except queue.Empty:
                return
            result_holder["error"] = RuntimeError(message)
            result_event.set()

    async def call(self, func, *args, timeout: float | None = None, **kwargs):
        """
        Schedule a function on the COM thread and await its result.

        Refuses to queue if another call is currently *awaited* by the caller
        path: the STA thread can only run one COM call at a time and slow
        operations should not silently stack up. The caller should retry after
        a short delay.

        Note that on ``ComBridgeTimeoutError`` the in-flight COM call keeps
        running on the COM thread (COM has no cancellation). The bridge lock
        is released when the caller's ``call`` returns, so subsequent calls
        will queue behind the still-running operation; if you need hard
        back-pressure, catch the timeout and refuse new work yourself.

        ``timeout`` defaults to ``DEFAULT_CALL_TIMEOUT`` (60s).
        """
        timeout_val = DEFAULT_CALL_TIMEOUT if timeout is None else float(timeout)
        label = getattr(func, "__name__", "<anonymous>")

        if not self._in_flight_lock.acquire(blocking=False):
            raise ComBridgeBusyError(
                f"COM thread is busy with previous request "
                f"({self._in_flight_label or 'unknown'}). Retry shortly."
            )
        try:
            self._in_flight_label = label
            result_event = threading.Event()
            result_holder: dict = {}
            self._request_queue.put((func, args, kwargs, result_event, result_holder))

            loop = asyncio.get_running_loop()
            signaled = await loop.run_in_executor(
                None, lambda: result_event.wait(timeout=timeout_val)
            )
            if not signaled:
                raise ComBridgeTimeoutError(
                    f"Outlook COM operation '{label}' timed out after "
                    f"{timeout_val:.0f}s. Outlook may be waiting on a dialog."
                )
            if "error" in result_holder:
                raise result_holder["error"]
            return result_holder.get("value")
        finally:
            self._in_flight_label = None
            self._in_flight_lock.release()

    def stop(self):
        """Signal the COM thread to shut down and reject pending requests."""
        self._shutdown.set()
        self._drain_pending("OutlookBridge stopped")
        if self._thread:
            self._thread.join(timeout=5)
