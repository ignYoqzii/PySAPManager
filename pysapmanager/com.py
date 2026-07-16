"""COM utilities for SAP automation.

This module provides helper functions to initialize and deinitialize COM
(Component Object Model) in the current thread. This step is often necessary
before any interaction with COM objects, particularly when automating SAP GUI
Scripting under Windows.

COM initialization is required for multi-threaded scenarios where secondary
threads do not have COM initialized by default.
"""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager

import pythoncom


@contextmanager
def com_initialized() -> Iterator[None]:
    """Context manager to initialize COM in the current thread.

    This context manager facilitates safe interaction with COM objects by
    ensuring that the calling thread is properly initialized. It calls
    `pythoncom.CoInitialize()` upon entering the context and guarantees that
    `pythoncom.CoUninitialize()` is invoked upon exit, ensuring cleanup
    regardless of whether the code block completes successfully or raises
    an exception.

    This is mandatory for multi-threaded applications, as each thread that
    interacts with COM objects must perform its own initialization.

    Yields:
        None: The context manager does not return a value; it strictly
            manages the lifecycle of the thread's COM state.

    Raises:
        pythoncom.com_error: If the underlying COM initialization fails or
            if there is a conflict in the apartment threading model.

    Examples:
        **Standard Synchronous Usage:**

        ```python
        with com_initialized():
            # Perform COM calls here
            import win32com.client as win32
            app = win32.GetObject("SAPGUI")
        ```

        **Multi-threaded Usage:**

        ```python
        import threading
        import win32com.client as win32

        def worker_thread():
            with com_initialized():
                # COM calls are now safe in this background thread
                app = win32.GetObject("SAPGUI")

        thread = threading.Thread(target=worker_thread)
        thread.start()
        thread.join()
        ```
    """
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()