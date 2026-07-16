"""Utility functions and shared logging infrastructure for PySAPManager.

This module centralizes two things used throughout the package:

- ``default_logger``: the no-frills console logger used whenever a caller
  does not supply their own logging function.
- A **global logging switch** (:func:`set_logging_enabled` /
  :func:`is_logging_enabled`) that lets a user of the library silence *all*
  log output produced by PySAPManager in one call, regardless of whether the
  default logger or a custom logger callback is in use.

Every internal component (``SAPManager``, ``SAPApplication``,
``SAPConnection``, ``VBScriptRunner``, ``SAPProcessService``) accepts an
optional ``logger`` callable. Instead of storing that callable directly,
components wrap it with :func:`resolve_logger`, which is a thin gate that
checks the global switch before delegating to the real logger. This means
the global switch works even if you pass your own custom logging function.

Example:
    Disable all PySAPManager log output for the remainder of the process:

    ```python
    from pysapmanager import set_logging_enabled

    set_logging_enabled(False)
    ```

    Re-enable it later:

    ```python
    set_logging_enabled(True)
    ```
"""

from __future__ import annotations

from collections.abc import Callable

LoggerFunc = Callable[[str, bool], None]
"""Type alias for the logging callable expected throughout the library.

The callable must accept a message string and a ``critical`` boolean flag,
e.g. ``def my_logger(message: str, critical: bool = False) -> None: ...``.
"""

_logging_enabled: bool = True
"""Module-level flag controlling whether any PySAPManager logger call produces output.

This is intentionally a simple module-level flag rather than a class
attribute so that the switch is process-wide and trivially accessible from
anywhere in the library (``from .utils import is_logging_enabled``) without
needing to thread a configuration object through every constructor.
"""


def set_logging_enabled(enabled: bool) -> None:
    """Globally enables or disables all logging produced by PySAPManager.

    This affects every component in the library, including instances that
    were constructed with a custom ``logger`` callable: when disabled, no
    logger (default or custom) will be invoked by any PySAPManager component
    until logging is re-enabled.

    This is a process-wide (module-level) setting. It affects all
    ``SAPManager``, ``SAPApplication``, ``SAPConnection``, ``VBScriptRunner``,
    and ``SAPProcessService`` instances currently alive or created after this
    call, since each of them resolves the flag dynamically on every log call
    rather than caching it at construction time.

    Args:
        enabled: ``True`` to allow log messages to be emitted (default
            library behavior). ``False`` to suppress all PySAPManager log output.

    Example:
        ```python
        from pysapmanager import SAPManager, SAPConfig, set_logging_enabled

        set_logging_enabled(False)  # silence the library

        with SAPManager(config=SAPConfig(...)) as manager:
            ...  # no log output will be produced

        set_logging_enabled(True)  # restore normal logging
        ```
    """
    global _logging_enabled
    _logging_enabled = bool(enabled)


def is_logging_enabled() -> bool:
    """Returns whether PySAPManager logging is currently enabled.

    Returns:
        bool: ``True`` if logging is currently enabled (the default),
            ``False`` if it has been disabled via :func:`set_logging_enabled`.
    """
    return _logging_enabled


def default_logger(message: str, critical: bool = False) -> None:
    """Logs a message to standard output.

    This is the default logging mechanism for the library, providing basic
    console output with severity level indication. It is used automatically
    by every PySAPManager component that is not given an explicit ``logger``
    argument.

    Note:
        This function does not check the global logging switch itself;
        gating is applied by :func:`resolve_logger`, which wraps whichever
        logger (default or custom) a component ends up using. Calling
        ``default_logger`` directly will always print, by design, so it
        remains usable standalone (e.g. for debugging).

    Args:
        message: The message to be logged.
        critical: If True, prefixes the log with '[CRITICAL]', otherwise
            '[INFO]'. Defaults to False.
    """
    print(f"[{'CRITICAL' if critical else 'INFO'}] {message}")


def resolve_logger(logger: LoggerFunc | None) -> LoggerFunc:
    """Wraps a logger callable (or the default one) with the global on/off switch.

    Every PySAPManager component should call this function once, at
    construction time, and store the *returned* callable as its logger
    rather than the raw ``logger`` argument. This guarantees that
    :func:`set_logging_enabled` silences the component regardless of whether
    it was built with a custom logger or fell back to :func:`default_logger`.

    Args:
        logger: A user-supplied logging callable, or ``None`` to fall back
            to :func:`default_logger`.

    Returns:
        LoggerFunc: A callable with the same signature as ``logger`` that
            first checks :func:`is_logging_enabled` and, only if logging is
            currently enabled, delegates to the resolved underlying logger.

    Example:
        ```python
        class MyComponent:
            def __init__(self, logger=None):
                self._logger = resolve_logger(logger)

            def do_work(self):
                self._logger("Working...", False)
        ```
    """
    base_logger = logger or default_logger

    def _gated_logger(message: str, critical: bool = False) -> None:
        if _logging_enabled:
            base_logger(message, critical)

    return _gated_logger
