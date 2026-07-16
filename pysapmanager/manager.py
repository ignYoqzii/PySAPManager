"""High-level facade for controlling SAP GUI.

This module exposes a high-level facade for orchestrating the various
components necessary for SAP GUI automation.

The SAPManager class centralizes:

- Launching and stopping the SAP Logon process
- Access to the SAP GUI Scripting application
- Preparation of an environment compatible with a target connection
- Waiting for asynchronous tasks
- Proper release of internal resources
- Automatic closing of the connection opened in its context

This facade simplifies library usage by hiding the details of assembling
technical services and business objects.
"""

from __future__ import annotations

import time
from collections.abc import Callable
from concurrent.futures import Future, wait
from pathlib import Path
from types import TracebackType
from typing import TYPE_CHECKING, Any

from .application import SAPApplication
from .exceptions import SAPConnectionException
from .models import SAPConfig
from .process import SAPProcessService
from .script_runner import VBScriptRunner
from .utils import resolve_logger, set_logging_enabled

if TYPE_CHECKING:
    from .connection import SAPConnection


class SAPManager:
    """Main facade for SAP GUI automation.

    This class coordinates the different internal components necessary for
    controlling SAP GUI:

    - SAPProcessService: Manages the SAP Logon process
    - SAPApplication: Accesses the SAP GUI Scripting application
    - VBScriptRunner: Executes scripts in parallel

    The manager prepares the SAP environment and exposes the SAP GUI Scripting
    application, but does not itself open the connection. Connection opening
    is the responsibility of SAPApplication, consistent with the COM hierarchy
    of SAP GUI Scripting.

    When a connection is opened via the application attached to this manager,
    the manager remembers it to enable automatic closing when exiting the
    `with` context.

    Note:
        A single `SAPManager` can track **at most one** open connection at a
        time. Calling `application.open_connection()` a second time on the
        same manager before closing the first (via `close_connection()`)
        raises `SAPConnectionException`. If you need several simultaneous
        connections, create multiple sessions on one connection instead
        (`connection.create_session()`), or use separate `SAPManager`
        instances.

    Attributes:
        _config (SAPConfig): Configuration settings for SAP and automation.
        _logger (Callable[[str, bool], None]): The logging function.
        _process (SAPProcessService): Service for managing SAP Logon processes.
        _runner (VBScriptRunner): Service for asynchronous script execution.
        _application (SAPApplication | None): Cached SAP GUI Scripting application.
        _connection (SAPConnection | None): The currently active/managed connection.

    Example:
        Typical usage with a context manager:

        ```python
        with SAPManager(config) as manager:
            application = manager.get_application()
            connection = application.open_connection()
            session = connection.first_session()
        ```
    """

    def __init__(
        self,
        config: SAPConfig,
        *,
        logger: Callable[[str, bool], None] | None = None,
        file_validator: Callable[[Path], None] | None = None,
        max_parallel_scripts: int = 4,
        enable_logging: bool | None = None,
    ) -> None:
        """Initializes the main SAP GUI manager.

        This method constructs the different internal services necessary for
        SAP automation from the provided configuration.

        Args:
            config: Global SAP configuration used for launching SAP Logon and
                preparing the target connection.
            logger: Logging function. If None, :func:`~pysapmanager.utils.default_logger`
                is used. Expected signature:
                `logger(message: str, critical: bool = False) -> None`.
            file_validator: File validator eventually used by the script runner.
                Expected signature: `file_validator(path: Path) -> None`.
            max_parallel_scripts: Maximum number of scripts executed in parallel
                by VBScriptRunner. Defaults to 4.
            enable_logging: Convenience flag equivalent to calling
                `pysapmanager.set_logging_enabled(enable_logging)` before
                constructing the manager. Since the underlying switch is
                library-wide (see :func:`pysapmanager.set_logging_enabled`),
                passing `enable_logging=False` here silences **all**
                PySAPManager components for the lifetime of the process, not
                just this manager instance - including ones created before
                or after this call, and it stays disabled until something
                re-enables it. Defaults to `None`, which leaves the current
                global logging state untouched (so this constructor never
                surprises you by silently re-enabling logging you disabled
                elsewhere). Pass `True` explicitly to force logging back on.
        """
        if enable_logging is not None:
            set_logging_enabled(enable_logging)

        self._config = config
        self._logger = resolve_logger(logger)

        self._process = SAPProcessService(
            config.sap_logon_path,
            logger=self._logger,
        )
        self._runner = VBScriptRunner(
            max_workers=max_parallel_scripts,
            file_validator=file_validator,
            logger=self._logger,
        )
        self._application: SAPApplication | None = None
        self._connection: SAPConnection | None = None

    @property
    def process(self) -> SAPProcessService:
        """SAPProcessService: The SAP Logon process management service."""
        return self._process

    @property
    def connection(self) -> SAPConnection | None:
        """SAPConnection | None: The connection currently registered in the manager, 
        or None if no connection is registered.
        """
        return self._connection

    def get_application(
        self, retry_count: int = 20, retry_delay: float = 0.5
    ) -> SAPApplication:
        """Returns the SAP GUI Scripting application associated with the manager.

        The application is instantiated on demand and then cached for subsequent
        accesses. This method serves as the normal entry point for obtaining
        the SAPApplication object from the manager.

        Args:
            retry_count: Maximum number of attempts to get the SAP GUI Scripting
                application. Defaults to 20.
            retry_delay: Delay in seconds between attempts to access the SAP
                GUI Scripting application. Defaults to 0.5.

        Returns:
            SAPApplication: Instance of the SAP GUI Scripting application.
        """
        if self._application is None:
            self._application = SAPApplication(
                config=self._config,
                script_runner=self._runner,
                retry_count=retry_count,
                retry_delay=retry_delay,
                logger=self._logger,
                on_connection_opened=self._register_connection,
            )
        return self._application

    def _register_connection(self, connection: SAPConnection) -> None:
        """Registers the connection opened in the manager's context.

        This method is called by the application when it opens a connection
        in the manager's context. The manager can remember only one connection
        at a time.

        Args:
            connection: SAP connection opened via the manager's application.

        Raises:
            SAPConnectionException: If a connection is already registered in
                this manager.
        """
        if self._connection is not None:
            raise SAPConnectionException(
                "A SAP connection is already registered in this manager.",
                critical=True,
            )
        self._connection = connection

    @staticmethod
    def wait_until(
        condition: Callable[[], bool],
        timeout: float = 10,
        interval: float = 0.2,
    ) -> bool:
        """Waits for a condition to become true within a maximum delay.

        This method periodically evaluates a condition until it returns True or
        until the maximum delay expires. Exceptions raised by the condition are
        ignored to tolerate transient states, such as during process startup or
        shutdown.

        Args:
            condition: Function with no arguments called repeatedly.
            timeout: Maximum wait duration in seconds.
            interval: Delay in seconds between condition evaluations.

        Returns:
            bool: True if the condition becomes true before timeout expiration,
                False otherwise.

        Example:
            Wait for a process to stop running:

            ```python
            SAPManager.wait_until(
                lambda: not process.is_running(),
                timeout=10
            )
            ```
        """
        start = time.time()
        while time.time() - start < timeout:
            try:
                if condition():
                    return True
            except Exception:
                pass
            time.sleep(interval)
        return False

    def _reset_sap_logon(self) -> None:
        """Closes then relaunches SAP Logon.

        This method is used when the target connection defined in the
        configuration is already open in SAP. It forces a complete restart
        of SAP Logon to start from a clean state before reopening the connection.

        Raises:
            TimeoutError: If SAP Logon does not close or relaunch within the
                allotted time.
        """
        self._logger("Closing SAP Logon.", False)
        self._process.kill()

        if not self.wait_until(lambda: not self._process.is_running()):
            raise TimeoutError("SAP Logon did not close within the allotted time.")

        self._logger("Reopening SAP Logon.", False)
        self._process.launch()

        if not self.wait_until(lambda: self._process.is_running()):
            raise TimeoutError("SAP Logon did not open within the allotted time.")

        self._application = None
        self._connection = None

    def prepare_sap_environment(self) -> None:
        """Prepares the SAP environment for the configured target connection.

        Behavior is as follows:

        1. If SAP Logon is not running, it is started.
        2. If SAP Logon is already running, open connections are inspected.
        3. If the target connection is already open, SAP Logon is closed and
           relaunched to ensure a clean state.
        4. If the target connection is not open, the environment is preserved.

        Raises:
            TimeoutError: If SAP Logon does not open, close, or relaunch within
                the allotted time.
        """
        if not self._process.is_running():
            self._logger("SAP Logon is not open. Opening.", False)
            self._process.launch()

            if not self.wait_until(lambda: self._process.is_running()):
                raise TimeoutError(
                    "SAP Logon did not open within the allotted time."
                )

            self._application = None
            return

        self._application = None
        application = self.get_application()

        existing = application.find_connection(self._config.connection_description)
        if existing is not None:
            self._logger(
                f"Target connection '{self._config.connection_description}' "
                "is already open. Restarting SAP Logon.",
                False,
            )
            self._reset_sap_logon()
            return

        self._logger(
            f"Target connection '{self._config.connection_description}' "
            "is not open. Current environment can be used.",
            False,
        )

    def close_connection(self) -> None:
        """Closes the connection registered in the manager, if any.

        Raises:
            SAPConnectionException: If the connection closing fails.
        """
        if self._connection is None:
            return

        self._connection.close()
        self._connection = None
        self._logger("Manager connection closed.", False)

    def wait_all(self, *futures: Future[Any]) -> None:
        """Waits for all asynchronous tasks to complete and reraises any exceptions.

        Args:
            *futures: Set of asynchronous task Future objects to wait for.

        Raises:
            Exception: Any exception raised in one of the tasks is propagated
                when calling `future.result()`.
        """
        done, _ = wait(futures)
        for future in done:
            future.result()

    def shutdown(self) -> None:
        """Releases the manager's internal resources.

        Gracefully stops components requiring explicit termination, such as
        the script runner's thread pool.
        """
        self._runner.shutdown()

    def __enter__(self) -> SAPManager:
        """Context manager entry.

        Prepares the SAP environment and returns the manager instance.

        Returns:
            SAPManager: The initialized manager instance.
        """
        self.prepare_sap_environment()
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        """Context manager exit.

        Closes the registered connection then releases internal resources,
        ensuring cleanup regardless of errors within the `with` block.
        """
        try:
            self.close_connection()
        finally:
            self.shutdown()