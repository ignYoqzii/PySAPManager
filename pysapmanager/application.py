"""Access to the SAP GUI Scripting application.

This module provides an abstraction of the SAP GUI Scripting
application exposed via COM. It serves as the entry point to the GuiApplication
object obtained from GetObject("SAPGUI") and the GetScriptingEngine property.

The SAPApplication class enables:

- Retrieving the SAP GUI Scripting application with retry strategy
- Iterating over open SAP connections
- Searching for an open connection by its description
- Opening the target connection defined in configuration
"""

from __future__ import annotations

import time
from collections.abc import Callable, Iterator
from typing import Any

import win32com.client as win32

from .connection import SAPConnection
from .exceptions import SAPApplicationException, SAPConnectionException
from .models import SAPConfig
from .script_runner import VBScriptRunner
from .utils import resolve_logger


class SAPApplication:
    """Represents the SAP GUI Scripting application (GuiApplication).

    This class encapsulates access to the SAP GUI Scripting application
    exposed by COM and provides the main operations related to open or
    unopened SAP connections.

    The instance automatically retrieves the COM GuiApplication object upon
    creation, with a configurable retry mechanism to tolerate availability
    delays in SAP GUI immediately after SAP Logon startup.

    Attributes:
        _config (SAPConfig): Global configuration for SAP connection parameters.
        _script_runner (VBScriptRunner): Executor for VBScript files.
        _retry_count (int): Maximum attempts to retrieve COM interface.
        _retry_delay (float): Seconds to wait between retry attempts.
        _logger (Callable[[str, bool], None]): Logging function for reporting status.
        _on_connection_opened (Callable[[SAPConnection], None] | None): Callback for new connections.
        _com_application (Any): The underlying COM GuiApplication object.
    """

    def __init__(
        self,
        *,
        config: SAPConfig,
        script_runner: VBScriptRunner,
        retry_count: int = 20,
        retry_delay: float = 0.5,
        logger: Callable[[str, bool], None] | None = None,
        on_connection_opened: Callable[[SAPConnection], None] | None = None,
    ) -> None:
        """Initializes the SAP GUI Scripting application.

        Args:
            config: Global SAP configuration used particularly for opening the
                target connection.
            script_runner: VBScript executor used by connections and sessions
                created from this application.
            retry_count: Maximum number of attempts to retrieve the SAP GUI Scripting
                application. Defaults to 20.
            retry_delay: Delay in seconds between successive retry attempts.
                Defaults to 0.5.
            logger: Logging function. If None, :func:`~pysapmanager.utils.default_logger`
                is used. Regardless of which logger is used, all output is
                subject to the library-wide switch controlled by
                :func:`pysapmanager.set_logging_enabled`.
            on_connection_opened: Callback invoked when a connection is opened
                successfully. Allows a calling component like the manager to
                register the opened connection.

        Raises:
            SAPApplicationException: If the SAP GUI Scripting application cannot
                be retrieved after the specified retry attempts.
        """
        self._config = config
        self._script_runner = script_runner
        self._retry_count = retry_count
        self._retry_delay = retry_delay
        self._logger = resolve_logger(logger)
        self._on_connection_opened = on_connection_opened
        self._com_application = self._acquire_com_application()

    def _acquire_com_application(self) -> Any:
        """Retrieves the COM GuiApplication object from SAP GUI Scripting.

        This method attempts to retrieve the SAPGUI COM object via
        win32com.client.GetObject, then accesses its GetScriptingEngine
        property to obtain the GuiApplication object.

        On failure, multiple attempts are made according to the configuration,
        with a delay between each attempt.

        Returns:
            The COM GuiApplication object returned by SAP GUI Scripting.

        Raises:
            SAPApplicationException: If the SAPGUI object is not found, if the
                scripting engine is unavailable, or if all retry attempts fail.
        """
        last_error: Exception | None = None

        for attempt in range(1, self._retry_count + 1):
            try:
                sap_gui = win32.GetObject("SAPGUI")

                application = sap_gui.GetScriptingEngine
                if application is None:
                    raise SAPApplicationException(
                        "SAP GUI Scripting engine is unavailable."
                    )

                self._logger("SAP GUI Scripting application retrieved successfully.", False)
                return application

            except Exception as exc:
                last_error = exc
                self._logger(
                    f"Attempt {attempt}/{self._retry_count} to access SAP GUI Scripting failed: {exc}",
                    True,
                )
                if attempt < self._retry_count:
                    time.sleep(self._retry_delay)

        raise SAPApplicationException(
            f"Unable to access SAP GUI Scripting engine. Ensure SAP GUI Scripting is enabled. "
            f"Detailed error: {last_error}",
            critical=True,
        ) from last_error

    @property
    def com_object(self) -> Any:
        """Returns the raw COM GuiApplication object.

        This property exposes the underlying COM object directly to enable
        advanced use cases not covered by the library's business API.

        Returns:
            The raw COM object representing the SAP GUI Scripting application.
        """
        return self._com_application

    def iter_connections(self) -> Iterator[SAPConnection]:
        """Iterates over open SAP connections via native COM enumeration.

        This method relies on the COM collection Application.Children as
        exposed by SAP GUI Scripting. Each COM connection retrieved is wrapped
        in a SAPConnection instance.

        Yields:
            SAPConnection: Each open connection currently managed by the application.

        Raises:
            SAPApplicationException: If iteration over open connections fails due to
                COM communication errors.
        """
        try:
            for com_connection in self._com_application.Children:
                yield SAPConnection(
                    application=self,
                    com_connection=com_connection,
                    config=self._config,
                    script_runner=self._script_runner,
                    logger=self._logger,
                )
        except Exception as exc:
            raise SAPApplicationException(
                f"Failed to iterate over open SAP connections: {exc}",
                critical=True,
            ) from exc

    def find_connection(self, description: str) -> SAPConnection | None:
        """Searches for an open connection by its description.

        The comparison is performed on the Description property of each open
        SAP connection. The description matching is case-sensitive after
        stripping whitespace.

        Args:
            description: The exact description of the connection to search for.

        Returns:
            The matching SAPConnection instance if found, otherwise None.
        """
        expected = description.strip()

        for connection in self.iter_connections():
            try:
                current = connection.description.strip()
                if current == expected:
                    return connection
            except Exception:
                continue

        return None

    def open_connection(self, sync: bool = True) -> SAPConnection:
        """Opens the target SAP connection defined in the configuration.

        This method calls OpenConnection() on the COM application using the
        connection description defined in SAPConfig. Once the COM connection
        is obtained, it is wrapped in a SAPConnection instance.

        If the configuration requests it, SAP credentials are automatically
        injected. If a registration callback was provided, the opened connection
        is notified to the calling component so it can be remembered.

        Args:
            sync: Whether the connection opening should be synchronous (waits
                for the window to fully open). Defaults to True.

        Returns:
            The open SAP connection wrapped in a business-logic object.

        Raises:
            SAPConnectionException: If OpenConnection returns None, if the
                connection description is invalid, or if an underlying COM error
                occurs during the connection process.
        """
        try:
            com_connection = self._com_application.OpenConnection(
                self._config.connection_description,
                sync,
            )

            if com_connection is None:
                raise SAPConnectionException(
                    "OpenConnection returned None (connection failed).",
                    critical=True,
                )

            connection = SAPConnection(
                application=self,
                com_connection=com_connection,
                config=self._config,
                script_runner=self._script_runner,
                logger=self._logger,
            )
            connection.fill_credentials_if_needed()

            if self._on_connection_opened is not None:
                self._on_connection_opened(connection)

            self._logger(f"SAP connection opened: {connection.description}", False)
            return connection

        except SAPConnectionException:
            raise
        except Exception as exc:
            raise SAPConnectionException(
                f"Failed to open SAP connection '{self._config.connection_description}': {exc}",
                critical=True,
            ) from exc