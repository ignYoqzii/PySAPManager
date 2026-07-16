"""Management of a SAP GUI connection.

This module defines a high-level abstraction for manipulating an open SAP GUI
connection via the SAP/COM scripting interface.

The SAPConnection class encapsulates a COM object of type GuiConnection and
exposes the main operations expected on a SAP connection:

- Access to its main properties
- Retrieve one or more existing sessions
- Create a new SAP session
- Close the connection
- Inject credentials if configuration requires it
"""

from __future__ import annotations

import time
from collections.abc import Callable
from typing import TYPE_CHECKING, Any

from .exceptions import SAPConnectionException, SAPSessionException
from .models import SAPConfig
from .script_runner import VBScriptRunner
from .session import SAPSession
from .utils import resolve_logger

if TYPE_CHECKING:
    from .application import SAPApplication


class SAPConnection:
    """Represents an open SAP GUI connection.

    This class encapsulates a SAP connection exposed via COM and provides a
    simplified business API for manipulating its sessions and performing
    standard operations like closing the connection or logging in.

    It relies on a combination of raw COM objects, custom session wrappers,
    and business configurations.

    Attributes:
        MAX_SESSIONS (int): The maximum number of sessions allowed for a single 
            SAP connection (SAP GUI limit is 6).
        _application (SAPApplication): The parent SAP application instance.
        _com_connection (Any): The underlying COM GuiConnection object.
        _config (SAPConfig): The configuration settings for this connection.
        _script_runner (VBScriptRunner): Executor for VBScript files.
        _logger (Callable[[str, bool], None]): The logging function.

    Example:
        >>> connection = application.open_connection()
        >>> session = connection.first_session()
        >>> new_session = connection.create_session()
        >>> connection.close()
    """

    MAX_SESSIONS = 6
    """The maximum number of sessions a user can have open simultaneously on a single connection."""

    def __init__(
        self,
        *,
        application: SAPApplication,
        com_connection: Any,
        config: SAPConfig,
        script_runner: VBScriptRunner,
        logger: Callable[[str, bool], None] | None = None,
    ) -> None:
        """Initializes a SAP connection wrapper.

        Args:
            application: Parent SAP application instance that spawned this connection.
            com_connection: COM object representing an existing active SAP connection.
            config: SAP configuration associated with this specific connection.
            script_runner: Component responsible for executing VBScripts.
            logger: Logging function. If None, :func:`~pysapmanager.utils.default_logger`
                is used. All output is subject to the library-wide switch
                controlled by :func:`pysapmanager.set_logging_enabled`.
        """
        self._application = application
        self._com_connection = com_connection
        self._config = config
        self._script_runner = script_runner
        self._logger = resolve_logger(logger)

    @property
    def application(self) -> SAPApplication:
        """SAPApplication: The parent SAP GUI Scripting application."""
        return self._application

    @property
    def com_object(self) -> Any:
        """Any: The raw COM GuiConnection object."""
        return self._com_connection

    @property
    def name(self) -> str:
        """str: The technical name of the SAP connection."""
        return self._com_connection.Name

    @property
    def description(self) -> str:
        """str: The business description of the SAP connection."""
        return self._com_connection.Description

    @property
    def session_count(self) -> int:
        """int: Number of sessions currently open in the connection."""
        return self._com_connection.Children.Count

    def sessions(self) -> list[SAPSession]:
        """Returns all sessions currently open in the connection.

        This method materializes the collection of SAP sessions as a Python list.
        It is appropriate for the SAP GUI context, where a connection can contain
        a maximum of six sessions.

        Returns:
            list[SAPSession]: A list of all active SAP sessions.

        Raises:
            SAPSessionException: If a session cannot be retrieved or wrapped 
                correctly due to a COM communication issue.
        """
        return [self.session(i) for i in range(self.session_count)]

    def session(self, index: int) -> SAPSession:
        """Returns the session at the specified index.

        Args:
            index: The 0-based index of the session to retrieve.

        Returns:
            SAPSession: The SAP session at the given index.

        Raises:
            SAPSessionException: If the session cannot be retrieved or if 
                the index is out of bounds.
        """
        try:
            com_session = self._com_connection.Children(index)
            return SAPSession(
                connection=self,
                com_session=com_session,
                script_runner=self._script_runner,
            )
        except Exception as exc:
            raise SAPSessionException(
                f"Failed to retrieve SAP session at index {index}: {exc}",
                critical=True,
            ) from exc

    def first_session(self) -> SAPSession:
        """Returns the first session of the connection.

        Returns:
            SAPSession: The first active session (index 0).

        Raises:
            SAPSessionException: If the first session cannot be accessed.
        """
        return self.session(0)

    def create_session(
        self,
        *,
        timeout: float = 10.0,
        poll_delay: float = 0.2,
    ) -> SAPSession:
        """Creates and returns a new SAP session.

        Session creation is triggered from the first session of the connection.
        The method then performs a polling loop to wait for the new session to 
        appear in the connection's children collection.

        Args:
            timeout: Maximum time in seconds allowed for the new session to appear.
                Defaults to 10.0.
            poll_delay: Interval in seconds between checks for session existence.
                Defaults to 0.2.

        Returns:
            SAPSession: The newly created SAP session.

        Raises:
            SAPSessionException: If the maximum session limit (MAX_SESSIONS) is 
                reached, if the operation times out, or if a COM error occurs.
        """
        try:
            initial_count = self.session_count

            if initial_count >= self.MAX_SESSIONS:
                raise SAPSessionException(
                    f"Maximum number of sessions ({self.MAX_SESSIONS}) reached.",
                    critical=True,
                )

            base_session = self._com_connection.Children(0)
            base_session.CreateSession()

            deadline = time.time() + timeout
            while time.time() < deadline:
                current_count = self.session_count
                if current_count > initial_count:
                    new_session = self.session(initial_count)
                    self._logger(
                        f"New SAP session created successfully: {new_session.name}.",
                        False
                    )
                    return new_session
                time.sleep(poll_delay)

            raise SAPSessionException(
                "Timeout during new SAP session creation.",
                critical=True,
            )

        except SAPSessionException:
            raise
        except Exception as exc:
            raise SAPSessionException(
                f"Failed to create a new SAP session: {exc}",
                critical=True,
            ) from exc

    def close(self) -> None:
        """Closes the current SAP connection.

        The method first attempts to close the connection via the native 
        COM API `CloseConnection()`. If that fails, it attempts a fallback 
        method by entering the '/nex' command into the SAP command field 
        of the first session.

        Raises:
            SAPConnectionException: If all attempts to close the connection fail.
        """
        try:
            self._com_connection.CloseConnection()
            self._logger("SAP connection closed via CloseConnection().", False)
            return

        except Exception as close_exc:
            self._logger(
                f"CloseConnection() failed: {close_exc}. Attempting via '/nex'.", True
            )

            try:
                com_session = self.first_session().com_object
                com_session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
                com_session.findById("wnd[0]").sendVKey(0)
                self._logger("SAP connection closed via '/nex' command.", False)

            except Exception as nex_exc:
                raise SAPConnectionException(
                    "Failed to close SAP connection. "
                    f"CloseConnection() failed ({close_exc}) and "
                    f"'/nex' failed ({nex_exc}).",
                    critical=True,
                ) from nex_exc

    def fill_credentials_if_needed(self) -> None:
        """Injects SAP credentials if the configuration requires it.

        This method checks if the configuration mandates a manual login. If 
        so, it locates the standard SAP fields (Client, User, Password, Language) 
        in the first session and populates them.

        Raises:
            SAPConnectionException: If the credential fields cannot be found 
                on the active screen or if an injection error occurs.
        """
        if not self._config.manual_login_required:
            return

        credentials = self._config.credentials
        if credentials is None:
            return

        try:
            com_session: Any = self.first_session().com_object

            com_session.findById("wnd[0]/usr/txtRSYST-MANDT").text = credentials.client
            com_session.findById("wnd[0]/usr/txtRSYST-BNAME").text = credentials.username
            com_session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = credentials.password
            com_session.findById("wnd[0]/usr/txtRSYST-LANGU").text = credentials.language

            self._logger("SAP credentials injected successfully.", False)
        except Exception as exc:
            raise SAPConnectionException(
                f"Failed to inject SAP credentials: {exc}",
                critical=True,
            ) from exc