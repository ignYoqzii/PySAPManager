"""Representation of a SAP GUI session.

This module defines an abstraction of a single SAP GUI session.

The SAPSession class encapsulates a COM object representing a SAP session and
exposes a simplified API enabling:

- Access to the session name
- Access to the parent connection
- Retrieval of the underlying raw COM object
- Determining if the session is busy
- Resolving technical session and connection indices
- Executing a VBScript on the session synchronously or asynchronously
"""

from __future__ import annotations

from concurrent.futures import Future
from typing import TYPE_CHECKING, Any

from .models import ScriptExecutionRequest
from .script_runner import VBScriptRunner

if TYPE_CHECKING:
    from .connection import SAPConnection


class SAPSession:
    """Represents a single SAP GUI session.

    This class encapsulates an individual SAP session exposed by COM and
    stores the information necessary for executing scripts targeted at
    this session.

    A SAPSession instance is typically created from a SAP connection via
    `SAPConnection`.

    Attributes:
        _connection (SAPConnection): Parent SAP connection.
        _com_session (Any): Raw COM object representing the SAP session.
        _script_runner (VBScriptRunner): Script executor.

    Example:
        ```python
        session = connection.first_session()
        output = session.execute_script(request)
        future = session.execute_script_async(request)
        ```
    """

    def __init__(
        self,
        *,
        connection: SAPConnection,
        com_session: Any,
        script_runner: VBScriptRunner,
    ) -> None:
        """Initializes a SAP session.

        Args:
            connection: Parent SAP connection.
            com_session: COM object representing the SAP session.
            script_runner: Script executor used to run VBScripts.
        """
        self._connection = connection
        self._com_session = com_session
        self._script_runner = script_runner

    @property
    def connection(self) -> SAPConnection:
        """Parent SAP connection."""
        return self._connection

    @property
    def com_object(self) -> Any:
        """Raw COM `GuiSession` object.

        This property exposes the underlying COM object directly to enable
        advanced interactions not covered by the SAPSession business API.

        Returns:
            Any: COM object representing the SAP session.
        """
        return self._com_session

    @property
    def name(self) -> str:
        """SAP session name."""
        return str(self._com_session.Name)

    @property
    def is_busy(self) -> bool:
        """Whether the SAP session is busy.

        Returns:
            bool: True if the session is busy, False otherwise.
        """
        return bool(self._com_session.Busy)

    @property
    def session_index(self) -> int:
        """Index of the session in its parent connection.

        This index is used to pass the correct session context to
        VBScripts executed via `cscript.exe`.

        Returns:
            int: Index of the session in the parent connection's session collection.

        Raises:
            ValueError: If the session index cannot be determined.
        """
        parent = self._connection.com_object
        session_id = self._com_session.Id

        for i in range(parent.Children.Count):
            if parent.Children(i).Id == session_id:
                return i

        raise ValueError("Failed to determine SAP session index.")

    @property
    def connection_index(self) -> int:
        """Index of the parent connection in the SAP application.

        This index is used to pass the correct connection context to
        VBScripts executed via `cscript.exe`.

        Returns:
            int: Index of the parent connection in the SAP application's connection
                collection.

        Raises:
            ValueError: If the connection index cannot be determined.
        """
        app = self._connection.application.com_object
        connection_id = self._connection.com_object.Id

        for i in range(app.Children.Count):
            if app.Children(i).Id == connection_id:
                return i

        raise ValueError("Failed to determine SAP connection index.")

    def execute_script(self, request: ScriptExecutionRequest) -> str:
        """Executes a VBScript on this session.

        This method delegates execution to `VBScriptRunner`, passing the context
        of the current session and connection along with parameters from the request.

        Args:
            request: Script execution request parameters.

        Returns:
            str: Standard output returned by the VBScript.

        Raises:
            SAPScriptExecutionException: If script execution fails, if the
                script cannot be decoded with the resolved encoding, or if
                `request.script_path` is already being executed concurrently
                by another in-flight call (see `VBScriptRunner`'s concurrency
                Warning).
        """
        return self._script_runner.execute(
            script_path=request.script_path,
            script_args=request.script_args,
            session_index=self.session_index,
            connection_index=self.connection_index,
            timeout=request.timeout,
            creates_excel=request.creates_excel,
            created_file_path=request.created_file_path,
            encoding=request.encoding,
        )

    def execute_script_async(self, request: ScriptExecutionRequest) -> Future[str]:
        """Executes a VBScript on this session asynchronously.

        Note:
            If `request.script_path` is already in use by another in-flight
            execution (on this session or another), the concurrency guard in
            `VBScriptRunner.execute` raises `SAPScriptExecutionException` -
            but since this happens on the worker thread, that error surfaces
            when you call `future.result()`, not from this method directly.

        Args:
            request: Script execution request parameters.

        Returns:
            Future[str]: Future object representing the asynchronous task.
        """
        return self._script_runner.execute_async(
            script_path=request.script_path,
            script_args=request.script_args,
            session_index=self.session_index,
            connection_index=self.connection_index,
            timeout=request.timeout,
            creates_excel=request.creates_excel,
            created_file_path=request.created_file_path,
            encoding=request.encoding,
        )