"""Management of the SAP Logon process.

This module provides a service to supervise the lifecycle of the SAP Logon
process on the operating system. It centralizes technical operations related
to the SAP client to avoid scattering this logic throughout the library.

Key responsibilities:
- Detecting running SAP Logon instances
- Launching SAP Logon if not already running
- Forcefully terminating SAP Logon instances
"""

from __future__ import annotations

import subprocess
from collections.abc import Callable
from pathlib import Path

import psutil

from .exceptions import SAPProcessException
from .utils import resolve_logger


class SAPProcessService:
    """Manages the lifecycle of the SAP Logon process.

    This class encapsulates the system operations necessary to interact with
    the SAP Logon executable. It relies on `psutil` for process detection and
    `subprocess` for launching the program.

    Upon initialization, the service validates that the provided path to the
    SAP executable exists. If the path is invalid, a business exception is
    immediately raised.

    Attributes:
        PROCESS_NAME (str): The name of the SAP Logon executable as it appears
            in the system process list.
        _sap_logon_path (Path): Path to the SAP Logon executable.
        _logger (Callable[[str, bool], None]): The logging function.

    Example:
        ```python
        service = SAPProcessService(r"C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPGUI\\saplogon.exe")
        if not service.is_running():
            service.launch()
        ```
    """

    PROCESS_NAME = "saplogon.exe"

    def __init__(
        self,
        sap_logon_path: Path | str,
        logger: Callable[[str, bool], None] | None = None,
    ) -> None:
        """Initializes the SAP Logon process management service.

        Args:
            sap_logon_path: Path to the SAP Logon executable.
            logger: Logging function. If None, :func:`~pysapmanager.utils.default_logger`
                is used. All output is subject to the library-wide switch
                controlled by :func:`pysapmanager.set_logging_enabled`.

        Raises:
            SAPProcessException: If the SAP executable does not exist at the
                specified path.
        """
        self._sap_logon_path = Path(sap_logon_path)
        self._logger = resolve_logger(logger)

        if not self._sap_logon_path.is_file():
            raise SAPProcessException(
                f"SAP executable not found: {self._sap_logon_path}",
                critical=True,
            )

    def is_running(self) -> bool:
        """Checks if SAP Logon is currently running.

        The detection is based on enumerating system processes via
        `psutil.process_iter()` and performing a case-insensitive comparison
        with the expected SAP Logon executable name.

        Returns:
            bool: True if at least one saplogon.exe process is detected,
                False otherwise.
        """
        return any(
            proc.info["name"] and proc.info["name"].lower() == self.PROCESS_NAME
            for proc in psutil.process_iter(["name"])
        )

    def launch(self) -> None:
        """Launches SAP Logon if it is not already running.

        If SAP Logon is already detected as active, this method does not launch
        a new instance and simply logs the information.

        Raises:
            SAPProcessException: If launching the process fails.
        """
        if self.is_running():
            self._logger("SAP Logon is already running.", False)
            return

        try:
            subprocess.Popen([str(self._sap_logon_path)])
            self._logger("SAP Logon launched successfully.", False)
        except Exception as exc:
            raise SAPProcessException(
                f"Failed to launch SAP Logon: {exc}",
                critical=True,
            ) from exc

    def kill(self) -> None:
        """Forcefully terminates all detected SAP Logon instances.

        This method iterates through all system processes and forces the
        termination of those matching saplogon.exe.

        If at least one instance is found and terminated, a success message
        is logged. Otherwise, an information message indicates no SAP Logon
        process was running.

        Raises:
            SAPProcessException: If an error occurs during process traversal or
                termination.
        """
        try:
            found = False
            for proc in psutil.process_iter(["name"]):
                name = proc.info["name"]
                if name and name.lower() == self.PROCESS_NAME:
                    proc.kill()
                    found = True

            if found:
                self._logger("SAP Logon closed.", False)
            else:
                self._logger("No SAP Logon process to close.", False)
        except Exception as exc:
            raise SAPProcessException(
                f"Failed to close SAP Logon: {exc}",
                critical=True,
            ) from exc