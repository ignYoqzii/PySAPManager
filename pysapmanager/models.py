"""Data models for SAP configuration and script execution.

This module centralizes immutable data structures used to describe SAP
configuration and script execution parameters. Built on dataclasses, these
models provide simple, readable, and safe objects for transporting configuration
throughout the application.

The main entities are:

- SAPCredentials: Encapsulates SAP login credentials
- SAPConfig: Describes overall SAP connection configuration
- ScriptExecutionRequest: Encapsulates a VBScript execution request on a session
"""

from __future__ import annotations

import codecs
from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class SAPCredentials:
    """SAP login credentials for manual authentication.

    This immutable structure contains the information necessary to populate
    the SAP login screen when automatic credential injection is required.

    Security:
        `password` is excluded from this class's auto-generated `repr()`
        (it prints as `password=<hidden>` instead of the plaintext value).
        This matters because `SAPCredentials` instances are commonly embedded
        in a `SAPConfig` and end up inside exception messages, tracebacks, or
        ad-hoc `print()`/logging calls for debugging - without this,
        the plaintext password would be trivially leaked into logs.
        `str(credentials)` uses the same repr and is therefore also safe to
        log by accident, but avoid printing `credentials.password` directly.

    Attributes:
        username: SAP username.
        password: Password associated with the username. Hidden from `repr()`.
        client: SAP client number.
        language: SAP login language.

    Raises:
        ValueError: If any required value is empty or invalid.
    """

    username: str
    password: str = field(repr=False)
    client: str
    language: str

    def __post_init__(self) -> None:
        """Validates SAP credentials data.

        Raises:
            ValueError: If username, password, client, or language are empty.
        """
        if not self.username or not self.username.strip():
            raise ValueError("username cannot be empty.")

        if not self.password or not self.password.strip():
            raise ValueError("password cannot be empty.")

        if not self.client or not self.client.strip():
            raise ValueError("client cannot be empty.")

        if not self.language or not self.language.strip():
            raise ValueError("language cannot be empty.")


@dataclass(frozen=True)
class SAPConfig:
    """Main SAP GUI configuration.

    This immutable structure groups the parameters necessary to control SAP GUI
    from the application, including the path to the SAP Logon executable, the
    description of the target connection, and optionally the credentials to inject.

    Attributes:
        sap_logon_path: Path to the SAP Logon executable.
        connection_description: Description of the SAP connection as it appears
            in SAP Logon Pad.
        credentials: SAP credentials to use for automatic login screen filling.
            If None, no automatic credential injection is performed.

    Raises:
        ValueError: If SAP Logon path or connection description are empty.
    """

    sap_logon_path: str | Path
    connection_description: str
    credentials: SAPCredentials | None = None

    def __post_init__(self) -> None:
        """Validates and normalizes the SAP configuration.

        Raises:
            ValueError: If SAP Logon path or connection description are empty.
        """
        if not str(self.sap_logon_path).strip():
            raise ValueError("sap_logon_path cannot be empty.")

        if not self.connection_description or not self.connection_description.strip():
            raise ValueError("connection_description cannot be empty.")

    @property
    def manual_login_required(self) -> bool:
        """Whether credentials need to be injected manually.

        Returns:
            bool: True if credentials are provided, False otherwise.
        """
        return self.credentials is not None


@dataclass(frozen=True)
class ScriptExecutionRequest:
    """Request to execute a VBScript on a SAP session.

    This immutable structure encapsulates the parameters necessary to launch
    a VBScript to control a specific SAP session during automation. It describes
    the script to execute, its maximum execution time, and any expected effects
    on the file system.

    Attributes:
        script_path: Path to the VBScript file to execute.
        script_args: Additional arguments to pass to the script. Defaults to empty tuple.
        timeout: Maximum execution duration in seconds. Defaults to 120.
        creates_excel: Whether the script is expected to create or manipulate
            an Excel file. Defaults to False.
        created_file_path: Path to the file expected after execution, if applicable.
            Defaults to None.
        encoding: Text encoding to use when reading/patching this script's
            `.vbs` file (e.g. `"utf-16"`, `"utf-8"`, `"utf-8-sig"`, `"cp1252"`).
            Set this explicitly if a particular script is known to be saved
            with a non-default encoding, to avoid silently mis-decoded content.

    Raises:
        ValueError: If request parameters are inconsistent.
    """

    script_path: str | Path
    script_args: tuple[str, ...] = ()
    timeout: int = 120
    creates_excel: bool = False
    created_file_path: str | Path | None = None
    encoding: str = "utf-16"

    def __post_init__(self) -> None:
        """Validates and normalizes the script execution request.

        Raises:
            ValueError: If script path is empty, timeout is invalid,
                creates_excel is True without created_file_path, or encoding
                is provided but is not a valid/known text codec name.
        """
        if not str(self.script_path).strip():
            raise ValueError("script_path cannot be empty.")

        if self.timeout <= 0:
            raise ValueError("timeout must be greater than 0.")

        if self.creates_excel and self.created_file_path is None:
            raise ValueError("created_file_path is required when creates_excel=True.")

        if not self.encoding.strip():
            raise ValueError("encoding cannot be empty when provided.")
        try:
            codecs.lookup(self.encoding)
        except LookupError as exc:
            raise ValueError(f"Unknown text encoding: '{self.encoding}'.") from exc

        if self.created_file_path is not None:
            # Bypass frozen=True to set normalized Path object
            object.__setattr__(
                self,
                "created_file_path",
                Path(self.created_file_path),
            )

            if not str(self.created_file_path).strip():
                raise ValueError("created_file_path cannot be empty.")