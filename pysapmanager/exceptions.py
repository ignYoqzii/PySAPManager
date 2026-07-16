"""SAP automation exception classes.

This module defines custom exception classes specific to the PySAPManager library.
All exceptions inherit from SAPException and include a 'critical' flag to
indicate whether an error should be treated as blocking by calling code.

This separation allows for distinguishing SAP-specific business errors from
generic Python exceptions, enabling more precise error handling and recovery
strategies.
"""


class SAPException(Exception):
    """Base exception for SAP GUI automation.

    This class serves as the root for all library-specific exceptions. It adds
    a 'critical' attribute allowing calling code to distinguish critical errors
    from potentially recoverable errors.

    Attributes:
        critical (bool): Indicates whether this error is considered critical
            or blocking by the automation workflow.
    """

    def __init__(self, message: str, *, critical: bool = False) -> None:
        """Initializes the base SAP exception.

        Args:
            message: Descriptive error message explaining the cause of the exception.
            critical: Whether this error should be treated as critical/blocking.
                Defaults to False.
        """
        super().__init__(message)
        self.critical: bool = critical


class SAPProcessException(SAPException):
    """Exception raised during SAP Logon process operations.

    This exception is used when errors occur while interacting with the SAP 
    Logon executable or the lifecycle management of the GUI process.
    """


class SAPApplicationException(SAPException):
    """Exception raised during SAP GUI Scripting application access.

    This exception is used when there are issues retrieving or communicating
    with the COM `GuiApplication` object.
    """


class SAPConnectionException(SAPException):
    """Exception raised during SAP connection operations.

    This exception covers issues such as failed connection opening, invalid
    connection descriptions, or errors occurring during connection shutdown.
    """


class SAPSessionException(SAPException):
    """Exception raised during SAP session operations.

    This exception occurs when manipulating active SAP sessions, such as 
    session creation, session retrieval, or interactions within a session.
    """


class SAPScriptExecutionException(SAPException):
    """Exception raised during VBScript execution on a SAP session.

    This exception indicates that a VBScript file triggered by the automation 
    failed to execute or returned an error status.
    """