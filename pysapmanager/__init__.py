"""# PySAPManager - Python SAP GUI Automation Library.

PySAPManager is a professional Python library for automating SAP operations via SAP GUI
Scripting. It provides an object-oriented API to automate SAP tasks from Python by
leveraging the SAP GUI Scripting engine. The package enables environment preparation,
opening target SAP connections, manipulating sessions, executing VBScript scripts directly
recorded from SAP GUI, and orchestrating parallel executions while respecting COM and SAP GUI constraints.

## Key Features

- Automated startup and supervision of the **SAP Logon** process
- Opening and management of target **SAP** connections
- Creation and retrieval of **SAP** sessions
- Synchronous **SAP** script execution
- Asynchronous **SAP** script execution with parallel control
- Support for **COM** initialization in multi-threaded environments
- Simple integration into automation workflows
- Type hints and comprehensive error handling
- Built-in support for Excel file operations within scripts

## Prerequisites

- **SAP GUI** installed on the workstation
- **SAP GUI Scripting** enabled on the client side
- **SAP GUI Scripting** authorized on the server side
- **Windows** environment with **COM** support

## Quick Start Example

```python
from pysapmanager import SAPManager, SAPConfig, ScriptExecutionRequest

config = SAPConfig(
    sap_logon_path=r"C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPGUI\\saplogon.exe",
    connection_description="description",
)

with SAPManager(config=config) as manager:
    application = manager.get_application()
    connection = application.open_connection()
    session = connection.first_session()

    output = session.execute_script(
        ScriptExecutionRequest(
            script_path=r"C:\\path\\to\\script.vbs",
            creates_excel=True,
            created_file_path=r"C:\\path\\to\\output.xlsx",
        )
    )
```

## Disabling Logging

By default, PySAPManager components log status messages via a simple console
logger (or a custom logger you provide). To silence all log output from the
library process-wide:

```python
from pysapmanager import set_logging_enabled

set_logging_enabled(False)  # silence PySAPManager everywhere, including custom loggers
```

`SAPManager` also accepts an `enable_logging` constructor argument as a
shortcut for the same global switch:

```python
with SAPManager(config=config, enable_logging=False) as manager:
    ...  # no PySAPManager log output during this block or afterwards
```

## Multi-Threaded Example

```python
from pysapmanager import SAPManager, SAPConfig, ScriptExecutionRequest, com_initialized

config = SAPConfig(
    sap_logon_path=r"C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPGUI\\saplogon.exe",
    connection_description="description",
)

with com_initialized():
    with SAPManager(config=config, max_parallel_scripts=2) as manager:
        application = manager.get_application()
        connection = application.open_connection()
        first_session = connection.first_session()
        other_session = connection.create_session()

        # Submit scripts asynchronously
        first_task = first_session.execute_script_async(
            ScriptExecutionRequest(
                script_path=r"C:\\path\\to\\script1.vbs",
                creates_excel=True,
                created_file_path=r"C:\\path\\to\\output1.xlsx",
            )
        )

        other_task = other_session.execute_script_async(
            ScriptExecutionRequest(
                script_path=r"C:\\path\\to\\script2.vbs",
                creates_excel=True,
                created_file_path=r"C:\\path\\to\\output2.xlsx",
            )
        )

        # Your code executes in parallel with the scripts

        # Wait for all scripts to complete
        manager.wait_all(first_task, other_task)

        # Your code executes after scripts complete
```

## Public API

The package exposes the following main classes and functions:

- `SAPManager`: Main facade for SAP automation
- `SAPApplication`: Access to SAP GUI Scripting application
- `SAPConnection`: Represents an open SAP connection
- `SAPSession`: Represents a SAP session
- `SAPConfig`: Configuration for SAP connections
- `SAPCredentials`: Login credentials for SAP
- `ScriptExecutionRequest`: Script execution parameters
- `com_initialized()`: Context manager for COM initialization
- `set_logging_enabled(enabled)`: Globally enable/disable all library logging
- `is_logging_enabled()`: Check the current global logging state

Attributes:
    __title__ (str): The official name of the package.
    __version__ (str): The current version number of the package.
    __author__ (str): The author and maintainer of the package.
    __license__ (str): The license under which the package is distributed.
    __all__ (list of str): The public API exposed by the package for `from pysapmanager import *` usage.
"""

from .application import SAPApplication
from .com import com_initialized
from .connection import SAPConnection
from .manager import SAPManager
from .models import SAPConfig, SAPCredentials, ScriptExecutionRequest
from .session import SAPSession
from .utils import is_logging_enabled, set_logging_enabled

__title__ = "PySAPManager"
__version__ = "0.1.0"
__author__ = "yoqzii"
__license__ = "MIT License"

__all__ = [
    "SAPManager",
    "SAPApplication",
    "SAPConnection",
    "SAPSession",
    "SAPConfig",
    "SAPCredentials",
    "ScriptExecutionRequest",
    "com_initialized",
    "set_logging_enabled",
    "is_logging_enabled",
]