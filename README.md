# PySAPManager

A professional Python library for automating SAP operations via SAP GUI Scripting. PySAPManager provides an object-oriented API to automate SAP tasks from Python by leveraging the SAP GUI Scripting engine.

## Overview

PySAPManager enables you to:

- Automate startup and supervision of the SAP Logon process
- Open and manage target SAP connections
- Create and retrieve SAP sessions
- Execute VBScript scripts directly recorded from SAP GUI (both synchronously and asynchronously)
- Orchestrate parallel script executions while respecting COM and SAP GUI constraints
- Handle multi-threaded automation workflows
- Work with Excel files within automated scripts

## Prerequisites

- **Windows** environment with COM support
- **SAP GUI** installed on the workstation
- **SAP GUI Scripting** enabled on the client side
- **SAP GUI Scripting** authorized on the server side

## Installation

Install the latest version from PyPI:

*In the future...*

Or install from source:

```bash
git clone https://github.com/ignYoqzii/pysapmanager.git
cd pysapmanager
pip install -e .
```

### Dependencies

- `psutil` - Cross-platform process management
- `pywin32` - Windows COM interface support

These are automatically installed with the package.

## Quick Start

### Basic Usage

```python
from pysapmanager import SAPManager, SAPConfig, ScriptExecutionRequest

# Configure SAP connection
config = SAPConfig(
    sap_logon_path=r"C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe",
    connection_description="your_connection_name",
)

# Use context manager for automatic cleanup
with SAPManager(config=config) as manager:
    application = manager.get_application()
    connection = application.open_connection()
    session = connection.first_session()
    
    # Execute a script
    output = session.execute_script(
        ScriptExecutionRequest(
            script_path=r"C:\path\to\your_script.vbs",
            creates_excel=True,
            created_file_path=r"C:\path\to\output.xlsx",
        )
    )
```

### Multi-Threaded Automation

For parallel script execution across multiple sessions:

```python
from pysapmanager import SAPManager, SAPConfig, ScriptExecutionRequest, com_initialized

config = SAPConfig(
    sap_logon_path=r"C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe",
    connection_description="your_connection_name",
)

# Use com_initialized() for multi-threaded environments
with com_initialized():
    with SAPManager(config=config, max_parallel_scripts=2) as manager:
        application = manager.get_application()
        connection = application.open_connection()
        
        # Create multiple sessions
        session1 = connection.first_session()
        session2 = connection.create_session()
        
        # Submit scripts asynchronously
        task1 = session1.execute_script_async(
            ScriptExecutionRequest(
                script_path=r"C:\path\to\script1.vbs",
                creates_excel=True,
                created_file_path=r"C:\path\to\output1.xlsx",
            )
        )
        
        task2 = session2.execute_script_async(
            ScriptExecutionRequest(
                script_path=r"C:\path\to\script2.vbs",
                creates_excel=True,
                created_file_path=r"C:\path\to\output2.xlsx",
            )
        )
        
        # Wait for all scripts to complete
        manager.wait_all(task1, task2)
```

## API Reference

### Core Classes

**SAPManager**
Main facade for SAP automation. Handles SAP Logon startup, connection management, and script execution orchestration.

- `get_application()` - Access SAP GUI Scripting application
- `execute_script()` - Execute a script synchronously
- `execute_script_async()` - Execute a script asynchronously
- `wait_all(*tasks)` - Wait for all async tasks to complete

**SAPConfig**
Configuration for SAP connections.

- `sap_logon_path` - Path to saplogon.exe
- `connection_description` - SAP connection description
- Additional optional parameters for advanced configurations

**SAPApplication**
Access to SAP GUI Scripting application interface.

- `open_connection()` - Open a SAP connection
- `close()` - Close the application

**SAPConnection**
Represents an open SAP connection.

- `first_session()` - Get the first available session
- `create_session()` - Create a new session in the connection
- `close()` - Close the connection

**SAPSession**
Represents a SAP session for script execution.

- `execute_script(request)` - Execute a script synchronously
- `execute_script_async(request)` - Execute a script asynchronously

**ScriptExecutionRequest**
Parameters for script execution.

- `script_path` - Path to the VBScript file
- `creates_excel` - Whether the script creates an Excel file
- `created_file_path` - Path where the output Excel file will be saved

### Utility Functions

**com_initialized()**
Context manager for proper COM initialization in multi-threaded environments.

```python
with com_initialized():
    # COM-dependent code here
    pass
```

**set_logging_enabled(enabled)**
Globally enable or disable all library logging.

```python
from pysapmanager import set_logging_enabled

set_logging_enabled(False)  # Silence all PySAPManager log output
```

**is_logging_enabled()**
Check the current global logging state.

```python
from pysapmanager import is_logging_enabled

if is_logging_enabled():
    print("Logging is enabled")
```

## Configuration

### Disabling Logging

By default, PySAPManager logs status messages via a console logger. To silence all output:

```python
from pysapmanager import set_logging_enabled

set_logging_enabled(False)
```

Or pass the `enable_logging` parameter to SAPManager:

```python
with SAPManager(config=config, enable_logging=False) as manager:
    # No logging output during this block
    pass
```

### Custom SAP Credentials

For automated login with credentials:

```python
from pysapmanager import SAPConfig, SAPCredentials

credentials = SAPCredentials(
    username="your_username",
    password="your_password",
)

config = SAPConfig(
    sap_logon_path=r"C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe",
    connection_description="your_connection",
    credentials=credentials,
)
```

## Features

- **Type Hints** - Full type annotations for IDE support and type checking
- **Error Handling** - Comprehensive exception hierarchy for easy error management
- **Context Managers** - Automatic resource cleanup with `with` statements
- **Async Execution** - Non-blocking script execution with parallel control
- **COM Support** - Proper handling of COM in multi-threaded environments
- **Process Management** - Automatic SAP Logon startup and supervision
- **Plug and Play** - Use recorded scripts from SAP and use them with the manager easily

## Error Handling

The library provides specific exceptions for different error scenarios:

```python
from pysapmanager import SAPManager, SAPConfig

config = SAPConfig(...)

try:
    with SAPManager(config=config) as manager:
        # Your SAP automation code
        pass
except Exception as e:
    print(f"SAP automation failed: {e}")
```

## Best Practices

1. **Always use context managers** - Ensures proper cleanup of SAP connections and processes
   ```python
   with SAPManager(config=config) as manager:
       # Your code here
   ```

2. **Use `com_initialized()` for multi-threaded code** - Required for proper COM initialization in threads
   ```python
   with com_initialized():
       with SAPManager(config=config) as manager:
           # Your multi-threaded code
   ```

3. **Set realistic parallelism limits** - Don't exceed your SAP system's capacity and SAP session's limit
   ```python
   manager = SAPManager(config=config, max_parallel_scripts=2)
   ```

4. **Handle async tasks properly** - Always wait for async tasks before closing the manager
   ```python
   task = session.execute_script_async(request)
   manager.wait_all(task)
   ```

5. **Store credentials securely** - Never hardcode credentials; use environment variables or secure vaults
   ```python
   import os
   
   username = os.getenv("SAP_USERNAME")
   password = os.getenv("SAP_PASSWORD")
   ```

## Troubleshooting

### SAP GUI Scripting Not Enabled

If you get an error about SAP GUI Scripting not being available:

1. Verify SAP GUI Scripting is enabled on your client

2. Ensure SAP GUI Scripting is authorized on the server side
   - Contact your SAP system administrator

### COM Errors in Multi-Threaded Code

Always wrap multi-threaded code with `com_initialized()`:

```python
from pysapmanager import com_initialized

with com_initialized():
    # Your multi-threaded SAP automation code
    pass
```

### Scripts Timeouts

Increase timeout values in ScriptExecutionRequest if your script takes longer to finish:

```python
request = ScriptExecutionRequest(
    script_path=r"Path\to\your\script.vbs",
    timeout=180, # 3 minutes
)
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Related Projects

- [pywin32](https://github.com/pywin32/pywin32) - Python for Windows Extensions
- [psutil](https://github.com/giampaolo/psutil) - Cross-platform process management

## Disclaimer

This library is not affiliated with or endorsed by SAP SE. SAP and all its products are trademarks or registered trademarks of SAP SE in Germany and other countries.
