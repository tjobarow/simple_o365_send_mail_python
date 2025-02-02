
# Simple Send Mail for Microsoft Graph API

A lightweight (requiring only one external dependency - ```requests``` to run) python wrapper over the [user: sendMail](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http) API endpoint, allowing you to easily integrate mail functionality into any Python workflow. Supports including file attachments, CC recipients, and BCC recipients.

_Note: While the wrapper does only require ```requests``` to run, you do need to have ```build``` and ```wheel``` installed to build the package from source._

## Requirements

### Python Dependencies

To build and import the package you need to install:

```
build==1.2.2.post1
wheel==0.45.1
requests==2.32.3
```

or rather

```
pip install -r requirements.txt
```

### Microsoft Graph API Requirements
- Enterprise application with OAuth credentials, and permission to use the [SendMail Graph API endpoint](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http#permissions)
- An existing user account to source the email from

## Installation (Build from source)
1. To build SimpleO365SendMail from source, first clone the repo to a local directory.

2. Create a new python virtual environment (if you have Python's ```venv``` module installed)

```python -m venv .my-venv```

3. Activate the environment (depends on OS)

- __Linux/MacOS:__

```source ./.my-venv/bin/activate```

- __Windows (PowerShell|CMD prompt)__

```./.my-venv/Scripts/[Activate.ps1|Activate.bat]```

4. Install buildtools & wheel

```pip install build wheel```

5. Build package from within cloned repo

```python -m build --wheel```

6. Install SimpleO365SendMail using pip from within root directory of SimpleO365SendMail
   
```pip install .``` 
    
## Quick Start

Please check out the [example_usage.py](https://github.com/tjobarow/simple_o365_send_mail_python/blob/7e4fd986a3011da5eba693505c9b1e9decf335bd/example_usage.py) file to learn how to use the package! It's fairly simple overall.

