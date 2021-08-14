# xlwings installer

## Create installer

Go to `Releases` > `Draft/Create a new release`. Add a version like `1.0.0` to `Tag version`, then hit `Publish release`.

Wait a few minutes until the installer will appear under the release. You can follow the progress under the `Actions` tab.

## Setup

### Excel file

You can add your Excel file to this repository if you like.

* Add the standalone xlwings VBA module and the Dictionary class module, e.g. via `xlwings quickstart project --standalone`. In case you would like to start with an existing Excel file, it may still be the easiest to create the standalone quickstart file, then go the the VBA editor (`Alt-F11`) and export the two modules, then import them again in the existing file.
* Make sure that in the VBA editor (`Alt-F11`) under `Tools` > `References` xlwings is unchecked
* Rename the `_xlwings.conf` sheet into `xlwings.conf`
* As `Interpreter_Win`, set the following value: `%LOCALAPPDATA%\project\pythonw.exe` while replacing `project` with the name of your project (Windows)
* As `Interpreter_Mac`, set the following value: `$HOME/project/bin/python` while replacing `project` with the name of your project (macOS)
* If you like, you can hide `xlwings.conf` 

### Source code

Source code can either be embedded in the Excel file or added to the `src` directory here. The first option requires xlwings PRO, the second option will also work with xlwings CE.

### Data files

It is recommended to store your data files in the `data` directory. The directory gets copied next to the Python executable and can be accessed like this:

```
import sys
from pathlib import Path
DATA_DIR = Path(sys.executable).resolve().parent / 'data'
```

### Dependencies

Add your dependencies to `requirements.txt`.

If you have dependencies from **public** Git repos, you can add them by using the `https` version of the Git clone URL (`#subdirectory=<DIRECTORY>` is only required if your `setup.py` file is not in the root directory):

```
git+https://github.com/<USERNAME>/<REPO>.git@<BRANCH OR TAG OR SHA>#subdirectory=<DIRECTORY>
```

If you have dependencies from **private** Git repos, it's a bit more complicated:

**In your private repo with the source for Python package(s)**:  
* Create a deploy key via by executing the following command on a Terminal: `ssh-keygen -t rsa -b 4096 -m PEM` - it will ask you for a passphrase: Hit Enter without entering one and same for the confirmation. Note that your key needs to be in the PEM format. If you want to use an existing one, make sure it starts with `----BEGIN RSA PRIVATE KEY-----` and not with `-----BEGIN OPENSSH PRIVATE KEY-----`. You can convert an existing key into PEM format like this: `ssh-keygen -p -m PEM -f /path/to/id_rsa`.
* Go to `Settings` > `Deploy keys` and add the public key as deploy key (usually `id_rsa.pub`).


**In this repo**:  
* Use the `ssh` version of the Git clone url, however, **you need to replace the `:` with `/`**, it should look like this:

```
git+ssh://git@github.com/<USERNAME>/<REPO>.git@<BRANCH OR TAG OR SHA>#subdirectory=<DIRECTORY>
```
* Uncomment the `Install SSH key` section in `.github/workflows/main.yml`
* On this installer repo, go to `Settings` > `Secrets` and add the secret variable called `SSH_KEY`: Paste the private ssh deploy key from above (usually `id_rsa`).


### Code signing

* Store your code sign certificate as `sign_cert_file` in the root of this repository (make sure your repo is private).
* Go to `Settings` > `Secrets` and add the password as `code_sign_password`.

### Project details

Update the following under `.github/workflows/main.yml`:

```
PROJECT: 
APP_ID: 
APP_PUBLISHER: 
```

Comment out the `Code signing` section if you don't have a certificate.

### Python version

Set your Python version under `.github/workflows/main.yml`:

```
python-version: '3.8'
architecture: 'x64'
```
