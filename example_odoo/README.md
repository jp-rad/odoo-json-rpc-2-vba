# odoo-addons-dev

A development environment template for Odoo custom modules using VSCode and Docker.

## Overview

This template lets you quickly set up a development environment for Odoo custom modules with VSCode and Docker. You can also test standard Odoo applications.

### Prerequisites

- [git](https://git-scm.com/)
- [VSCode](https://code.visualstudio.com/download)
- [Docker](https://www.docker.com/)

## Getting Started

### Using the Template

1. Open the repository on GitHub:  
   https://github.com/jp-rad/odoo-addons-dev
2. Click `Use this template`:  
   [Generate a new repository](https://github.com/jp-rad/odoo-addons-dev/generate)

### Cloning the Repository

Open Command Prompt and run:

```cmd
mkdir c:\workgit
cd c:\workgit
git clone <your GitHub repository URL>
```

### Opening in VSCode Container

1. Open the cloned folder in VSCode.
2. Install the "Remote - Containers" extension.
3. Click the "Remote - Containers" icon in the bottom left.
4. Select "Remote-Containers: Reopen in Container".
5. The Docker container will start, and you can develop remotely in VSCode.

### Registering the Root CA Certificate on Your Local Machine

This development environment uses nginx with a self‑signed Root CA to provide HTTPS.  
To access the site securely, you must register the Root CA certificate on your OS.

1. Open the following URL in your browser:  
   **https://localhost/trustme**
2. The file `dev-root-certificate.crt` will be downloaded.
3. Register `dev-root-certificate.crt` in your OS as a **Trusted Root Certification Authority**.

#### Example (Windows)

1. Right‑click `dev-root-certificate.crt`.
2. Select **Install Certificate**.
3. Choose **Local Machine**.
4. Select **Trusted Root Certification Authorities** as the certificate store.
5. Restart your browser.

After this, you can safely access **https://localhost/**.

### custom_addons

Develop your own add-ons in the `custom_addons` folder.  
See the [Odoo developer tutorial](https://www.odoo.com/documentation/master/developer/howtos/backend.html).

To create a new module, run:

```
/opt/odoo/repo/odoo/odoo-bin scaffold <module_name>
```

### Logging in to Odoo

1. Click the "Run and Debug" icon in VSCode.
2. Select "Odoo addons" and click "Start Debugging" to launch Odoo.
3. Open your browser and go to [http://localhost/](http://localhost/).
4. Log in with `admin` / `admin`.

## License

### odoo-addons-dev

This repository is licensed under the [MIT License](https://opensource.org/licenses/MIT).  
You may use, modify, and redistribute it for commercial or non-commercial purposes.

### License for Your Custom Modules

For modules under `custom_addons`, choose either [LGPLv3](https://www.gnu.org/licenses/lgpl-3.0.html) or [AGPLv3](https://www.gnu.org/licenses/agpl-3.0.html):

- **LGPLv3**: Standard for Odoo core. Allows commercial use and closed-source integration, but requires publishing changes to LGPL parts.
- **AGPLv3**: Standard for OCA modules. Requires source code disclosure, even for server-side use.

Select the license based on your distribution policy and compatibility needs.

### Odoo and OCA Licenses

Odoo core is mainly LGPLv3; OCA modules are AGPLv3 or LGPLv3.  
Always check the license terms for each project.

For details, see:  
- [OCA / Odoo Meeting on Licenses](https://odoo-community.org/blog/news-updates-1/oca-odoo-meeting-on-licenses-21)
