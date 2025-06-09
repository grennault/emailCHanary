# 🕵️‍♂️📬👨🏻‍💻 Email🇨🇭anary

![License](https://img.shields.io/github/license/grennault/emailCHanary)
![Issues](https://img.shields.io/github/issues/grennault/emailCHanary)
![Stars](https://img.shields.io/github/stars/grennault/emailCHanary)
![Last Commit](https://img.shields.io/github/last-commit/grennault/emailCHanary)

> **EmailCHanary**: Canary-style email beaconing for detecting your (and your clients') compromised Azure users.

## 🔍 What is EmailCHanary?

As date of today, **EmailCHanary** is a lightweight and stealthy experiment script designed to **detect compromised Microsoft Azure email clients**. It works by deploying decoy (canary) email addresses that alert you when they receive email, signaling potential user account compromise.

Think of it as a tripwire for your cloud-based email clients. If someone send an email to the email address bait, you’ll know.

---

## ✨ Features

- 🛡️ **Early breach detection** for Microsoft 365 / Azure email environments
- 📨 **Deploys decoy emails** that silently trigger alerts when opened
- 🔗 **Email-based alerting**
- 👥 **Internal and external email baits**: both own-company users and external users bait.
- 📊 **Logging & Reporting** support for incident response
- 🧰 Easy integration with SIEM or SOC pipelines

---

## 🚀 Getting Started

### Prerequisites

- Azure AD / Microsoft 365
- Basic knowledge of Powershell and Azure security best practices
- Sufficient privilege to create shared mailbox (the canaray mailbox) and to add a forwarding rule to it (to some monitoring email addresses)

### Installation

##### Fast installation

> **IMPORTANT: This script only supports to be executed as a User in the Azure Portal Powershell terminal in a browser. You need to have sufficient privilege to create shared mailbox or AD User (i.e. the canaray mailbox) and to add a forwarding rule to it (to some monitoring email addresses). This is not recommended or please review the script [script.ps1](./script.ps1) before executing it.**

1. Open [Azure Portal](https://portal.azure.com/#home)
2. Open the Azure Powershell:

![Open Azure Powershell](./img/openAzurePS.jpg)

3. Execute the below command:

```
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/grennault/emailCHanary/main/entrypoint.ps1" -OutFile "entrypoint.ps1"; & "entrypoint.ps1"
```

4. Follow the instructions


##### Manual installation

1. Create a canary shared mailbox __OR__ a canary AD user:

    Create a shared mailbox:

        ```
        New-Mailbox -Shared
            -Name $SharedMailboxName
            -Alias $SharedMailboxAlias
            -PrimarySmtpAddress $SharedMailboxEmail
        ```

    Ensure GAL visibility is enabled so that the shared email address is visible for user in your organization

    __OR__

    Create an AD user:

        ```
        New-Mailbox -Shared
            -Name $SharedMailboxName
            -Alias $SharedMailboxAlias
            -PrimarySmtpAddress $SharedMailboxEmail
        ```

2. Add an Outlook Exchange licence to the user (to obtain a working email address)

2. Add forwarding rule from the previously created canaray email to a monitoring email

---

## Help us

#### Support us

<a href="https://www.buymeacoffee.com/grennault" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/default-orange.png" alt="Buy Me A Coffee" height="41" width="174"></a>

#### Community & contact

- [Github Discussion](https://github.com/grennault/emailCHanary/discussions): sharing feedbacks and asking questions.
- [Github Issue](https://github.com/grennault/emailCHanary/issues): bugs reporting and feature requests.

**Contributors**

<a href="https://github.com/grennault/emailCHanary/graphs/contributors">
  <img src="https://contrib.rocks/image?repo=grennault/emailCHanary" />
</a>

#### Star History

<a href="https://www.star-history.com/#grennault/emailCHanary&Timeline">
 <picture>
   <source media="(prefers-color-scheme: dark)" srcset="https://api.star-history.com/svg?repos=grennault/emailCHanary&type=Timeline&theme=dark" />
   <source media="(prefers-color-scheme: light)" srcset="https://api.star-history.com/svg?repos=grennault/emailCHanary&type=Timeline" />
   <img alt="Star History Chart" src="https://api.star-history.com/svg?repos=grennault/emailCHanary&type=Timeline" />
 </picture>
</a>