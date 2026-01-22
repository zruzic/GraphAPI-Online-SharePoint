# Graph/REST API – SharePoint Online  
### Practical Examples & Hands-On Guide for Devs & Admins

This repository is a **personal, practical collection of working Microsoft Graph/REST API examples for SharePoint Online**.

It is intended as a **guide and reference** for developers and administrators who struggle with:

* SharePoint Online management
* SharePoint integrations
* Microsoft Graph/REST API complexity
* Translating Microsoft documentation into working code

This is **not a framework** and **not a polished SDK**.
It is a **hands-on survival kit** built from real-world scenarios.

---

## 🎯 Who This Repository Is For

This repo is useful if you are:

* A **SharePoint admin** trying to automate tasks
* A **Power Platform / Power Automate developer** integrating with SharePoint
* A **developer** who prefers seeing raw REST calls instead of SDK abstractions
* Someone who wants **working examples**, not theory

You do **not** need to use everything in this repository.

Most users will:

* Open one file
* Copy part of the code
* Adapt it to their tenant
* Move on

That is exactly the intended usage.

---

## 🚫 What This Repository Is NOT

To avoid confusion, this repository is **not**:

* A production-ready library
* A complete error-handling framework
* A replacement for official Microsoft SDKs
* A “one-click” automation solution

Clarity and learning value are prioritized over perfection.

---

## 📂 Repository Contents (High Level)

### 🔹 Postman / Bruno Collections

* Ready-to-import API collections
* Ideal for:

  * Exploring Graph endpoints
  * Exploring SharePoint REST endpoints
  * Testing permissions
  * Understanding request/response structure

### 🔹 PowerShell Scripts

* Focused on admin and automation scenarios
* Explicit REST calls (no SDK magic)
* Useful for:

  * Tenant automation
  * Scheduled tasks
  * Troubleshooting Graph behavior

### 🔹 Python Scripts

* Cross-platform examples
* Easy to integrate into custom tools or pipelines
* Uses simple `requests`-based Graph calls

### 🔹 Documentation

* HTML documentation describing endpoints and scenarios
* Meant as a **local reference**, not official documentation

---

## 🚀 How to Start (Recommended Path)

If you are new to Graph/REST + SharePoint:

1. **Start with authentication**

   * Understand how Client Credentials flow works
2. **Use Postman first**

   * Run GET requests
   * Inspect responses
3. **Move to scripts**

   * Copy PowerShell or Python examples
   * Adjust URLs and payloads
4. **Focus on one scenario**

   * Sites
   * Lists
   * Files
   * Permissions
   * Pages

You do not need to understand the entire repository to use it.

---

## 🔐 Permissions Matter (Important)

Most scripts require **Application permissions**, not Delegated ones.

Each script typically needs permissions such as:

* `Sites.Read.All`
* `Sites.ReadWrite.All`
* `Files.ReadWrite.All`
* `Sites.Selected` (recommended for least-privilege access)

**Recommendation:**
Whenever possible, prefer using `Sites.Selected` instead of tenant-wide permissions. This allows you to grant an application access only to specific SharePoint sites, reducing security risk and aligning with the principle of least privilege.

When using `Sites.Selected`, access is granted per site with one of the following SharePoint access roles:

* **Read** – Read-only access to site content (lists, libraries, files)
* **Write** – Create, update, and delete content within the site
* **Manage** – Manage lists, libraries, and permissions within the site
* **Owner** – Full control of the site, including site settings and permissions

These roles are assigned after consent, either via SharePoint Admin Center or Microsoft Graph, and determine what the application can do on each selected site.

👉 Always verify required permissions in Azure AD / Entra ID before running scripts.

If a script fails, **permissions are the first thing to check**.

---

## ⚠ Notes & Limitations

* Some scripts are **experimental**
* Graph API behavior may vary:

  * Between tenants
  * Between classic vs modern sites
  * Depending on enabled SharePoint features
* Error handling is intentionally simple to keep examples readable

Use this repository as a **learning and reference tool**, not a drop-in production dependency.

---

## 🤝 Contributions

This is a personal collection, but:

* Improvements
* Fixes
* Additional examples
  are welcome if they align with the **practical, example-first** philosophy.

---

## 📌 Final Note

If Microsoft documentation left you thinking:

> “This should work, but it doesn’t…”

This repository exists because of that exact feeling.

Use what you need, ignore the rest, and adapt freely.


---

*Please ensure you do not commit any real secrets or credentials. Use environment variables or local config files ignored by Git.*

<img width="596" height="685" alt="GraphAPI-Online-SharePoint" src="https://github.com/user-attachments/assets/f59f6468-62dd-4a87-afe4-70d674e663cb" />
