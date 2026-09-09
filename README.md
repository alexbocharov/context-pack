# 📦 context-pack

![GitHub License](https://img.shields.io/github/license/alexbocharov/context-pack)

**`context-pack`** is a universal, ultra-lightweight PowerShell CLI tool designed to instantly pack an entire source code repository into a single, beautifully structured file (**Microsoft Word `.docx`** or **Plain Text `.txt`**).

By flattening complex directory trees while preserving absolute structural context, it generates the **ultimate context injection file** for Large Language Models (ChatGPT, Claude, local LLMs), technical compliance audits, and offline architectural reviews.

Out-of-the-box optimization for **.NET Aspire**, standard .NET solutions, modern frontend frameworks, and large monorepos.

## ⚡ Key Features

* 🤖 **Perfect for LLMs:** Stop dragging 50 separate files into Claude or ChatGPT. Feed it a single `context-pack` file, and the AI will instantly grasp your whole project architecture.
* 🎛️ **Dual-Format Output:** Choose between rich Microsoft Word layouts (`.docx`) for corporate reporting or lightweight `.txt` files for instant clipboard copy-pasting.
* 📂 **Path-Prepended Context:** Every single code block is automatically labeled with its exact relative repository path (e.g., `📄 File: src/Kiparis.AppHost/Program.cs`).
* 🛑 **Smart Dependency Stripping:** Automatically bypasses bloated build artifacts (`bin`, `obj`, `target`), heavy package folders (`node_modules`, `vendor`), and IDE caches.
* 🛠️ **Zero External Dependencies:** Pure PowerShell script. No heavy npm modules, pip installs, or binary setup required.
* 🌍 **Cross-Platform Friendly:** The `-Format txt` mode works anywhere PowerShell Core (`pwsh`) runs (Windows, macOS, Linux).

## 📋 System Requirements

* **For `.txt` output:** PowerShell Core (`pwsh`) or classic Windows PowerShell on any OS (Windows, Linux, macOS).
* **For `.docx` output:** Windows 10/11 with Microsoft Word installed (utilizes the local COM automation engine).

## 🚀 Quick Start

### 1. Download
Grab the `pack.ps1` script from this repository.

### 2. Drop into Root
Move `pack.ps1` into the **root directory** of the repository you want to pack (where your main `.sln`, `package.json`, or project root folder lives).

### 3. Run the Pack
Open your terminal in that folder. If Windows blocks script execution, bypass it for the current session and run the tool:

#### Option A: Pack into Formatted Word Document (Default)
Generates a highly-structured Microsoft Word document with clean typography, page breaks, and proper Consolas code fonts:
```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
.\pack.ps1 -Format docx
```

#### Option B: Pack into Lightweight Text File
Fast, native text file with zero external dependencies. Ideal for headless environments or immediate prompt feeding:
```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
.\pack.ps1 -Format txt
```

## ⚙️ Customization

`context-pack` adapts to any tech stack instantly. Open `pack.ps1` in any code editor to modify the top configuration arrays:

* **Include more file types:** Add custom extensions to the `$AllowedExtensions` array.
* **Ignore extra folders:** Add specific directory names to the `$ExcludeFolders` array.

## 📄 License

This project is licensed under the **MIT License** — use it, modify it, share it, or integrate it into your workflows freely.

## 🤝 Contributing & Support

Got ideas to make `context-pack` even faster? Want to add native Markdown (`.md`) syntax highlighting support or automated `.gitignore` parsing? Contributions, issues, and feature requests are welcome!

Give a ⭐️ if this project helped you pack your context!