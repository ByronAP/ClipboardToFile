# Clipboard To File for Windows

<p align="center">
  <img src="assets/icon.svg" alt="Application Icon" width="128">
</p>

A lightweight, high-performance, and dependency-free Windows utility that intelligently creates files and their content in your active File Explorer window, directly from your clipboard.

Written in pure C++ with the Win32 API for maximum efficiency and minimal resource usage.

<p align="center">
  <!-- License Badge - Static -->
  <a href="https://opensource.org/licenses/MIT">
    <img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License: MIT">
  </a>
  <!-- CI Pipeline Status Badge - Dynamic -->
  <a href="https://github.com/ByronAP/ClipboardToFile/actions/workflows/ci-build.yml">
    <img src="https://github.com/ByronAP/ClipboardToFile/actions/workflows/ci-build.yml/badge.svg?branch=dev" alt="CI Build Status">
  </a>
  <!-- Latest Release Version Badge - Dynamic -->
  <a href="https://github.com/ByronAP/ClipboardToFile/releases/latest">
    <img src="https://img.shields.io/github/v/release/ByronAP/ClipboardToFile" alt="Latest Release">
  </a>
</p>

## The Problem It Solves

As a developer, you often need to create new files (`new_component.js`, `style.css`, `debug.log`, etc.) in your current working directory. You might also have code snippets or entire file contents on your clipboard that you want to save quickly.

This utility streamlines both workflows into a single action: **Copy & Done**.

## How It Works

The application runs silently in your system tray and monitors the clipboard for text that looks like a file you want to create. It offers two powerful, independently-togglable features:

### 1. Create Empty File
If you copy a simple filename like `new_style.css`, the app will instantly create that empty file in your active File Explorer window.

### 2. Create File with Content
If you copy a block of text where the **first line is a filename** (like `my_script.js`) and the rest is the content, the app will create the file and populate it with that content, all in one go. This is perfect for pasting code snippets, logs, or any text artifact.

The app uses a configurable, intelligent system (including regex) to detect the filename on the first line. In all cases, it will only act if there is **exactly one** File Explorer window open, as a safety measure to ensure files are never created in the wrong place.

## Features

-   **Dual Functionality:** Independently enable/disable "Create Empty File" and "Create File with Content" from the tray menu.
-   **Powerful Content Creation:** Uses a list of configurable regular expressions to intelligently detect filenames, even in complex formats like `// --- START OF FILE: my_app.cpp ---`.
-   **Centralized JSON Configuration:** All settings are managed in a single, easy-to-edit `config.json` file.
-   **Extremely Lightweight:** Written in C++ with the native Win32 API. No .NET, no Electron, no other runtimes. The executable is tiny and uses virtually no memory.
-   **Fully Automated & User-Friendly:**
    -   **"Start with Windows"** toggle in the tray menu for convenience.
    -   **"Edit Config..."** menu option opens `config.json` in your default editor.
    -   Automatically checks for new versions once every 24 hours.
    -   Automatically creates `config.json` on first run with sensible defaults.
    -   Live-reloads settings when `config.json` is modified.
-   **Safe and Informative:**
    -   Only works when a single File Explorer window is active to prevent ambiguity.
    -   Provides distinct toast notifications for empty files vs. files with content.

## Installation & Usage

1.  Go to the [**Releases**](https://github.com/ByronAP/ClipboardToFile/releases) page.
2.  Download the latest `ClipboardToFile-Setup.exe` file and run it.
3.  Alternatively, download the `ClipboardToFile-Portable.zip` for a version that requires no installation.

After running the application for the first time, a `config.json` file will be created in your user's AppData directory (`%APPDATA%\ClipboardToFile\`). You can customize all application behavior by editing this file via the **"Edit Config..."** option in the tray menu.

For convenience, right-click the tray icon and select **"Start with Windows"** to have the application launch automatically when you log in.

## Building from Source

This project uses Git Submodules for its dependencies.

### Prerequisites
-   **Visual Studio 2022** (or 2019) with the **"Desktop development with C++"** workload installed.
-   Windows 10 or 11 SDK.
-   Git installed.

### Steps
1.  Clone the repository and its submodules:
    ```bash
    git clone --recurse-submodules https://github.com/ByronAP/ClipboardToFile.git
    ```
2.  Open the `ClipboardToFile.sln` file in Visual Studio.
3.  Set the build configuration to **Release** and the platform to **x64**.
4.  From the menu, select **Build > Build Solution**.

The final executable will be located in the `x64/Release` folder.

## Contributing

This project was built for a specific purpose, but suggestions and improvements are welcome. Feel free to open an issue to discuss a potential feature or submit a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.