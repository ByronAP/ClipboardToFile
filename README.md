# Clipboard To File for Windows

<p align="center">
  <img src="assets/icon.svg" alt="Application Icon" width="128">
</p>

A lightweight, high-performance, and dependency-free Windows utility that creates files in the active File Explorer window from text copied to the clipboard. Written in pure C++ with the Win32 API for maximum efficiency and minimal resource usage.

## The Problem It Solves

As a developer, you often need to create new files (`new_component.js`, `style.css`, `debug.log`, etc.) in your current working directory. The standard workflow involves right-clicking in File Explorer, selecting "New", choosing the file type, and then renaming the file.

This utility streamlines that process into a single step: **Copy & Done**.

## How It Works

1.  **Runs in the background:** The application lives as a small icon in your system tray.
2.  **Monitors the clipboard:** It efficiently listens for clipboard changes using a low-level Windows hook (not polling).
3.  **Checks the text:** If you copy a piece of text (e.g., `mynewfile.txt`), the application checks if it looks like a valid filename and if its extension is listed in a configuration file.
4.  **Finds File Explorer:** It checks to see if there is **exactly one** File Explorer window open. This is a safety measure to ensure files are never created in the wrong place.
5.  **Creates the file:** If all conditions are met, it instantly creates an empty file with that name in the directory currently open in File Explorer.

A toast notification from the tray icon confirms the file creation. If there are zero or multiple Explorer windows open, it does nothing.

## Features

-   **Extremely Lightweight:** Written in C++ using the Win32 API. No .NET, no Electron, no other runtimes. The executable is tiny and uses virtually no memory or CPU.
-   **Instantaneous:** Uses a Windows hook (`SetClipboardViewer`) and a file system watcher (`ReadDirectoryChangesW`), making it event-driven and highly responsive.
-   **User-Friendly Configuration:**
    -   **"Start with Windows"** toggle in the tray menu for convenience.
    -   **"Edit Extensions..."** menu option opens the config file in your default editor.
    -   The list of file extensions is managed in a simple `extensions.txt` file.
    -   Automatically creates `extensions.txt` on first run with some defaults.
    -   Automatically reloads settings when `extensions.txt` is modified.
-   **Safe and Informative:**
    -   Designed to only work when a single File Explorer window is active to prevent ambiguity.
    -   Provides clear toast notifications for created files, reloaded settings, or configuration errors.

## Installation & Usage

1.  Go to the [**Releases**](https://github.com/ByronAP/ClipboardToFile/releases) page.
2.  Download the latest `ClipboardToFile.exe` file.
3.  Place the executable in any folder you like and run it.

On the first run, the application will create an `extensions.txt` file in the same directory. You can edit this file to add or remove the file extensions you want the application to monitor (one per line).

For convenience, right-click the tray icon and select **"Start with Windows"** to have the application launch automatically when you log in.

## Building from Source

This project was built using Visual Studio 2022 and the Windows 11 SDK.

### Prerequisites
-   **Visual Studio 2022** (or 2019) with the **"Desktop development with C++"** workload installed.
-   Windows 10 or 11 SDK.

### Steps
1.  Clone the repository:
    ```bash
    git clone https://github.com/ByronAP/ClipboardToFile.git
    ```
2.  Open the `ClipboardToFile.sln` file in Visual Studio.
3.  Set the build configuration to **Release** and the platform to **x64**.
4.  From the menu, select **Build > Build Solution**.

The final executable will be located in the `x64/Release` folder.

## Contributing

This project was built for a specific purpose, but suggestions and improvements are welcome. Feel free to open an issue to discuss a potential feature or submit a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.