//================================================================================================//
//                                     Clipboard To File                                          //
//                                                                                                //
//  A lightweight, dependency-free C++ utility that creates files in the active File Explorer     //
//  window from text copied to the clipboard.                                                     //
//================================================================================================//


//------------------------------------------------------------------------------------------------//
//                                         INCLUDES                                               //
//------------------------------------------------------------------------------------------------//
#include <windows.h>
#include <shlobj.h>     // For SHGetFolderPathW
#include <shlwapi.h>    // For Path... functions
#include <wininet.h>    // For HTTP requests to check for updates
#include <string>
#include <vector>
#include <algorithm>
#include <fstream>
#include <mutex>
#include <sstream>      // For wstringstream
#include "resource.h"


//------------------------------------------------------------------------------------------------//
//                                    LINKER DIRECTIVES                                           //
//------------------------------------------------------------------------------------------------//
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "wininet.lib")
#pragma comment(lib, "version.lib")


//------------------------------------------------------------------------------------------------//
//                              GLOBAL STATE & CONSTANTS                                          //
//------------------------------------------------------------------------------------------------//
#define WM_TRAY_ICON_MSG            (WM_USER + 1)   // Message for tray icon events
#define WM_APP_RELOAD_EXTENSIONS    (WM_USER + 2)   // Message from watcher thread to trigger reload
#define WM_APP_UPDATE_FOUND         (WM_USER + 3)   // Message for application updates
#define ID_TRAY_ICON                1
#define ID_MENU_TOGGLE              1001
#define ID_MENU_EDIT_EXTENSIONS     1002
#define ID_MENU_START_WITH_WINDOWS  1003
#define ID_MENU_EXIT                1004

const wchar_t CLASS_NAME[] = L"ClipboardToFileWindowClass";
HWND  g_hMainWnd = NULL;
HWND  g_hNextClipboardViewer = NULL;
bool  g_isEnabled = true;
std::vector<std::wstring> g_allowedExtensions;
std::mutex g_extensionsMutex;
HANDLE g_hWatcherThread = NULL;
HANDLE g_hShutdownEvent = NULL;


//------------------------------------------------------------------------------------------------//
//                                  FUNCTION PROTOTYPES                                           //
//------------------------------------------------------------------------------------------------//
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
void CreateTrayIcon(HWND);
void RemoveTrayIcon(HWND);
void ShowContextMenu(HWND);
void ShowToastNotification(HWND, const std::wstring&, const std::wstring&, DWORD);
std::wstring GetSettingsFilePath();
std::wstring GetSingleExplorerPath();
void ProcessClipboardChange();
DWORD WINAPI FileWatcherThread(LPVOID);
void InitializeAndLoadExtensions();
bool ReloadExtensions(bool);
bool IsStartupEnabled();
void SetStartup(bool);
void CheckForUpdatesIfNeeded();


//------------------------------------------------------------------------------------------------//
//                                  APPLICATION ENTRY POINT                                       //
//------------------------------------------------------------------------------------------------//
// Standard Windows application entry point. Creates a hidden window to handle messages.
int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE, _In_ LPWSTR, _In_ int)
{
    WNDCLASS wc = {};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_CLIPBOARDTOFILE));
    RegisterClass(&wc);

    g_hMainWnd = CreateWindowEx(0, CLASS_NAME, L"Clipboard To File Helper", WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, hInstance, NULL);

    if (g_hMainWnd == NULL) return 0;

    MSG msg = {};
    while (GetMessage(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return (int)msg.wParam;
}


//------------------------------------------------------------------------------------------------//
//                                  WINDOW PROCEDURE (WndProc)                                    //
//------------------------------------------------------------------------------------------------//
// Central message handler for the application's hidden window.
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
    case WM_CREATE:
        // This message is sent once when the window is first created.
        InitializeAndLoadExtensions();
        g_hNextClipboardViewer = SetClipboardViewer(hwnd);
        CreateTrayIcon(hwnd);
        g_hShutdownEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
        g_hWatcherThread = CreateThread(NULL, 0, FileWatcherThread, NULL, 0, NULL);
        CheckForUpdatesIfNeeded();
        break;
    case WM_DESTROY:
        // Performs cleanup in reverse order of creation to ensure safe shutdown.
        if (g_hShutdownEvent) SetEvent(g_hShutdownEvent);
        if (g_hWatcherThread) {
            WaitForSingleObject(g_hWatcherThread, 2000);
            CloseHandle(g_hWatcherThread);
        }
        if (g_hShutdownEvent) CloseHandle(g_hShutdownEvent);
        ChangeClipboardChain(hwnd, g_hNextClipboardViewer);
        RemoveTrayIcon(hwnd);
        PostQuitMessage(0);
        break;
    case WM_CHANGECBCHAIN:
        // A window is being removed from the clipboard viewer chain.
        if ((HWND)wParam == g_hNextClipboardViewer) g_hNextClipboardViewer = (HWND)lParam;
        else if (g_hNextClipboardViewer != NULL) SendMessage(g_hNextClipboardViewer, msg, wParam, lParam);
        break;
    case WM_DRAWCLIPBOARD:
        // The content of the clipboard has changed.
        ProcessClipboardChange();
        // CRITICAL: We must pass this message on to the next window in the chain.
        SendMessage(g_hNextClipboardViewer, msg, wParam, lParam);
        break;
    case WM_APP_RELOAD_EXTENSIONS:
        // Handles the reload request from our file watcher thread.
        Sleep(100); // Small delay to prevent race conditions with text editors.
        ReloadExtensions(false);
        break;
    case WM_APP_UPDATE_FOUND: {
        wchar_t* releaseUrl = (wchar_t*)lParam;
        if (releaseUrl) {
            std::wstring message = L"A new version is available!\n\nWould you like to open the download page?";
            if (MessageBoxW(hwnd, message.c_str(), L"Update Available", MB_YESNO | MB_ICONINFORMATION) == IDYES) {
                ShellExecuteW(NULL, L"open", releaseUrl, NULL, NULL, SW_SHOWNORMAL);
            }
            delete[] releaseUrl; // Free the memory allocated by the worker thread.
        }
        break;
    }
    case WM_TRAY_ICON_MSG:
        if (lParam == WM_RBUTTONUP) ShowContextMenu(hwnd);
        break;
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case ID_MENU_TOGGLE:
            g_isEnabled = !g_isEnabled;
            break;
        case ID_MENU_START_WITH_WINDOWS:
            SetStartup(!IsStartupEnabled());
            break;
        case ID_MENU_EDIT_EXTENSIONS: {
            std::wstring settingsPath = GetSettingsFilePath();
            // Use ShellExecute to open the file with its default registered application.
            ShellExecuteW(NULL, L"open", settingsPath.c_str(), NULL, NULL, SW_SHOWNORMAL);
            break;
        }
        case ID_MENU_EXIT:
            DestroyWindow(hwnd);
            break;
        }
        break;
    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}


//------------------------------------------------------------------------------------------------//
//                                     UPDATE CHECKER                                             //
//------------------------------------------------------------------------------------------------//
struct AppVersion { int major = 0, minor = 0, patch = 0, build = 0; };

AppVersion ParseVersionString(const std::wstring& versionStr) {
    AppVersion v;
    std::wstring s = (versionStr[0] == L'v') ? versionStr.substr(1) : versionStr;
    std::wstringstream ss(s);
    wchar_t dot;
    ss >> v.major >> dot >> v.minor >> dot >> v.patch >> dot >> v.build;
    return v;
}

AppVersion GetCurrentAppVersion() {
    wchar_t exePath[MAX_PATH];
    GetModuleFileNameW(NULL, exePath, MAX_PATH);
    DWORD handle = 0;
    DWORD versionSize = GetFileVersionInfoSizeW(exePath, &handle);
    if (versionSize == 0) return {};
    std::vector<BYTE> versionData(versionSize);
    if (!GetFileVersionInfoW(exePath, handle, versionSize, versionData.data())) return {};
    VS_FIXEDFILEINFO* pFileInfo = nullptr;
    UINT fileInfoSize = 0;
    if (VerQueryValueW(versionData.data(), L"\\", (LPVOID*)&pFileInfo, &fileInfoSize) && pFileInfo) {
        AppVersion v;
        v.major = HIWORD(pFileInfo->dwFileVersionMS);
        v.minor = LOWORD(pFileInfo->dwFileVersionMS);
        v.patch = HIWORD(pFileInfo->dwFileVersionLS);
        v.build = LOWORD(pFileInfo->dwFileVersionLS);
        return v;
    }
    return {};
}

const wchar_t* REG_APP_KEY = L"Software\\ByronAP\\ClipboardToFile";

// Background thread function that performs the network request to the GitHub API.
DWORD WINAPI PerformUpdateCheck(LPVOID) {
    HINTERNET hInternet = InternetOpenW(L"ClipboardToFile/1.0", INTERNET_OPEN_TYPE_DIRECT, NULL, NULL, 0);
    if (hInternet) {
        HINTERNET hConnect = InternetOpenUrlW(hInternet,
            L"https://api.github.com/repos/ByronAP/ClipboardToFile/releases/latest",
            L"User-Agent: ClipboardToFile-Update-Check\r\n", -1L,
            INTERNET_FLAG_SECURE | INTERNET_FLAG_RELOAD | INTERNET_FLAG_NO_CACHE_WRITE, 0);
        if (hConnect) {
            std::string response;
            char buffer[4096];
            DWORD bytesRead;
            while (InternetReadFile(hConnect, buffer, sizeof(buffer), &bytesRead) && bytesRead > 0) {
                response.append(buffer, bytesRead);
            }
            InternetCloseHandle(hConnect);
            size_t tagPos = response.find("\"tag_name\":");
            size_t urlPos = response.find("\"html_url\":");
            if (tagPos != std::string::npos && urlPos != std::string::npos) {
                size_t tagStart = response.find("\"", tagPos + 11) + 1;
                size_t tagEnd = response.find("\"", tagStart);
                std::string latestVersionStr = response.substr(tagStart, tagEnd - tagStart);
                size_t urlStart = response.find("\"", urlPos + 11) + 1;
                size_t urlEnd = response.find("\"", urlStart);
                std::string releaseUrlStr = response.substr(urlStart, urlEnd - urlStart);
                std::wstring latestVersionW(latestVersionStr.begin(), latestVersionStr.end());
                std::wstring releaseUrlW(releaseUrlStr.begin(), releaseUrlStr.end());
                AppVersion currentV = GetCurrentAppVersion();
                AppVersion latestV = ParseVersionString(latestVersionW);
                if (latestV.major > currentV.major ||
                    (latestV.major == currentV.major && latestV.minor > currentV.minor) ||
                    (latestV.major == currentV.major && latestV.minor == currentV.minor && latestV.patch > currentV.patch) ||
                    (latestV.major == currentV.major && latestV.minor == currentV.minor && latestV.patch == currentV.patch && latestV.build > currentV.build)) {
                    wchar_t* url_heap = new wchar_t[releaseUrlW.length() + 1];
                    wcscpy_s(url_heap, releaseUrlW.length() + 1, releaseUrlW.c_str());
                    PostMessage(g_hMainWnd, WM_APP_UPDATE_FOUND, 0, (LPARAM)url_heap);
                }
            }
        }
        InternetCloseHandle(hInternet);
    }
    // After any check, update the timestamp in the registry.
    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, REG_APP_KEY, 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS) {
        SYSTEMTIME st; GetSystemTime(&st);
        FILETIME now; SystemTimeToFileTime(&st, &now);
        RegSetValueExW(hKey, L"LastUpdateCheck", 0, REG_QWORD, (const BYTE*)&now, sizeof(now));
        RegCloseKey(hKey);
    }
    return 0;
}

// Checks the registry to see if 24 hours have passed since the last check.
void CheckForUpdatesIfNeeded() {
    HKEY hKey;
    FILETIME lastCheck = {};
    DWORD dataSize = sizeof(lastCheck);
    if (RegOpenKeyExW(HKEY_CURRENT_USER, REG_APP_KEY, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        RegQueryValueExW(hKey, L"LastUpdateCheck", NULL, NULL, (LPBYTE)&lastCheck, &dataSize);
        RegCloseKey(hKey);
    }
    FILETIME now;
    SYSTEMTIME st; GetSystemTime(&st);
    SystemTimeToFileTime(&st, &now);
    ULONGLONG lastCheckInt = ((ULONGLONG)lastCheck.dwHighDateTime << 32) + lastCheck.dwLowDateTime;
    ULONGLONG nowInt = ((ULONGLONG)now.dwHighDateTime << 32) + now.dwLowDateTime;
    ULONGLONG twentyFourHours = 24ULL * 60 * 60 * 1000 * 10000;
    if ((nowInt - lastCheckInt) > twentyFourHours) {
        HANDLE hThread = CreateThread(NULL, 0, PerformUpdateCheck, NULL, 0, NULL);
        if (hThread) CloseHandle(hThread); // Fire-and-forget the thread.
    }
}


//------------------------------------------------------------------------------------------------//
//                                  FILE WATCHER WORKER THREAD                                    //
//------------------------------------------------------------------------------------------------//
// Monitors the settings directory using asynchronous I/O and notifies the main thread of changes.
DWORD WINAPI FileWatcherThread(LPVOID)
{
    wchar_t appDataPath[MAX_PATH];
    if (FAILED(SHGetFolderPathW(NULL, CSIDL_APPDATA, NULL, 0, appDataPath))) return 1;

    std::wstring dirPath = std::wstring(appDataPath) + L"\\ClipboardToFile";
    CreateDirectoryW(dirPath.c_str(), NULL);

    HANDLE hDir = CreateFileW(dirPath.c_str(), FILE_LIST_DIRECTORY,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OVERLAPPED, NULL);
    if (hDir == INVALID_HANDLE_VALUE) return 1;

    BYTE buffer[1024];
    DWORD bytesReturned;
    OVERLAPPED overlapped = { 0 };
    overlapped.hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
    if (overlapped.hEvent == NULL) {
        CloseHandle(hDir);
        return 1;
    }

    HANDLE waitHandles[2] = { g_hShutdownEvent, overlapped.hEvent };

    while (true) {
        // Start the asynchronous directory change notification.
        ReadDirectoryChangesW(hDir, buffer, sizeof(buffer), FALSE,
            FILE_NOTIFY_CHANGE_LAST_WRITE | FILE_NOTIFY_CHANGE_FILE_NAME,
            &bytesReturned, &overlapped, NULL);

        // Wait efficiently until either the shutdown event is signaled or a file change occurs.
        DWORD waitStatus = WaitForMultipleObjects(2, waitHandles, FALSE, INFINITE);
        if (waitStatus == WAIT_OBJECT_0) break; // Shutdown event.

        if (waitStatus == WAIT_OBJECT_0 + 1) { // File change event.
            if (GetOverlappedResult(hDir, &overlapped, &bytesReturned, FALSE) && bytesReturned > 0) {
                FILE_NOTIFY_INFORMATION* pNotify = (FILE_NOTIFY_INFORMATION*)buffer;
                while (pNotify) {
                    std::wstring filename(pNotify->FileName, pNotify->FileNameLength / sizeof(wchar_t));
                    if (_wcsicmp(filename.c_str(), L"extensions.txt") == 0) {
                        // Post a message to the main UI thread to handle the reload safely.
                        PostMessage(g_hMainWnd, WM_APP_RELOAD_EXTENSIONS, 0, 0);
                        break;
                    }
                    pNotify = pNotify->NextEntryOffset > 0 ? (FILE_NOTIFY_INFORMATION*)((BYTE*)pNotify + pNotify->NextEntryOffset) : NULL;
                }
            }
            ResetEvent(overlapped.hEvent);
        }
        else {
            break; // An error occurred.
        }
    }

    CancelIo(hDir);
    CloseHandle(hDir);
    CloseHandle(overlapped.hEvent);
    return 0;
}


//------------------------------------------------------------------------------------------------//
//                          CORE LOGIC & FILE MANAGEMENT                                          //
//------------------------------------------------------------------------------------------------//
// Gets path to extensions.txt in %APPDATA%\ClipboardToFile, the correct location for user config.
std::wstring GetSettingsFilePath()
{
    wchar_t appDataPath[MAX_PATH];
    if (SUCCEEDED(SHGetFolderPathW(NULL, CSIDL_APPDATA, NULL, 0, appDataPath))) {
        std::wstring fullPath = std::wstring(appDataPath) + L"\\ClipboardToFile";
        CreateDirectoryW(fullPath.c_str(), NULL);
        return fullPath + L"\\extensions.txt";
    }
    return L"extensions.txt"; // Fallback to local directory.
}

// Creates a default config file on first run if needed, then performs initial load.
void InitializeAndLoadExtensions()
{
    std::wstring settingsPath = GetSettingsFilePath();
    std::wifstream checkFile(settingsPath);
    if (!checkFile.is_open()) {
        const std::vector<std::wstring> defaultExtensions = { L".txt", L".md", L".log", L".sql", L".cpp", L".h", L".js", L".json", L".xml", L".cs", L".c"};
        std::wofstream newSettingsFile(settingsPath);
        if (newSettingsFile.is_open()) {
            for (const auto& ext : defaultExtensions) newSettingsFile << ext << std::endl;
            newSettingsFile.close();
        }
    }
    else {
        checkFile.close();
    }
    ReloadExtensions(true);
}

// Reads extensions.txt and loads them into the global vector. Returns true on success.
bool ReloadExtensions(bool isInitialLoad)
{
    std::wstring settingsPath = GetSettingsFilePath();
    std::vector<std::wstring> tempExtensions;
    bool hadBadLines = false;
    std::wifstream settingsFile(settingsPath);
    if (!settingsFile.is_open()) {
        if (!isInitialLoad) ShowToastNotification(g_hMainWnd, L"Reload Failed", L"Could not open extensions.txt.", NIIF_ERROR);
        return false;
    }

    std::wstring line;
    while (std::getline(settingsFile, line)) {
        // Sanitize each line from the file.
        line.erase(0, line.find_first_not_of(L" \t\n\r"));
        line.erase(line.find_last_not_of(L" \t\n\r") + 1);
        if (line.empty() || line[0] == L'#' || (line[0] == L'/' && line[1] == L'/')) continue;
        if (wcspbrk(line.c_str(), L"\\/:*?\"<>|")) { hadBadLines = true; continue; }
        if (line[0] != L'.') line.insert(0, 1, L'.');
        std::transform(line.begin(), line.end(), line.begin(), ::towlower);
        tempExtensions.push_back(line);
    }
    settingsFile.close();

    // Atomically swap the new list for the old one for thread safety.
    {
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        g_allowedExtensions.swap(tempExtensions);
    }

    if (!isInitialLoad) {
        if (hadBadLines) ShowToastNotification(g_hMainWnd, L"Extensions Reloaded", L"Settings updated. Some invalid lines were skipped.", NIIF_WARNING);
        else ShowToastNotification(g_hMainWnd, L"Extensions Reloaded", L"Settings have been updated from extensions.txt.", NIIF_INFO);
    }
    return true;
}

// Called on every clipboard change to check for a valid filename and create the file.
void ProcessClipboardChange()
{
    if (!g_isEnabled || !IsClipboardFormatAvailable(CF_UNICODETEXT) || !OpenClipboard(g_hMainWnd)) return;

    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (hData == NULL) { CloseClipboard(); return; }

    wchar_t* pszText = static_cast<wchar_t*>(GlobalLock(hData));
    if (pszText == NULL) { CloseClipboard(); return; }

    std::wstring clipboardText(pszText);
    GlobalUnlock(hData);
    CloseClipboard();

    // Sanitize and validate the clipboard text.
    clipboardText.erase(0, clipboardText.find_first_not_of(L" \t\n\r"));
    clipboardText.erase(clipboardText.find_last_not_of(L" \t\n\r") + 1);
    if (clipboardText.empty() || wcspbrk(clipboardText.c_str(), L"\\/:*?\"<>|") != nullptr) return;

    wchar_t ext[_MAX_EXT];
    _wsplitpath_s(clipboardText.c_str(), NULL, 0, NULL, 0, NULL, 0, ext, _MAX_EXT);
    std::wstring extension(ext);
    std::transform(extension.begin(), extension.end(), extension.begin(), ::towlower);

    // Thread-safe read of the allowed extensions list.
    bool isAllowed = false;
    {
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        for (const auto& allowedExt : g_allowedExtensions) {
            if (extension == allowedExt) { isAllowed = true; break; }
        }
    }

    if (!isAllowed) return;

    // Safety check: only proceed if exactly one File Explorer window is open.
    std::wstring explorerPath = GetSingleExplorerPath();
    if (!explorerPath.empty()) {
        std::wstring fullPath = explorerPath + L"\\" + clipboardText;
        if (GetFileAttributesW(fullPath.c_str()) == INVALID_FILE_ATTRIBUTES) {
            try {
                HANDLE hFile = CreateFileW(fullPath.c_str(), GENERIC_WRITE, 0, NULL, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
                if (hFile != INVALID_HANDLE_VALUE) {
                    CloseHandle(hFile);
                    std::wstring successMessage = L"Created file: " + clipboardText;
                    ShowToastNotification(g_hMainWnd, L"File Created", successMessage, NIIF_INFO);
                }
            }
            catch (...) {}
        }
    }
}

// Uses COM to find and return the path of a single open File Explorer window.
std::wstring GetSingleExplorerPath()
{
    std::vector<std::wstring> paths;
    IShellWindows* pShellWindows = NULL;

    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) return L"";

    hr = CoCreateInstance(CLSID_ShellWindows, NULL, CLSCTX_ALL, IID_IShellWindows, (void**)&pShellWindows);
    if (SUCCEEDED(hr)) {
        long count;
        pShellWindows->get_Count(&count);
        for (long i = 0; i < count; i++) {
            VARIANT v; VariantInit(&v); v.vt = VT_I4; v.lVal = i;
            IDispatch* pDispatch;
            if (SUCCEEDED(pShellWindows->Item(v, &pDispatch))) {
                IWebBrowser2* pBrowser;
                if (SUCCEEDED(pDispatch->QueryInterface(IID_IWebBrowser2, (void**)&pBrowser))) {
                    HWND hwndBrowser;
                    if (SUCCEEDED(pBrowser->get_HWND((SHANDLE_PTR*)&hwndBrowser))) {
                        wchar_t className[256];
                        if (GetClassNameW(hwndBrowser, className, 256) && wcscmp(className, L"CabinetWClass") == 0) {
                            BSTR bstrURL;
                            if (SUCCEEDED(pBrowser->get_LocationURL(&bstrURL))) {
                                std::wstring url(bstrURL, SysStringLen(bstrURL));
                                SysFreeString(bstrURL);
                                wchar_t localPath[MAX_PATH];
                                DWORD pathLen = MAX_PATH;
                                if (SUCCEEDED(PathCreateFromUrlW(url.c_str(), localPath, &pathLen, 0))) paths.push_back(localPath);
                            }
                        }
                    }
                    pBrowser->Release();
                }
                pDispatch->Release();
            }
            VariantClear(&v);
        }
        pShellWindows->Release();
    }
    CoUninitialize();

    if (paths.size() == 1) return paths[0];
    return L"";
}


//------------------------------------------------------------------------------------------------//
//                                  TRAY ICON & UI MANAGEMENT                                     //
//------------------------------------------------------------------------------------------------//

// Adds the icon to the system tray.
void CreateTrayIcon(HWND hwnd)
{
    NOTIFYICONDATAW nid = {};
    nid.cbSize = sizeof(NOTIFYICONDATAW);
    nid.hWnd = hwnd;
    nid.uID = ID_TRAY_ICON;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.uCallbackMessage = WM_TRAY_ICON_MSG;
    nid.hIcon = (HICON)LoadImage(GetModuleHandle(NULL), MAKEINTRESOURCE(IDI_SMALL), IMAGE_ICON, 16, 16, 0);
    wcscpy_s(nid.szTip, L"Clipboard To File");
    Shell_NotifyIconW(NIM_ADD, &nid);
}

// Removes the icon from the system tray.
void RemoveTrayIcon(HWND hwnd)
{
    NOTIFYICONDATAW nid = {};
    nid.cbSize = sizeof(NOTIFYICONDATAW);
    nid.hWnd = hwnd;
    nid.uID = ID_TRAY_ICON;
    Shell_NotifyIconW(NIM_DELETE, &nid);
}

// Creates and displays the right-click context menu for the tray icon.
void ShowContextMenu(HWND hwnd)
{
    POINT pt; GetCursorPos(&pt);
    HMENU hMenu = CreatePopupMenu();
    if (hMenu) {
        UINT enabledFlags = g_isEnabled ? MF_STRING | MF_CHECKED : MF_STRING | MF_UNCHECKED;
        InsertMenu(hMenu, 0, MF_BYPOSITION | enabledFlags, ID_MENU_TOGGLE, L"Enabled");
        UINT startupFlags = IsStartupEnabled() ? MF_STRING | MF_CHECKED : MF_STRING | MF_UNCHECKED;
        InsertMenu(hMenu, 1, MF_BYPOSITION | startupFlags, ID_MENU_START_WITH_WINDOWS, L"Start with Windows");
        InsertMenu(hMenu, 2, MF_BYPOSITION | MF_STRING, ID_MENU_EDIT_EXTENSIONS, L"Edit Extensions...");
        InsertMenu(hMenu, 3, MF_SEPARATOR, 0, NULL);
        InsertMenu(hMenu, 4, MF_BYPOSITION | MF_STRING, ID_MENU_EXIT, L"Exit");

        SetForegroundWindow(hwnd);
        TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hwnd, NULL);
        DestroyMenu(hMenu);
    }
}

// Displays a toast notification from the tray icon.
void ShowToastNotification(HWND hwnd, const std::wstring& title, const std::wstring& msg, DWORD iconType)
{
    NOTIFYICONDATAW nid = {};
    nid.cbSize = sizeof(NOTIFYICONDATAW);
    nid.hWnd = hwnd;
    nid.uID = ID_TRAY_ICON;
    nid.uFlags = NIF_INFO;
    nid.dwInfoFlags = iconType;
    wcscpy_s(nid.szInfoTitle, title.c_str());
    wcscpy_s(nid.szInfo, msg.c_str());
    Shell_NotifyIconW(NIM_MODIFY, &nid);
}


//------------------------------------------------------------------------------------------------//
//                                      REGISTRY HELPERS                                          //
//------------------------------------------------------------------------------------------------//
const wchar_t* REG_RUN_KEY = L"Software\\Microsoft\\Windows\\CurrentVersion\\Run";
const wchar_t* REG_VALUE_NAME = L"ClipboardToFile";

// Checks the registry to see if the application is configured to run at startup.
bool IsStartupEnabled()
{
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, REG_RUN_KEY, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        LSTATUS status = RegQueryValueExW(hKey, REG_VALUE_NAME, NULL, NULL, NULL, NULL);
        RegCloseKey(hKey);
        return status == ERROR_SUCCESS;
    }
    return false;
}

// Adds or removes the application from the Windows startup registry key.
void SetStartup(bool enable)
{
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, REG_RUN_KEY, 0, KEY_WRITE, &hKey) == ERROR_SUCCESS) {
        if (enable) {
            wchar_t exePath[MAX_PATH];
            GetModuleFileNameW(NULL, exePath, MAX_PATH);
            RegSetValueExW(hKey, REG_VALUE_NAME, 0, REG_SZ, (const BYTE*)exePath, (DWORD)((wcslen(exePath) + 1) * sizeof(wchar_t)));
        }
        else {
            RegDeleteValueW(hKey, REG_VALUE_NAME);
        }
        RegCloseKey(hKey);
    }
}