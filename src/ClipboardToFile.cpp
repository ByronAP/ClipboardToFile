// ClipboardToFile.cpp
//
// A lightweight, dependency-free C++ utility to create files from clipboard text.
// - Runs in the system tray.
// - Uses a Win32 hook (Clipboard Viewer) for efficiency.
// - Uses a dedicated worker thread with ReadDirectoryChangesW to monitor extensions.txt.
// - Uses COM to find the active File Explorer path.
//

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "shlwapi.lib")

#include <windows.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <string>
#include <vector>
#include <algorithm>
#include <fstream>
#include <mutex>
#include "resource.h"

// --- Global State & Constants ---
#define WM_TRAY_ICON_MSG (WM_USER + 1)
#define WM_APP_RELOAD_EXTENSIONS (WM_USER + 2)
#define ID_TRAY_ICON 1
#define ID_MENU_TOGGLE 1001
#define ID_MENU_EDIT_EXTENSIONS 1002
#define ID_MENU_START_WITH_WINDOWS 1003
#define ID_MENU_EXIT   1004

const wchar_t CLASS_NAME[] = L"ClipboardToFileWindowClass";
HWND  g_hMainWnd = NULL;
HWND  g_hNextClipboardViewer = NULL;
bool  g_isEnabled = true;

std::vector<std::wstring> g_allowedExtensions;
std::mutex g_extensionsMutex;

HANDLE g_hWatcherThread = NULL;
HANDLE g_hShutdownEvent = NULL;


// --- Function Prototypes ---
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
void CreateTrayIcon(HWND hwnd);
void RemoveTrayIcon(HWND hwnd);
void ShowContextMenu(HWND hwnd);
void ShowToastNotification(HWND hwnd, const std::wstring& title, const std::wstring& msg, DWORD iconType);
std::wstring GetSingleExplorerPath();
void ProcessClipboardChange();
DWORD WINAPI FileWatcherThread(LPVOID lpParam);
void InitializeAndLoadExtensions();
bool ReloadExtensions(bool isInitialLoad);
bool IsStartupEnabled();
void SetStartup(bool enable);


// --- Application Entry Point ---
int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR    lpCmdLine,
    _In_ int       nCmdShow)
{
    WNDCLASS wc = {};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_CLIPBOARDTOFILE));
    RegisterClass(&wc);

    g_hMainWnd = CreateWindowEx(
        0, CLASS_NAME, L"Clipboard To File Helper", WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
        NULL, NULL, hInstance, NULL
    );

    if (g_hMainWnd == NULL) return 0;

    MSG msg = {};
    while (GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return (int)msg.wParam;
}


// --- Window Procedure: The Heart of the Application ---
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_CREATE:
        InitializeAndLoadExtensions();
        g_hNextClipboardViewer = SetClipboardViewer(hwnd);
        CreateTrayIcon(hwnd);
        g_hShutdownEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
        g_hWatcherThread = CreateThread(NULL, 0, FileWatcherThread, NULL, 0, NULL);
        break;

    case WM_DESTROY:
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
        if ((HWND)wParam == g_hNextClipboardViewer)
            g_hNextClipboardViewer = (HWND)lParam;
        else if (g_hNextClipboardViewer != NULL)
            SendMessage(g_hNextClipboardViewer, msg, wParam, lParam);
        break;

    case WM_DRAWCLIPBOARD:
        ProcessClipboardChange();
        SendMessage(g_hNextClipboardViewer, msg, wParam, lParam);
        break;

    case WM_APP_RELOAD_EXTENSIONS:
        Sleep(100);
        ReloadExtensions(false);
        break;

    case WM_TRAY_ICON_MSG:
        if (lParam == WM_RBUTTONUP) ShowContextMenu(hwnd);
        break;

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case ID_MENU_TOGGLE:
            g_isEnabled = !g_isEnabled;
            break;
        case ID_MENU_START_WITH_WINDOWS:
            SetStartup(!IsStartupEnabled());
            break;
        case ID_MENU_EDIT_EXTENSIONS:
        {
            wchar_t exePath[MAX_PATH];
            GetModuleFileNameW(NULL, exePath, MAX_PATH);
            PathRemoveFileSpecW(exePath);
            std::wstring settingsPath = std::wstring(exePath) + L"\\extensions.txt";
            ShellExecuteW(NULL, L"open", settingsPath.c_str(), NULL, NULL, SW_SHOWNORMAL);
        }
        break;
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


// --- File Watcher Worker Thread ---
DWORD WINAPI FileWatcherThread(LPVOID lpParam)
{
    wchar_t dirPath[MAX_PATH];
    GetModuleFileNameW(NULL, dirPath, MAX_PATH);
    PathRemoveFileSpecW(dirPath);

    HANDLE hDir = CreateFileW(dirPath, FILE_LIST_DIRECTORY,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, NULL);

    if (hDir == INVALID_HANDLE_VALUE) return 1;

    BYTE buffer[1024];
    DWORD bytesReturned;

    while (WaitForSingleObject(g_hShutdownEvent, 0) == WAIT_TIMEOUT)
    {
        if (ReadDirectoryChangesW(hDir, buffer, sizeof(buffer), FALSE,
            FILE_NOTIFY_CHANGE_LAST_WRITE | FILE_NOTIFY_CHANGE_FILE_NAME,
            &bytesReturned, NULL, NULL))
        {
            if (bytesReturned > 0)
            {
                FILE_NOTIFY_INFORMATION* pNotify = (FILE_NOTIFY_INFORMATION*)buffer;
                while (pNotify)
                {
                    std::wstring filename(pNotify->FileName, pNotify->FileNameLength / sizeof(wchar_t));
                    if (_wcsicmp(filename.c_str(), L"extensions.txt") == 0)
                    {
                        PostMessage(g_hMainWnd, WM_APP_RELOAD_EXTENSIONS, 0, 0);
                        break;
                    }
                    pNotify = pNotify->NextEntryOffset > 0 ?
                        (FILE_NOTIFY_INFORMATION*)((BYTE*)pNotify + pNotify->NextEntryOffset) : NULL;
                }
            }
        }
    }

    CloseHandle(hDir);
    return 0;
}


// --- Core Logic & File Management ---

void InitializeAndLoadExtensions()
{
    wchar_t exePath[MAX_PATH];
    GetModuleFileNameW(NULL, exePath, MAX_PATH);
    PathRemoveFileSpecW(exePath);
    std::wstring settingsPath = std::wstring(exePath) + L"\\extensions.txt";

    if (GetFileAttributesW(settingsPath.c_str()) == INVALID_FILE_ATTRIBUTES)
    {
        const std::vector<std::wstring> defaultExtensions = {
            L".txt", L".md", L".log", L".sql", L".cpp", L".h", L".js", L".json", L".xml"
        };
        std::wofstream newSettingsFile(settingsPath);
        if (newSettingsFile.is_open()) {
            for (const auto& ext : defaultExtensions) {
                newSettingsFile << ext << std::endl;
            }
            newSettingsFile.close();
        }
    }
    ReloadExtensions(true);
}

bool ReloadExtensions(bool isInitialLoad)
{
    wchar_t exePath[MAX_PATH];
    GetModuleFileNameW(NULL, exePath, MAX_PATH);
    PathRemoveFileSpecW(exePath);
    std::wstring settingsPath = std::wstring(exePath) + L"\\extensions.txt";

    std::vector<std::wstring> tempExtensions;
    bool hadBadLines = false;

    std::wifstream settingsFile(settingsPath);
    if (!settingsFile.is_open())
    {
        if (!isInitialLoad) ShowToastNotification(g_hMainWnd, L"Reload Failed", L"Could not open extensions.txt.", NIIF_ERROR);
        return false;
    }

    std::wstring line;
    while (std::getline(settingsFile, line))
    {
        line.erase(0, line.find_first_not_of(L" \t\n\r"));
        line.erase(line.find_last_not_of(L" \t\n\r") + 1);

        if (line.empty() || line[0] == L'#' || (line[0] == L'/' && line[1] == L'/')) continue;
        if (wcspbrk(line.c_str(), L"\\/:*?\"<>|")) {
            hadBadLines = true;
            continue;
        }

        if (line[0] != L'.') line.insert(0, 1, L'.');

        std::transform(line.begin(), line.end(), line.begin(), ::towlower);
        tempExtensions.push_back(line);
    }
    settingsFile.close();

    {
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        g_allowedExtensions.swap(tempExtensions);
    }

    if (!isInitialLoad)
    {
        if (hadBadLines) ShowToastNotification(g_hMainWnd, L"Extensions Reloaded", L"Settings updated. Some invalid lines were skipped.", NIIF_WARNING);
        else ShowToastNotification(g_hMainWnd, L"Extensions Reloaded", L"Settings have been updated from extensions.txt.", NIIF_INFO);
    }
    return true;
}

void ProcessClipboardChange()
{
    if (!g_isEnabled || !IsClipboardFormatAvailable(CF_UNICODETEXT)) return;
    if (!OpenClipboard(g_hMainWnd)) return;

    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (hData == NULL) { CloseClipboard(); return; }

    wchar_t* pszText = static_cast<wchar_t*>(GlobalLock(hData));
    if (pszText == NULL) { CloseClipboard(); return; }

    std::wstring clipboardText(pszText);
    GlobalUnlock(hData);
    CloseClipboard();

    clipboardText.erase(0, clipboardText.find_first_not_of(L" \t\n\r"));
    clipboardText.erase(clipboardText.find_last_not_of(L" \t\n\r") + 1);

    if (clipboardText.empty() || wcspbrk(clipboardText.c_str(), L"\\/:*?\"<>|") != nullptr) return;

    wchar_t ext[_MAX_EXT];
    _wsplitpath_s(clipboardText.c_str(), NULL, 0, NULL, 0, NULL, 0, ext, _MAX_EXT);
    std::wstring extension(ext);
    std::transform(extension.begin(), extension.end(), extension.begin(), ::towlower);

    bool isAllowed = false;
    {
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        for (const auto& allowedExt : g_allowedExtensions)
        {
            if (extension == allowedExt) {
                isAllowed = true;
                break;
            }
        }
    }

    if (!isAllowed) return;

    std::wstring explorerPath = GetSingleExplorerPath();
    if (!explorerPath.empty())
    {
        std::wstring fullPath = explorerPath + L"\\" + clipboardText;
        if (GetFileAttributesW(fullPath.c_str()) == INVALID_FILE_ATTRIBUTES)
        {
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

std::wstring GetSingleExplorerPath()
{
    std::vector<std::wstring> paths;
    IShellWindows* pShellWindows = NULL;

    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        return L"";
    }

    hr = CoCreateInstance(CLSID_ShellWindows, NULL, CLSCTX_ALL, IID_IShellWindows, (void**)&pShellWindows); 
    
    if (SUCCEEDED(hr))
    {
        long count;
        pShellWindows->get_Count(&count);
        for (long i = 0; i < count; i++)
        {
            VARIANT v; VariantInit(&v); v.vt = VT_I4; v.lVal = i;
            IDispatch* pDispatch;
            hr = pShellWindows->Item(v, &pDispatch);
            VariantClear(&v);
            if (SUCCEEDED(hr))
            {
                IWebBrowser2* pBrowser;
                hr = pDispatch->QueryInterface(IID_IWebBrowser2, (void**)&pBrowser);
                if (SUCCEEDED(hr))
                {
                    HWND hwndBrowser;
                    if (SUCCEEDED(pBrowser->get_HWND((SHANDLE_PTR*)&hwndBrowser)))
                    {
                        wchar_t className[256];
                        if (GetClassNameW(hwndBrowser, className, 256) && wcscmp(className, L"CabinetWClass") == 0)
                        {
                            BSTR bstrURL;
                            if (SUCCEEDED(pBrowser->get_LocationURL(&bstrURL)))
                            {
                                std::wstring url(bstrURL, SysStringLen(bstrURL));
                                SysFreeString(bstrURL);
                                DWORD pathLen = MAX_PATH;
                                wchar_t localPath[MAX_PATH];
                                if (SUCCEEDED(PathCreateFromUrlW(url.c_str(), localPath, &pathLen, 0)))
                                {
                                    paths.push_back(localPath);
                                }
                            }
                        }
                    }
                    pBrowser->Release();
                }
                pDispatch->Release();
            }
        }
        pShellWindows->Release();
    }
    CoUninitialize();

    if (paths.size() == 1) return paths[0];
    return L"";
}


// --- Tray Icon & UI Management ---

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

void RemoveTrayIcon(HWND hwnd)
{
    NOTIFYICONDATAW nid = {};
    nid.cbSize = sizeof(NOTIFYICONDATAW);
    nid.hWnd = hwnd;
    nid.uID = ID_TRAY_ICON;
    Shell_NotifyIconW(NIM_DELETE, &nid);
}

void ShowContextMenu(HWND hwnd)
{
    POINT pt;
    GetCursorPos(&pt);
    HMENU hMenu = CreatePopupMenu();
    if (hMenu)
    {
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


// --- Registry Helpers ---

const wchar_t* REG_RUN_KEY = L"Software\\Microsoft\\Windows\\CurrentVersion\\Run";
const wchar_t* REG_VALUE_NAME = L"ClipboardToFile";

bool IsStartupEnabled()
{
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, REG_RUN_KEY, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        LSTATUS status = RegQueryValueExW(hKey, REG_VALUE_NAME, NULL, NULL, NULL, NULL);
        RegCloseKey(hKey);
        return status == ERROR_SUCCESS;
    }
    return false;
}

void SetStartup(bool enable)
{
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, REG_RUN_KEY, 0, KEY_WRITE, &hKey) == ERROR_SUCCESS)
    {
        if (enable)
        {
            wchar_t exePath[MAX_PATH];
            GetModuleFileNameW(NULL, exePath, MAX_PATH);
            RegSetValueExW(hKey, REG_VALUE_NAME, 0, REG_SZ, (const BYTE*)exePath, (wcslen(exePath) + 1) * sizeof(wchar_t));
        }
        else
        {
            RegDeleteValueW(hKey, REG_VALUE_NAME);
        }
        RegCloseKey(hKey);
    }
}