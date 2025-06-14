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
#include <iomanip>      // For std::setw
#include <regex>        // For std::wregex
#include "nlohmann/json.hpp"     // For nlohmann/json library via submodule
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
#define WM_APP_RELOAD_CONFIG        (WM_USER + 2)   // Message from watcher thread to trigger reload
#define WM_APP_UPDATE_FOUND         (WM_USER + 3)   // Message for application updates
#define ID_TRAY_ICON                1
#define ID_MENU_TOGGLE_EMPTY        1001
#define ID_MENU_TOGGLE_CONTENT      1002
#define ID_MENU_EDIT_CONFIG         1003
#define ID_MENU_START_WITH_WINDOWS  1004
#define ID_MENU_EXIT                1005

const wchar_t CLASS_NAME[] = L"ClipboardToFileWindowClass";
HWND  g_hMainWnd = NULL;
HWND  g_hNextClipboardViewer = NULL;
HANDLE g_hWatcherThread = NULL;
HANDLE g_hShutdownEvent = NULL;
std::mutex g_extensionsMutex;

bool g_bIgnoreNextClipboard = true;  // Ignore first clipboard notification on startup

struct AppSettings {
    bool isCreateEmptyFileEnabled = true;
    bool isCreateWithContentEnabled = true;
    std::vector<std::wstring> allowedExtensions;
    std::vector<std::wstring> contentCreationRegexes;
    int heuristicWordCountLimit = 5;
};
AppSettings g_settings;

// Enum for file conflict resolution actions
enum class FileConflictAction {
    Replace,
    Skip,
    Rename
};


//------------------------------------------------------------------------------------------------//
//                                  FUNCTION PROTOTYPES                                           //
//------------------------------------------------------------------------------------------------//
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
void CreateTrayIcon(HWND);
void RemoveTrayIcon(HWND);
void ShowContextMenu(HWND);
void ShowToastNotification(HWND, const std::wstring&, const std::wstring&, DWORD);
std::wstring GetConfigFilePath();
std::wstring GetSingleExplorerPath();
void ProcessClipboardChange();
DWORD WINAPI FileWatcherThread(LPVOID);
void LoadSettings();
void SaveSettings();
bool IsStartupEnabled();
void SetStartup(bool);
void CheckForUpdatesIfNeeded();
bool TryFullFileGeneration(const std::wstring&);
void TryEmptyFileCreation(const std::wstring&);
int CountWords(const std::wstring&);
struct AppVersion { int major = 0, minor = 0, patch = 0, build = 0; };
AppVersion GetCurrentAppVersion();
AppVersion ParseVersionString(const std::wstring&);
FileConflictAction ShowFileConflictDialog(const std::wstring&);
std::wstring GenerateUniqueFilename(const std::wstring&);
bool CreateFileWithContentAtomic(const std::wstring&, const std::wstring&);
bool CreateEmptyFileAtomic(const std::wstring&);


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
        LoadSettings();
        // Use modern clipboard listener API (Vista+) instead of legacy viewer chain
        AddClipboardFormatListener(hwnd);
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

        // Clean up any pending WM_APP_UPDATE_FOUND messages to prevent memory leaks
        MSG pendingMsg;
        while (PeekMessageW(&pendingMsg, hwnd, WM_APP_UPDATE_FOUND, WM_APP_UPDATE_FOUND, PM_REMOVE)) {
            wchar_t* releaseUrl = (wchar_t*)pendingMsg.lParam;
            if (releaseUrl) {
                delete[] releaseUrl; // Free the memory allocated by PerformUpdateCheck
            }
        }

        // Remove modern clipboard listener (no chain management needed)
        RemoveClipboardFormatListener(hwnd);
        RemoveTrayIcon(hwnd);
        PostQuitMessage(0);
        break;
    case WM_CLIPBOARDUPDATE:
        // Modern clipboard change notification (Vista+) - more reliable than legacy chain
        if (g_bIgnoreNextClipboard) {
            g_bIgnoreNextClipboard = false;  // Reset flag after ignoring first notification
        }
        else {
            ProcessClipboardChange();
        }
        // No message forwarding needed with modern API - each listener gets direct notification
        break;
    case WM_APP_RELOAD_CONFIG:
        // Handles the reload request from our file watcher thread.
        Sleep(100); // Small delay to prevent race conditions with text editors.
        LoadSettings();
        ShowToastNotification(g_hMainWnd, L"Config Reloaded", L"Configuration has been updated from config.json.", NIIF_INFO);
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
        case ID_MENU_TOGGLE_EMPTY: {
            std::lock_guard<std::mutex> lock(g_extensionsMutex);
            g_settings.isCreateEmptyFileEnabled = !g_settings.isCreateEmptyFileEnabled;
            SaveSettings();
            break;
        }
        case ID_MENU_TOGGLE_CONTENT: {
            std::lock_guard<std::mutex> lock(g_extensionsMutex);
            g_settings.isCreateWithContentEnabled = !g_settings.isCreateWithContentEnabled;
            SaveSettings();
            break;
        }
        case ID_MENU_START_WITH_WINDOWS:
            SetStartup(!IsStartupEnabled());
            break;
        case ID_MENU_EDIT_CONFIG: {
            std::wstring settingsPath = GetConfigFilePath();
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
//                            CONFIGURATION & SETTINGS MANAGEMENT                                 //
//------------------------------------------------------------------------------------------------//
// Helper function to convert a UTF-8 std::string (from JSON library) to a std::wstring (for Win32 API).
std::wstring Utf8ToWstring(const std::string& str) {
    if (str.empty()) return std::wstring();
    int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), NULL, 0);
    std::wstring wstrTo(size_needed, 0);
    MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
    return wstrTo;
}

// Helper function to convert a std::wstring to a UTF-8 std::string for saving to JSON.
std::string WstringToUtf8(const std::wstring& wstr) {
    if (wstr.empty()) return std::string();
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), NULL, 0, NULL, NULL);
    std::string strTo(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), &strTo[0], size_needed, NULL, NULL);
    return strTo;
}

// Gets the full path to config.json in %APPDATA%\ClipboardToFile.
std::wstring GetConfigFilePath() {
    wchar_t appDataPath[MAX_PATH];
    if (SUCCEEDED(SHGetFolderPathW(NULL, CSIDL_APPDATA, NULL, 0, appDataPath))) {
        std::wstring fullPath = std::wstring(appDataPath) + L"\\ClipboardToFile";
        CreateDirectoryW(fullPath.c_str(), NULL);
        return fullPath + L"\\config.json";
    }
    return L"config.json"; // Fallback to local directory.
}

// Writes the current state of the g_settings struct to config.json, persisting user choices.
void SaveSettings() {
    std::wstring settingsPath = GetConfigFilePath();
    nlohmann::json j;
    {
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        j["createEmptyFileEnabled"] = g_settings.isCreateEmptyFileEnabled;
        j["createWithContentEnabled"] = g_settings.isCreateWithContentEnabled;
        std::vector<std::string> utf8_allowedExtensions;
        for (const auto& wstr : g_settings.allowedExtensions) utf8_allowedExtensions.push_back(WstringToUtf8(wstr));
        j["allowedExtensions"] = utf8_allowedExtensions;
        std::vector<std::string> utf8_regexes;
        for (const auto& wstr : g_settings.contentCreationRegexes) utf8_regexes.push_back(WstringToUtf8(wstr));
        j["contentCreationRegexes"] = utf8_regexes;
        j["heuristicWordCountLimit"] = g_settings.heuristicWordCountLimit;
    }
    std::ofstream o(settingsPath);
    o << std::setw(2) << j << std::endl;
}

// Reads config.json, creates a default if missing, and populates the global g_settings struct.
void LoadSettings() {
    std::wstring settingsPath = GetConfigFilePath();
    AppSettings defaults;
    defaults.allowedExtensions = { L".txt", L".md", L".log", L".sql", L".cpp", L".h", L".js", L".json", L".xml", L".cs", L".c" };
    defaults.contentCreationRegexes = {
        L"^// --- START OF FILE: (.*) ---$",
        L"^file: (.*)$",
        L"^(.*\\.[a-zA-Z0-9]+)$"
    };

    std::ifstream f(settingsPath);
    if (!f.is_open()) {
        {
            std::lock_guard<std::mutex> lock(g_extensionsMutex);
            g_settings = defaults;
        }
        SaveSettings(); // Save the new default file.
        return;
    }

    try {
        nlohmann::json j = nlohmann::json::parse(f);
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        g_settings.isCreateEmptyFileEnabled = j.value("createEmptyFileEnabled", defaults.isCreateEmptyFileEnabled);
        g_settings.isCreateWithContentEnabled = j.value("createWithContentEnabled", defaults.isCreateWithContentEnabled);
        if (j.contains("allowedExtensions")) {
            g_settings.allowedExtensions.clear();
            for (const auto& str : j["allowedExtensions"]) g_settings.allowedExtensions.push_back(Utf8ToWstring(str.get<std::string>()));
        }
        else { g_settings.allowedExtensions = defaults.allowedExtensions; }
        if (j.contains("contentCreationRegexes")) {
            g_settings.contentCreationRegexes.clear();
            for (const auto& str : j["contentCreationRegexes"]) g_settings.contentCreationRegexes.push_back(Utf8ToWstring(str.get<std::string>()));
        }
        else { g_settings.contentCreationRegexes = defaults.contentCreationRegexes; }
        g_settings.heuristicWordCountLimit = j.value("heuristicWordCountLimit", defaults.heuristicWordCountLimit);
    }
    catch (const nlohmann::json::parse_error&) {
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        g_settings = defaults;
        ShowToastNotification(g_hMainWnd, L"Config Error", L"Could not parse config.json. Loading defaults.", NIIF_ERROR);
    }
}


//------------------------------------------------------------------------------------------------//
//                                     UPDATE CHECKER                                             //
//------------------------------------------------------------------------------------------------//
// Parses a version string like "v1.2.3.4" into a struct for easy comparison.
AppVersion ParseVersionString(const std::wstring& versionStr) {
    AppVersion v;
    std::wstring s = (versionStr[0] == L'v') ? versionStr.substr(1) : versionStr;
    std::wstringstream ss(s);
    wchar_t dot;
    ss >> v.major >> dot >> v.minor >> dot >> v.patch >> dot >> v.build;
    return v;
}

// Reads the embedded VS_VERSION_INFO resource from the running executable to determine its own version.
AppVersion GetCurrentAppVersion() {
    wchar_t exePath[MAX_PATH];
    GetModuleFileNameW(NULL, exePath, MAX_PATH);
    DWORD handle = 0; // This is a dummy variable for GetFileVersionInfoSizeW.
    DWORD versionSize = GetFileVersionInfoSizeW(exePath, &handle);
    if (versionSize == 0) return {};
    std::vector<BYTE> versionData(versionSize);
    // The second parameter to GetFileVersionInfoW is ignored and per documentation should be 0.
    if (!GetFileVersionInfoW(exePath, 0, versionSize, versionData.data())) return {};
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
        // Explicitly define headers and calculate length to satisfy static analysis.
        const wchar_t* headers = L"User-Agent: ClipboardToFile-Update-Check\r\n";
        HINTERNET hConnect = InternetOpenUrlW(hInternet,
            L"https://api.github.com/repos/ByronAP/ClipboardToFile/releases/latest",
            headers, (DWORD)wcslen(headers),
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
    std::wstring configPath = GetConfigFilePath();
    wchar_t dirPath[MAX_PATH];
    wcscpy_s(dirPath, configPath.c_str());
    PathRemoveFileSpecW(dirPath);
    HANDLE hDir = CreateFileW(dirPath, FILE_LIST_DIRECTORY,
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
                    if (_wcsicmp(filename.c_str(), L"config.json") == 0) {
                        // Post a message to the main UI thread to handle the reload safely.
                        PostMessage(g_hMainWnd, WM_APP_RELOAD_CONFIG, 0, 0);
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
//                          FILE CONFLICT RESOLUTION                                              //
//------------------------------------------------------------------------------------------------//
// Shows a dialog asking the user what to do when a file already exists
FileConflictAction ShowFileConflictDialog(const std::wstring& filename)
{
    std::wstring message = L"The file '" + filename + L"' already exists.\n\n"
        L"What would you like to do?\n\n"
        L"Yes = Replace (overwrite the existing file)\n"
        L"No = Skip (do not create the file)\n"
        L"Cancel = Rename (create with a different name)";

    int result = MessageBoxW(NULL,
        message.c_str(),
        L"File Already Exists",
        MB_YESNOCANCEL | MB_ICONWARNING | MB_DEFBUTTON2);

    switch (result) {
    case IDYES: return FileConflictAction::Replace;
    case IDNO: return FileConflictAction::Skip;
    case IDCANCEL: return FileConflictAction::Rename;
    default: return FileConflictAction::Skip;
    }
}

// Generates a unique filename by appending a number to the base name
std::wstring GenerateUniqueFilename(const std::wstring& originalPath)
{
    if (GetFileAttributesW(originalPath.c_str()) == INVALID_FILE_ATTRIBUTES) {
        return originalPath; // Original doesn't exist, use it
    }

    wchar_t drive[_MAX_DRIVE];
    wchar_t dir[_MAX_DIR];
    wchar_t fname[_MAX_FNAME];
    wchar_t ext[_MAX_EXT];

    _wsplitpath_s(originalPath.c_str(), drive, _MAX_DRIVE, dir, _MAX_DIR, fname, _MAX_FNAME, ext, _MAX_EXT);

    int counter = 1;
    std::wstring newPath;

    do {
        std::wstringstream ss;
        ss << drive << dir << fname << L" (" << counter << L")" << ext;
        newPath = ss.str();
        counter++;
    } while (GetFileAttributesW(newPath.c_str()) != INVALID_FILE_ATTRIBUTES && counter < 1000);

    return newPath;
}


//------------------------------------------------------------------------------------------------//
//                          CORE LOGIC & FILE MANAGEMENT                                          //
//------------------------------------------------------------------------------------------------//
// A simple helper to count words in a string, used by the content-creation heuristic.
int CountWords(const std::wstring& str) {
    std::wstringstream ss(str);
    std::wstring word;
    int count = 0;
    while (ss >> word) count++;
    return count;
}

// Heuristically checks for a filename. It first tries matching against user-defined regexes,
// and if not found, falls back to a simpler "filename on first line" heuristic.
// Returns true if it successfully creates a file, preventing other logic from running.
bool TryFullFileGeneration(const std::wstring& clipboardText) {
    size_t first_line_end = clipboardText.find(L'\n');
    if (first_line_end == std::wstring::npos) return false;

    std::wstring firstLine = clipboardText.substr(0, first_line_end);
    firstLine.erase(0, firstLine.find_first_not_of(L" \t\r\n"));
    firstLine.erase(firstLine.find_last_not_of(L" \t\r\n") + 1);

    std::wstring filename;
    bool format_detected = false;

    // Priority 1: Iterate through user-defined regex patterns from config.
    std::vector<std::wstring> regexes;
    {
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        regexes = g_settings.contentCreationRegexes;
    }
    for (const auto& pattern : regexes) {
        try {
            std::wregex rx(pattern, std::regex::ECMAScript | std::regex::icase);
            std::wsmatch match;
            if (std::regex_match(firstLine, match, rx) && match.size() > 1) {
                filename = match[1].str();
                format_detected = true;
                break;
            }
        }
        catch (const std::regex_error&) { continue; } // Silently ignore invalid regex patterns.
    }

    // Priority 2: Fallback to the simpler word-count heuristic.
    if (!format_detected) {
        wchar_t ext[_MAX_EXT];
        _wsplitpath_s(firstLine.c_str(), NULL, 0, NULL, 0, NULL, 0, ext, _MAX_EXT);
        std::wstring extension(ext);
        std::transform(extension.begin(), extension.end(), extension.begin(), ::towlower);

        bool isAllowedExtension = false;
        int wordCountLimit = 5;
        {
            std::lock_guard<std::mutex> lock(g_extensionsMutex);
            for (const auto& allowedExt : g_settings.allowedExtensions) {
                if (extension == allowedExt) { isAllowedExtension = true; break; }
            }
            wordCountLimit = g_settings.heuristicWordCountLimit;
        }

        if (isAllowedExtension && CountWords(firstLine) <= wordCountLimit) {
            filename = firstLine;
            format_detected = true;
        }
    }

    // If a filename was found by any method, proceed with file creation.
    if (format_detected) {
        filename.erase(0, filename.find_first_not_of(L" \t\n\r"));
        filename.erase(filename.find_last_not_of(L" \t\n\r") + 1);
        if (filename.empty() || wcspbrk(filename.c_str(), L"\\/:*?\"<>|") != nullptr) {
            return true; // Detected a pattern but filename is invalid. Stop all further processing.
        }
        std::wstring content = clipboardText.substr(first_line_end + 1);
        std::wstring explorerPath = GetSingleExplorerPath();
        if (!explorerPath.empty()) {
            std::wstring fullPath = explorerPath + L"\\" + filename;
            std::wstring finalPath = fullPath;

            // Check if file exists and handle conflict
            if (GetFileAttributesW(fullPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
                FileConflictAction action = ShowFileConflictDialog(filename);

                switch (action) {
                case FileConflictAction::Skip:
                    return true; // User chose to skip, don't create file
                case FileConflictAction::Rename:
                    finalPath = GenerateUniqueFilename(fullPath);
                    // Extract just the filename for the success message
                    wchar_t fname[_MAX_FNAME];
                    wchar_t ext[_MAX_EXT];
                    _wsplitpath_s(finalPath.c_str(), NULL, 0, NULL, 0, fname, _MAX_FNAME, ext, _MAX_EXT);
                    filename = std::wstring(fname) + ext;
                    break;
                case FileConflictAction::Replace:
                    // Will use atomic replacement
                    break;
                }
            }

            bool success = false;

            if (GetFileAttributesW(finalPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
                // File exists and user chose to replace - use atomic replacement
                success = CreateFileWithContentAtomic(finalPath, content);
            }
            else {
                // File doesn't exist - create normally
                std::wofstream outFile(finalPath);
                if (outFile.is_open()) {
                    outFile << content;
                    outFile.close();
                    success = !outFile.fail();
                }
            }

            if (success) {
                std::wstring successMessage = L"Generated file with content: " + filename;
                ShowToastNotification(g_hMainWnd, L"File Generated", successMessage, NIIF_INFO);
                return true;
            }
        }
    }
    return false;
}


// Handles creating an empty file if the entire clipboard text is a valid filename.
void TryEmptyFileCreation(const std::wstring& clipboardText) {
    std::wstring filename = clipboardText;
    filename.erase(0, filename.find_first_not_of(L" \t\n\r"));
    filename.erase(filename.find_last_not_of(L" \t\n\r") + 1);
    if (filename.empty() || wcspbrk(filename.c_str(), L"\\/:*?\"<>|") != nullptr) return;

    // To avoid ambiguity, don't treat multi-line text as an empty filename.
    if (filename.find(L'\n') != std::wstring::npos) return;

    wchar_t ext[_MAX_EXT];
    _wsplitpath_s(filename.c_str(), NULL, 0, NULL, 0, NULL, 0, ext, _MAX_EXT);
    std::wstring extension(ext);
    std::transform(extension.begin(), extension.end(), extension.begin(), ::towlower);

    // Thread-safe read of the allowed extensions list.
    bool isAllowed = false;
    {
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        for (const auto& allowedExt : g_settings.allowedExtensions) {
            if (extension == allowedExt) { isAllowed = true; break; }
        }
    }

    if (!isAllowed) return;

    // Safety check: only proceed if exactly one File Explorer window is open.
    std::wstring explorerPath = GetSingleExplorerPath();
    if (!explorerPath.empty()) {
        std::wstring fullPath = explorerPath + L"\\" + filename;
        std::wstring finalPath = fullPath;

        // Check if file exists and handle conflict
        if (GetFileAttributesW(fullPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
            FileConflictAction action = ShowFileConflictDialog(filename);

            switch (action) {
            case FileConflictAction::Skip:
                return; // User chose to skip, don't create file
            case FileConflictAction::Rename:
                finalPath = GenerateUniqueFilename(fullPath);
                // Extract just the filename for the success message
                wchar_t fname[_MAX_FNAME];
                wchar_t ext[_MAX_EXT];
                _wsplitpath_s(finalPath.c_str(), NULL, 0, NULL, 0, fname, _MAX_FNAME, ext, _MAX_EXT);
                filename = std::wstring(fname) + ext;
                break;
            case FileConflictAction::Replace:
                // For replace, we'll use atomic replacement with temporary file
                break;
            }
        }

        bool success = false;

        if (GetFileAttributesW(finalPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
            // File exists and user chose to replace - use atomic replacement
            success = CreateEmptyFileAtomic(finalPath);
        }
        else {
            // File doesn't exist - create normally
            HANDLE hFile = CreateFileW(finalPath.c_str(), GENERIC_WRITE, 0, NULL, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
            if (hFile != INVALID_HANDLE_VALUE) {
                CloseHandle(hFile);
                success = true;
            }
        }

        if (success) {
            std::wstring successMessage = L"Created empty file: " + filename;
            ShowToastNotification(g_hMainWnd, L"File Created", successMessage, NIIF_INFO);
        }
    }
}

// Helper function for atomic file replacement with content
bool CreateFileWithContentAtomic(const std::wstring& targetPath, const std::wstring& content) {
    // Generate temporary filename in same directory
    wchar_t drive[_MAX_DRIVE];
    wchar_t dir[_MAX_DIR];
    wchar_t fname[_MAX_FNAME];
    wchar_t ext[_MAX_EXT];

    _wsplitpath_s(targetPath.c_str(), drive, _MAX_DRIVE, dir, _MAX_DIR, fname, _MAX_FNAME, ext, _MAX_EXT);

    // Create temporary filename
    std::wstring tempPath;
    int counter = 0;
    do {
        std::wstringstream ss;
        ss << drive << dir << fname << L"_tmp_" << counter << ext;
        tempPath = ss.str();
        counter++;
    } while (GetFileAttributesW(tempPath.c_str()) != INVALID_FILE_ATTRIBUTES && counter < 1000);

    if (counter >= 1000) {
        return false; // Couldn't generate unique temp name
    }

    // Create the temporary file with content
    std::wofstream tempFile(tempPath);
    if (!tempFile.is_open()) {
        return false;
    }

    tempFile << content;
    tempFile.close();

    // Check if write was successful
    if (tempFile.fail()) {
        DeleteFileW(tempPath.c_str());
        return false;
    }

    // Atomically replace the original file with the temporary file
    if (MoveFileExW(tempPath.c_str(), targetPath.c_str(), MOVEFILE_REPLACE_EXISTING)) {
        return true;
    }
    else {
        // If atomic replacement failed, clean up the temporary file
        DeleteFileW(tempPath.c_str());
        return false;
    }
}

// Helper function for atomic file replacement
bool CreateEmptyFileAtomic(const std::wstring& targetPath) {
    // Generate temporary filename in same directory
    wchar_t drive[_MAX_DRIVE];
    wchar_t dir[_MAX_DIR];
    wchar_t fname[_MAX_FNAME];
    wchar_t ext[_MAX_EXT];

    _wsplitpath_s(targetPath.c_str(), drive, _MAX_DRIVE, dir, _MAX_DIR, fname, _MAX_FNAME, ext, _MAX_EXT);

    // Create temporary filename
    std::wstring tempPath;
    int counter = 0;
    do {
        std::wstringstream ss;
        ss << drive << dir << fname << L"_tmp_" << counter << ext;
        tempPath = ss.str();
        counter++;
    } while (GetFileAttributesW(tempPath.c_str()) != INVALID_FILE_ATTRIBUTES && counter < 1000);

    if (counter >= 1000) {
        return false; // Couldn't generate unique temp name
    }

    // Create the temporary empty file
    HANDLE hTempFile = CreateFileW(tempPath.c_str(), GENERIC_WRITE, 0, NULL, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hTempFile == INVALID_HANDLE_VALUE) {
        return false;
    }
    CloseHandle(hTempFile);

    // Atomically replace the original file with the temporary file
    if (MoveFileExW(tempPath.c_str(), targetPath.c_str(), MOVEFILE_REPLACE_EXISTING)) {
        return true;
    }
    else {
        // If atomic replacement failed, clean up the temporary file
        DeleteFileW(tempPath.c_str());
        return false;
    }
}

// Main dispatcher called on every clipboard change.
void ProcessClipboardChange()
{
    bool emptyEnabled, contentEnabled;
    {
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        emptyEnabled = g_settings.isCreateEmptyFileEnabled;
        contentEnabled = g_settings.isCreateWithContentEnabled;
    }
    if (!emptyEnabled && !contentEnabled) return;

    if (!IsClipboardFormatAvailable(CF_UNICODETEXT) || !OpenClipboard(g_hMainWnd)) return;
    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (hData == NULL) { CloseClipboard(); return; }
    wchar_t* pszText = static_cast<wchar_t*>(GlobalLock(hData));
    if (pszText == NULL) { CloseClipboard(); return; }
    std::wstring clipboardText(pszText);
    GlobalUnlock(hData);
    CloseClipboard();

    // Give "Create with Content" priority if it's enabled.
    if (contentEnabled) {
        if (TryFullFileGeneration(clipboardText)) {
            return; // If it succeeded or handled the event, we are done.
        }
    }
    // Otherwise, try the "Create Empty File" logic if it's enabled.
    if (emptyEnabled) {
        TryEmptyFileCreation(clipboardText);
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
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        UINT emptyFlags = g_settings.isCreateEmptyFileEnabled ? MF_STRING | MF_CHECKED : MF_STRING | MF_UNCHECKED;
        InsertMenu(hMenu, 0, MF_BYPOSITION | emptyFlags, ID_MENU_TOGGLE_EMPTY, L"Create Empty File");
        UINT contentFlags = g_settings.isCreateWithContentEnabled ? MF_STRING | MF_CHECKED : MF_STRING | MF_UNCHECKED;
        InsertMenu(hMenu, 1, MF_BYPOSITION | contentFlags, ID_MENU_TOGGLE_CONTENT, L"Create File with Content");
        InsertMenu(hMenu, 2, MF_SEPARATOR, 0, NULL);
        UINT startupFlags = IsStartupEnabled() ? MF_STRING | MF_CHECKED : MF_STRING | MF_UNCHECKED;
        InsertMenu(hMenu, 3, MF_BYPOSITION | startupFlags, ID_MENU_START_WITH_WINDOWS, L"Start with Windows");
        InsertMenu(hMenu, 4, MF_BYPOSITION | MF_STRING, ID_MENU_EDIT_CONFIG, L"Edit Config...");
        InsertMenu(hMenu, 5, MF_SEPARATOR, 0, NULL);
        InsertMenu(hMenu, 6, MF_BYPOSITION | MF_STRING, ID_MENU_EXIT, L"Exit");
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