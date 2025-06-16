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
#include <queue>
#include <stack>
#include <memory>


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
#define ID_MENU_TOGGLE_DIRECTORY    1006
#define ID_MENU_DIRECTORY_OPTIONS   1007

const wchar_t CLASS_NAME[] = L"ClipboardToFileWindowClass";
const wchar_t* REG_APP_KEY = L"Software\\ByronAP\\ClipboardToFile";
HWND  g_hMainWnd = NULL;
HWND  g_hNextClipboardViewer = NULL;
HANDLE g_hWatcherThread = NULL;
HANDLE g_hShutdownEvent = NULL;
std::vector<std::wregex> g_compiledRegexes;
std::mutex g_extensionsMutex;

bool g_bComInitialized = false;  // Track COM initialization state

// Struct to hold both pattern and compiled regex for efficient reuse
struct CompiledRegex {
    std::wstring pattern;
    std::wregex compiled;
    bool isValid;

    CompiledRegex() : isValid(false) {}
    CompiledRegex(const std::wstring& pat) : pattern(pat), isValid(false) {
        try {
            compiled = std::wregex(pattern, std::regex::ECMAScript | std::regex::icase);
            isValid = true;
        }
        catch (const std::regex_error&) {
            isValid = false;
        }
    }
};

struct AppSettings {
    bool isCreateEmptyFileEnabled = true;
    bool isCreateWithContentEnabled = true;
    bool isCreateDirectoryStructureEnabled = true;
    std::vector<std::wstring> allowedExtensions;
    std::vector<std::wstring> contentCreationRegexes;
    int heuristicWordCountLimit = 5;
    bool createEmptyDirectories = true;
    bool skipExistingDirectories = true;
};
AppSettings g_settings;

// Enum for file conflict resolution actions
enum class FileConflictAction {
    Replace,
    Skip,
    Rename
};

struct TreeNode {
    std::wstring name;
    bool isDirectory;
    std::wstring content;  // For enhanced format with file contents
    std::vector<std::unique_ptr<TreeNode>> children;

    TreeNode(const std::wstring& n, bool isDir = false) : name(n), isDirectory(isDir) {}
};

enum class TreeFormat {
    Unknown,
    TreeCommand,      // Uses ├── └── characters
    Indentation,      // Uses spaces/tabs
    PathList,         // Full paths like path/to/file.txt
    Enhanced          // With file content markers
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
bool TryFileGeneration(const std::wstring&);
int CountWords(const std::wstring&);
struct AppVersion { int major = 0, minor = 0, patch = 0, build = 0; };
AppVersion GetCurrentAppVersion();
AppVersion ParseVersionString(const std::wstring&);
FileConflictAction ShowFileConflictDialog(const std::wstring&);
std::wstring GenerateUniqueFilename(const std::wstring&);
bool CreateFileWithContentAtomic(const std::wstring&, const std::wstring&);
bool CreateEmptyFileAtomic(const std::wstring&);
bool IsValidFilename(const std::wstring&);
std::vector<std::wstring> FindAdditionalFilenames(const std::wstring&, size_t);
bool TryDirectoryStructureCreation(const std::wstring& clipboardText);
TreeFormat DetectTreeFormat(const std::wstring& text);
std::unique_ptr<TreeNode> ParseTreeStructure(const std::wstring& text, TreeFormat format);
std::unique_ptr<TreeNode> ParseTreeCommandFormat(const std::vector<std::wstring>& lines);
std::unique_ptr<TreeNode> ParseIndentationFormat(const std::vector<std::wstring>& lines);
std::unique_ptr<TreeNode> ParsePathListFormat(const std::vector<std::wstring>& lines);
std::unique_ptr<TreeNode> ParseEnhancedFormat(const std::vector<std::wstring>& lines);
bool CreateDirectoryStructure(const TreeNode* root, const std::wstring& basePath);
bool IsPathSafe(const std::wstring& path);
void GetTreeSummary(const TreeNode* node, int& dirCount, int& fileCount);


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
        // Initialize COM once at startup for the main thread
        if (SUCCEEDED(CoInitialize(NULL))) {
            g_bComInitialized = true;
        }

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
        // Uninitialize COM once at shutdown
        if (g_bComInitialized) {
            CoUninitialize();
            g_bComInitialized = false;
        }
        PostQuitMessage(0);
        break;
    case WM_CLIPBOARDUPDATE:
        // Modern clipboard change notification (Vista+) - more reliable than legacy chain
        ProcessClipboardChange();

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
        case ID_MENU_TOGGLE_DIRECTORY: {
            std::lock_guard<std::mutex> lock(g_extensionsMutex);
            g_settings.isCreateDirectoryStructureEnabled = !g_settings.isCreateDirectoryStructureEnabled;
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

// Helper function to precompile regex patterns (call with mutex already held)
void CompileRegexPatterns() {
    g_compiledRegexes.clear();
    for (const auto& pattern : g_settings.contentCreationRegexes) {
        try {
            g_compiledRegexes.emplace_back(pattern, std::regex::ECMAScript | std::regex::icase);
        }
        catch (const std::regex_error&) {
            // Skip invalid regex patterns - don't add to compiled list
            continue;
        }
    }
}

// Writes the current state of the g_settings struct to config.json, persisting user choices.
void SaveSettings() {
    std::wstring settingsPath = GetConfigFilePath();
    nlohmann::json j;
    {
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        j["createEmptyFileEnabled"] = g_settings.isCreateEmptyFileEnabled;
        j["createWithContentEnabled"] = g_settings.isCreateWithContentEnabled;
        j["createDirectoryStructureEnabled"] = g_settings.isCreateDirectoryStructureEnabled;
        j["createEmptyDirectories"] = g_settings.createEmptyDirectories;
        j["skipExistingDirectories"] = g_settings.skipExistingDirectories;

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
    defaults.isCreateDirectoryStructureEnabled = true;
    defaults.createEmptyDirectories = true;
    defaults.skipExistingDirectories = true;

    std::ifstream f(settingsPath);
    if (!f.is_open()) {
        {
            std::lock_guard<std::mutex> lock(g_extensionsMutex);
            g_settings = defaults;
            CompileRegexPatterns();
        }
        
        SaveSettings(); // Save the new default file.
        return;
    }

    try {
        nlohmann::json j = nlohmann::json::parse(f);
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        g_settings.isCreateEmptyFileEnabled = j.value("createEmptyFileEnabled", defaults.isCreateEmptyFileEnabled);
        g_settings.isCreateWithContentEnabled = j.value("createWithContentEnabled", defaults.isCreateWithContentEnabled);
        g_settings.isCreateDirectoryStructureEnabled = j.value("createDirectoryStructureEnabled", defaults.isCreateDirectoryStructureEnabled);
        g_settings.createEmptyDirectories = j.value("createEmptyDirectories", defaults.createEmptyDirectories);
        g_settings.skipExistingDirectories = j.value("skipExistingDirectories", defaults.skipExistingDirectories);

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
        CompileRegexPatterns();
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

bool TryDirectoryStructureCreation(const std::wstring& clipboardText) {
    bool enabled;
    {
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        enabled = g_settings.isCreateDirectoryStructureEnabled;
    }

    if (!enabled) return false;

    // Detect format
    TreeFormat format = DetectTreeFormat(clipboardText);
    if (format == TreeFormat::Unknown) return false;

    // Parse the structure
    auto root = ParseTreeStructure(clipboardText, format);
    if (!root) return false;

    // Get Explorer path
    std::wstring explorerPath = GetSingleExplorerPath();
    if (explorerPath.empty()) {
        ShowToastNotification(g_hMainWnd, L"Error", L"No File Explorer window found.", NIIF_ERROR);
        return false;
    }

    // Count items for user confirmation
    int dirCount = 0, fileCount = 0;
    GetTreeSummary(root.get(), dirCount, fileCount);

    // Show confirmation dialog for large structures
    if (dirCount + fileCount > 10) {
        std::wstring message = L"Create directory structure with:\n\n";
        message += L"• " + std::to_wstring(dirCount) + L" directories\n";
        message += L"• " + std::to_wstring(fileCount) + L" files\n\n";
        message += L"Continue?";

        if (MessageBoxW(NULL, message.c_str(), L"Confirm Directory Structure",
            MB_YESNO | MB_ICONQUESTION) != IDYES) {
            return true; // User cancelled, but we handled the clipboard
        }
    }

    // Create the structure
    if (CreateDirectoryStructure(root.get(), explorerPath)) {
        std::wstring msg = L"Created " + std::to_wstring(dirCount) + L" directories and " +
            std::to_wstring(fileCount) + L" files";
        ShowToastNotification(g_hMainWnd, L"Structure Created", msg, NIIF_INFO);
        return true;
    }
    else {
        ShowToastNotification(g_hMainWnd, L"Error", L"Failed to create directory structure", NIIF_ERROR);
        return false;
    }
}

TreeFormat DetectTreeFormat(const std::wstring& text) {
    // Check for tree command characters (using Unicode code points)
    // 0x251C = '├', 0x2514 = '└', 0x2502 = '│'
    if (text.find(0x251C) != std::wstring::npos || text.find(0x2514) != std::wstring::npos ||
        text.find(0x2502) != std::wstring::npos) {
        return TreeFormat::TreeCommand;
    }

    // Check for enhanced format markers
    if (text.find(L"---START:") != std::wstring::npos || text.find(L"---END:") != std::wstring::npos) {
        return TreeFormat::Enhanced;
    }

    // Split into lines for further analysis
    std::vector<std::wstring> lines;
    std::wstringstream ss(text);
    std::wstring line;
    while (std::getline(ss, line)) {
        if (!line.empty()) lines.push_back(line);
    }

    if (lines.empty()) return TreeFormat::Unknown;

    // Check for path list format (contains forward or back slashes)
    bool hasSlashes = false;
    for (const auto& l : lines) {
        if (l.find(L'/') != std::wstring::npos || l.find(L'\\') != std::wstring::npos) {
            hasSlashes = true;
            break;
        }
    }

    // Check for consistent indentation
    bool hasIndentation = false;
    for (const auto& l : lines) {
        if (l[0] == L' ' || l[0] == L'\t') {
            hasIndentation = true;
            break;
        }
    }

    if (hasSlashes && !hasIndentation) return TreeFormat::PathList;
    if (hasIndentation) return TreeFormat::Indentation;

    return TreeFormat::Unknown;
}

std::unique_ptr<TreeNode> ParseTreeStructure(const std::wstring& text, TreeFormat format) {
    // Split into lines
    std::vector<std::wstring> lines;
    std::wstringstream ss(text);
    std::wstring line;
    while (std::getline(ss, line)) {
        lines.push_back(line);
    }

    switch (format) {
    case TreeFormat::TreeCommand:
        return ParseTreeCommandFormat(lines);
    case TreeFormat::Indentation:
        return ParseIndentationFormat(lines);
    case TreeFormat::PathList:
        return ParsePathListFormat(lines);
    case TreeFormat::Enhanced:
        return ParseEnhancedFormat(lines);
    default:
        return nullptr;
    }
}

std::unique_ptr<TreeNode> ParseTreeCommandFormat(const std::vector<std::wstring>& lines) {
    auto root = std::make_unique<TreeNode>(L"root", true);
    std::vector<TreeNode*> stack;
    stack.push_back(root.get());

    for (const auto& line : lines) {
        if (line.empty()) continue;

        // Count depth by tree characters
        int depth = 0;
        size_t pos = 0;
        while (pos < line.length()) {
            if (line[pos] == 0x2502 || line[pos] == L' ') {  // 0x2502 is '│'
                depth++;
                pos += 4; // Tree characters are usually followed by 3 spaces
            }
            else {
                break;
            }
        }

        // Find the actual content after tree characters
        // 0x2502 = '│', 0x251C = '├', 0x2514 = '└', 0x2500 = '─'
        std::wstring treeChars = L" \t";
        treeChars += static_cast<wchar_t>(0x2502);  // │
        treeChars += static_cast<wchar_t>(0x251C);  // ├
        treeChars += static_cast<wchar_t>(0x2514);  // └
        treeChars += static_cast<wchar_t>(0x2500);  // ─
        size_t contentStart = line.find_first_not_of(treeChars, pos);
        if (contentStart == std::wstring::npos) continue;

        std::wstring name = line.substr(contentStart);
        name.erase(0, name.find_first_not_of(L" \t"));
        name.erase(name.find_last_not_of(L" \t\r") + 1);

        if (name.empty()) continue;

        // Check if it's a directory (ends with /)
        bool isDir = name.back() == L'/';
        if (isDir) name.pop_back();

        // Adjust stack to current depth
        while (stack.size() > depth + 1) stack.pop_back();

        // Create node and add to parent
        auto node = std::make_unique<TreeNode>(name, isDir);
        TreeNode* nodePtr = node.get();
        stack.back()->children.push_back(std::move(node));

        if (isDir) stack.push_back(nodePtr);
    }

    return root;
}

std::unique_ptr<TreeNode> ParseIndentationFormat(const std::vector<std::wstring>& lines) {
    auto root = std::make_unique<TreeNode>(L"root", true);
    std::vector<std::pair<TreeNode*, int>> stack; // node, indent level
    stack.push_back({ root.get(), -1 });

    for (const auto& line : lines) {
        if (line.empty()) continue;

        // Count leading spaces/tabs
        int indent = 0;
        for (wchar_t c : line) {
            if (c == L' ') indent++;
            else if (c == L'\t') indent += 4; // treat tab as 4 spaces
            else break;
        }

        // Extract name
        std::wstring name = line.substr(indent);
        name.erase(0, name.find_first_not_of(L" \t"));
        name.erase(name.find_last_not_of(L" \t\r") + 1);

        if (name.empty()) continue;

        // Check if directory
        bool isDir = name.back() == L'/';
        if (isDir) name.pop_back();

        // Find parent based on indentation
        while (!stack.empty() && stack.back().second >= indent) {
            stack.pop_back();
        }

        // Create node
        auto node = std::make_unique<TreeNode>(name, isDir);
        TreeNode* nodePtr = node.get();
        stack.back().first->children.push_back(std::move(node));

        if (isDir) stack.push_back({ nodePtr, indent });
    }

    return root;
}

std::unique_ptr<TreeNode> ParsePathListFormat(const std::vector<std::wstring>& lines) {
    auto root = std::make_unique<TreeNode>(L"root", true);

    for (const auto& line : lines) {
        std::wstring path = line;
        path.erase(0, path.find_first_not_of(L" \t"));
        path.erase(path.find_last_not_of(L" \t\r") + 1);

        if (path.empty()) continue;

        // Normalize path separators
        std::replace(path.begin(), path.end(), L'\\', L'/');

        // Split path into components
        std::vector<std::wstring> components;
        std::wstringstream ss(path);
        std::wstring component;
        while (std::getline(ss, component, L'/')) {
            if (!component.empty()) components.push_back(component);
        }

        if (components.empty()) continue;

        // Navigate/create path in tree
        TreeNode* current = root.get();
        for (size_t i = 0; i < components.size(); ++i) {
            const auto& comp = components[i];
            bool isLastComponent = (i == components.size() - 1);
            bool isDir = isLastComponent ? (path.back() == L'/') : true;

            // Check for file extension in last component
            if (isLastComponent && !isDir) {
                size_t dotPos = comp.find_last_of(L'.');
                if (dotPos != std::wstring::npos && dotPos > 0) {
                    isDir = false;
                }
                else {
                    isDir = true; // No extension, assume directory
                }
            }

            // Find or create child
            TreeNode* child = nullptr;
            for (auto& c : current->children) {
                if (c->name == comp) {
                    child = c.get();
                    break;
                }
            }

            if (!child) {
                auto newChild = std::make_unique<TreeNode>(comp, isDir);
                child = newChild.get();
                current->children.push_back(std::move(newChild));
            }

            if (isDir) current = child;
        }
    }

    return root;
}

std::unique_ptr<TreeNode> ParseEnhancedFormat(const std::vector<std::wstring>& lines) {
    auto root = ParseIndentationFormat(lines); // Start with basic indentation parsing

    // Now look for content markers
    std::wstring currentFile;
    std::wstring currentContent;
    bool inContent = false;

    for (size_t i = 0; i < lines.size(); ++i) {
        const auto& line = lines[i];

        // Check for content start marker
        if (line.find(L"---START:") != std::wstring::npos) {
            size_t start = line.find(L"---START:") + 9;
            size_t end = line.find(L"---", start);
            if (end != std::wstring::npos) {
                currentFile = line.substr(start, end - start);
                currentFile.erase(0, currentFile.find_first_not_of(L" \t"));
                currentFile.erase(currentFile.find_last_not_of(L" \t") + 1);
                inContent = true;
                currentContent.clear();
            }
        }
        // Check for content end marker
        else if (line.find(L"---END:") != std::wstring::npos && inContent) {
            inContent = false;
            // Find the file node and set its content
            std::function<void(TreeNode*)> setContent = [&](TreeNode* node) {
                if (!node->isDirectory && node->name == currentFile) {
                    node->content = currentContent;
                    return;
                }
                for (auto& child : node->children) {
                    setContent(child.get());
                }
                };
            setContent(root.get());
        }
        // Collect content
        else if (inContent) {
            if (!currentContent.empty()) currentContent += L"\n";
            currentContent += line;
        }
    }

    return root;
}

bool CreateDirectoryStructure(const TreeNode* root, const std::wstring& basePath) {
    if (!root || root->children.empty()) return false;

    bool skipExisting;
    {
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        skipExisting = g_settings.skipExistingDirectories;
    }

    std::function<bool(const TreeNode*, const std::wstring&)> createNode =
        [&](const TreeNode* node, const std::wstring& parentPath) -> bool {

        std::wstring fullPath = parentPath + L"\\" + node->name;

        // Security check
        if (!IsPathSafe(fullPath)) {
            ShowToastNotification(g_hMainWnd, L"Security Error",
                L"Invalid path detected: " + node->name, NIIF_ERROR);
            return false;
        }

        if (node->isDirectory) {
            // Create directory
            DWORD attrs = GetFileAttributesW(fullPath.c_str());
            if (attrs == INVALID_FILE_ATTRIBUTES) {
                if (!CreateDirectoryW(fullPath.c_str(), NULL)) {
                    DWORD error = GetLastError();
                    if (error != ERROR_ALREADY_EXISTS) {
                        return false;
                    }
                }
            }
            else if (!(attrs & FILE_ATTRIBUTE_DIRECTORY)) {
                // File exists with same name
                if (!skipExisting) {
                    ShowToastNotification(g_hMainWnd, L"Error",
                        L"File exists with directory name: " + node->name, NIIF_ERROR);
                    return false;
                }
            }

            // Create children
            for (const auto& child : node->children) {
                if (!createNode(child.get(), fullPath)) {
                    return false;
                }
            }
        }
        else {
            // Create file
            bool createEmptyDirs;
            {
                std::lock_guard<std::mutex> lock(g_extensionsMutex);
                createEmptyDirs = g_settings.createEmptyDirectories;
            }

            // Ensure parent directory exists
            std::wstring parentDir = fullPath.substr(0, fullPath.find_last_of(L"\\"));
            DWORD attrs = GetFileAttributesW(parentDir.c_str());
            if (attrs == INVALID_FILE_ATTRIBUTES && createEmptyDirs) {
                // Create parent directories if needed
                std::wstring currentPath;
                std::wstringstream ss(parentDir);
                std::wstring segment;
                while (std::getline(ss, segment, L'\\')) {
                    if (!currentPath.empty()) currentPath += L"\\";
                    currentPath += segment;
                    CreateDirectoryW(currentPath.c_str(), NULL);
                }
            }

            // Create the file
            if (GetFileAttributesW(fullPath.c_str()) == INVALID_FILE_ATTRIBUTES) {
                if (!node->content.empty()) {
                    // Create with content
                    std::wofstream file(fullPath);
                    if (file.is_open()) {
                        file << node->content;
                        file.close();
                    }
                    else {
                        return false;
                    }
                }
                else {
                    // Create empty file
                    HANDLE hFile = CreateFileW(fullPath.c_str(), GENERIC_WRITE, 0, NULL,
                        CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
                    if (hFile != INVALID_HANDLE_VALUE) {
                        CloseHandle(hFile);
                    }
                    else {
                        return false;
                    }
                }
            }
        }

        return true;
        };

    // Create all children of root (skip the root node itself)
    for (const auto& child : root->children) {
        if (!createNode(child.get(), basePath)) {
            return false;
        }
    }

    return true;
}

// Unified function that handles both empty file generation and file generation with content
bool TryFileGeneration(const std::wstring& clipboardText) {
    bool emptyEnabled, contentEnabled;
    {
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        emptyEnabled = g_settings.isCreateEmptyFileEnabled;
        contentEnabled = g_settings.isCreateWithContentEnabled;
    }

    if (!emptyEnabled && !contentEnabled) return false;

    size_t first_line_end = clipboardText.find(L'\n');

    std::wstring firstLine;
    std::wstring content;
    bool isMultiLine = (first_line_end != std::wstring::npos);

    if (isMultiLine) {
        // Multi-line content: split at newline
        firstLine = clipboardText.substr(0, first_line_end);
        content = clipboardText.substr(first_line_end + 1);

        // If content creation is disabled, don't process multi-line content
        if (!contentEnabled) return false;
    }
    else {
        // Single-line content: treat entire clipboard as "first line" initially
        firstLine = clipboardText;
        content = L"";
    }

    // Trim the first line
    firstLine.erase(0, firstLine.find_first_not_of(L" \t\r\n"));
    firstLine.erase(firstLine.find_last_not_of(L" \t\r\n") + 1);

    std::wstring filename;
    bool format_detected = false;
    size_t filename_end_pos = 0;

    // Priority 1: Use pre-compiled regex patterns from config (if content creation is enabled)
    if (contentEnabled) {
        std::lock_guard<std::mutex> lock(g_extensionsMutex);
        for (const auto& compiledRegex : g_compiledRegexes) {
            try {
                std::wsmatch match;
                if (std::regex_match(firstLine, match, compiledRegex) && match.size() > 1) {
                    filename = match[1].str();
                    format_detected = true;
                    filename_end_pos = first_line_end != std::wstring::npos ? first_line_end + 1 : clipboardText.length();
                    break;
                }
            }
            catch (const std::regex_error&) {
                continue; // Silently ignore runtime regex errors.
            }
        }
    }

    // Priority 2: Check if first word is a filename with content following (single-line only)
    if (!format_detected && !isMultiLine) {
        std::wstringstream ss(firstLine);
        std::wstring firstWord;
        if (ss >> firstWord) {
            // Check if first word looks like a filename
            wchar_t ext[_MAX_EXT];
            _wsplitpath_s(firstWord.c_str(), NULL, 0, NULL, 0, NULL, 0, ext, _MAX_EXT);
            std::wstring extension(ext);
            std::transform(extension.begin(), extension.end(), extension.begin(), ::towlower);

            bool isAllowedExtension = false;
            {
                std::lock_guard<std::mutex> lock(g_extensionsMutex);
                for (const auto& allowedExt : g_settings.allowedExtensions) {
                    if (extension == allowedExt) {
                        isAllowedExtension = true;
                        break;
                    }
                }
            }

            if (isAllowedExtension) {
                // Extract content after the filename
                size_t firstWordEnd = firstLine.find(firstWord) + firstWord.length();
                if (firstWordEnd < firstLine.length()) {
                    // There's content after the filename
                    content = firstLine.substr(firstWordEnd);
                    content.erase(0, content.find_first_not_of(L" \t")); // Trim leading whitespace
                    filename = firstWord;
                    format_detected = true;
                    filename_end_pos = firstWordEnd;

                    // For this case, we need content creation enabled since we found content
                    if (!contentEnabled) {
                        return false;
                    }
                }
                else {
                    // Just the filename, no content - treat as empty file case
                    filename = firstWord;
                    content = L"";
                    format_detected = true;
                    filename_end_pos = firstWordEnd;

                    // For empty file, we need empty file creation enabled
                    if (!emptyEnabled) {
                        return false;
                    }
                }
            }
        }
    }

    // Priority 3: Fallback to the simpler word-count heuristic (for both modes)
    if (!format_detected) {
        wchar_t ext[_MAX_EXT];
        _wsplitpath_s(firstLine.c_str(), NULL, 0, NULL, 0, NULL, 0, ext, _MAX_EXT);
        std::wstring extension(ext);
        std::transform(extension.begin(), extension.end(), extension.begin(), ::towlower);

        bool isAllowedExtension = false;
        int wordCountLimit;

        {
            std::lock_guard<std::mutex> lock(g_extensionsMutex);
            wordCountLimit = g_settings.heuristicWordCountLimit;
            for (const auto& allowedExt : g_settings.allowedExtensions) {
                if (extension == allowedExt) {
                    isAllowedExtension = true;
                    break;
                }
            }
        }

        if (isAllowedExtension && CountWords(firstLine) <= wordCountLimit) {
            filename = firstLine;
            format_detected = true;
            filename_end_pos = first_line_end != std::wstring::npos ? first_line_end + 1 : clipboardText.length();

            // Priority 3 creates empty files, so check if empty file creation is enabled
            if (!emptyEnabled) {
                return false;
            }
        }
    }

    // If we found a filename, check if there are more filenames following it
    if (format_detected && emptyEnabled) {
        std::vector<std::wstring> allFilenames;
        allFilenames.push_back(filename);

        // Look for additional filenames using smart line-based logic
        std::vector<std::wstring> additionalFilenames = FindAdditionalFilenames(clipboardText, filename_end_pos);
        allFilenames.insert(allFilenames.end(), additionalFilenames.begin(), additionalFilenames.end());

        // If we found multiple filenames, handle as batch creation
        if (allFilenames.size() >= 2) {
            std::wstring explorerPath = GetSingleExplorerPath();
            if (explorerPath.empty()) {
                ShowToastNotification(g_hMainWnd, L"Error", L"No File Explorer window found.", NIIF_ERROR);
                return false;
            }

            // Separate files into existing and new
            std::vector<std::wstring> newFiles;
            std::vector<std::wstring> existingFiles;

            for (const auto& fname : allFilenames) {
                std::wstring fullPath = explorerPath + L"\\" + fname;
                if (GetFileAttributesW(fullPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
                    existingFiles.push_back(fname);
                }
                else {
                    newFiles.push_back(fname);
                }
            }

            // Handle existing files if any
            FileConflictAction conflictAction = FileConflictAction::Skip;
            if (!existingFiles.empty()) {
                std::wstring conflictMessage = L"The following files already exist:\n\n";
                for (size_t i = 0; i < existingFiles.size() && i < 10; ++i) {
                    conflictMessage += existingFiles[i] + L"\n";
                }
                if (existingFiles.size() > 10) {
                    conflictMessage += L"... and " + std::to_wstring(existingFiles.size() - 10) + L" more\n";
                }
                conflictMessage += L"\nChoose action for ALL existing files:\n\n";
                conflictMessage += L"Yes = Replace all existing files\n";
                conflictMessage += L"No = Skip all existing files\n";
                conflictMessage += L"Cancel = Rename all existing files";

                int result = MessageBoxW(NULL, conflictMessage.c_str(), L"Multiple File Conflicts",
                    MB_YESNOCANCEL | MB_ICONWARNING | MB_DEFBUTTON2);

                switch (result) {
                case IDYES: conflictAction = FileConflictAction::Replace; break;
                case IDNO: conflictAction = FileConflictAction::Skip; break;
                case IDCANCEL: conflictAction = FileConflictAction::Rename; break;
                default: conflictAction = FileConflictAction::Skip; break;
                }
            }

            // Create all files
            int successCount = 0;
            int skipCount = 0;
            std::vector<std::wstring> failedFiles;

            // Create new files first
            for (const auto& fname : newFiles) {
                std::wstring fullPath = explorerPath + L"\\" + fname;
                HANDLE hFile = CreateFileW(fullPath.c_str(), GENERIC_WRITE, 0, NULL,
                    CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
                if (hFile != INVALID_HANDLE_VALUE) {
                    CloseHandle(hFile);
                    successCount++;
                }
                else {
                    failedFiles.push_back(fname);
                }
            }

            // Handle existing files based on user choice
            for (const auto& fname : existingFiles) {
                if (conflictAction == FileConflictAction::Skip) {
                    skipCount++;
                    continue;
                }

                std::wstring fullPath = explorerPath + L"\\" + fname;
                std::wstring finalPath = fullPath;

                if (conflictAction == FileConflictAction::Rename) {
                    finalPath = GenerateUniqueFilename(fullPath);
                    HANDLE hFile = CreateFileW(finalPath.c_str(), GENERIC_WRITE, 0, NULL,
                        CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
                    if (hFile != INVALID_HANDLE_VALUE) {
                        CloseHandle(hFile);
                        successCount++;
                    }
                    else {
                        failedFiles.push_back(fname);
                    }
                }
                else if (conflictAction == FileConflictAction::Replace) {
                    if (CreateEmptyFileAtomic(finalPath)) {
                        successCount++;
                    }
                    else {
                        failedFiles.push_back(fname);
                    }
                }
            }

            // Show results to user
            std::wstring resultMessage;
            if (successCount > 0) {
                resultMessage = L"Successfully created " + std::to_wstring(successCount) + L" files";
                if (skipCount > 0) {
                    resultMessage += L", skipped " + std::to_wstring(skipCount) + L" existing files";
                }
                if (!failedFiles.empty()) {
                    resultMessage += L", failed to create " + std::to_wstring(failedFiles.size()) + L" files";
                }
                ShowToastNotification(g_hMainWnd, L"Multiple Files Created", resultMessage, NIIF_INFO);
            }
            else {
                resultMessage = L"No files were created";
                if (skipCount > 0) {
                    resultMessage += L" (" + std::to_wstring(skipCount) + L" files were skipped)";
                }
                if (!failedFiles.empty()) {
                    resultMessage += L" (" + std::to_wstring(failedFiles.size()) + L" files failed)";
                }
                ShowToastNotification(g_hMainWnd, L"File Creation", resultMessage, NIIF_WARNING);
            }

            return successCount > 0;
        }
    }

    // If not multiple files or only found one filename, proceed with original single file logic
    if (format_detected) {
        filename.erase(0, filename.find_first_not_of(L" \t\n\r"));
        filename.erase(filename.find_last_not_of(L" \t\n\r") + 1);
        if (!IsValidFilename(filename)) {
            return true; // Detected a pattern but filename is invalid. Stop all further processing.
        }

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

            if (content.empty()) {
                // Create empty file (single-line content or multi-line with empty content)
                if (GetFileAttributesW(finalPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
                    success = CreateEmptyFileAtomic(finalPath);
                }
                else {
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
            else {
                // Create file with content (multi-line content)
                if (GetFileAttributesW(finalPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
                    success = CreateFileWithContentAtomic(finalPath, content);
                }
                else {
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
                }
            }

            return success;
        }
    }
    return false;
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
    if (!IsClipboardFormatAvailable(CF_UNICODETEXT) || !OpenClipboard(g_hMainWnd)) return;
    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (hData == NULL) { CloseClipboard(); return; }

    wchar_t* pszText = static_cast<wchar_t*>(GlobalLock(hData));
    if (pszText == NULL) { CloseClipboard(); return; }
    std::wstring clipboardText(pszText);
    GlobalUnlock(hData);
    CloseClipboard();

    // Try directory structure creation first
    if (TryDirectoryStructureCreation(clipboardText)) {
        return;
    }

    // Fall back to file generation
    TryFileGeneration(clipboardText);
}

// Uses COM to find and return the path of a single open File Explorer window.
std::wstring GetSingleExplorerPath()
{
    std::vector<std::wstring> paths;
    IShellWindows* pShellWindows = NULL;

    // Check if COM is initialized, fallback to per-call initialization if not
    if (!g_bComInitialized) {
        if (SUCCEEDED(CoInitialize(NULL))) {
            g_bComInitialized = true; // Set flag so future calls know COM is ready
            // Don't call CoUninitialize() - leave COM initialized for future calls!
        }
        else {
            return L""; // COM initialization failed
        }
    }

    HRESULT hr = CoCreateInstance(CLSID_ShellWindows, NULL, CLSCTX_ALL, IID_IShellWindows, (void**)&pShellWindows);
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

    if (paths.size() >= 1) return paths[0];
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

        UINT dirFlags = g_settings.isCreateDirectoryStructureEnabled ? MF_STRING | MF_CHECKED : MF_STRING | MF_UNCHECKED;
        InsertMenu(hMenu, 2, MF_BYPOSITION | dirFlags, ID_MENU_TOGGLE_DIRECTORY, L"Create Directory Structure");

        InsertMenu(hMenu, 3, MF_SEPARATOR, 0, NULL);

        UINT startupFlags = IsStartupEnabled() ? MF_STRING | MF_CHECKED : MF_STRING | MF_UNCHECKED;
        InsertMenu(hMenu, 4, MF_BYPOSITION | startupFlags, ID_MENU_START_WITH_WINDOWS, L"Start with Windows");
        InsertMenu(hMenu, 5, MF_BYPOSITION | MF_STRING, ID_MENU_EDIT_CONFIG, L"Edit Config...");
        InsertMenu(hMenu, 6, MF_SEPARATOR, 0, NULL);
        InsertMenu(hMenu, 7, MF_BYPOSITION | MF_STRING, ID_MENU_EXIT, L"Exit");

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

//------------------------------------------------------------------------------------------------//
//                          FILE/DIRECTORY SECURITY & PATH VALIDATION                             //
//------------------------------------------------------------------------------------------------//

bool IsPathSafe(const std::wstring& path) {
    // Check for path traversal
    if (path.find(L"..\\") != std::wstring::npos || path.find(L"../") != std::wstring::npos) {
        return false;
    }

    // Check for absolute paths
    if (path.length() >= 2 && path[1] == L':') return false;
    if (path[0] == L'\\' || path[0] == L'/') return false;

    // Check for UNC paths
    if (path.length() >= 2 && path[0] == L'\\' && path[1] == L'\\') return false;

    return true;
}

void GetTreeSummary(const TreeNode* node, int& dirCount, int& fileCount) {
    if (!node) return;

    if (node->isDirectory && node->name != L"root") {
        dirCount++;
    }
    else if (!node->isDirectory) {
        fileCount++;
    }

    for (const auto& child : node->children) {
        GetTreeSummary(child.get(), dirCount, fileCount);
    }
}

// Comprehensive filename validation to prevent security issues and filesystem errors
bool IsValidFilename(const std::wstring& filename)
{
    // Check for empty filename
    if (filename.empty()) {
        return false;
    }

    // Check filename length (Windows has a 255 character limit for filenames)
    if (filename.length() > 255) {
        return false;
    }

    // Check for path traversal attempts
    if (filename.find(L"../") != std::wstring::npos || filename.find(L"..\\") != std::wstring::npos) {
        return false;
    }

    // Check for absolute paths (should only be relative filenames)
    if (filename.length() >= 2 && filename[1] == L':') { // Drive letter (C:, D:, etc.)
        return false;
    }
    if (filename[0] == L'\\' || filename[0] == L'/') { // UNC paths or root paths
        return false;
    }

    // Check for invalid characters in Windows filenames
    const wchar_t* invalidChars = L"\\/:*?\"<>|";
    if (filename.find_first_of(invalidChars) != std::wstring::npos) {
        return false;
    }

    // Check for control characters (0x00-0x1F)
    for (wchar_t c : filename) {
        if (c >= 0x00 && c <= 0x1F) {
            return false;
        }
    }

    // Check for reserved Windows device names (case-insensitive)
    std::wstring upperFilename = filename;
    std::transform(upperFilename.begin(), upperFilename.end(), upperFilename.begin(), ::towupper);

    // Extract base name without extension for reserved name checking
    std::wstring baseName = upperFilename;
    size_t dotPos = baseName.find_last_of(L'.');
    if (dotPos != std::wstring::npos) {
        baseName = baseName.substr(0, dotPos);
    }

    // Check basic reserved device names
    const std::vector<std::wstring> basicReservedNames = {
        L"CON", L"PRN", L"AUX", L"NUL"
    };

    for (const auto& reserved : basicReservedNames) {
        if (baseName == reserved) {
            return false;
        }
    }

    // Check COM ports (COMx where x is any number)
    if (baseName.length() >= 4 && baseName.substr(0, 3) == L"COM") {
        std::wstring numberPart = baseName.substr(3);
        if (!numberPart.empty() && std::all_of(numberPart.begin(), numberPart.end(), ::iswdigit)) {
            return false;
        }
    }

    // Check LPT ports (LPTx where x is any number)
    if (baseName.length() >= 4 && baseName.substr(0, 3) == L"LPT") {
        std::wstring numberPart = baseName.substr(3);
        if (!numberPart.empty() && std::all_of(numberPart.begin(), numberPart.end(), ::iswdigit)) {
            return false;
        }
    }

    // Check for filenames ending with period (not allowed in Windows)
    // Note: Leading/trailing spaces are handled by trimming before this function is called
    if (filename.back() == L'.') {
        return false;
    }

    // Additional check: ensure filename doesn't contain only dots
    bool onlyDots = true;
    for (wchar_t c : filename) {
        if (c != L'.') {
            onlyDots = false;
            break;
        }
    }
    if (onlyDots) {
        return false;
    }

    return true;
}

// Smart search for additional filenames using line logic
std::vector<std::wstring> FindAdditionalFilenames(const std::wstring& text, size_t startPos) {
    std::vector<std::wstring> filenames;

    // Split remaining text into lines
    std::vector<std::wstring> lines;
    std::wstringstream ss(text.substr(startPos));
    std::wstring line;

    while (std::getline(ss, line)) {
        // Trim whitespace from line
        line.erase(0, line.find_first_not_of(L" \t\r"));
        line.erase(line.find_last_not_of(L" \t\r") + 1);
        lines.push_back(line);
    }

    if (lines.empty()) return filenames;

    // Check first line for multiple space-separated filenames
    std::wstringstream firstLineStream(lines[0]);
    std::wstring word;
    int wordsInFirstLine = 0;
    std::vector<std::wstring> firstLineFilenames;

    while (firstLineStream >> word) {
        wordsInFirstLine++;
        if (IsValidFilename(word)) {
            // Check if it has a valid extension
            wchar_t ext[_MAX_EXT];
            _wsplitpath_s(word.c_str(), NULL, 0, NULL, 0, NULL, 0, ext, _MAX_EXT);
            std::wstring extension(ext);
            std::transform(extension.begin(), extension.end(), extension.begin(), ::towlower);

            bool isAllowedExtension = false;
            int wordCountLimit;
            {
                std::lock_guard<std::mutex> lock(g_extensionsMutex);
                wordCountLimit = g_settings.heuristicWordCountLimit;
                for (const auto& allowedExt : g_settings.allowedExtensions) {
                    if (extension == allowedExt) {
                        isAllowedExtension = true;
                        break;
                    }
                }
            }

            if (isAllowedExtension && CountWords(word) <= wordCountLimit) {
                firstLineFilenames.push_back(word);
            }
        }
    }

    // If we found multiple space-separated filenames in first line, return those
    if (firstLineFilenames.size() > 1) {
        return firstLineFilenames;
    }

    // If first line had exactly one valid filename, add it and continue checking other lines
    if (firstLineFilenames.size() == 1) {
        filenames.push_back(firstLineFilenames[0]);
    }

    // Check subsequent lines one by one
    for (size_t i = 1; i < lines.size(); ++i) {
        if (lines[i].empty()) {
            // Empty line - skip and continue checking
            continue;
        }
        else {
            // Line has content - check if it's a valid filename
            if (IsValidFilename(lines[i])) {
                // Check if it has a valid extension
                wchar_t ext[_MAX_EXT];
                _wsplitpath_s(lines[i].c_str(), NULL, 0, NULL, 0, NULL, 0, ext, _MAX_EXT);
                std::wstring extension(ext);
                std::transform(extension.begin(), extension.end(), extension.begin(), ::towlower);

                bool isAllowedExtension = false;
                int wordCountLimit;
                {
                    std::lock_guard<std::mutex> lock(g_extensionsMutex);
                    wordCountLimit = g_settings.heuristicWordCountLimit;
                    for (const auto& allowedExt : g_settings.allowedExtensions) {
                        if (extension == allowedExt) {
                            isAllowedExtension = true;
                            break;
                        }
                    }
                }

                if (isAllowedExtension && CountWords(lines[i]) <= wordCountLimit) {
                    filenames.push_back(lines[i]);
                    // Continue checking next lines
                }
                else {
                    // Not a valid filename - stop searching
                    break;
                }
            }
            else {
                // Line has content but isn't a valid filename - stop searching
                break;
            }
        }
    }

    return filenames;
}