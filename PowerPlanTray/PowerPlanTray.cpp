// PowerPlanTray.cpp: Tray app to switch Windows power plans.

#include "framework.h"
#include "PowerPlanTray.h"

#include <shellapi.h>
#include <strsafe.h>
#include <vector>
#include <string>

#include <powrprof.h>
#pragma comment(lib, "PowrProf.lib")

#define WM_TRAYICON (WM_APP + 1)
#define TRAY_ID 1
#define ID_BASE_PLAN 10000
#define IDM_STARTUP 40001
#define IDM_REFRESH 40002
// AFK feature command IDs
#define IDM_AFK_OFF           40100
#define IDM_AFK_INTERVAL_BASE 40200 // 6 entries: 5,10,15,30,45,60
#define IDM_AFK_TARGET_BASE   40300 // dynamic per plan list
// Timer events
#define TIMER_EVENT_POLL_ACTIVE 1
#define TIMER_EVENT_AFK_CHECK 2

HINSTANCE g_hInst = nullptr;
HWND g_hWnd = nullptr;
UINT g_uTaskbarCreated = 0;
HPOWERNOTIFY g_hPowerNotify = nullptr;
GUID g_lastActiveGuid{};
HICON g_hTrayIcon = nullptr;
HANDLE g_hInstanceMutex = nullptr;
// AFK feature globals
int g_afkTimeoutMinutes = 0; // 0 = Off
GUID g_afkTargetGuid{};      // Target plan when AFK
GUID g_afkPrevGuid{};        // Plan before AFK switch
bool g_afkApplied = false;   // Whether AFK plan is currently applied

struct PlanItem {
    GUID guid;
    std::wstring name;
};

static const wchar_t* kClassName = L"PowerPlanTrayHiddenWindow";

ATOM RegisterTrayWindowClass(HINSTANCE hInstance);
BOOL CreateHiddenWindow(HINSTANCE hInstance);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

std::vector<PlanItem> EnumeratePlans();
bool GetActivePlanGuid(GUID& outGuid);
bool SetActivePlan(const GUID& guid);
void ShowTrayMenu(HWND hWnd);
void AddOrUpdateTrayIcon(HWND hWnd);
void RemoveTrayIcon(HWND hWnd);
void UpdateTrayTooltip(HWND hWnd);
void EnableDpiAwareness();
bool IsStartupEnabled();
bool SetStartupEnabled(bool enable);
// AFK helpers
void AfkLoadSettings();
void AfkSaveSettings();
void AfkCheckTick(HWND hWnd);
DWORD GetIdleSeconds();
static std::wstring LoadResString(UINT id)
{
    wchar_t buf[256] = {};
    int n = LoadStringW(g_hInst, id, buf, ARRAYSIZE(buf));
    if (n <= 0) return L"";
    return std::wstring(buf, n);
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE /*hPrevInstance*/,
                     _In_ LPWSTR    /*lpCmdLine*/,
                     _In_ int       /*nCmdShow*/)
{
    g_hInst = hInstance;
    EnableDpiAwareness();
    g_uTaskbarCreated = RegisterWindowMessage(L"TaskbarCreated");

    // Single instance mutex
    g_hInstanceMutex = CreateMutexW(nullptr, TRUE, L"Local\\PowerPlanTray_SingleInstance");
    if (g_hInstanceMutex && GetLastError() == ERROR_ALREADY_EXISTS)
    {
        std::wstring title = LoadResString(IDS_MSG_ALREADY_RUNNING_TITLE);
        std::wstring text  = LoadResString(IDS_MSG_ALREADY_RUNNING_TEXT);
        if (title.empty()) title = L"PowerPlanTray";
        if (text.empty())  text  = L"PowerPlanTray is already running.";
        MessageBoxW(nullptr, text.c_str(), title.c_str(), MB_OK | MB_ICONINFORMATION);
        // Release initial ownership and close handle
        ReleaseMutex(g_hInstanceMutex);
        CloseHandle(g_hInstanceMutex);
        g_hInstanceMutex = nullptr;
        return 0;
    }

    if (!RegisterTrayWindowClass(hInstance))
        return 0;
    if (!CreateHiddenWindow(hInstance))
        return 0;

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return (int)msg.wParam;
}

ATOM RegisterTrayWindowClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex = {};
    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_POWERPLANTRAY));
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcex.lpszClassName = kClassName;
    wcex.hIconSm = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SMALL));
    return RegisterClassExW(&wcex);
}

BOOL CreateHiddenWindow(HINSTANCE hInstance)
{
    g_hWnd = CreateWindowW(kClassName, L"PowerPlanTray", WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);
    if (!g_hWnd)
        return FALSE;

    // Do not show any window; use only tray icon
    AddOrUpdateTrayIcon(g_hWnd);
    UpdateTrayTooltip(g_hWnd);

    // Initialize last known scheme
    GetActivePlanGuid(g_lastActiveGuid);

    // Subscribe to power setting change for personality changes
    g_hPowerNotify = RegisterPowerSettingNotification(g_hWnd, &GUID_POWERSCHEME_PERSONALITY, DEVICE_NOTIFY_WINDOW_HANDLE);

    // Fallback: poll for active plan changes (covers custom plans with same personality)
    SetTimer(g_hWnd, TIMER_EVENT_POLL_ACTIVE, 2000, nullptr);

    // Load AFK settings and start AFK timer
    AfkLoadSettings();
    // If no target set yet, default to current active plan to avoid surprise switches
    if (IsEqualGUID(g_afkTargetGuid, GUID{}))
    {
        GUID cur{}; if (GetActivePlanGuid(cur)) g_afkTargetGuid = cur;
    }
    SetTimer(g_hWnd, TIMER_EVENT_AFK_CHECK, 1000, nullptr); // AFK checker, 1s cadence

    return TRUE;
}

static UINT GetWindowDpi(HWND hWnd)
{
    HMODULE hUser32 = GetModuleHandleW(L"user32.dll");
    if (hUser32)
    {
        typedef UINT (WINAPI *GetDpiForWindowFn)(HWND);
        auto p = reinterpret_cast<GetDpiForWindowFn>(GetProcAddress(hUser32, "GetDpiForWindow"));
        if (p) return p(hWnd);
    }
    return 96; // fallback
}

static HICON CreateTrayIconForDpi(HWND hWnd)
{
    UINT dpi = GetWindowDpi(hWnd);
    int cx = GetSystemMetrics(SM_CXSMICON);
    int cy = GetSystemMetrics(SM_CYSMICON);

    HMODULE hUser32 = GetModuleHandleW(L"user32.dll");
    if (hUser32)
    {
        typedef int (WINAPI *GetSystemMetricsForDpiFn)(int, UINT);
        auto p = reinterpret_cast<GetSystemMetricsForDpiFn>(GetProcAddress(hUser32, "GetSystemMetricsForDpi"));
        if (p)
        {
            cx = p(SM_CXSMICON, dpi);
            cy = p(SM_CYSMICON, dpi);
        }
    }
    // Use the main app icon resource and scale to desired size
    return (HICON)LoadImageW(g_hInst, MAKEINTRESOURCE(IDI_POWERPLANTRAY), IMAGE_ICON, cx, cy, LR_DEFAULTCOLOR);
}

void AddOrUpdateTrayIcon(HWND hWnd)
{
    NOTIFYICONDATA nid = {};
    nid.cbSize = sizeof(nid);
    nid.hWnd = hWnd;
    nid.uID = TRAY_ID;
    nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
    nid.uCallbackMessage = WM_TRAYICON;
    if (g_hTrayIcon) { DestroyIcon(g_hTrayIcon); g_hTrayIcon = nullptr; }
    g_hTrayIcon = CreateTrayIconForDpi(hWnd);
    nid.hIcon = g_hTrayIcon ? g_hTrayIcon : LoadIcon(g_hInst, MAKEINTRESOURCE(IDI_SMALL));
    auto tip = LoadResString(IDS_TRAY_TOOLTIP_DEFAULT);
    StringCchCopy(nid.szTip, ARRAYSIZE(nid.szTip), tip.c_str());
    Shell_NotifyIcon(NIM_ADD, &nid);

    // Opt into modern behavior and DPI handling for tray icons
    nid.uVersion = NOTIFYICON_VERSION_4;
    Shell_NotifyIcon(NIM_SETVERSION, &nid);
}

void RemoveTrayIcon(HWND hWnd)
{
    NOTIFYICONDATA nid = {};
    nid.cbSize = sizeof(nid);
    nid.hWnd = hWnd;
    nid.uID = TRAY_ID;
    Shell_NotifyIcon(NIM_DELETE, &nid);
    if (g_hTrayIcon) { DestroyIcon(g_hTrayIcon); g_hTrayIcon = nullptr; }
}

void UpdateTrayTooltip(HWND hWnd)
{
    GUID active;
    std::wstring tip = LoadResString(IDS_TRAY_TOOLTIP_DEFAULT);
    if (GetActivePlanGuid(active))
    {
        auto items = EnumeratePlans();
        for (const auto& it : items)
        {
            if (IsEqualGUID(it.guid, active)) { tip = it.name; break; }
        }
    }

    NOTIFYICONDATA nid = {};
    nid.cbSize = sizeof(nid);
    nid.hWnd = hWnd;
    nid.uID = TRAY_ID;
    nid.uFlags = NIF_TIP;
    StringCchCopy(nid.szTip, ARRAYSIZE(nid.szTip), tip.c_str());
    Shell_NotifyIcon(NIM_MODIFY, &nid);
}

void ShowTrayMenu(HWND hWnd)
{
    auto plans = EnumeratePlans();
    GUID active{};
    GetActivePlanGuid(active);

    HMENU hMenu = CreatePopupMenu();
    // 1) Power plans first
    UINT id = ID_BASE_PLAN;
    for (const auto& p : plans)
    {
        UINT flags = MF_STRING | MF_ENABLED;
        if (IsEqualGUID(p.guid, active))
            flags |= MF_CHECKED;
        AppendMenu(hMenu, flags, id++, p.name.c_str());
    }
    // Separator after plans
    AppendMenu(hMenu, MF_SEPARATOR, 0, nullptr);

    // AFK submenu
    HMENU hAfk = CreatePopupMenu();
    auto sAfk = LoadResString(IDS_MENU_AFK);
    auto sAfkOff = LoadResString(IDS_MENU_AFK_OFF);
    auto sAfkTimeout = LoadResString(IDS_MENU_AFK_TIMEOUT);
    auto sAfkTarget = LoadResString(IDS_MENU_AFK_TARGET);

    // Timeout options
    HMENU hAfkTimeout = CreatePopupMenu();
    // "Off" inside the timeout submenu
    AppendMenu(hAfkTimeout, MF_STRING | (g_afkTimeoutMinutes == 0 ? MF_CHECKED : 0), IDM_AFK_OFF, sAfkOff.c_str());
    const int intervals[7] = { 1, 5, 10, 15, 30, 45, 60 };
    const UINT intervalStrings[7] = {
        IDS_MENU_AFK_1MIN,
        IDS_MENU_AFK_5MIN,
        IDS_MENU_AFK_10MIN,
        IDS_MENU_AFK_15MIN,
        IDS_MENU_AFK_30MIN,
        IDS_MENU_AFK_45MIN,
        IDS_MENU_AFK_60MIN
    };
    for (int i = 0; i < 7; ++i)
    {
        auto s = LoadResString(intervalStrings[i]);
        UINT flags = MF_STRING;
        if (g_afkTimeoutMinutes == intervals[i]) flags |= MF_CHECKED;
        AppendMenu(hAfkTimeout, flags, IDM_AFK_INTERVAL_BASE + i, s.c_str());
    }
    AppendMenu(hAfk, MF_POPUP, (UINT_PTR)hAfkTimeout, sAfkTimeout.c_str());

    // Target plan submenu
    HMENU hAfkTarget = CreatePopupMenu();
    UINT tid = IDM_AFK_TARGET_BASE;
    for (const auto& p : plans)
    {
        UINT flags = MF_STRING | MF_ENABLED;
        if (IsEqualGUID(p.guid, g_afkTargetGuid))
            flags |= MF_CHECKED;
        AppendMenu(hAfkTarget, flags, tid++, p.name.c_str());
    }
    AppendMenu(hAfk, MF_POPUP, (UINT_PTR)hAfkTarget, sAfkTarget.c_str());

    AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hAfk, sAfk.c_str());

    // 2) Other options follow
    auto sRefresh = LoadResString(IDS_MENU_REFRESH);
    AppendMenu(hMenu, MF_STRING, IDM_REFRESH, sRefresh.c_str());
    bool startup = IsStartupEnabled();
    auto sStartup = LoadResString(IDS_MENU_STARTUP);
    AppendMenu(hMenu, MF_STRING | (startup ? MF_CHECKED : 0), IDM_STARTUP, sStartup.c_str());

    // 3) Exit at the end
    AppendMenu(hMenu, MF_SEPARATOR, 0, nullptr);
    auto sExit = LoadResString(IDS_MENU_EXIT);
    AppendMenu(hMenu, MF_STRING, IDM_EXIT, sExit.c_str());

    POINT pt; GetCursorPos(&pt);
    SetForegroundWindow(hWnd);
    TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_BOTTOMALIGN, pt.x, pt.y, 0, hWnd, nullptr);
    DestroyMenu(hMenu);
}

std::vector<PlanItem> EnumeratePlans()
{
    std::vector<PlanItem> result;
    DWORD index = 0;
    for (;; ++index)
    {
        GUID guid{};
        DWORD size = sizeof(GUID);
        DWORD status = PowerEnumerate(nullptr, nullptr, nullptr, ACCESS_SCHEME, index, reinterpret_cast<UCHAR*>(&guid), &size);
        if (status != ERROR_SUCCESS)
            break;

        DWORD nameSize = 0;
        if (PowerReadFriendlyName(nullptr, &guid, nullptr, nullptr, nullptr, &nameSize) != ERROR_SUCCESS)
            continue;
        std::wstring name; name.resize(nameSize / sizeof(wchar_t));
        if (PowerReadFriendlyName(nullptr, &guid, nullptr, nullptr, reinterpret_cast<UCHAR*>(&name[0]), &nameSize) != ERROR_SUCCESS)
            continue;
        // Ensure null-termination trimming
        if (!name.empty() && name.back() == L'\0') name.pop_back();

        result.push_back({ guid, name });
    }
    return result;
}

bool GetActivePlanGuid(GUID& outGuid)
{
    GUID* pGuid = nullptr;
    if (PowerGetActiveScheme(nullptr, &pGuid) == ERROR_SUCCESS && pGuid)
    {
        outGuid = *pGuid;
        LocalFree(pGuid);
        return true;
    }
    return false;
}

bool SetActivePlan(const GUID& guid)
{
    return PowerSetActiveScheme(nullptr, &guid) == ERROR_SUCCESS;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    if (message == g_uTaskbarCreated)
    {
        AddOrUpdateTrayIcon(hWnd);
        UpdateTrayTooltip(hWnd);
        return 0;
    }

    switch (message)
    {
    case WM_CREATE:
        return 0;
    case WM_COMMAND:
    {
        const UINT cmd = LOWORD(wParam);
        if (cmd == IDM_EXIT)
        {
            DestroyWindow(hWnd);
            return 0;
        }
        if (cmd == IDM_REFRESH)
        {
            UpdateTrayTooltip(hWnd);
            return 0;
        }
        if (cmd == IDM_STARTUP)
        {
            bool now = IsStartupEnabled();
            SetStartupEnabled(!now);
            // No persistent menu to update; next open reflects new state
            return 0;
        }
        if (cmd == IDM_AFK_OFF)
        {
            // Disable AFK switching; if currently applied, revert now
            g_afkTimeoutMinutes = 0;
            if (g_afkApplied)
            {
                GUID cur{}; GetActivePlanGuid(cur);
                if (!IsEqualGUID(cur, g_afkPrevGuid) && !IsEqualGUID(g_afkPrevGuid, GUID{}))
                {
                    SetActivePlan(g_afkPrevGuid);
                    g_lastActiveGuid = g_afkPrevGuid;
                }
                g_afkApplied = false;
                UpdateTrayTooltip(hWnd);
            }
            AfkSaveSettings();
            return 0;
        }
        if (cmd >= IDM_AFK_INTERVAL_BASE && cmd < IDM_AFK_INTERVAL_BASE + 100)
        {
            const int intervals[7] = { 1, 5, 10, 15, 30, 45, 60 };
            int idx = (int)cmd - IDM_AFK_INTERVAL_BASE;
            if (idx >= 0 && idx < 7)
            {
                g_afkTimeoutMinutes = intervals[idx];
                AfkSaveSettings();
            }
            return 0;
        }
        if (cmd >= IDM_AFK_TARGET_BASE && cmd < IDM_AFK_TARGET_BASE + 10000)
        {
            UINT index = cmd - IDM_AFK_TARGET_BASE;
            auto plans = EnumeratePlans();
            if (index < plans.size())
            {
                g_afkTargetGuid = plans[index].guid;
                AfkSaveSettings();
            }
            return 0;
        }
        if (cmd >= ID_BASE_PLAN && cmd < ID_BASE_PLAN + 10000)
        {
            UINT index = cmd - ID_BASE_PLAN;
            auto plans = EnumeratePlans();
            if (index < plans.size())
            {
                SetActivePlan(plans[index].guid);
                UpdateTrayTooltip(hWnd);
            }
            return 0;
        }
        break;
    }
    case WM_TRAYICON:
        if (LOWORD(lParam) == WM_RBUTTONUP || LOWORD(lParam) == WM_CONTEXTMENU)
        {
            ShowTrayMenu(hWnd);
            return 0;
        }
        break;
    case WM_POWERBROADCAST:
        if (wParam == PBT_POWERSETTINGCHANGE)
        {
            // Power scheme likely changed; refresh tooltip
            UpdateTrayTooltip(hWnd);
            return TRUE;
        }
        break;
    case WM_DPICHANGED:
        // Recreate or refresh tray icon to ensure crisp rendering
        RemoveTrayIcon(hWnd);
        AddOrUpdateTrayIcon(hWnd);
        UpdateTrayTooltip(hWnd);
        return 0;
    case WM_TIMER:
        if (wParam == TIMER_EVENT_POLL_ACTIVE)
        {
            GUID now{};
            if (GetActivePlanGuid(now) && !IsEqualGUID(now, g_lastActiveGuid))
            {
                g_lastActiveGuid = now;
                UpdateTrayTooltip(hWnd);
            }
            return 0;
        }
        else if (wParam == TIMER_EVENT_AFK_CHECK)
        {
            AfkCheckTick(hWnd);
            return 0;
        }
        break;
    case WM_DESTROY:
        if (g_hPowerNotify)
        {
            UnregisterPowerSettingNotification(g_hPowerNotify);
            g_hPowerNotify = nullptr;
        }
        KillTimer(hWnd, 1);
        KillTimer(hWnd, 2);
        RemoveTrayIcon(hWnd);
        if (g_hInstanceMutex)
        {
            ReleaseMutex(g_hInstanceMutex);
            CloseHandle(g_hInstanceMutex);
            g_hInstanceMutex = nullptr;
        }
        PostQuitMessage(0);
        return 0;
    default:
        break;
    }
    return DefWindowProc(hWnd, message, wParam, lParam);
}

bool IsStartupEnabled()
{
    HKEY hKey;
    const wchar_t* subkey = L"Software\\Microsoft\\Windows\\CurrentVersion\\Run";
    if (RegOpenKeyExW(HKEY_CURRENT_USER, subkey, 0, KEY_READ, &hKey) != ERROR_SUCCESS)
        return false;
    DWORD type = 0; DWORD size = 0;
    LONG rc = RegGetValueW(hKey, nullptr, L"PowerPlanTray", RRF_RT_REG_SZ, &type, nullptr, &size);
    RegCloseKey(hKey);
    return rc == ERROR_SUCCESS;
}

bool SetStartupEnabled(bool enable)
{
    const wchar_t* subkey = L"Software\\Microsoft\\Windows\\CurrentVersion\\Run";
    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, subkey, 0, nullptr, 0, KEY_SET_VALUE | KEY_QUERY_VALUE, nullptr, &hKey, nullptr) != ERROR_SUCCESS)
        return false;

    bool ok = false;
    if (enable)
    {
        wchar_t path[MAX_PATH] = {};
        GetModuleFileNameW(nullptr, path, ARRAYSIZE(path));
        // Quote path to handle spaces
        std::wstring value = L"\""; value += path; value += L"\"";
        ok = (RegSetValueExW(hKey, L"PowerPlanTray", 0, REG_SZ, reinterpret_cast<const BYTE*>(value.c_str()), static_cast<DWORD>((value.size() + 1) * sizeof(wchar_t))) == ERROR_SUCCESS);
    }
    else
    {
        ok = (RegDeleteValueW(hKey, L"PowerPlanTray") == ERROR_SUCCESS) || (GetLastError() == ERROR_FILE_NOT_FOUND);
    }
    RegCloseKey(hKey);
    return ok;
}

void EnableDpiAwareness()
{
    // Try Per Monitor V2 if available
    HMODULE hUser32 = GetModuleHandleW(L"user32.dll");
    if (hUser32)
    {
        typedef BOOL (WINAPI *SetCtxFn)(HANDLE);
        auto pSetCtx = reinterpret_cast<SetCtxFn>(GetProcAddress(hUser32, "SetProcessDpiAwarenessContext"));
        // Negative HANDLE values correspond to predefined DPI contexts
        const HANDLE PER_MONITOR_V2 = (HANDLE)-4;
        const HANDLE PER_MONITOR = (HANDLE)-3;
        const HANDLE SYSTEM_AWARE = (HANDLE)-2;
        if (pSetCtx)
        {
            if (pSetCtx(PER_MONITOR_V2)) return;
            if (pSetCtx(PER_MONITOR)) return;
            pSetCtx(SYSTEM_AWARE);
            return;
        }
    }

    // Fallback to Shcore per-monitor DPI awareness on older systems
    HMODULE hShcore = LoadLibraryW(L"Shcore.dll");
    if (hShcore)
    {
        enum PROCESS_DPI_AWARENESS { PROCESS_DPI_UNAWARE = 0, PROCESS_SYSTEM_DPI_AWARE = 1, PROCESS_PER_MONITOR_DPI_AWARE = 2 };
        typedef HRESULT (WINAPI *SetAwFn)(PROCESS_DPI_AWARENESS);
        auto pSetAw = reinterpret_cast<SetAwFn>(GetProcAddress(hShcore, "SetProcessDpiAwareness"));
        if (pSetAw)
        {
            if (SUCCEEDED(pSetAw(PROCESS_PER_MONITOR_DPI_AWARE))) { FreeLibrary(hShcore); return; }
        }
        FreeLibrary(hShcore);
    }

    // Legacy system-DPI aware as a last resort
    if (hUser32)
    {
        typedef BOOL (WINAPI *SetLegacyFn)(void);
        auto pSetLegacy = reinterpret_cast<SetLegacyFn>(GetProcAddress(hUser32, "SetProcessDPIAware"));
        if (pSetLegacy) pSetLegacy();
    }
}

// ===== AFK helpers =====
static const wchar_t* kAfkRegPath = L"Software\\PowerPlanTray";

void AfkLoadSettings()
{
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, kAfkRegPath, 0, KEY_QUERY_VALUE, &hKey) == ERROR_SUCCESS)
    {
        DWORD dw = 0; DWORD size = sizeof(dw);
        if (RegGetValueW(hKey, nullptr, L"AfkTimeoutMinutes", RRF_RT_REG_DWORD, nullptr, &dw, &size) == ERROR_SUCCESS)
        {
            g_afkTimeoutMinutes = (int)dw;
        }
        GUID g{}; size = sizeof(g);
        if (RegGetValueW(hKey, nullptr, L"AfkTargetPlan", RRF_RT_REG_BINARY, nullptr, &g, &size) == ERROR_SUCCESS && size == sizeof(GUID))
        {
            g_afkTargetGuid = g;
        }
        RegCloseKey(hKey);
    }
}

void AfkSaveSettings()
{
    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, kAfkRegPath, 0, nullptr, 0, KEY_SET_VALUE, nullptr, &hKey, nullptr) == ERROR_SUCCESS)
    {
        DWORD dw = (DWORD)g_afkTimeoutMinutes;
        RegSetValueExW(hKey, L"AfkTimeoutMinutes", 0, REG_DWORD, reinterpret_cast<const BYTE*>(&dw), sizeof(dw));
        RegSetValueExW(hKey, L"AfkTargetPlan", 0, REG_BINARY, reinterpret_cast<const BYTE*>(&g_afkTargetGuid), sizeof(GUID));
        RegCloseKey(hKey);
    }
}

DWORD GetIdleSeconds()
{
    LASTINPUTINFO li{}; li.cbSize = sizeof(li);
    if (!GetLastInputInfo(&li)) return 0;
    ULONGLONG now = GetTickCount64();
    ULONGLONG then = (ULONGLONG)li.dwTime;
    if (now < then) return 0; // extremely unlikely with GetTickCount64
    ULONGLONG diff = now - then;
    return (DWORD)(diff / 1000ULL);
}

void AfkCheckTick(HWND hWnd)
{
    if (g_afkTimeoutMinutes <= 0)
        return; // feature disabled

    DWORD idle = GetIdleSeconds();
    const DWORD threshold = (DWORD)g_afkTimeoutMinutes * 60U;

    if (idle >= threshold)
    {
        if (!g_afkApplied)
        {
            GUID cur{}; if (GetActivePlanGuid(cur)) g_afkPrevGuid = cur; else g_afkPrevGuid = GUID{};
            if (!IsEqualGUID(cur, g_afkTargetGuid) && !IsEqualGUID(g_afkTargetGuid, GUID{}))
            {
                SetActivePlan(g_afkTargetGuid);
                g_lastActiveGuid = g_afkTargetGuid;
                UpdateTrayTooltip(hWnd);
            }
            g_afkApplied = true;
        }
    }
    else
    {
        if (g_afkApplied)
        {
            GUID cur{}; GetActivePlanGuid(cur);
            if (!IsEqualGUID(cur, g_afkPrevGuid) && !IsEqualGUID(g_afkPrevGuid, GUID{}))
            {
                SetActivePlan(g_afkPrevGuid);
                g_lastActiveGuid = g_afkPrevGuid;
                UpdateTrayTooltip(hWnd);
            }
            g_afkApplied = false;
        }
    }
}
