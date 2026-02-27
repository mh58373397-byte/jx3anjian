#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN
#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <mmsystem.h>
#include <shellapi.h>
#include <commctrl.h>
#include <string.h>
#include <stdio.h>
#include <stdlib.h>

#define INTERCEPTION_STATIC
#include "interception.h"
#include "resource.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "winmm.lib")
#pragma comment(lib, "comctl32.lib")

/* ================================================================== */
/*          DYNAMIC LOADING - interception.dll from resource           */
/* ================================================================== */

typedef InterceptionContext      (*PFN_create_context)(void);
typedef void                     (*PFN_destroy_context)(InterceptionContext);
typedef void                     (*PFN_set_filter)(InterceptionContext, InterceptionPredicate, InterceptionFilter);
typedef unsigned int             (*PFN_get_hardware_id)(InterceptionContext, InterceptionDevice, void*, unsigned int);
typedef InterceptionDevice       (*PFN_wait_with_timeout)(InterceptionContext, unsigned long);
typedef int                      (*PFN_receive)(InterceptionContext, InterceptionDevice, InterceptionStroke*, unsigned int);
typedef int                      (*PFN_send)(InterceptionContext, InterceptionDevice, const InterceptionStroke*, unsigned int);
typedef int                      (*PFN_is_keyboard)(InterceptionDevice);
typedef int                      (*PFN_is_mouse)(InterceptionDevice);

static PFN_create_context     pfn_create_context;
static PFN_destroy_context    pfn_destroy_context;
static PFN_set_filter         pfn_set_filter;
static PFN_get_hardware_id    pfn_get_hardware_id;
static PFN_wait_with_timeout  pfn_wait_with_timeout;
static PFN_receive            pfn_receive;
static PFN_send               pfn_send;
static PFN_is_keyboard        pfn_is_keyboard;
static PFN_is_mouse           pfn_is_mouse;

#define interception_create_context     pfn_create_context
#define interception_destroy_context    pfn_destroy_context
#define interception_set_filter         pfn_set_filter
#define interception_get_hardware_id    pfn_get_hardware_id
#define interception_wait_with_timeout  pfn_wait_with_timeout
#define interception_receive            pfn_receive
#define interception_send               pfn_send
#define interception_is_keyboard        pfn_is_keyboard
#define interception_is_mouse          pfn_is_mouse

static HMODULE g_dll = NULL;

/* ================================================================== */
/*                       RESOURCE EXTRACTION                           */
/* ================================================================== */

static BOOL extract_resource(int id, const WCHAR *destPath) {
    HRSRC hRes = FindResourceW(NULL, MAKEINTRESOURCEW(id), RT_RCDATA);
    if (!hRes) return FALSE;
    HGLOBAL hData = LoadResource(NULL, hRes);
    if (!hData) return FALSE;
    void *data = LockResource(hData);
    DWORD size = SizeofResource(NULL, hRes);
    if (!data || !size) return FALSE;
    HANDLE hFile = CreateFileW(destPath, GENERIC_WRITE, 0, NULL,
                               CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE) return FALSE;
    DWORD written;
    WriteFile(hFile, data, size, &written, NULL);
    CloseHandle(hFile);
    return written == size;
}

static WCHAR g_dll_path[MAX_PATH];

static BOOL load_interception_dll(void) {
    WCHAR tempDir[MAX_PATH];
    GetTempPathW(MAX_PATH, tempDir);
    wsprintfW(g_dll_path, L"%sinterception.dll", tempDir);
    extract_resource(IDR_INTERCEPTION_DLL, g_dll_path);
    g_dll = LoadLibraryW(g_dll_path);
    if (!g_dll) return FALSE;
    pfn_create_context    = (PFN_create_context)   GetProcAddress(g_dll, "interception_create_context");
    pfn_destroy_context   = (PFN_destroy_context)  GetProcAddress(g_dll, "interception_destroy_context");
    pfn_set_filter        = (PFN_set_filter)       GetProcAddress(g_dll, "interception_set_filter");
    pfn_get_hardware_id   = (PFN_get_hardware_id)  GetProcAddress(g_dll, "interception_get_hardware_id");
    pfn_wait_with_timeout = (PFN_wait_with_timeout)GetProcAddress(g_dll, "interception_wait_with_timeout");
    pfn_receive           = (PFN_receive)          GetProcAddress(g_dll, "interception_receive");
    pfn_send              = (PFN_send)             GetProcAddress(g_dll, "interception_send");
    pfn_is_keyboard       = (PFN_is_keyboard)      GetProcAddress(g_dll, "interception_is_keyboard");
    pfn_is_mouse          = (PFN_is_mouse)         GetProcAddress(g_dll, "interception_is_mouse");
    return pfn_create_context && pfn_destroy_context && pfn_set_filter &&
           pfn_get_hardware_id && pfn_wait_with_timeout && pfn_receive &&
           pfn_send && pfn_is_keyboard && pfn_is_mouse;
}

/* ================================================================== */
/*                         GLOBAL STATE                                */
/* ================================================================== */

#define ID_HK_TOGGLE 1
#define ID_HK_GAME   2
#define IDT_UI       3
#define IDT_REPEAT   4

#define IDC_RADIO_HOLD     100
#define IDC_RADIO_TOGGLE   101
#define IDC_BTN_DEC        102
#define IDC_BTN_INC        103
#define IDC_BTN_STARTSTOP  105
#define IDC_BTN_GAMEMODE   106
#define IDC_BTN_EXIT       107
#define IDC_LABEL_STATUS   108
#define IDC_LABEL_SPEED    109
#define IDC_LABEL_DRIVER   110
#define IDC_LABEL_REPEAT   111
#define IDC_LABEL_HKNAME   112
#define IDC_BTN_SETHK      113
#define IDC_LABEL_BURST    114
#define IDC_LABEL_GMHKNAME 115
#define IDC_BTN_SETGMHK   116
#define IDC_LABEL_DELAY    117
#define IDC_RADIO_HYBRID   118
#define IDC_BTN_ABOUT      119
#define IDC_CHK_KEYLOCK    121
#define IDC_BTN_UNINSTALL  122
#define IDC_CHK_TURBO      123
#define IDM_GM_TOGGLE      200
#define IDM_GM_SHOWMAIN    201
#define IDM_GM_EXIT        202

#define KID(code, st) ((int)((code) & 0xFF) | (((st) & INTERCEPTION_KEY_E0) ? 0x100 : 0))
#define MID_LBUTTON   0x200
#define MID_RBUTTON   0x201
#define MID_MBUTTON   0x202
#define MID_XBUTTON1  0x203
#define MID_XBUTTON2  0x204
#define MAX_KID       0x205
#define MODE_HOLD   0
#define MODE_CUSTOM 1
#define MODE_HYBRID 2

#define GAME_ALPHA          90
#define UI_TIMER_MS         200
#define TOGGLE_DEBOUNCE_MS  300
#define INTERCEPT_POLL_MS   200

static CRITICAL_SECTION g_cs_active;

static InterceptionContext g_ctx      = NULL;
static BOOL                g_drv_ok   = FALSE;
static volatile BOOL       g_active   = FALSE;
static volatile BOOL       g_quit     = FALSE;
static volatile int        g_delay    = 1000;
static volatile int        g_mode     = MODE_HOLD;
static HANDLE              g_ithread  = NULL;
static HANDLE              g_rthread  = NULL;
static HWND                g_hwnd     = NULL;

static volatile BOOL               g_held[MAX_KID];
static volatile InterceptionDevice g_hdev[MAX_KID];
static volatile unsigned short     g_hflags[MAX_KID];
static volatile BOOL               g_toggled[MAX_KID];
static volatile InterceptionDevice g_tdev[MAX_KID];
static volatile unsigned short     g_tflags[MAX_KID];

static BOOL  g_exclude[256];
static BOOL  g_custom_keys[256];
static BOOL  g_game_mode    = FALSE;
static HHOOK g_kbhook       = NULL;
static RECT  g_normal_rect;
static int   g_game_x = 1295, g_game_y = 253, g_game_w = 200, g_game_h = 80;

static int   g_hk_toggle_vk = VK_F1;
static int   g_hk_game_vk   = VK_F9;
static int   g_setting_hk   = 0;

static BOOL  g_key_lock     = FALSE;
static BOOL  g_turbo        = FALSE;

static HWND g_lbl_driver, g_lbl_status, g_lbl_delay, g_lbl_speed;
static HWND g_lbl_repeat, g_lbl_hkname, g_lbl_gmhkname;
static HWND g_btn_startstop, g_radio_hold, g_radio_toggle, g_radio_hybrid;
static HWND g_chk_keylock, g_chk_turbo;

static int     g_repeat_btn   = 0;
static int     g_repeat_count = 0;
static WNDPROC g_orig_btn_proc = NULL;

static HFONT  g_font_kb      = NULL;
static HFONT  g_font_legend  = NULL;
static HFONT  g_font_ui      = NULL;
static HFONT  g_font_game    = NULL;
static HPEN   g_pen_border   = NULL;
static HBRUSH g_br_excl      = NULL;
static HBRUSH g_br_hold      = NULL;
static HBRUSH g_br_custom    = NULL;
static HBRUSH g_br_unsel     = NULL;
static HBRUSH g_br_skip      = NULL;
static HBRUSH g_br_kbbg      = NULL;

static HDC     g_paint_memdc  = NULL;
static HBITMAP g_paint_bitmap = NULL;
static HBITMAP g_paint_oldbm  = NULL;
static int     g_paint_w      = 0;
static int     g_paint_h      = 0;

static void init_gdi_cache(void) {
    g_font_kb     = CreateFontW(12, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET, 0, 0, CLEARTYPE_QUALITY, DEFAULT_PITCH|FF_SWISS, L"Segoe UI");
    g_font_legend = CreateFontW(13, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET, 0, 0, CLEARTYPE_QUALITY, DEFAULT_PITCH|FF_SWISS, L"Segoe UI");
    g_font_ui     = CreateFontW(15, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET, 0, 0, CLEARTYPE_QUALITY, DEFAULT_PITCH|FF_SWISS, L"Segoe UI");
    g_font_game   = CreateFontW(17, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET, 0, 0, CLEARTYPE_QUALITY, DEFAULT_PITCH|FF_SWISS, L"Segoe UI");
    g_pen_border  = CreatePen(PS_SOLID, 1, RGB(60,60,60));
    g_br_excl     = CreateSolidBrush(RGB(200,60,60));
    g_br_hold     = CreateSolidBrush(RGB(220,140,20));
    g_br_custom   = CreateSolidBrush(RGB(50,170,50));
    g_br_unsel    = CreateSolidBrush(RGB(180,180,180));
    g_br_skip     = CreateSolidBrush(RGB(100,100,100));
    g_br_kbbg     = CreateSolidBrush(RGB(240,240,240));
}

static void cleanup_gdi_cache(void) {
    HGDIOBJ objs[] = { g_font_kb, g_font_legend, g_font_ui, g_font_game, g_pen_border,
                       g_br_excl, g_br_hold, g_br_custom, g_br_unsel, g_br_skip, g_br_kbbg };
    for (int i = 0; i < (int)(sizeof(objs)/sizeof(objs[0])); i++)
        if (objs[i]) DeleteObject(objs[i]);
    if (g_paint_memdc) {
        if (g_paint_oldbm) SelectObject(g_paint_memdc, g_paint_oldbm);
        DeleteDC(g_paint_memdc);
    }
    if (g_paint_bitmap) DeleteObject(g_paint_bitmap);
}

static int kid_to_vk(int kid);

#define MAX_ACTIVE 16
typedef struct {
    int kid, vk;
    BOOL is_mouse;
    unsigned short scan, flags_dn, flags_up;
    unsigned short mouse_dn, mouse_up;
} ActiveSlot;
static volatile ActiveSlot g_aslots[MAX_ACTIVE];
static volatile int        g_active_count = 0;
static volatile LONG       g_pps = 0;
static LARGE_INTEGER       g_qpc_freq;

static void active_add_ex(int kid, BOOL is_mouse, unsigned short mouse_dn, unsigned short mouse_up) {
    EnterCriticalSection(&g_cs_active);
    for (int i = 0; i < g_active_count; i++)
        if (g_aslots[i].kid == kid) { LeaveCriticalSection(&g_cs_active); return; }
    if (g_active_count >= MAX_ACTIVE) { LeaveCriticalSection(&g_cs_active); return; }
    int vk = kid_to_vk(kid);
    if (vk <= 0) { LeaveCriticalSection(&g_cs_active); return; }
    ActiveSlot s; memset(&s, 0, sizeof(s));
    s.kid = kid; s.vk = vk; s.is_mouse = is_mouse;
    if (is_mouse) {
        s.mouse_dn = mouse_dn; s.mouse_up = mouse_up;
    } else {
        s.scan = (unsigned short)(kid & 0xFF);
        s.flags_dn = (kid & 0x100) ? INTERCEPTION_KEY_E0 : 0;
        s.flags_up = s.flags_dn | INTERCEPTION_KEY_UP;
    }
    g_aslots[g_active_count] = s;
    g_active_count++;
    LeaveCriticalSection(&g_cs_active);
}

static void active_add(int kid) {
    active_add_ex(kid, FALSE, 0, 0);
}

static void active_add_mouse(int mid, int vk, unsigned short dn, unsigned short up) {
    (void)vk;
    active_add_ex(mid, TRUE, dn, up);
}

static void active_remove(int kid) {
    EnterCriticalSection(&g_cs_active);
    for (int i = 0; i < g_active_count; i++) {
        if (g_aslots[i].kid == kid) {
            g_aslots[i] = g_aslots[g_active_count - 1];
            g_active_count--;
            LeaveCriticalSection(&g_cs_active);
            return;
        }
    }
    LeaveCriticalSection(&g_cs_active);
}

/* ================================================================== */
/*                    VISUAL KEYBOARD LAYOUT                           */
/* ================================================================== */

#define KU 32
#define KG 2
#define KB_X 10
#define KB_Y 12

typedef struct { int vk; short c4,r4,w4,h4; const WCHAR *label; } KeySpec;

static const KeySpec g_keyspecs[] = {
    /* Row 0: Function keys (r4=0) */
    {VK_ESCAPE,   0, 0, 4,4, L"Esc"},
    {VK_F1,       8, 0, 4,4, L"F1"}, {VK_F2,      12, 0, 4,4, L"F2"},
    {VK_F3,      16, 0, 4,4, L"F3"}, {VK_F4,      20, 0, 4,4, L"F4"},
    {VK_F5,      26, 0, 4,4, L"F5"}, {VK_F6,      30, 0, 4,4, L"F6"},
    {VK_F7,      34, 0, 4,4, L"F7"}, {VK_F8,      38, 0, 4,4, L"F8"},
    {VK_F9,      44, 0, 4,4, L"F9"}, {VK_F10,     48, 0, 4,4, L"F10"},
    {VK_F11,     52, 0, 4,4, L"F11"},{VK_F12,     56, 0, 4,4, L"F12"},
    {VK_SNAPSHOT,62, 0, 4,4, L"Prt"},{VK_SCROLL,  66, 0, 4,4, L"Scr"},
    {VK_PAUSE,   70, 0, 4,4, L"Pau"},

    /* Row 1: Number row (r4=6) */
    {VK_OEM_3,    0, 6, 4,4, L"`"},
    {'1',         4, 6, 4,4, L"1"}, {'2',         8, 6, 4,4, L"2"},
    {'3',        12, 6, 4,4, L"3"}, {'4',        16, 6, 4,4, L"4"},
    {'5',        20, 6, 4,4, L"5"}, {'6',        24, 6, 4,4, L"6"},
    {'7',        28, 6, 4,4, L"7"}, {'8',        32, 6, 4,4, L"8"},
    {'9',        36, 6, 4,4, L"9"}, {'0',        40, 6, 4,4, L"0"},
    {VK_OEM_MINUS,44,6, 4,4, L"-"}, {VK_OEM_PLUS,48,6, 4,4, L"="},
    {VK_BACK,    52, 6, 8,4, L"Bksp"},
    {VK_INSERT,  62, 6, 4,4, L"Ins"},{VK_HOME,   66, 6, 4,4, L"Hm"},
    {VK_PRIOR,   70, 6, 4,4, L"PU"},
    {VK_NUMLOCK, 76, 6, 4,4, L"NL"}, {VK_DIVIDE, 80, 6, 4,4, L"/"},
    {VK_MULTIPLY,84, 6, 4,4, L"*"},  {VK_SUBTRACT,88,6, 4,4, L"-"},

    /* Row 2: QWERTY (r4=10) */
    {VK_TAB,      0,10, 6,4, L"Tab"},
    {'Q',         6,10, 4,4, L"Q"}, {'W',        10,10, 4,4, L"W"},
    {'E',        14,10, 4,4, L"E"}, {'R',        18,10, 4,4, L"R"},
    {'T',        22,10, 4,4, L"T"}, {'Y',        26,10, 4,4, L"Y"},
    {'U',        30,10, 4,4, L"U"}, {'I',        34,10, 4,4, L"I"},
    {'O',        38,10, 4,4, L"O"}, {'P',        42,10, 4,4, L"P"},
    {VK_OEM_4,   46,10, 4,4, L"["}, {VK_OEM_6,  50,10, 4,4, L"]"},
    {VK_OEM_5,   54,10, 6,4, L"\\"},
    {VK_DELETE,  62,10, 4,4, L"Del"},{VK_END,    66,10, 4,4, L"End"},
    {VK_NEXT,    70,10, 4,4, L"PD"},
    {VK_NUMPAD7, 76,10, 4,4, L"7"}, {VK_NUMPAD8,80,10, 4,4, L"8"},
    {VK_NUMPAD9, 84,10, 4,4, L"9"}, {VK_ADD,    88,10, 4,8, L"+"},

    /* Row 3: ASDF (r4=14) */
    {VK_CAPITAL,  0,14, 7,4, L"Caps"},
    {'A',         7,14, 4,4, L"A"}, {'S',        11,14, 4,4, L"S"},
    {'D',        15,14, 4,4, L"D"}, {'F',        19,14, 4,4, L"F"},
    {'G',        23,14, 4,4, L"G"}, {'H',        27,14, 4,4, L"H"},
    {'J',        31,14, 4,4, L"J"}, {'K',        35,14, 4,4, L"K"},
    {'L',        39,14, 4,4, L"L"}, {VK_OEM_1,  43,14, 4,4, L";"},
    {VK_OEM_7,   47,14, 4,4, L"'"},
    {VK_RETURN,  51,14, 9,4, L"Enter"},
    {VK_NUMPAD4, 76,14, 4,4, L"4"}, {VK_NUMPAD5,80,14, 4,4, L"5"},
    {VK_NUMPAD6, 84,14, 4,4, L"6"},

    /* Row 4: ZXCV (r4=18) */
    {VK_LSHIFT,   0,18, 9,4, L"Shift"},
    {'Z',         9,18, 4,4, L"Z"}, {'X',        13,18, 4,4, L"X"},
    {'C',        17,18, 4,4, L"C"}, {'V',        21,18, 4,4, L"V"},
    {'B',        25,18, 4,4, L"B"}, {'N',        29,18, 4,4, L"N"},
    {'M',        33,18, 4,4, L"M"}, {VK_OEM_COMMA,37,18,4,4, L","},
    {VK_OEM_PERIOD,41,18,4,4,L"."}, {VK_OEM_2,  45,18, 4,4, L"/"},
    {VK_RSHIFT,  49,18,11,4, L"Shift"},
    {VK_UP,      66,18, 4,4, L"\x2191"},
    {VK_NUMPAD1, 76,18, 4,4, L"1"}, {VK_NUMPAD2,80,18, 4,4, L"2"},
    {VK_NUMPAD3, 84,18, 4,4, L"3"},
    {VK_RETURN,  88,18, 4,8, L"Ent"},

    /* Row 5: Bottom (r4=22) */
    {VK_LCONTROL, 0,22, 5,4, L"Ctrl"},
    {VK_LWIN,     5,22, 5,4, L"Win"},
    {VK_LMENU,   10,22, 5,4, L"Alt"},
    {VK_SPACE,   15,22,25,4, L"Space"},
    {VK_RMENU,   40,22, 5,4, L"Alt"},
    {VK_RWIN,    45,22, 5,4, L"Win"},
    {VK_APPS,    50,22, 5,4, L"Mn"},
    {VK_RCONTROL,55,22, 5,4, L"Ctrl"},
    {VK_LEFT,    62,22, 4,4, L"\x2190"},
    {VK_DOWN,    66,22, 4,4, L"\x2193"},
    {VK_RIGHT,   70,22, 4,4, L"\x2192"},
    {VK_NUMPAD0, 76,22, 8,4, L"0"},
    {VK_DECIMAL, 84,22, 4,4, L"."},

    /* Mouse buttons (r4=0, above numpad) */
    {VK_LBUTTON,  76, 0, 4,4, L"ML"},
    {VK_RBUTTON,  80, 0, 4,4, L"MR"},
    {VK_MBUTTON,  84, 0, 4,4, L"MM"},
    {VK_XBUTTON1, 88, 0, 4,4, L"X1"},
    {VK_XBUTTON2, 92, 0, 4,4, L"X2"},
};

#define N_KEYS (sizeof(g_keyspecs)/sizeof(g_keyspecs[0]))

typedef struct { int vk; RECT rc; const WCHAR *label; } KeyRect;
static KeyRect g_krects[sizeof(g_keyspecs)/sizeof(g_keyspecs[0])];

static void init_keyrects(void) {
    for (int i = 0; i < (int)N_KEYS; i++) {
        const KeySpec *s = &g_keyspecs[i];
        g_krects[i].vk = s->vk;
        g_krects[i].label = s->label;
        g_krects[i].rc.left   = KB_X + s->c4 * KU / 4;
        g_krects[i].rc.top    = KB_Y + s->r4 * KU / 4;
        g_krects[i].rc.right  = g_krects[i].rc.left + s->w4 * KU / 4 - KG;
        g_krects[i].rc.bottom = g_krects[i].rc.top  + s->h4 * KU / 4 - KG;
    }
}

/* ================================================================== */
/*                       CONFIG SAVE / LOAD                            */
/* ================================================================== */

static void get_exe_dir(WCHAR *buf, int buflen) {
    GetModuleFileNameW(NULL, buf, buflen);
    WCHAR *p = wcsrchr(buf, L'\\');
    if (p) *(p + 1) = 0;
}

static void save_config(void) {
    WCHAR path[MAX_PATH];
    get_exe_dir(path, MAX_PATH);
    lstrcatW(path, L"config.json");
    char buf[8192]; int pos = 0; int rem = (int)sizeof(buf);
    #define SCFG(fmt, ...) do { int n = snprintf(buf+pos, rem, fmt, __VA_ARGS__); if (n>0){pos+=n; rem-=n;} } while(0)
    SCFG("{\n  \"delay_us\": %d,\n  \"mode\": %d,\n", g_delay, g_mode);
    SCFG("  \"exclude\": [%s", "");
    int first = 1;
    for (int i = 1; i < 256 && rem > 8; i++) { if (!g_exclude[i]) continue; if (!first) SCFG(",%s",""); SCFG("%d", i); first = 0; }
    SCFG("],\n  \"custom_keys\": [%s", "");
    first = 1;
    for (int i = 1; i < 256 && rem > 8; i++) { if (!g_custom_keys[i]) continue; if (!first) SCFG(",%s",""); SCFG("%d", i); first = 0; }
    SCFG("],\n%s", "");
    SCFG("  \"game_x\": %d, \"game_y\": %d, \"game_w\": %d, \"game_h\": %d,\n",
         g_game_x, g_game_y, g_game_w, g_game_h);
    SCFG("  \"hk_toggle\": %d, \"hk_game\": %d,\n", g_hk_toggle_vk, g_hk_game_vk);
    SCFG("  \"key_lock\": %d,\n  \"turbo\": %d\n}\n", g_key_lock ? 1 : 0, g_turbo ? 1 : 0);
    #undef SCFG
    HANDLE hFile = CreateFileW(path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE) return;
    DWORD written; WriteFile(hFile, buf, (DWORD)pos, &written, NULL); CloseHandle(hFile);
}

static void parse_int_array(const char *start, BOOL *arr, int arrlen) {
    const char *p = strchr(start, '[');
    if (!p) return; p++;
    while (*p && *p != ']') {
        while (*p == ' ' || *p == ',' || *p == '\n' || *p == '\r') p++;
        if (*p >= '0' && *p <= '9') { int val = atoi(p); if (val > 0 && val < arrlen) arr[val] = TRUE; while (*p >= '0' && *p <= '9') p++; }
        else if (*p != ']') p++;
    }
}

static void load_config(void) {
    WCHAR path[MAX_PATH]; get_exe_dir(path, MAX_PATH); lstrcatW(path, L"config.json");
    HANDLE hFile = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, NULL);
    if (hFile == INVALID_HANDLE_VALUE) { save_config(); return; }
    char buf[8192]; DWORD bytesRead;
    ReadFile(hFile, buf, sizeof(buf)-1, &bytesRead, NULL); CloseHandle(hFile);
    buf[bytesRead] = 0;
    char *p;
    p = strstr(buf, "\"delay_us\"");
    if (p) { p = strchr(p, ':'); if (p) g_delay = atoi(p+1); }
    else { p = strstr(buf, "\"delay\""); if (p) { p = strchr(p, ':'); if (p) g_delay = atoi(p+1)*1000; } }
    p = strstr(buf, "\"mode\"");
    if (p) { p = strchr(p, ':'); if (p) g_mode = atoi(p+1); }
    p = strstr(buf, "\"exclude\"");
    if (p) { memset(g_exclude, 0, sizeof(g_exclude)); parse_int_array(p, g_exclude, 256); }
    p = strstr(buf, "\"custom_keys\"");
    if (p) { memset(g_custom_keys, 0, sizeof(g_custom_keys)); parse_int_array(p, g_custom_keys, 256); }
    p = strstr(buf, "\"game_x\""); if (p) { p = strchr(p, ':'); if (p) g_game_x = atoi(p+1); }
    p = strstr(buf, "\"game_y\""); if (p) { p = strchr(p, ':'); if (p) g_game_y = atoi(p+1); }
    p = strstr(buf, "\"game_w\""); if (p) { p = strchr(p, ':'); if (p) { g_game_w = atoi(p+1); if (g_game_w<100) g_game_w=200; } }
    p = strstr(buf, "\"game_h\""); if (p) { p = strchr(p, ':'); if (p) { g_game_h = atoi(p+1); if (g_game_h<40) g_game_h=80; } }
    p = strstr(buf, "\"hk_toggle\""); if (p) { p = strchr(p, ':'); if (p) { int v=atoi(p+1); if (v>0 && v<256) g_hk_toggle_vk=v; } }
    p = strstr(buf, "\"hk_game\"");   if (p) { p = strchr(p, ':'); if (p) { int v=atoi(p+1); if (v>0 && v<256) g_hk_game_vk=v; } }
    p = strstr(buf, "\"key_lock\""); if (p) { p = strchr(p, ':'); if (p) g_key_lock = (atoi(p+1) != 0); }
    p = strstr(buf, "\"turbo\"");    if (p) { p = strchr(p, ':'); if (p) g_turbo = (atoi(p+1) != 0); }
}

/* ================================================================== */
/*                    DRIVER INSTALLATION HELPERS                      */
/* ================================================================== */

static void con_print(HANDLE h, const WCHAR *s) { DWORD w; WriteConsoleW(h, s, (DWORD)lstrlenW(s), &w, NULL); }

static int do_install(void) {
    AllocConsole();
    SetConsoleTitleW(L"\x5B89\x88C5 Interception \x9A71\x52A8");
    HANDLE hCon = GetStdHandle(STD_OUTPUT_HANDLE);
    con_print(hCon, L"=== \x5B89\x88C5 Interception \x9A71\x52A8 ===\n\n");
    WCHAR tempDir[MAX_PATH], installerPath[MAX_PATH];
    GetTempPathW(MAX_PATH, tempDir);
    wsprintfW(installerPath, L"%sinstall-interception.exe", tempDir);
    con_print(hCon, L"[1/2] \x6B63\x5728\x91CA\x653E\x5B89\x88C5\x7A0B\x5E8F...\n");
    if (!extract_resource(IDR_INSTALL_EXE, installerPath)) {
        con_print(hCon, L"  [\x5931\x8D25]\n\n\x6309\x56DE\x8F66\x952E\x9000\x51FA...\n");
        WCHAR c; DWORD r; ReadConsoleW(GetStdHandle(STD_INPUT_HANDLE), &c, 1, &r, NULL); return 1;
    }
    con_print(hCon, L"  [\x6210\x529F]\n\n[2/2] \x6B63\x5728\x5B89\x88C5\x9A71\x52A8...\n");
    WCHAR cmdline[MAX_PATH+32]; wsprintfW(cmdline, L"\"%s\" /install", installerPath);
    STARTUPINFOW si={0}; si.cb=sizeof(si); PROCESS_INFORMATION pi={0};
    if (!CreateProcessW(NULL, cmdline, NULL, NULL, FALSE, 0, NULL, tempDir, &si, &pi)) {
        con_print(hCon, L"  [\x5931\x8D25]\n"); DeleteFileW(installerPath);
        WCHAR c; DWORD r; ReadConsoleW(GetStdHandle(STD_INPUT_HANDLE), &c, 1, &r, NULL); return 1;
    }
    WaitForSingleObject(pi.hProcess, 30000); CloseHandle(pi.hProcess); CloseHandle(pi.hThread);
    DeleteFileW(installerPath);
    con_print(hCon, L"  [\x6210\x529F]\n\n\x5B89\x88C5\x5B8C\x6210\xFF01\x8BF7\x91CD\x542F\x7535\x8111\x3002\n\n\x6309\x56DE\x8F66\x952E\x5173\x95ED...\n");
    WCHAR c; DWORD rd; ReadConsoleW(GetStdHandle(STD_INPUT_HANDLE), &c, 1, &rd, NULL);
    FreeConsole(); return 0;
}

static int do_uninstall(void) {
    AllocConsole();
    SetConsoleTitleW(L"\x5378\x8F7D Interception \x9A71\x52A8");
    HANDLE hCon = GetStdHandle(STD_OUTPUT_HANDLE);
    con_print(hCon, L"=== \x5378\x8F7D Interception \x9A71\x52A8 ===\n\n");
    WCHAR tempDir[MAX_PATH], installerPath[MAX_PATH];
    GetTempPathW(MAX_PATH, tempDir);
    wsprintfW(installerPath, L"%sinstall-interception.exe", tempDir);
    con_print(hCon, L"[1/2] \x6B63\x5728\x91CA\x653E\x5378\x8F7D\x7A0B\x5E8F...\n");
    if (!extract_resource(IDR_INSTALL_EXE, installerPath)) {
        con_print(hCon, L"  [\x5931\x8D25]\n\n\x6309\x56DE\x8F66\x952E\x9000\x51FA...\n");
        WCHAR c; DWORD r; ReadConsoleW(GetStdHandle(STD_INPUT_HANDLE), &c, 1, &r, NULL); return 1;
    }
    con_print(hCon, L"  [\x6210\x529F]\n\n[2/2] \x6B63\x5728\x5378\x8F7D\x9A71\x52A8...\n");
    WCHAR cmdline[MAX_PATH+32]; wsprintfW(cmdline, L"\"%s\" /uninstall", installerPath);
    STARTUPINFOW si={0}; si.cb=sizeof(si); PROCESS_INFORMATION pi={0};
    if (!CreateProcessW(NULL, cmdline, NULL, NULL, FALSE, 0, NULL, tempDir, &si, &pi)) {
        con_print(hCon, L"  [\x5931\x8D25]\n"); DeleteFileW(installerPath);
        WCHAR c; DWORD r; ReadConsoleW(GetStdHandle(STD_INPUT_HANDLE), &c, 1, &r, NULL); return 1;
    }
    WaitForSingleObject(pi.hProcess, 30000); CloseHandle(pi.hProcess); CloseHandle(pi.hThread);
    DeleteFileW(installerPath);
    con_print(hCon, L"  [\x6210\x529F]\n\n\x5378\x8F7D\x5B8C\x6210\xFF01\x8BF7\x91CD\x542F\x7535\x8111\x3002\n\n\x6309\x56DE\x8F66\x952E\x5173\x95ED...\n");
    WCHAR c; DWORD rd; ReadConsoleW(GetStdHandle(STD_INPUT_HANDLE), &c, 1, &rd, NULL);
    FreeConsole(); return 0;
}

static void init_interception(void);

static void try_uninstall(void) {
    int r = MessageBoxW(g_hwnd,
        L"\x786E\x5B9A\x8981\x5378\x8F7D Interception \x9A71\x52A8\x5417\xFF1F\n\n"
        L"\x5378\x8F7D\x540E\x9700\x91CD\x542F\x7535\x8111\xFF0C\x8FDE\x53D1\x529F\x80FD\x5C06\x4E0D\x53EF\x7528\x3002",
        L"\x5378\x8F7D\x9A71\x52A8", MB_YESNO | MB_ICONWARNING);
    if (r != IDYES) return;
    if (g_active) { g_active = FALSE; if (g_rthread) { WaitForSingleObject(g_rthread,500); CloseHandle(g_rthread); g_rthread=NULL; } }
    WCHAR exe[MAX_PATH]; GetModuleFileNameW(NULL, exe, MAX_PATH);
    SHELLEXECUTEINFOW sei; memset(&sei, 0, sizeof(sei));
    sei.cbSize=sizeof(sei); sei.fMask=SEE_MASK_NOCLOSEPROCESS;
    sei.lpVerb=L"runas"; sei.lpFile=exe; sei.lpParameters=L"/uninstall"; sei.nShow=SW_SHOW;
    if (!ShellExecuteExW(&sei)) { MessageBoxW(g_hwnd, L"\x65E0\x6CD5\x83B7\x53D6\x7BA1\x7406\x5458\x6743\x9650", L"\x9519\x8BEF", MB_ICONERROR); return; }
    WaitForSingleObject(sei.hProcess, 180000); CloseHandle(sei.hProcess);
    MessageBoxW(g_hwnd, L"\x9A71\x52A8\x5378\x8F7D\x5B8C\x6210\xFF0C\x8BF7\x91CD\x542F\x7535\x8111\x3002", L"\x63D0\x793A", MB_ICONWARNING);
    g_drv_ok = FALSE;
    InvalidateRect(g_hwnd, NULL, TRUE);
}

static void try_auto_install(void) {
    int r = MessageBoxW(NULL,
        L"Interception \x9A71\x52A8\x672A\x5B89\x88C5\x6216\x672A\x52A0\x8F7D\x3002\n\n"
        L"\x662F\x5426\x81EA\x52A8\x5B89\x88C5\xFF1F\xFF08\x9700\x7BA1\x7406\x5458\x6743\x9650\xFF09\n"
        L"\x5B89\x88C5\x540E\x9700\x91CD\x542F\x7535\x8111\x3002",
        L"\x5B89\x88C5\x9A71\x52A8", MB_YESNO | MB_ICONQUESTION);
    if (r != IDYES) return;
    WCHAR exe[MAX_PATH]; GetModuleFileNameW(NULL, exe, MAX_PATH);
    SHELLEXECUTEINFOW sei; memset(&sei, 0, sizeof(sei));
    sei.cbSize=sizeof(sei); sei.fMask=SEE_MASK_NOCLOSEPROCESS;
    sei.lpVerb=L"runas"; sei.lpFile=exe; sei.lpParameters=L"/install"; sei.nShow=SW_SHOW;
    if (!ShellExecuteExW(&sei)) { MessageBoxW(NULL, L"\x65E0\x6CD5\x83B7\x53D6\x7BA1\x7406\x5458\x6743\x9650", L"\x9519\x8BEF", MB_ICONERROR); return; }
    WaitForSingleObject(sei.hProcess, 180000); CloseHandle(sei.hProcess);
    if (g_ctx) { interception_destroy_context(g_ctx); g_ctx = NULL; }
    g_drv_ok = FALSE; Sleep(2000); init_interception();
    if (g_drv_ok) MessageBoxW(NULL, L"\x9A71\x52A8\x5B89\x88C5\x6210\x529F\xFF01", L"\x6210\x529F", MB_ICONINFORMATION);
    else MessageBoxW(NULL, L"\x9A71\x52A8\x5B89\x88C5\x5B8C\x6210\xFF0C\x8BF7\x91CD\x542F\x7535\x8111\x3002", L"\x63D0\x793A", MB_ICONWARNING);
}

/* ================================================================== */
/*                           APP LOGIC                                 */
/* ================================================================== */

static BOOL is_skippable(int vk) {
    if (vk == 0) return TRUE;
    if (vk == g_hk_toggle_vk || vk == g_hk_game_vk) return TRUE;
    return FALSE;
}

static BOOL is_excluded(int vk) { return (vk > 0 && vk < 256 && g_exclude[vk]); }

static int kid_to_vk(int kid) {
    switch (kid) {
    case MID_LBUTTON:  return VK_LBUTTON;
    case MID_RBUTTON:  return VK_RBUTTON;
    case MID_MBUTTON:  return VK_MBUTTON;
    case MID_XBUTTON1: return VK_XBUTTON1;
    case MID_XBUTTON2: return VK_XBUTTON2;
    }
    unsigned short sc = (unsigned short)(kid & 0xFF);
    if (kid & 0x100) {
        switch (sc) {
        case 0x1D: return VK_RCONTROL; case 0x38: return VK_RMENU;
        case 0x47: return VK_HOME;     case 0x48: return VK_UP;
        case 0x49: return VK_PRIOR;    case 0x4B: return VK_LEFT;
        case 0x4D: return VK_RIGHT;    case 0x4F: return VK_END;
        case 0x50: return VK_DOWN;     case 0x51: return VK_NEXT;
        case 0x52: return VK_INSERT;   case 0x53: return VK_DELETE;
        case 0x5B: return VK_LWIN;     case 0x5C: return VK_RWIN;
        case 0x35: return VK_DIVIDE;   case 0x1C: return VK_RETURN;
        }
    } else {
        switch (sc) {
        case 0x1D: return VK_LCONTROL;
        case 0x38: return VK_LMENU;
        case 0x2A: return VK_LSHIFT;
        case 0x36: return VK_RSHIFT;
        case 0x47: return VK_NUMPAD7;
        case 0x48: return VK_NUMPAD8;
        case 0x49: return VK_NUMPAD9;
        case 0x4B: return VK_NUMPAD4;
        case 0x4D: return VK_NUMPAD6;
        case 0x4F: return VK_NUMPAD1;
        case 0x50: return VK_NUMPAD2;
        case 0x51: return VK_NUMPAD3;
        case 0x52: return VK_NUMPAD0;
        case 0x53: return VK_DECIMAL;
        }
    }
    return (int)MapVirtualKeyW(sc, MAPVK_VSC_TO_VK);
}

static void get_vk_name(int vk, WCHAR *buf, int buflen) {
    switch (vk) {
    case VK_LBUTTON:  lstrcpynW(buf, L"M.Left", buflen); return;
    case VK_RBUTTON:  lstrcpynW(buf, L"M.Right", buflen); return;
    case VK_MBUTTON:  lstrcpynW(buf, L"M.Mid", buflen); return;
    case VK_XBUTTON1: lstrcpynW(buf, L"M.X1", buflen); return;
    case VK_XBUTTON2: lstrcpynW(buf, L"M.X2", buflen); return;
    }
    UINT sc = MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    if (sc && GetKeyNameTextW((LONG)(sc << 16), buf, buflen) > 0) return;
    wsprintfW(buf, L"VK 0x%02X", vk);
}

static void init_exclude_defaults(void) {
    memset(g_exclude, 0, sizeof(g_exclude));
    static const int def_excl[] = {
        VK_LBUTTON, VK_RBUTTON, VK_MBUTTON, VK_XBUTTON1, VK_XBUTTON2,
        VK_BACK, VK_TAB, VK_RETURN, VK_SHIFT, VK_PAUSE, VK_CAPITAL,
        VK_ESCAPE, VK_SPACE,
        VK_PRIOR, VK_NEXT, VK_END, VK_HOME,
        VK_LEFT, VK_UP, VK_RIGHT, VK_DOWN,
        VK_SNAPSHOT, VK_INSERT, VK_DELETE,
        '0', '7', '8', '9',
        'A', 'B', 'D', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P',
        'S', 'U', 'V', 'W', 'Y',
        VK_LWIN, VK_RWIN, VK_APPS,
        VK_NUMPAD0, VK_NUMPAD1, VK_NUMPAD2, VK_NUMPAD3, VK_NUMPAD4,
        VK_NUMPAD5, VK_NUMPAD6, VK_NUMPAD7, VK_NUMPAD8, VK_NUMPAD9,
        VK_MULTIPLY, VK_ADD, VK_SUBTRACT, VK_DECIMAL, VK_DIVIDE,
        VK_F1, VK_F2, VK_F3, VK_F4, VK_F5, VK_F6,
        VK_F7, VK_F8, VK_F9, VK_F10, VK_F11, VK_F12,
        VK_NUMLOCK, VK_SCROLL,
        VK_LSHIFT, VK_RSHIFT, VK_LCONTROL, VK_RCONTROL, VK_LMENU, VK_RMENU,
        VK_OEM_1, VK_OEM_PLUS, VK_OEM_COMMA, VK_OEM_MINUS,
        VK_OEM_PERIOD, VK_OEM_2, VK_OEM_4, VK_OEM_5, VK_OEM_6, VK_OEM_7,
    };
    for (int i = 0; i < (int)(sizeof(def_excl)/sizeof(def_excl[0])); i++)
        if (def_excl[i] > 0 && def_excl[i] < 256) g_exclude[def_excl[i]] = TRUE;
}

static void init_custom_keys_defaults(void) {
    memset(g_custom_keys, 0, sizeof(g_custom_keys));
}

static int build_active_list(WCHAR *buf, int bufmax, int max_show) {
    int count=0, pos=0;
    for (int kid=0; kid<MAX_KID && count<max_show; kid++) {
        if (!g_held[kid]) continue;
        int vk = kid_to_vk(kid);
        if (is_skippable(vk) || is_excluded(vk)) continue;
        WCHAR name[64]; get_vk_name(vk, name, 64);
        if (count>0) { buf[pos++]=L','; buf[pos++]=L' '; }
        int len=lstrlenW(name);
        for (int i=0; i<len && pos<bufmax-1; i++) buf[pos++]=name[i];
        buf[pos]=0; count++;
    }
    return count;
}

static int build_toggled_list(WCHAR *buf, int bufmax, int max_show) {
    int count=0, pos=0;
    for (int kid=0; kid<MAX_KID && count<max_show; kid++) {
        if (!g_toggled[kid]) continue;
        int vk = kid_to_vk(kid);
        WCHAR name[64]; get_vk_name(vk, name, 64);
        if (count>0) { buf[pos++]=L','; buf[pos++]=L' '; }
        int len=lstrlenW(name);
        for (int i=0; i<len && pos<bufmax-1; i++) buf[pos++]=name[i];
        buf[pos]=0; count++;
    }
    return count;
}

/* ================================================================== */
/*                          THREADS                                    */
/* ================================================================== */

static unsigned short mid_to_mouse_up(int mid) {
    switch (mid) {
    case MID_LBUTTON:  return INTERCEPTION_MOUSE_LEFT_BUTTON_UP;
    case MID_RBUTTON:  return INTERCEPTION_MOUSE_RIGHT_BUTTON_UP;
    case MID_MBUTTON:  return INTERCEPTION_MOUSE_MIDDLE_BUTTON_UP;
    case MID_XBUTTON1: return INTERCEPTION_MOUSE_BUTTON_4_UP;
    case MID_XBUTTON2: return INTERCEPTION_MOUSE_BUTTON_5_UP;
    }
    return 0;
}

static void send_toggled_keyups(void) {
    for (int kid=0; kid<MAX_KID; kid++) {
        if (g_toggled[kid] && g_tdev[kid] && g_drv_ok) {
            if (kid >= MID_LBUTTON && kid <= MID_XBUTTON2) {
                InterceptionMouseStroke ms; memset(&ms, 0, sizeof(ms));
                ms.state = mid_to_mouse_up(kid);
                interception_send(g_ctx, g_tdev[kid], (InterceptionStroke *)&ms, 1);
            } else {
                InterceptionKeyStroke up_ks;
                up_ks.code = (unsigned short)(kid & 0xFF);
                up_ks.state = INTERCEPTION_KEY_UP | g_tflags[kid];
                up_ks.information = 0;
                interception_send(g_ctx, g_tdev[kid], (InterceptionStroke *)&up_ks, 1);
            }
        }
    }
}

static void release_toggled(void) {
    for (int kid=0; kid<MAX_KID; kid++) {
        if (g_toggled[kid] && g_tdev[kid] && g_drv_ok) {
            if (kid >= MID_LBUTTON && kid <= MID_XBUTTON2) {
                InterceptionMouseStroke ms; memset(&ms, 0, sizeof(ms));
                ms.state = mid_to_mouse_up(kid);
                interception_send(g_ctx, g_tdev[kid], (InterceptionStroke *)&ms, 1);
            } else {
                InterceptionKeyStroke up_ks;
                up_ks.code = (unsigned short)(kid & 0xFF);
                up_ks.state = INTERCEPTION_KEY_UP | g_tflags[kid];
                up_ks.information = 0;
                interception_send(g_ctx, g_tdev[kid], (InterceptionStroke *)&up_ks, 1);
            }
        }
    }
    memset((void *)g_toggled, 0, sizeof(g_toggled));
    memset((void *)g_tdev, 0, sizeof(g_tdev));
    EnterCriticalSection(&g_cs_active);
    g_active_count = 0;
    LeaveCriticalSection(&g_cs_active);
}

static const struct { int mid; int vk; unsigned short dn; unsigned short up; } g_mouse_btns[] = {
    {MID_LBUTTON,  VK_LBUTTON,  INTERCEPTION_MOUSE_LEFT_BUTTON_DOWN,   INTERCEPTION_MOUSE_LEFT_BUTTON_UP},
    {MID_RBUTTON,  VK_RBUTTON,  INTERCEPTION_MOUSE_RIGHT_BUTTON_DOWN,  INTERCEPTION_MOUSE_RIGHT_BUTTON_UP},
    {MID_MBUTTON,  VK_MBUTTON,  INTERCEPTION_MOUSE_MIDDLE_BUTTON_DOWN, INTERCEPTION_MOUSE_MIDDLE_BUTTON_UP},
    {MID_XBUTTON1, VK_XBUTTON1, INTERCEPTION_MOUSE_BUTTON_4_DOWN,      INTERCEPTION_MOUSE_BUTTON_4_UP},
    {MID_XBUTTON2, VK_XBUTTON2, INTERCEPTION_MOUSE_BUTTON_5_DOWN,      INTERCEPTION_MOUSE_BUTTON_5_UP},
};
#define N_MOUSE_BTNS (sizeof(g_mouse_btns)/sizeof(g_mouse_btns[0]))

static DWORD WINAPI intercept_proc(LPVOID p) {
    (void)p;
    interception_set_filter(g_ctx, interception_is_keyboard, INTERCEPTION_FILTER_KEY_ALL);
    interception_set_filter(g_ctx, interception_is_mouse,
        INTERCEPTION_FILTER_MOUSE_LEFT_BUTTON_DOWN | INTERCEPTION_FILTER_MOUSE_LEFT_BUTTON_UP |
        INTERCEPTION_FILTER_MOUSE_RIGHT_BUTTON_DOWN | INTERCEPTION_FILTER_MOUSE_RIGHT_BUTTON_UP |
        INTERCEPTION_FILTER_MOUSE_MIDDLE_BUTTON_DOWN | INTERCEPTION_FILTER_MOUSE_MIDDLE_BUTTON_UP |
        INTERCEPTION_FILTER_MOUSE_BUTTON_4_DOWN | INTERCEPTION_FILTER_MOUSE_BUTTON_4_UP |
        INTERCEPTION_FILTER_MOUSE_BUTTON_5_DOWN | INTERCEPTION_FILTER_MOUSE_BUTTON_5_UP);
    while (!g_quit) {
        InterceptionDevice dev = interception_wait_with_timeout(g_ctx, INTERCEPT_POLL_MS);
        if (dev == 0) continue;
        InterceptionStroke stroke;
        if (interception_receive(g_ctx, dev, &stroke, 1) <= 0) continue;
        if (interception_is_keyboard(dev)) {
            InterceptionKeyStroke *ks = (InterceptionKeyStroke *)&stroke;
            int kid = KID(ks->code, ks->state);
            if (kid >= 0 && kid < 0x200) {
                int vk = kid_to_vk(kid);
                BOOL is_hk = (vk == g_hk_toggle_vk || vk == g_hk_game_vk);
                if (ks->state & INTERCEPTION_KEY_UP) {
                    if (!is_hk) {
                        g_held[kid]=FALSE; g_hdev[kid]=0;
                        if (g_mode==MODE_HOLD || g_mode==MODE_HYBRID) {
                            if (!(g_mode==MODE_HYBRID && g_toggled[kid]))
                                active_remove(kid);
                        }
                    }
                } else if (!is_hk) {
                    if ((g_mode==MODE_CUSTOM || g_mode==MODE_HYBRID) && g_active && !g_held[kid]) {
                        if (!is_skippable(vk) && vk>0 && vk<256 && g_custom_keys[vk]) {
                            if (g_toggled[kid]) { g_toggled[kid]=FALSE; g_tdev[kid]=0; active_remove(kid); }
                            else { g_toggled[kid]=TRUE; g_tdev[kid]=dev; g_tflags[kid]=ks->state&(unsigned short)~INTERCEPTION_KEY_UP; active_add(kid); }
                        }
                    }
                    g_held[kid]=TRUE; g_hdev[kid]=dev;
                    g_hflags[kid]=ks->state&(unsigned short)~INTERCEPTION_KEY_UP;
                    if (g_mode==MODE_HOLD || g_mode==MODE_HYBRID) active_add(kid);
                }
            }
        } else if (interception_is_mouse(dev)) {
            InterceptionMouseStroke *ms = (InterceptionMouseStroke *)&stroke;
            for (int b = 0; b < (int)N_MOUSE_BTNS; b++) {
                int mid = g_mouse_btns[b].mid, vk = g_mouse_btns[b].vk;
                if (ms->state & g_mouse_btns[b].up) {
                    g_held[mid]=FALSE; g_hdev[mid]=0;
                    if (g_mode==MODE_HOLD || g_mode==MODE_HYBRID) {
                        if (!(g_mode==MODE_HYBRID && g_toggled[mid]))
                            active_remove(mid);
                    }
                }
                if (ms->state & g_mouse_btns[b].dn) {
                    if ((g_mode==MODE_CUSTOM || g_mode==MODE_HYBRID) && g_active && !g_held[mid]) {
                        if (!is_excluded(vk) && vk>0 && vk<256 && g_custom_keys[vk]) {
                            if (g_toggled[mid]) { g_toggled[mid]=FALSE; g_tdev[mid]=0; active_remove(mid); }
                            else { g_toggled[mid]=TRUE; g_tdev[mid]=dev; active_add_mouse(mid, vk, g_mouse_btns[b].dn, g_mouse_btns[b].up); }
                        }
                    }
                    g_held[mid]=TRUE; g_hdev[mid]=dev;
                    if (g_mode==MODE_HOLD || g_mode==MODE_HYBRID)
                        active_add_mouse(mid, vk, g_mouse_btns[b].dn, g_mouse_btns[b].up);
                }
            }
        }
        interception_send(g_ctx, dev, &stroke, 1);
    }
    return 0;
}

static void interruptible_sleep_us(int us) {
    LARGE_INTEGER start, now;
    long long target, elapsed;
    int ms_sleep;
    if (us <= 0 || !g_active) return;
    QueryPerformanceCounter(&start);
    target = (long long)us * g_qpc_freq.QuadPart / 1000000LL;
    ms_sleep = us / 1000 - 2;
    if (ms_sleep < 0) ms_sleep = 0;
    while (ms_sleep > 0 && g_active) {
        int chunk = (ms_sleep > 4) ? 4 : ms_sleep;
        Sleep(chunk);
        ms_sleep -= chunk;
        QueryPerformanceCounter(&now);
        if (now.QuadPart - start.QuadPart >= target) return;
    }
    do {
        YieldProcessor();
        QueryPerformanceCounter(&now);
        elapsed = now.QuadPart - start.QuadPart;
    } while (elapsed < target && g_active);
}

static DWORD WINAPI repeat_proc(LPVOID p) {
    (void)p;
    LARGE_INTEGER last;
    InterceptionKeyStroke ibatch[2];
    InterceptionMouseStroke mbatch[2];
    LONG local_count = 0;
    ActiveSlot local_slots[MAX_ACTIVE];
    BOOL local_toggled[MAX_ACTIVE];
    BOOL local_held[MAX_ACTIVE];
    InterceptionDevice local_tdev[MAX_ACTIVE];
    InterceptionDevice local_hdev[MAX_ACTIVE];

    timeBeginPeriod(1);
    SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL);
    memset(ibatch, 0, sizeof(ibatch));
    memset(mbatch, 0, sizeof(mbatch));
    QueryPerformanceCounter(&last);

    while (g_active) {
        int cnt, mode, delay, i;
        BOOL any_sent = FALSE;
        EnterCriticalSection(&g_cs_active);
        cnt = g_active_count;
        if (cnt > 0) {
            memcpy(local_slots, (void*)g_aslots, cnt * sizeof(ActiveSlot));
            for (i = 0; i < cnt; i++) {
                int kid = local_slots[i].kid;
                local_toggled[i] = g_toggled[kid];
                local_held[i] = g_held[kid];
                local_tdev[i] = g_tdev[kid];
                local_hdev[i] = g_hdev[kid];
            }
        }
        LeaveCriticalSection(&g_cs_active);
        if (cnt == 0) { Sleep(1); continue; }
        mode = g_mode; delay = g_delay;
        for (i = 0; i < cnt && g_active; i++) {
            int vk = local_slots[i].vk;
            BOOL is_toggled_on = local_toggled[i];
            BOOL is_held_on = local_held[i];
            BOOL is_on;
            InterceptionDevice idev;
            if (mode==MODE_HYBRID) is_on = is_toggled_on || (is_held_on && !is_toggled_on);
            else if (mode==MODE_HOLD) is_on = is_held_on;
            else is_on = is_toggled_on;
            if (!is_on) continue;
            if ((mode==MODE_HOLD || (mode==MODE_HYBRID && !is_toggled_on)) && (is_skippable(vk)||is_excluded(vk))) continue;
            if (mode==MODE_HYBRID) idev = is_toggled_on ? local_tdev[i] : local_hdev[i];
            else if (mode==MODE_HOLD) idev = local_hdev[i];
            else idev = local_tdev[i];
            if (!idev) continue;
            any_sent = TRUE;
            if (local_slots[i].is_mouse) {
                mbatch[0].state = local_slots[i].mouse_dn;
                mbatch[1].state = local_slots[i].mouse_up;
                interception_send(g_ctx, idev, (InterceptionStroke*)mbatch, 2);
            } else {
                ibatch[0].code = local_slots[i].scan; ibatch[0].state = local_slots[i].flags_dn;
                ibatch[1].code = local_slots[i].scan; ibatch[1].state = local_slots[i].flags_up;
                interception_send(g_ctx, idev, (InterceptionStroke*)ibatch, 2);
            }
            local_count++;
        }
        {
            LARGE_INTEGER now;
            QueryPerformanceCounter(&now);
            if (now.QuadPart - last.QuadPart >= g_qpc_freq.QuadPart) {
                InterlockedExchange(&g_pps, local_count); local_count = 0; last = now;
            }
        }
        if (!any_sent) { Sleep(1); continue; }
        interruptible_sleep_us(g_turbo ? (delay<1?1:delay) : (delay<1000?1000:delay));
    }
    InterlockedExchange(&g_pps, 0); timeEndPeriod(1); return 0;
}

/* ================================================================== */
/*                      TOGGLE / GAME MODE                             */
/* ================================================================== */

static DWORD g_last_toggle_tick = 0;

static void toggle_active(void) {
    DWORD now = GetTickCount();
    if (now - g_last_toggle_tick < TOGGLE_DEBOUNCE_MS) return;
    g_last_toggle_tick = now;
    if (!g_drv_ok) {
        MessageBoxW(g_hwnd, L"Interception \x9A71\x52A8\x672A\x5C31\x7EEA\xFF01", L"\x9519\x8BEF", MB_ICONERROR);
        return;
    }
    if (g_active) {
        g_active = FALSE;
        if (g_rthread) { WaitForSingleObject(g_rthread,500); CloseHandle(g_rthread); g_rthread=NULL; }
        if (g_key_lock)
            send_toggled_keyups();
        else
            release_toggled();
    } else {
        if (!g_key_lock) release_toggled();
        g_active = TRUE;
        g_rthread = CreateThread(NULL, 0, repeat_proc, NULL, 0, NULL);
    }
    InvalidateRect(g_hwnd, NULL, TRUE);
}

static BOOL CALLBACK show_child_proc(HWND child, LPARAM lp) {
    ShowWindow(child, (int)lp); return TRUE;
}

static void toggle_game_mode(HWND hwnd) {
    if (IsIconic(hwnd)) ShowWindow(hwnd, SW_RESTORE);
    g_game_mode = !g_game_mode;
    if (g_game_mode) {
        GetWindowRect(hwnd, &g_normal_rect);
        EnumChildWindows(hwnd, show_child_proc, SW_HIDE);
        SetWindowLongW(hwnd, GWL_STYLE, WS_POPUP | WS_VISIBLE);
        SetWindowLongW(hwnd, GWL_EXSTYLE, WS_EX_TOPMOST | WS_EX_LAYERED | WS_EX_TOOLWINDOW);
        SetLayeredWindowAttributes(hwnd, 0, GAME_ALPHA, LWA_ALPHA);
        int x=g_game_x, y=g_game_y;
        if (x<0||y<0) { x=g_normal_rect.left; y=g_normal_rect.top; }
        SetWindowPos(hwnd, HWND_TOPMOST, x, y, g_game_w, g_game_h, SWP_FRAMECHANGED);
    } else {
        RECT gr; GetWindowRect(hwnd, &gr);
        g_game_x=gr.left; g_game_y=gr.top; g_game_w=gr.right-gr.left; g_game_h=gr.bottom-gr.top;
        save_config();
        SetWindowLongW(hwnd, GWL_STYLE, WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU|WS_MINIMIZEBOX|WS_VISIBLE);
        SetWindowLongW(hwnd, GWL_EXSTYLE, WS_EX_TOPMOST);
        int w=g_normal_rect.right-g_normal_rect.left, h=g_normal_rect.bottom-g_normal_rect.top;
        SetWindowPos(hwnd, HWND_TOPMOST, g_normal_rect.left, g_normal_rect.top, w, h, SWP_FRAMECHANGED);
        EnumChildWindows(hwnd, show_child_proc, SW_SHOW);
    }
    InvalidateRect(hwnd, NULL, TRUE);
}

/* ================================================================== */
/*                       KEYBOARD HOOK                                 */
/* ================================================================== */

static void update_hotkey_labels(void);

static LRESULT CALLBACK kb_hook_proc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode == HC_ACTION && g_setting_hk && wParam == WM_KEYDOWN) {
        KBDLLHOOKSTRUCT *p = (KBDLLHOOKSTRUCT *)lParam;
        int vk = (int)p->vkCode;
        if (vk == VK_ESCAPE) { g_setting_hk = 0; update_hotkey_labels(); return 1; }
        if (g_setting_hk == 1) {
            UnregisterHotKey(g_hwnd, ID_HK_TOGGLE);
            g_hk_toggle_vk = vk;
            RegisterHotKey(g_hwnd, ID_HK_TOGGLE, 0, g_hk_toggle_vk);
        } else if (g_setting_hk == 2) {
            UnregisterHotKey(g_hwnd, ID_HK_GAME);
            g_hk_game_vk = vk;
            RegisterHotKey(g_hwnd, ID_HK_GAME, 0, g_hk_game_vk);
        }
        g_setting_hk = 0;
        save_config();
        update_hotkey_labels();
        InvalidateRect(g_hwnd, NULL, TRUE);
        return 1;
    }
    return CallNextHookEx(g_kbhook, nCode, wParam, lParam);
}

/* ================================================================== */
/*                        UI UPDATE HELPERS                            */
/* ================================================================== */

static void update_hotkey_labels(void) {
    WCHAR name[64];
    if (g_setting_hk == 1) lstrcpyW(name, L"\x8BF7\x6309\x952E...");
    else get_vk_name(g_hk_toggle_vk, name, 64);
    SetWindowTextW(g_lbl_hkname, name);
    if (g_setting_hk == 2) lstrcpyW(name, L"\x8BF7\x6309\x952E...");
    else get_vk_name(g_hk_game_vk, name, 64);
    SetWindowTextW(g_lbl_gmhkname, name);
}

static void update_ui_labels(void) {
    WCHAR t[256];
    SetWindowTextW(g_lbl_driver, g_drv_ok
        ? L"\x2714  Interception \x9A71\x52A8\x5C31\x7EEA"
        : L"\x2716  Interception \x9A71\x52A8\x672A\x5C31\x7EEA");

    if (g_active) {
        SetWindowTextW(g_lbl_status, L"\x72B6\x6001:  \x2705 \x5DF2\x5F00\x542F");
        SetWindowTextW(g_btn_startstop, L"\x5173\x95ED\x8FDE\x53D1");
    } else {
        SetWindowTextW(g_lbl_status, L"\x72B6\x6001:  \x5DF2\x5173\x95ED");
        SetWindowTextW(g_btn_startstop, L"\x5F00\x542F\x8FDE\x53D1");
    }

    LONG pps = g_pps;
    if (g_active && pps > 0) wsprintfW(t, L"\x901F\x5EA6:  %ld \x6B21/\x79D2", pps);
    else lstrcpyW(t, L"\x901F\x5EA6:  --");
    SetWindowTextW(g_lbl_speed, t);

    if (g_delay >= 10000) wsprintfW(t, L"%dms", g_delay/1000);
    else if (g_delay < 1000) wsprintfW(t, L"%d\x03BCs", g_delay);
    else { int mw=g_delay/1000, mf=(g_delay%1000)/100; wsprintfW(t, L"%d\x03BCs (%d.%dms)", g_delay, mw, mf); }
    SetWindowTextW(g_lbl_delay, t);

    WCHAR keys[256]={0}; int count;
    if (g_active) {
        if (g_mode==MODE_HYBRID) {
            int c1 = build_toggled_list(keys,240,3);
            if (c1>0 && c1<3) { int len=lstrlenW(keys); keys[len]=L','; keys[len+1]=L' '; keys[len+2]=0; }
            int c2 = build_active_list(keys+lstrlenW(keys), 240-lstrlenW(keys), 3);
            count = c1+c2;
        } else count = (g_mode==MODE_HOLD) ? build_active_list(keys,240,5) : build_toggled_list(keys,240,5);
        if (count>0) wsprintfW(t, L"\x8FDE\x53D1\x4E2D:  %s", keys);
        else lstrcpyW(t, L"\x8FDE\x53D1\x4E2D:  --");
    } else lstrcpyW(t, L"\x8FDE\x53D1\x4E2D:  --");
    SetWindowTextW(g_lbl_repeat, t);
}

/* ================================================================== */
/*                   DELAY ADJUST + BUTTON REPEAT                      */
/* ================================================================== */

static void do_delay_adjust(int btn_id) {
    int min_delay = g_turbo ? 1 : 1000;
    if (btn_id == IDC_BTN_DEC) {
        if (g_delay > 10000) g_delay -= 1000;
        else if (g_delay > 100) g_delay -= 100;
        else if (g_turbo && g_delay > 10) g_delay -= 10;
        else if (g_turbo && g_delay > 1) g_delay -= 1;
        if (g_delay < min_delay) g_delay = min_delay;
    } else {
        if (g_turbo && g_delay < 10) g_delay += 1;
        else if (g_turbo && g_delay < 100) g_delay += 10;
        else if (g_delay < 10000) g_delay += 100;
        else if (g_delay < 1000000) g_delay += 1000;
    }
    save_config();
    update_ui_labels();
}

static LRESULT CALLBACK btn_repeat_proc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    int id = GetDlgCtrlID(hwnd);
    switch (msg) {
    case WM_LBUTTONDOWN:
        do_delay_adjust(id);
        g_repeat_btn = id;
        g_repeat_count = 0;
        SetTimer(g_hwnd, IDT_REPEAT, 200, NULL);
        break;
    case WM_LBUTTONUP:
    case WM_CAPTURECHANGED:
        if (g_repeat_btn == id) {
            KillTimer(g_hwnd, IDT_REPEAT);
            g_repeat_btn = 0;
        }
        break;
    }
    return CallWindowProcW(g_orig_btn_proc, hwnd, msg, wp, lp);
}

/* ================================================================== */
/*                          PAINTING                                   */
/* ================================================================== */

static void paint_keyboard(HDC hdc) {
    HFONT old_font = (HFONT)SelectObject(hdc, g_font_kb);
    HPEN old_pen = (HPEN)SelectObject(hdc, g_pen_border);
    SetBkMode(hdc, TRANSPARENT);

    for (int i = 0; i < (int)N_KEYS; i++) {
        KeyRect *kr = &g_krects[i];
        int vk = kr->vk;
        HBRUSH fill_br; COLORREF text;

        if (is_skippable(vk)) {
            fill_br = g_br_skip; text = RGB(180,180,180);
        } else if (g_mode == MODE_HOLD) {
            if (vk > 0 && vk < 256 && g_exclude[vk]) { fill_br = g_br_excl; text = RGB(255,255,255); }
            else { fill_br = g_br_hold; text = RGB(255,255,255); }
        } else if (g_mode == MODE_CUSTOM) {
            if (vk > 0 && vk < 256 && g_custom_keys[vk]) { fill_br = g_br_custom; text = RGB(255,255,255); }
            else { fill_br = g_br_unsel; text = RGB(40,40,40); }
        } else {
            BOOL is_cust = (vk > 0 && vk < 256 && g_custom_keys[vk]);
            BOOL is_excl = (vk > 0 && vk < 256 && g_exclude[vk]);
            if (is_cust)       { fill_br = g_br_custom; text = RGB(255,255,255); }
            else if (!is_excl) { fill_br = g_br_hold;   text = RGB(255,255,255); }
            else               { fill_br = g_br_excl;   text = RGB(255,255,255); }
        }

        RECT r = kr->rc;
        HBRUSH old_br = (HBRUSH)SelectObject(hdc, fill_br);
        Rectangle(hdc, r.left, r.top, r.right, r.bottom);
        SelectObject(hdc, old_br);

        SetTextColor(hdc, text);
        DrawTextW(hdc, kr->label, -1, &r, DT_CENTER|DT_VCENTER|DT_SINGLELINE);
    }

    SelectObject(hdc, old_pen);
    SelectObject(hdc, old_font);
}

static void paint_keyboard_legend(HDC hdc, int y) {
    HFONT old = (HFONT)SelectObject(hdc, g_font_legend);
    SetBkMode(hdc, TRANSPARENT);
    int x = KB_X;
    RECT sq;

    sq.left=x; sq.top=y; sq.right=x+12; sq.bottom=y+12;
    SetTextColor(hdc, RGB(0,0,0));

    if (g_mode == MODE_HOLD) {
        FillRect(hdc, &sq, g_br_hold);
        TextOutW(hdc, x+16, y-1, L"\x8FDE\x53D1", 2);
        x += 70; sq.left=x; sq.right=x+12;
        FillRect(hdc, &sq, g_br_excl);
        TextOutW(hdc, x+16, y-1, L"\x4E0D\x8FDE\x53D1", 3);
    } else if (g_mode == MODE_CUSTOM) {
        FillRect(hdc, &sq, g_br_custom);
        TextOutW(hdc, x+16, y-1, L"\x5DF2\x9009", 2);
        x += 70; sq.left=x; sq.right=x+12;
        FillRect(hdc, &sq, g_br_unsel);
        TextOutW(hdc, x+16, y-1, L"\x672A\x9009", 2);
    } else {
        FillRect(hdc, &sq, g_br_custom);
        TextOutW(hdc, x+16, y-1, L"\x6307\x5B9A", 2);
        x += 60; sq.left=x; sq.right=x+12;
        FillRect(hdc, &sq, g_br_hold);
        TextOutW(hdc, x+16, y-1, L"\x6309\x4F4F", 2);
        x += 60; sq.left=x; sq.right=x+12;
        FillRect(hdc, &sq, g_br_excl);
        TextOutW(hdc, x+16, y-1, L"\x6392\x9664", 2);
    }

    x += 85; sq.left=x; sq.right=x+12;
    FillRect(hdc, &sq, g_br_skip);
    TextOutW(hdc, x+16, y-1, L"\x70ED\x952E", 2);

    SelectObject(hdc, old);
}

static void paint(HWND hwnd) {
    PAINTSTRUCT ps;
    HDC hdcScreen = BeginPaint(hwnd, &ps);
    RECT rc; GetClientRect(hwnd, &rc);
    int w = rc.right - rc.left, h = rc.bottom - rc.top;

    if (w != g_paint_w || h != g_paint_h || !g_paint_memdc) {
        if (g_paint_memdc) {
            if (g_paint_oldbm) SelectObject(g_paint_memdc, g_paint_oldbm);
            DeleteDC(g_paint_memdc);
        }
        if (g_paint_bitmap) DeleteObject(g_paint_bitmap);
        g_paint_memdc = CreateCompatibleDC(hdcScreen);
        g_paint_bitmap = CreateCompatibleBitmap(hdcScreen, w, h);
        g_paint_oldbm = (HBITMAP)SelectObject(g_paint_memdc, g_paint_bitmap);
        g_paint_w = w;
        g_paint_h = h;
    }
    HDC hdc = g_paint_memdc;

    if (g_game_mode) {
        FillRect(hdc, &rc, (HBRUSH)GetStockObject(BLACK_BRUSH));
        SetBkMode(hdc, TRANSPARENT);
        HFONT old = (HFONT)SelectObject(hdc, g_font_game);
        WCHAR t[256]; int y=10, lh=22, x=8;
        if (g_active) { SetTextColor(hdc, RGB(0,255,0)); lstrcpyW(t, L"\x72B6\x6001: ON"); }
        else { SetTextColor(hdc, RGB(255,60,60)); lstrcpyW(t, L"\x72B6\x6001: OFF"); }
        TextOutW(hdc, x, y, t, lstrlenW(t)); y+=lh;
        SetTextColor(hdc, RGB(200,200,200));
        WCHAR keys[256]={0}; int count;
        if (g_active) {
            if (g_mode==MODE_HYBRID) { int c1=build_toggled_list(keys,240,3); if(c1>0&&c1<3){int l=lstrlenW(keys);keys[l]=L',';keys[l+1]=L' ';keys[l+2]=0;} build_active_list(keys+lstrlenW(keys),240-lstrlenW(keys),3); count=lstrlenW(keys)>0?1:0; }
            else { count=(g_mode==MODE_HOLD)?build_active_list(keys,240,5):build_toggled_list(keys,240,5); }
            if(count>0) wsprintfW(t,L"\x8FDE\x53D1: %s",keys); else lstrcpyW(t,L"\x8FDE\x53D1: --"); }
        else lstrcpyW(t, L"\x8FDE\x53D1: --");
        TextOutW(hdc, x, y, t, lstrlenW(t)); y+=lh;
        LONG pps=g_pps;
        if (g_active&&pps>0) { SetTextColor(hdc,RGB(0,255,100)); wsprintfW(t,L"\x901F\x5EA6: %ld/s",pps); }
        else { SetTextColor(hdc,RGB(120,120,120)); lstrcpyW(t,L"\x901F\x5EA6: --"); }
        TextOutW(hdc, x, y, t, lstrlenW(t));
        SelectObject(hdc, old);
    } else {
        FillRect(hdc, &rc, (HBRUSH)GetStockObject(WHITE_BRUSH));

        RECT kbBg;
        kbBg.left = KB_X - 5; kbBg.top = KB_Y - 8;
        kbBg.right = KB_X + 96*KU/4 + 5; kbBg.bottom = KB_Y + 26*KU/4 + 30;
        FillRect(hdc, &kbBg, g_br_kbbg);

        paint_keyboard(hdc);
        paint_keyboard_legend(hdc, KB_Y + 26*KU/4 + 8);
    }

    BitBlt(hdcScreen, 0, 0, w, h, hdc, 0, 0, SRCCOPY);
    EndPaint(hwnd, &ps);
}

/* ================================================================== */
/*                      WINDOW PROCEDURE                               */
/* ================================================================== */

#define LP_X 15
#define LP_W 730

static HWND make_label(HWND parent, int id, const WCHAR *text, int x, int y, int w, int h, DWORD style) {
    return CreateWindowW(L"STATIC", text, WS_CHILD|WS_VISIBLE|style, x, y, w, h, parent, (HMENU)(INT_PTR)id, NULL, NULL);
}

static HWND make_button(HWND parent, int id, const WCHAR *text, int x, int y, int w, int h) {
    return CreateWindowW(L"BUTTON", text, WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, x, y, w, h, parent, (HMENU)(INT_PTR)id, NULL, NULL);
}

static void create_controls(HWND hwnd) {
    HFONT hf = g_font_ui;
    HWND lbl;
    int y = KB_Y + 26*KU/4 + 35;

    g_lbl_driver = make_label(hwnd, IDC_LABEL_DRIVER, L"", LP_X, y, LP_W, 20, 0);
    SendMessageW(g_lbl_driver, WM_SETFONT, (WPARAM)hf, TRUE);
    y += 24;
    make_label(hwnd, 0, NULL, LP_X, y, LP_W, 2, SS_ETCHEDHORZ);
    y += 8;

    lbl = make_label(hwnd, 0, L"\x8FDE\x53D1\x6A21\x5F0F:", LP_X, y, 80, 18, 0);
    SendMessageW(lbl, WM_SETFONT, (WPARAM)hf, TRUE);
    lbl = make_label(hwnd, 0, L"\x6309\x952E\x95F4\x9694:", 510, y, 80, 18, 0);
    SendMessageW(lbl, WM_SETFONT, (WPARAM)hf, TRUE);
    y += 22;

    g_radio_hold = CreateWindowW(L"BUTTON", L"\x6309\x4F4F\x8FDE\x53D1 (Hold)",
        WS_CHILD|WS_VISIBLE|BS_AUTORADIOBUTTON|WS_GROUP, LP_X+5, y, 148, 20, hwnd, (HMENU)IDC_RADIO_HOLD, NULL, NULL);
    SendMessageW(g_radio_hold, WM_SETFONT, (WPARAM)hf, TRUE);
    g_radio_toggle = CreateWindowW(L"BUTTON", L"\x6307\x5B9A\x8FDE\x53D1 (Toggle)",
        WS_CHILD|WS_VISIBLE|BS_AUTORADIOBUTTON, LP_X+158, y, 155, 20, hwnd, (HMENU)IDC_RADIO_TOGGLE, NULL, NULL);
    SendMessageW(g_radio_toggle, WM_SETFONT, (WPARAM)hf, TRUE);
    g_radio_hybrid = CreateWindowW(L"BUTTON", L"\x6DF7\x5408\x6A21\x5F0F (Hybrid)",
        WS_CHILD|WS_VISIBLE|BS_AUTORADIOBUTTON, LP_X+318, y, 160, 20, hwnd, (HMENU)IDC_RADIO_HYBRID, NULL, NULL);
    SendMessageW(g_radio_hybrid, WM_SETFONT, (WPARAM)hf, TRUE);
    {
        HWND sel = (g_mode==MODE_HOLD) ? g_radio_hold : (g_mode==MODE_CUSTOM) ? g_radio_toggle : g_radio_hybrid;
        SendMessageW(sel, BM_SETCHECK, BST_CHECKED, 0);
    }

    HWND bdec = make_button(hwnd, IDC_BTN_DEC, L"-", 510, y-1, 35, 24);
    SendMessageW(bdec, WM_SETFONT, (WPARAM)hf, TRUE);
    g_lbl_delay = make_label(hwnd, IDC_LABEL_DELAY, L"", 550, y, 100, 20, SS_CENTER);
    SendMessageW(g_lbl_delay, WM_SETFONT, (WPARAM)hf, TRUE);
    HWND binc = make_button(hwnd, IDC_BTN_INC, L"+", 655, y-1, 35, 24);
    SendMessageW(binc, WM_SETFONT, (WPARAM)hf, TRUE);
    g_orig_btn_proc = (WNDPROC)SetWindowLongPtrW(bdec, GWLP_WNDPROC, (LONG_PTR)btn_repeat_proc);
    SetWindowLongPtrW(binc, GWLP_WNDPROC, (LONG_PTR)btn_repeat_proc);
    y += 26;
    make_label(hwnd, 0, NULL, LP_X, y, LP_W, 2, SS_ETCHEDHORZ);
    y += 8;

    g_chk_keylock = CreateWindowW(L"BUTTON", L"\x6682\x505C\x6A21\x5F0F",
        WS_CHILD|WS_VISIBLE|BS_AUTOCHECKBOX, LP_X+5, y, 100, 20, hwnd, (HMENU)(INT_PTR)IDC_CHK_KEYLOCK, NULL, NULL);
    SendMessageW(g_chk_keylock, WM_SETFONT, (WPARAM)hf, TRUE);
    if (g_key_lock) SendMessageW(g_chk_keylock, BM_SETCHECK, BST_CHECKED, 0);
    g_chk_turbo = CreateWindowW(L"BUTTON", L"\x6781\x901F\x6A21\x5F0F",
        WS_CHILD|WS_VISIBLE|BS_AUTOCHECKBOX, LP_X+115, y, 100, 20, hwnd, (HMENU)(INT_PTR)IDC_CHK_TURBO, NULL, NULL);
    SendMessageW(g_chk_turbo, WM_SETFONT, (WPARAM)hf, TRUE);
    if (g_turbo) SendMessageW(g_chk_turbo, BM_SETCHECK, BST_CHECKED, 0);
    y += 24;

    g_lbl_status = make_label(hwnd, IDC_LABEL_STATUS, L"", LP_X, y, 190, 18, 0);
    SendMessageW(g_lbl_status, WM_SETFONT, (WPARAM)hf, TRUE);
    g_lbl_repeat = make_label(hwnd, IDC_LABEL_REPEAT, L"", LP_X+195, y, 290, 18, 0);
    SendMessageW(g_lbl_repeat, WM_SETFONT, (WPARAM)hf, TRUE);
    g_lbl_speed = make_label(hwnd, IDC_LABEL_SPEED, L"", LP_X+490, y, 230, 18, 0);
    SendMessageW(g_lbl_speed, WM_SETFONT, (WPARAM)hf, TRUE);
    y += 22;
    make_label(hwnd, 0, NULL, LP_X, y, LP_W, 2, SS_ETCHEDHORZ);
    y += 8;

    g_btn_startstop = make_button(hwnd, IDC_BTN_STARTSTOP, L"\x5F00\x542F\x8FDE\x53D1", LP_X, y, 140, 28);
    SendMessageW(g_btn_startstop, WM_SETFONT, (WPARAM)hf, TRUE);
    HWND bgm = make_button(hwnd, IDC_BTN_GAMEMODE, L"\x6E38\x620F\x6A21\x5F0F", LP_X+150, y, 110, 28);
    SendMessageW(bgm, WM_SETFONT, (WPARAM)hf, TRUE);
    lbl = make_label(hwnd, 0, L"\x542F\x52A8:", LP_X+280, y+5, 40, 18, 0);
    SendMessageW(lbl, WM_SETFONT, (WPARAM)hf, TRUE);
    g_lbl_hkname = make_label(hwnd, IDC_LABEL_HKNAME, L"", LP_X+322, y+4, 75, 20, SS_CENTER|WS_BORDER);
    SendMessageW(g_lbl_hkname, WM_SETFONT, (WPARAM)hf, TRUE);
    HWND bsh = make_button(hwnd, IDC_BTN_SETHK, L"\x6539", LP_X+400, y+2, 30, 24);
    SendMessageW(bsh, WM_SETFONT, (WPARAM)hf, TRUE);
    lbl = make_label(hwnd, 0, L"\x6E38\x620F:", LP_X+450, y+5, 40, 18, 0);
    SendMessageW(lbl, WM_SETFONT, (WPARAM)hf, TRUE);
    g_lbl_gmhkname = make_label(hwnd, IDC_LABEL_GMHKNAME, L"", LP_X+492, y+4, 75, 20, SS_CENTER|WS_BORDER);
    SendMessageW(g_lbl_gmhkname, WM_SETFONT, (WPARAM)hf, TRUE);
    HWND bsg = make_button(hwnd, IDC_BTN_SETGMHK, L"\x6539", LP_X+570, y+2, 30, 24);
    SendMessageW(bsg, WM_SETFONT, (WPARAM)hf, TRUE);
    HWND buninst = make_button(hwnd, IDC_BTN_UNINSTALL, L"\x5378\x8F7D\x9A71\x52A8", LP_X+610, y, 80, 28);
    SendMessageW(buninst, WM_SETFONT, (WPARAM)hf, TRUE);
    HWND babout = make_button(hwnd, IDC_BTN_ABOUT, L"\x5173\x4E8E", LP_X+695, y, 50, 28);
    SendMessageW(babout, WM_SETFONT, (WPARAM)hf, TRUE);

    update_hotkey_labels();
    update_ui_labels();
}

static int keyboard_hittest(int mx, int my) {
    for (int i = 0; i < (int)N_KEYS; i++) {
        RECT *r = &g_krects[i].rc;
        if (mx >= r->left && mx < r->right && my >= r->top && my < r->bottom)
            return i;
    }
    return -1;
}

/* ================================================================== */
/*                         ABOUT DIALOG                               */
/* ================================================================== */

#define ABOUT_CW  390
#define ABOUT_CH  180

static HWND g_about_hwnd = NULL;

static LRESULT CALLBACK about_wndproc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_COMMAND:
        if (LOWORD(wp) == IDOK || LOWORD(wp) == IDCANCEL)
            { DestroyWindow(hwnd); return 0; }
        break;
    case WM_NOTIFY: {
        NMHDR *nm = (NMHDR *)lp;
        if (nm->code == NM_CLICK || nm->code == NM_RETURN) {
            NMLINK *nml = (NMLINK *)lp;
            ShellExecuteW(NULL, L"open", nml->item.szUrl, NULL, NULL, SW_SHOWNORMAL);
            return 0;
        }
        break;
    }
    case WM_CLOSE:
        DestroyWindow(hwnd);
        return 0;
    case WM_DESTROY:
        g_about_hwnd = NULL;
        EnableWindow(g_hwnd, TRUE);
        SetForegroundWindow(g_hwnd);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

static void show_about_dialog(HWND hwndParent) {
    WNDCLASSW wc;
    RECT rc;
    int wx, wy;
    HWND lbl, link1, lbl2, link2, btn;
    DWORD style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU;

    if (g_about_hwnd) { SetForegroundWindow(g_about_hwnd); return; }

    memset(&wc, 0, sizeof(wc));
    wc.lpfnWndProc = about_wndproc;
    wc.hInstance = GetModuleHandleW(NULL);
    wc.hCursor = LoadCursorW(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = L"AboutWndClass";
    RegisterClassW(&wc);

    rc.left = 0; rc.top = 0; rc.right = ABOUT_CW; rc.bottom = ABOUT_CH;
    AdjustWindowRectEx(&rc, style, FALSE, 0);
    wx = rc.right - rc.left;
    wy = rc.bottom - rc.top;

    {
        RECT pr;
        int px, py;
        GetWindowRect(hwndParent, &pr);
        px = pr.left + ((pr.right - pr.left) - wx) / 2;
        py = pr.top + ((pr.bottom - pr.top) - wy) / 2;
        g_about_hwnd = CreateWindowExW(0, L"AboutWndClass",
            L"\x5173\x4E8E", style,
            px, py, wx, wy,
            hwndParent, NULL, GetModuleHandleW(NULL), NULL);
    }
    if (!g_about_hwnd) return;

    HFONT hf = g_font_ui;

    lbl = CreateWindowW(L"STATIC",
        L"by\x8106\x76AE\x5377\xFF0C\x5F00\x6E90\x5730\x5740:",
        WS_CHILD | WS_VISIBLE, 20, 15, 350, 20, g_about_hwnd, NULL, NULL, NULL);
    SendMessageW(lbl, WM_SETFONT, (WPARAM)hf, TRUE);

    link1 = CreateWindowW(L"SysLink",
        L"<a href=\"https://github.com/mh58373397-byte/jx3anjian\">"
        L"https://github.com/mh58373397-byte/jx3anjian</a>",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        20, 38, 350, 20, g_about_hwnd, (HMENU)501, NULL, NULL);
    SendMessageW(link1, WM_SETFONT, (WPARAM)hf, TRUE);

    lbl2 = CreateWindowW(L"STATIC",
        L"jx3box\x5730\x5740:",
        WS_CHILD | WS_VISIBLE, 20, 72, 350, 20, g_about_hwnd, NULL, NULL, NULL);
    SendMessageW(lbl2, WM_SETFONT, (WPARAM)hf, TRUE);

    link2 = CreateWindowW(L"SysLink",
        L"<a href=\"https://www.jx3box.com/tool/106371\">"
        L"https://www.jx3box.com/tool/106371</a>",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        20, 95, 350, 20, g_about_hwnd, (HMENU)502, NULL, NULL);
    SendMessageW(link2, WM_SETFONT, (WPARAM)hf, TRUE);

    btn = CreateWindowW(L"BUTTON", L"\x786E\x5B9A",
        WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON | WS_TABSTOP,
        155, 135, 80, 30, g_about_hwnd, (HMENU)IDOK, NULL, NULL);
    SendMessageW(btn, WM_SETFONT, (WPARAM)hf, TRUE);

    EnableWindow(hwndParent, FALSE);
    ShowWindow(g_about_hwnd, SW_SHOW);
    UpdateWindow(g_about_hwnd);
}

static LRESULT CALLBACK wndproc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_NCHITTEST:
        if (g_game_mode) {
            LRESULT hit = DefWindowProcW(hwnd, msg, wp, lp);
            if (hit == HTCLIENT) return HTCAPTION;
            return hit;
        }
        break;

    case WM_CREATE:
        RegisterHotKey(hwnd, ID_HK_TOGGLE, 0, g_hk_toggle_vk);
        RegisterHotKey(hwnd, ID_HK_GAME, 0, g_hk_game_vk);
        SetTimer(hwnd, IDT_UI, UI_TIMER_MS, NULL);
        create_controls(hwnd);
        return 0;

    case WM_TIMER:
        if (wp == IDT_UI) {
            update_ui_labels();
            if (g_active) InvalidateRect(hwnd, NULL, FALSE);
        } else if (wp == IDT_REPEAT) {
            if (g_repeat_btn && (GetAsyncKeyState(VK_LBUTTON) & 0x8000)) {
                do_delay_adjust(g_repeat_btn);
                g_repeat_count++;
                int interval = (g_repeat_count > 20) ? 10 :
                               (g_repeat_count > 10) ? 20 :
                               (g_repeat_count > 3)  ? 40 : 80;
                SetTimer(hwnd, IDT_REPEAT, interval, NULL);
            } else {
                KillTimer(hwnd, IDT_REPEAT);
                g_repeat_btn = 0;
            }
        }
        return 0;

    case WM_HOTKEY:
        if (wp == ID_HK_TOGGLE) toggle_active();
        else if (wp == ID_HK_GAME) toggle_game_mode(hwnd);
        return 0;

    case WM_COMMAND:
        switch (LOWORD(wp)) {
        case IDC_RADIO_HOLD:
            release_toggled(); g_mode = MODE_HOLD; save_config();
            InvalidateRect(hwnd, NULL, TRUE); break;
        case IDC_RADIO_TOGGLE:
            release_toggled(); g_mode = MODE_CUSTOM; save_config();
            InvalidateRect(hwnd, NULL, TRUE); break;
        case IDC_RADIO_HYBRID:
            release_toggled(); g_mode = MODE_HYBRID; save_config();
            InvalidateRect(hwnd, NULL, TRUE); break;
        case IDC_BTN_DEC:
        case IDC_BTN_INC:
            break;
        case IDC_BTN_STARTSTOP:
            toggle_active(); break;
        case IDC_BTN_GAMEMODE:
            toggle_game_mode(hwnd); break;
        case IDC_BTN_EXIT:
            DestroyWindow(hwnd); break;
        case IDC_BTN_SETHK:
            g_setting_hk = 1; update_hotkey_labels(); break;
        case IDC_BTN_SETGMHK:
            g_setting_hk = 2; update_hotkey_labels(); break;
        case IDC_CHK_KEYLOCK: {
            BOOL want = (SendMessageW(g_chk_keylock, BM_GETCHECK, 0, 0) == BST_CHECKED);
            if (want) {
                int r = MessageBoxW(hwnd,
                    L"\x5F00\x542F\x6682\x505C\x6A21\x5F0F\x540E\xFF0C\x542F\x52A8/\x505C\x6B62"
                    L"\x7684\x903B\x8F91\x4F1A\x53D8\x6210\x542F\x52A8/\x6682\x505C\xFF0C"
                    L"\x53EF\x4EE5\x7528\x4E8E\x8BB0\x5FC6\x6307\x5B9A\x8FDE\x53D1\x7684\x6309\x952E\x3002\n\n"
                    L"\x786E\x8BA4\x5F00\x542F\x5417\xFF1F",
                    L"\x6682\x505C\x6A21\x5F0F", MB_YESNO | MB_ICONQUESTION);
                if (r == IDYES) { g_key_lock = TRUE; }
                else { g_key_lock = FALSE; SendMessageW(g_chk_keylock, BM_SETCHECK, BST_UNCHECKED, 0); }
            } else {
                g_key_lock = FALSE;
                if (!g_active) release_toggled();
            }
            save_config();
            break;
        }
        case IDC_CHK_TURBO: {
            BOOL want = (SendMessageW(g_chk_turbo, BM_GETCHECK, 0, 0) == BST_CHECKED);
            if (want) {
                int r = MessageBoxW(hwnd,
                    L"\x6781\x901F\x6A21\x5F0F\x5141\x8BB8\x6309\x952E\x95F4\x9694\x4F4E\x4E8E" L"1ms\xFF0C"
                    L"\x4F46\x8FC7\x4F4E\x7684\x95F4\x9694\x53EF\x80FD\x5BFC\x81F4\x7535\x8111\x6B7B\x673A\x3002\n\n"
                    L"\x786E\x5B9A\x8981\x5F00\x542F\x5417\xFF1F",
                    L"\x6781\x901F\x6A21\x5F0F", MB_YESNO | MB_ICONWARNING);
                if (r == IDYES) { g_turbo = TRUE; }
                else { g_turbo = FALSE; SendMessageW(g_chk_turbo, BM_SETCHECK, BST_UNCHECKED, 0); }
            } else {
                g_turbo = FALSE;
                if (g_delay < 1000) { g_delay = 1000; update_ui_labels(); }
            }
            save_config();
            break;
        }
        case IDC_BTN_UNINSTALL:
            try_uninstall();
            break;
        case IDC_BTN_ABOUT:
            show_about_dialog(hwnd);
            break;
        }
        return 0;

    case WM_NCRBUTTONDOWN:
        if (g_game_mode) return 0;
        break;
    case WM_NCRBUTTONUP:
        if (g_game_mode) {
            HMENU hMenu = CreatePopupMenu();
            AppendMenuW(hMenu, MF_STRING, IDM_GM_TOGGLE,
                g_active ? L"\x5173\x95ED\x8FDE\x53D1" : L"\x5F00\x542F\x8FDE\x53D1");
            AppendMenuW(hMenu, MF_STRING, IDM_GM_SHOWMAIN, L"\x663E\x793A\x4E3B\x754C\x9762");
            AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
            AppendMenuW(hMenu, MF_STRING, IDM_GM_EXIT, L"\x9000\x51FA");
            POINT pt; GetCursorPos(&pt);
            SetForegroundWindow(hwnd);
            int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_NONOTIFY, pt.x, pt.y, 0, hwnd, NULL);
            DestroyMenu(hMenu);
            if (cmd == IDM_GM_TOGGLE) toggle_active();
            else if (cmd == IDM_GM_SHOWMAIN) toggle_game_mode(hwnd);
            else if (cmd == IDM_GM_EXIT) DestroyWindow(hwnd);
            return 0;
        }
        break;

    case WM_LBUTTONDOWN: {
        int mx = (short)LOWORD(lp), my = (short)HIWORD(lp);
        if (!g_game_mode) {
            int idx = keyboard_hittest(mx, my);
            if (idx >= 0) {
                int vk = g_krects[idx].vk;
                if (!is_skippable(vk) && vk > 0 && vk < 256) {
                    if (g_mode == MODE_HOLD) {
                        g_exclude[vk] = !g_exclude[vk];
                    } else if (g_mode == MODE_CUSTOM) {
                        g_custom_keys[vk] = !g_custom_keys[vk];
                    } else {
                        if (g_custom_keys[vk])          { g_custom_keys[vk]=FALSE; g_exclude[vk]=TRUE; }
                        else if (!g_exclude[vk])         { g_custom_keys[vk]=TRUE; }
                        else                             { g_exclude[vk]=FALSE; }
                    }
                    save_config();
                    InvalidateRect(hwnd, NULL, TRUE);
                }
            }
        }
        return 0;
    }

    case WM_CTLCOLORSTATIC: {
        HDC hdcCtl = (HDC)wp;
        SetBkColor(hdcCtl, RGB(255,255,255));
        SetBkMode(hdcCtl, TRANSPARENT);
        return (LRESULT)GetStockObject(WHITE_BRUSH);
    }

    case WM_ERASEBKGND:
        return 1;

    case WM_PAINT:
        paint(hwnd);
        return 0;

    case WM_CLOSE:
        g_active = FALSE;
        if (g_rthread) { WaitForSingleObject(g_rthread,2000); CloseHandle(g_rthread); g_rthread=NULL; }
        release_toggled();
        if (g_game_mode) { RECT gr; GetWindowRect(hwnd,&gr); g_game_x=gr.left; g_game_y=gr.top; g_game_w=gr.right-gr.left; g_game_h=gr.bottom-gr.top; }
        save_config();
        DestroyWindow(hwnd);
        return 0;

    case WM_DESTROY:
        KillTimer(hwnd, IDT_UI);
        UnregisterHotKey(hwnd, ID_HK_TOGGLE);
        UnregisterHotKey(hwnd, ID_HK_GAME);
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

/* ================================================================== */
/*                      INTERCEPTION INIT                              */
/* ================================================================== */

static void init_interception(void) {
    g_ctx = interception_create_context();
    if (!g_ctx) return;
    wchar_t hwid[256];
    for (int i = 0; i < INTERCEPTION_MAX_KEYBOARD; i++) {
        InterceptionDevice dev = INTERCEPTION_KEYBOARD(i);
        if (interception_get_hardware_id(g_ctx, dev, hwid, sizeof(hwid)) > 0) { g_drv_ok = TRUE; break; }
    }
}

/* ================================================================== */
/*                         ENTRY POINT                                 */
/* ================================================================== */

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE hPrev, LPWSTR cmd, int nShow) {
    (void)hPrev;
    LPWSTR p = cmd;
    while (*p == L' ' || *p == L'\t') p++;
    if (wcsncmp(p, L"/install", 8) == 0) return do_install();
    if (wcsncmp(p, L"/uninstall", 10) == 0) return do_uninstall();

    INITCOMMONCONTROLSEX icex = { sizeof(icex), ICC_STANDARD_CLASSES | ICC_LINK_CLASS };
    InitCommonControlsEx(&icex);

    if (!load_interception_dll()) {
        MessageBoxW(NULL, L"\x65E0\x6CD5\x52A0\x8F7D interception.dll", L"\x9519\x8BEF", MB_ICONERROR);
        return 1;
    }

    InitializeCriticalSection(&g_cs_active);
    QueryPerformanceFrequency(&g_qpc_freq);
    init_exclude_defaults();
    init_custom_keys_defaults();
    load_config();
    init_interception();
    init_keyrects();
    init_gdi_cache();

    if (!g_drv_ok) try_auto_install();
    if (g_ctx) g_ithread = CreateThread(NULL, 0, intercept_proc, NULL, 0, NULL);

    WNDCLASSW wc = {0};
    wc.lpfnWndProc   = wndproc;
    wc.hInstance      = hInst;
    wc.hCursor       = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH);
    wc.lpszClassName = L"AutoKeyClass";
    RegisterClassW(&wc);

    int cw = KB_X + 96*KU/4 + 10;
    int ch = 434;
    RECT wr = {0, 0, cw, ch};
    AdjustWindowRectEx(&wr, WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU|WS_MINIMIZEBOX, FALSE, WS_EX_TOPMOST);
    int ww = wr.right - wr.left, wh = wr.bottom - wr.top;
    int sx = (GetSystemMetrics(SM_CXSCREEN) - ww) / 2;
    int sy = (GetSystemMetrics(SM_CYSCREEN) - wh) / 2;

    /* UTF-8: 丐帮高手v3 - convert at runtime to avoid source encoding issues */
    static const char title_utf8[] = "\xE4\xB8\x90\xE5\xB8\xAE\xE9\xAB\x98\xE6\x89\x8Bv3.2";
    WCHAR title_w[32];
    MultiByteToWideChar(CP_UTF8, 0, title_utf8, -1, title_w, 32);
    g_hwnd = CreateWindowExW(WS_EX_TOPMOST, L"AutoKeyClass",
        title_w,
        WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU|WS_MINIMIZEBOX,
        sx, sy, ww, wh, NULL, NULL, hInst, NULL);

    g_kbhook = SetWindowsHookExW(WH_KEYBOARD_LL, kb_hook_proc, hInst, 0);

    ShowWindow(g_hwnd, nShow);
    UpdateWindow(g_hwnd);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) { TranslateMessage(&msg); DispatchMessage(&msg); }

    g_quit = TRUE; g_active = FALSE;
    if (g_kbhook) UnhookWindowsHookEx(g_kbhook);
    release_toggled();
    if (g_rthread) { WaitForSingleObject(g_rthread, 2000); CloseHandle(g_rthread); }
    if (g_ithread) { WaitForSingleObject(g_ithread, 1000); CloseHandle(g_ithread); }
    if (g_ctx) interception_destroy_context(g_ctx);
    cleanup_gdi_cache();
    DeleteCriticalSection(&g_cs_active);
    if (g_dll) { FreeLibrary(g_dll); g_dll = NULL; }
    DeleteFileW(g_dll_path);
    return 0;
}
