#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN
#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <commdlg.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <string.h>
#include <stdio.h>
#include <stdlib.h>

#include "macro.h"

#pragma comment(lib, "comdlg32.lib")

/* Externs from main.c */
extern InterceptionContext g_ctx;
extern BOOL                g_drv_ok;
extern HFONT               g_font_ui;
extern LARGE_INTEGER       g_qpc_freq;
extern HWND                g_hwnd;
extern int                 g_hk_toggle_vk;
extern int                 g_hk_game_vk;
extern void                save_config(void);

typedef int (*PFN_send)(InterceptionContext, InterceptionDevice, const InterceptionStroke*, unsigned int);
typedef int (*PFN_is_keyboard)(InterceptionDevice);
typedef int (*PFN_is_mouse)(InterceptionDevice);
typedef unsigned int (*PFN_get_hardware_id)(InterceptionContext, InterceptionDevice, void*, unsigned int);
extern PFN_send              pfn_send;
extern PFN_is_keyboard       pfn_is_keyboard;
extern PFN_is_mouse          pfn_is_mouse;
extern PFN_get_hardware_id   pfn_get_hardware_id;

/* ================================================================== */
/*                        CONTROL IDS                                  */
/* ================================================================== */

#define IDC_M_COMBO_SLOT        300
#define IDC_M_BTN_NEW           301
#define IDC_M_BTN_DEL           302
#define IDC_M_LBL_MACRO_HK     303
#define IDC_M_BTN_SET_MACRO_HK  304
#define IDC_M_BTN_REC           305
#define IDC_M_BTN_STOPREC       306
#define IDC_M_LBL_REC_HK       307
#define IDC_M_BTN_SET_REC_HK   308
#define IDC_M_EDIT_LOOPS        309
#define IDC_M_CHK_INFINITE      310
#define IDC_M_CHK_MOUSE_TRACK   311
#define IDC_M_BTN_PLAY          312
#define IDC_M_BTN_STOPPLAY      313
#define IDC_M_LBL_PLAY_HK      314
#define IDC_M_BTN_SET_PLAY_HK   315
#define IDC_M_BTN_SAVE          316
#define IDC_M_BTN_LOAD          317
#define IDC_M_EDIT              318
#define IDC_M_LABEL_STATUS      319

#define IDC_M_LBL_STOP_HK      320
#define IDC_M_BTN_SET_STOP_HK   321
#define IDC_M_CHK_PRESS_MODE    322
#define IDC_M_CHK_ENABLED       323
#define IDC_M_BTN_RENAME        324
#define IDC_M_RENAME_EDIT       325

#define ID_M_HK_REC    80
#define ID_M_HK_STOP   81

#define IDT_M_STATUS     350
#define IDT_M_MOUSE_POLL 351

#define MACRO_WND_W  700
#define MACRO_WND_H  560
#define WM_PLAY_TOGGLE_SOUND (WM_APP + 10)

#define MAX_MACRO_EVENTS 50000
#define MAX_MACRO_CMDS   50000

/* ================================================================== */
/*                     SCRIPT COMMAND TYPES                            */
/* ================================================================== */

enum {
    MC_KEY_DOWN, MC_KEY_UP, MC_DELAY,
    MC_MOUSE_MOVE, MC_MOUSE_LDOWN, MC_MOUSE_LUP,
    MC_MOUSE_RDOWN, MC_MOUSE_RUP, MC_MOUSE_MDOWN, MC_MOUSE_MUP
};

typedef struct {
    int cmd;
    int vk;
    unsigned short scan;
    int delay_ms;
    int x, y;
} MacroStep;

/* ================================================================== */
/*                        GLOBAL STATE                                 */
/* ================================================================== */

MacroSlot g_macro_slots[MAX_MACRO_SLOTS];
int g_macro_slot_count = 0;
int g_macro_hk_rec_vk  = VK_F11;
int g_macro_hk_stop_vk = 0;
BOOL g_macro_press_mode = FALSE;
volatile BOOL g_macro_recording = FALSE;

static volatile BOOL s_macro_playing = FALSE;
static HANDLE s_play_thread  = NULL;
static int    s_playing_slot = -1;
static HWND   s_macro_hwnd   = NULL;
static HWND   s_edit         = NULL;
static HWND   s_lbl_status   = NULL;
static HWND   s_edit_loops   = NULL;
static HWND   s_chk_infinite = NULL;
static HWND   s_chk_mouse_track = NULL;
static HWND   s_combo        = NULL;
static HWND   s_lbl_macro_hk = NULL;
static HWND   s_lbl_rec_hk   = NULL;
static HWND   s_lbl_stop_hk  = NULL;
static HWND   s_chk_press_mode = NULL;
static HWND   s_chk_enabled  = NULL;
static HWND   s_rename_edit  = NULL;
static HWND   s_btn_new      = NULL;
static WNDPROC s_rename_origproc = NULL;
static BOOL   s_renaming     = FALSE;
static int    s_current_slot  = 0;
static int    s_press_poll_vk = 0;

static BOOL   s_mouse_track_active = FALSE;
static int    s_last_mouse_x = -1, s_last_mouse_y = -1;
static WCHAR  s_last_evt_desc[128] = {0};

static CRITICAL_SECTION s_cs_rec;
static BOOL   s_cs_initialized = FALSE;
static MacroEvent s_rec_buf[MAX_MACRO_EVENTS];
static int        s_rec_count = 0;
static LARGE_INTEGER s_rec_stop_time;

static MacroStep s_play_cmds[MAX_MACRO_CMDS];
static int       s_play_cmd_count = 0;
static int       s_play_loop_cur = 0;
static int       s_play_loop_total = 0;

static HHOOK  s_macro_kbhook = NULL;
static HHOOK  s_macro_mousehook = NULL;
static int    s_macro_setting_hk = 0;

static void save_current_slot_text(void);

static const WCHAR s_edit_hint[] = L"\x8BF7\x5148\x5F55\x5236\x7B2C\x4E00\x4E2A\x5B8F\xFF0C\x53D6\x6D88\x751F\x6548\x590D\x9009\x6846\x540E\x53EF\x7F16\x8F91\x6B64\x5185\x5BB9";

static BOOL is_edit_hint(void) {
    if (!s_edit) return FALSE;
    WCHAR buf[128];
    GetWindowTextW(s_edit, buf, 128);
    return lstrcmpW(buf, s_edit_hint) == 0;
}

/* ================================================================== */
/*                     VK NAME LOOKUP (for script)                     */
/* ================================================================== */

static int vk_from_name(const WCHAR *name) {
    static const struct { const WCHAR *n; int vk; } tbl[] = {
        {L"A",0x41},{L"B",0x42},{L"C",0x43},{L"D",0x44},{L"E",0x45},{L"F",0x46},{L"G",0x47},
        {L"H",0x48},{L"I",0x49},{L"J",0x4A},{L"K",0x4B},{L"L",0x4C},{L"M",0x4D},{L"N",0x4E},
        {L"O",0x4F},{L"P",0x50},{L"Q",0x51},{L"R",0x52},{L"S",0x53},{L"T",0x54},{L"U",0x55},
        {L"V",0x56},{L"W",0x57},{L"X",0x58},{L"Y",0x59},{L"Z",0x5A},
        {L"0",0x30},{L"1",0x31},{L"2",0x32},{L"3",0x33},{L"4",0x34},
        {L"5",0x35},{L"6",0x36},{L"7",0x37},{L"8",0x38},{L"9",0x39},
        {L"F1",VK_F1},{L"F2",VK_F2},{L"F3",VK_F3},{L"F4",VK_F4},
        {L"F5",VK_F5},{L"F6",VK_F6},{L"F7",VK_F7},{L"F8",VK_F8},
        {L"F9",VK_F9},{L"F10",VK_F10},{L"F11",VK_F11},{L"F12",VK_F12},
        {L"Esc",VK_ESCAPE},{L"Tab",VK_TAB},{L"CapsLock",VK_CAPITAL},
        {L"Shift",VK_SHIFT},{L"LShift",VK_LSHIFT},{L"RShift",VK_RSHIFT},
        {L"Ctrl",VK_CONTROL},{L"LCtrl",VK_LCONTROL},{L"RCtrl",VK_RCONTROL},
        {L"Alt",VK_MENU},{L"LAlt",VK_LMENU},{L"RAlt",VK_RMENU},
        {L"Space",VK_SPACE},{L"Enter",VK_RETURN},{L"Backspace",VK_BACK},
        {L"Insert",VK_INSERT},{L"Delete",VK_DELETE},
        {L"Home",VK_HOME},{L"End",VK_END},{L"PageUp",VK_PRIOR},{L"PageDown",VK_NEXT},
        {L"Up",VK_UP},{L"Down",VK_DOWN},{L"Left",VK_LEFT},{L"Right",VK_RIGHT},
        {L"NumLock",VK_NUMLOCK},{L"Num0",VK_NUMPAD0},{L"Num1",VK_NUMPAD1},
        {L"Num2",VK_NUMPAD2},{L"Num3",VK_NUMPAD3},{L"Num4",VK_NUMPAD4},
        {L"Num5",VK_NUMPAD5},{L"Num6",VK_NUMPAD6},{L"Num7",VK_NUMPAD7},
        {L"Num8",VK_NUMPAD8},{L"Num9",VK_NUMPAD9},
        {L"Num*",VK_MULTIPLY},{L"Num+",VK_ADD},{L"Num-",VK_SUBTRACT},
        {L"Num.",VK_DECIMAL},{L"Num/",VK_DIVIDE},
        {L"PrintScreen",VK_SNAPSHOT},{L"ScrollLock",VK_SCROLL},{L"Pause",VK_PAUSE},
        {L"LWin",VK_LWIN},{L"RWin",VK_RWIN},{L"Apps",VK_APPS},
        {L";",VK_OEM_1},{L"=",VK_OEM_PLUS},{L",",VK_OEM_COMMA},
        {L"-",VK_OEM_MINUS},{L".",VK_OEM_PERIOD},{L"/",VK_OEM_2},
        {L"`",VK_OEM_3},{L"[",VK_OEM_4},{L"\\",VK_OEM_5},{L"]",VK_OEM_6},{L"'",VK_OEM_7},
    };
    for (int i = 0; i < (int)(sizeof(tbl)/sizeof(tbl[0])); i++)
        if (lstrcmpiW(name, tbl[i].n) == 0) return tbl[i].vk;
    if (wcsncmp(name, L"VK", 2) == 0) return _wtoi(name + 2);
    return 0;
}

static void vk_to_script_name(int vk, WCHAR *buf, int sz) {
    static const struct { int vk; const WCHAR *n; } tbl[] = {
        {VK_ESCAPE,L"Esc"},{VK_TAB,L"Tab"},{VK_CAPITAL,L"CapsLock"},
        {VK_SHIFT,L"Shift"},{VK_LSHIFT,L"LShift"},{VK_RSHIFT,L"RShift"},
        {VK_CONTROL,L"Ctrl"},{VK_LCONTROL,L"LCtrl"},{VK_RCONTROL,L"RCtrl"},
        {VK_MENU,L"Alt"},{VK_LMENU,L"LAlt"},{VK_RMENU,L"RAlt"},
        {VK_SPACE,L"Space"},{VK_RETURN,L"Enter"},{VK_BACK,L"Backspace"},
        {VK_INSERT,L"Insert"},{VK_DELETE,L"Delete"},
        {VK_HOME,L"Home"},{VK_END,L"End"},{VK_PRIOR,L"PageUp"},{VK_NEXT,L"PageDown"},
        {VK_UP,L"Up"},{VK_DOWN,L"Down"},{VK_LEFT,L"Left"},{VK_RIGHT,L"Right"},
        {VK_F1,L"F1"},{VK_F2,L"F2"},{VK_F3,L"F3"},{VK_F4,L"F4"},
        {VK_F5,L"F5"},{VK_F6,L"F6"},{VK_F7,L"F7"},{VK_F8,L"F8"},
        {VK_F9,L"F9"},{VK_F10,L"F10"},{VK_F11,L"F11"},{VK_F12,L"F12"},
        {VK_NUMLOCK,L"NumLock"},{VK_NUMPAD0,L"Num0"},{VK_NUMPAD1,L"Num1"},
        {VK_NUMPAD2,L"Num2"},{VK_NUMPAD3,L"Num3"},{VK_NUMPAD4,L"Num4"},
        {VK_NUMPAD5,L"Num5"},{VK_NUMPAD6,L"Num6"},{VK_NUMPAD7,L"Num7"},
        {VK_NUMPAD8,L"Num8"},{VK_NUMPAD9,L"Num9"},
        {VK_MULTIPLY,L"Num*"},{VK_ADD,L"Num+"},{VK_SUBTRACT,L"Num-"},
        {VK_DECIMAL,L"Num."},{VK_DIVIDE,L"Num/"},
        {VK_SNAPSHOT,L"PrintScreen"},{VK_SCROLL,L"ScrollLock"},{VK_PAUSE,L"Pause"},
        {VK_LWIN,L"LWin"},{VK_RWIN,L"RWin"},{VK_APPS,L"Apps"},
        {VK_OEM_1,L";"},{VK_OEM_PLUS,L"="},{VK_OEM_COMMA,L","},
        {VK_OEM_MINUS,L"-"},{VK_OEM_PERIOD,L"."},{VK_OEM_2,L"/"},
        {VK_OEM_3,L"`"},{VK_OEM_4,L"["},{VK_OEM_5,L"\\"},{VK_OEM_6,L"]"},{VK_OEM_7,L"'"},
    };
    if (vk >= 0x41 && vk <= 0x5A) { buf[0] = (WCHAR)vk; buf[1] = 0; return; }
    if (vk >= 0x30 && vk <= 0x39) { buf[0] = (WCHAR)vk; buf[1] = 0; return; }
    for (int i = 0; i < (int)(sizeof(tbl)/sizeof(tbl[0])); i++)
        if (tbl[i].vk == vk) { lstrcpynW(buf, tbl[i].n, sz); return; }
    wsprintfW(buf, L"VK%d", vk);
}

/* ================================================================== */
/*                     RECORDING                                       */
/* ================================================================== */

static MacroEvent s_pending_shift;
static BOOL s_has_pending_shift = FALSE;
static BOOL s_prev_was_e0 = FALSE;

static void macro_record_event(const MacroEvent *evt) {
    EnterCriticalSection(&s_cs_rec);
    if (s_rec_count < MAX_MACRO_EVENTS)
        s_rec_buf[s_rec_count++] = *evt;
    LeaveCriticalSection(&s_cs_rec);
}

void macro_push_event(const MacroEvent *evt) {
    if (!g_macro_recording) return;
    if (evt->type == MACRO_EVT_KEY_DOWN || evt->type == MACRO_EVT_KEY_UP) {
        if (evt->vk == g_macro_hk_rec_vk || evt->vk == g_macro_hk_stop_vk)
            return;
        for (int i = 0; i < g_macro_slot_count; i++)
            if (g_macro_slots[i].hk_play_vk > 0 && evt->vk == g_macro_slots[i].hk_play_vk)
                return;

        BOOL is_shift = (evt->vk == VK_SHIFT || evt->vk == VK_LSHIFT || evt->vk == VK_RSHIFT);
        BOOL is_e0 = (evt->flags & INTERCEPTION_KEY_E0) != 0;
        BOOL is_up = (evt->type == MACRO_EVT_KEY_UP);

        if (is_shift && is_e0) {
            return;
        }
        if (is_shift && s_prev_was_e0 && !is_up) {
            s_prev_was_e0 = FALSE;
            return;
        }
        if (is_shift && is_up) {
            if (s_has_pending_shift)
                macro_record_event(&s_pending_shift);
            s_pending_shift = *evt;
            s_has_pending_shift = TRUE;
            s_prev_was_e0 = FALSE;
            goto update_desc;
        }
        if (s_has_pending_shift) {
            if (!is_e0)
                macro_record_event(&s_pending_shift);
            s_has_pending_shift = FALSE;
        }
        s_prev_was_e0 = is_e0;
    } else {
        if (s_has_pending_shift) {
            macro_record_event(&s_pending_shift);
            s_has_pending_shift = FALSE;
        }
        s_prev_was_e0 = FALSE;
    }
    macro_record_event(evt);
update_desc:;

    WCHAR kn[64];
    switch (evt->type) {
    case MACRO_EVT_KEY_DOWN:
        vk_to_script_name(evt->vk, kn, 64);
        wsprintfW(s_last_evt_desc, L"\u6309\u4E0B %s", kn);
        break;
    case MACRO_EVT_KEY_UP:
        vk_to_script_name(evt->vk, kn, 64);
        wsprintfW(s_last_evt_desc, L"\u5F39\u8D77 %s", kn);
        break;
    case MACRO_EVT_MOUSE_MOVE:
        wsprintfW(s_last_evt_desc, L"\u9F20\u6807\u79FB\u52A8 %d,%d", evt->mouse_x, evt->mouse_y);
        break;
    case MACRO_EVT_MOUSE_BTN:
        if (evt->mouse_state & INTERCEPTION_MOUSE_LEFT_BUTTON_DOWN)        lstrcpyW(s_last_evt_desc, L"\u9F20\u6807\u5DE6\u952E\u6309\u4E0B");
        else if (evt->mouse_state & INTERCEPTION_MOUSE_LEFT_BUTTON_UP)     lstrcpyW(s_last_evt_desc, L"\u9F20\u6807\u5DE6\u952E\u5F39\u8D77");
        else if (evt->mouse_state & INTERCEPTION_MOUSE_RIGHT_BUTTON_DOWN)  lstrcpyW(s_last_evt_desc, L"\u9F20\u6807\u53F3\u952E\u6309\u4E0B");
        else if (evt->mouse_state & INTERCEPTION_MOUSE_RIGHT_BUTTON_UP)    lstrcpyW(s_last_evt_desc, L"\u9F20\u6807\u53F3\u952E\u5F39\u8D77");
        else if (evt->mouse_state & INTERCEPTION_MOUSE_MIDDLE_BUTTON_DOWN) lstrcpyW(s_last_evt_desc, L"\u9F20\u6807\u4E2D\u952E\u6309\u4E0B");
        else if (evt->mouse_state & INTERCEPTION_MOUSE_MIDDLE_BUTTON_UP)   lstrcpyW(s_last_evt_desc, L"\u9F20\u6807\u4E2D\u952E\u5F39\u8D77");
        break;
    }
}

static void build_script_from_recording(void) {
    EnterCriticalSection(&s_cs_rec);
    int count = s_rec_count;
    if (count <= 0) { LeaveCriticalSection(&s_cs_rec); return; }
    if (count > MAX_MACRO_EVENTS) count = MAX_MACRO_EVENTS;
    MacroEvent *local = (MacroEvent*)malloc(count * sizeof(MacroEvent));
    if (!local) { LeaveCriticalSection(&s_cs_rec); return; }
    memcpy(local, (void*)s_rec_buf, count * sizeof(MacroEvent));
    LeaveCriticalSection(&s_cs_rec);

    WCHAR *script = (WCHAR*)malloc(65536 * sizeof(WCHAR));
    if (!script) { free(local); return; }
    int pos = 0, rem = 65536;
    #define SAPP(fmt, ...) do { int n = wsprintfW(script+pos, fmt, __VA_ARGS__); if(n>0){pos+=n; rem-=n;} } while(0)
    #define SAPP0(s) do { int n = lstrlenW(s); if(n<rem){lstrcpyW(script+pos,s); pos+=n; rem-=n;} } while(0)

    for (int i = 0; i < count && rem > 100; i++) {
        if (i > 0) {
            long long dt = (local[i].timestamp.QuadPart - local[i-1].timestamp.QuadPart)
                           * 1000 / g_qpc_freq.QuadPart;
            if (dt < 1) dt = 1;
            SAPP(L"\u5EF6\u8FDF:%d\r\n", (int)dt);
        }
        WCHAR name[64];
        switch (local[i].type) {
        case MACRO_EVT_KEY_DOWN:
            vk_to_script_name(local[i].vk, name, 64);
            SAPP(L"\u6309\u4E0B:%s\r\n", name);
            break;
        case MACRO_EVT_KEY_UP:
            vk_to_script_name(local[i].vk, name, 64);
            SAPP(L"\u5F39\u8D77:%s\r\n", name);
            break;
        case MACRO_EVT_MOUSE_MOVE:
            SAPP(L"\u9F20\u6807\u79FB\u52A8\u5230:%d,%d\r\n", local[i].mouse_x, local[i].mouse_y);
            break;
        case MACRO_EVT_MOUSE_BTN: {
            unsigned short st = local[i].mouse_state;
            int mx = local[i].mouse_x, my = local[i].mouse_y;
            if (st & INTERCEPTION_MOUSE_LEFT_BUTTON_DOWN)   { SAPP(L"\u9F20\u6807\u79FB\u52A8\u5230:%d,%d\r\n", mx, my); SAPP0(L"\u9F20\u6807\u5DE6\u952E\u6309\u4E0B\r\n"); }
            if (st & INTERCEPTION_MOUSE_LEFT_BUTTON_UP)     SAPP0(L"\u9F20\u6807\u5DE6\u952E\u5F39\u8D77\r\n");
            if (st & INTERCEPTION_MOUSE_RIGHT_BUTTON_DOWN)  { SAPP(L"\u9F20\u6807\u79FB\u52A8\u5230:%d,%d\r\n", mx, my); SAPP0(L"\u9F20\u6807\u53F3\u952E\u6309\u4E0B\r\n"); }
            if (st & INTERCEPTION_MOUSE_RIGHT_BUTTON_UP)    SAPP0(L"\u9F20\u6807\u53F3\u952E\u5F39\u8D77\r\n");
            if (st & INTERCEPTION_MOUSE_MIDDLE_BUTTON_DOWN) { SAPP(L"\u9F20\u6807\u79FB\u52A8\u5230:%d,%d\r\n", mx, my); SAPP0(L"\u9F20\u6807\u4E2D\u952E\u6309\u4E0B\r\n"); }
            if (st & INTERCEPTION_MOUSE_MIDDLE_BUTTON_UP)   SAPP0(L"\u9F20\u6807\u4E2D\u952E\u5F39\u8D77\r\n");
            break;
        }
        }
    }
    if (count > 0 && rem > 40) {
        long long dt = (s_rec_stop_time.QuadPart - local[count-1].timestamp.QuadPart)
                       * 1000 / g_qpc_freq.QuadPart;
        if (dt > 0) SAPP(L"\u5EF6\u8FDF:%d\r\n", (int)dt);
    }
    #undef SAPP
    #undef SAPP0
    if (s_edit) SetWindowTextW(s_edit, script);
    free(script);
    free(local);
}

static void start_recording(void) {
    if (g_macro_recording || s_macro_playing) return;
    EnterCriticalSection(&s_cs_rec);
    s_rec_count = 0;
    LeaveCriticalSection(&s_cs_rec);
    g_macro_recording = TRUE;
    s_last_evt_desc[0] = 0;
    s_has_pending_shift = FALSE;
    s_prev_was_e0 = FALSE;

    if (s_chk_mouse_track)
        s_mouse_track_active = (SendMessageW(s_chk_mouse_track, BM_GETCHECK, 0, 0) == BST_CHECKED);
    else
        s_mouse_track_active = (s_current_slot >= 0 && s_current_slot < g_macro_slot_count) ?
                               g_macro_slots[s_current_slot].mouse_track : FALSE;
    if (s_mouse_track_active && s_macro_hwnd) {
        POINT pt; GetCursorPos(&pt);
        s_last_mouse_x = pt.x; s_last_mouse_y = pt.y;
        SetTimer(s_macro_hwnd, IDT_M_MOUSE_POLL, 100, NULL);
    }
    if (s_lbl_status) SetWindowTextW(s_lbl_status, L"\u72B6\u6001: \u6B63\u5728\u5F55\u5236...");
}

static void stop_recording(void) {
    if (!g_macro_recording) return;
    QueryPerformanceCounter(&s_rec_stop_time);
    g_macro_recording = FALSE;
    if (s_has_pending_shift) {
        macro_record_event(&s_pending_shift);
        s_has_pending_shift = FALSE;
    }
    if (s_mouse_track_active && s_macro_hwnd) {
        KillTimer(s_macro_hwnd, IDT_M_MOUSE_POLL);
        s_mouse_track_active = FALSE;
    }
    build_script_from_recording();
    save_current_slot_text();
    save_config();
    if (s_lbl_status) {
        WCHAR t[64];
        wsprintfW(t, L"\u72B6\u6001: \u5DF2\u5F55\u5236 %d \u6761\u4E8B\u4EF6", s_rec_count);
        SetWindowTextW(s_lbl_status, t);
    }
}

/* ================================================================== */
/*                     PARSING                                         */
/* ================================================================== */

static int parse_script(const WCHAR *text, MacroStep *cmds, int max_cmds) {
    int count = 0;
    const WCHAR *p = text;
    while (*p && count < max_cmds) {
        while (*p == L'\r' || *p == L'\n' || *p == L' ' || *p == L'\t') p++;
        if (!*p) break;
        const WCHAR *line_end = p;
        while (*line_end && *line_end != L'\r' && *line_end != L'\n') line_end++;
        int line_len = (int)(line_end - p);
        if (line_len <= 0) { p = line_end; continue; }
        WCHAR line[512];
        if (line_len >= 512) line_len = 511;
        memcpy(line, p, line_len * sizeof(WCHAR));
        line[line_len] = 0;
        p = line_end;
        MacroStep s; memset(&s, 0, sizeof(s));
        if (wcsncmp(line, L"\u6309\u4E0B:", 3) == 0) {
            s.cmd = MC_KEY_DOWN; s.vk = vk_from_name(line + 3);
            if (s.vk > 0) { s.scan = (unsigned short)MapVirtualKeyW(s.vk, MAPVK_VK_TO_VSC); cmds[count++] = s; }
        } else if (wcsncmp(line, L"\u5F39\u8D77:", 3) == 0) {
            s.cmd = MC_KEY_UP; s.vk = vk_from_name(line + 3);
            if (s.vk > 0) { s.scan = (unsigned short)MapVirtualKeyW(s.vk, MAPVK_VK_TO_VSC); cmds[count++] = s; }
        } else if (wcsncmp(line, L"\u5EF6\u8FDF:", 3) == 0) {
            s.cmd = MC_DELAY; s.delay_ms = _wtoi(line + 3);
            if (s.delay_ms < 0) s.delay_ms = 0;
            cmds[count++] = s;
        } else if (wcsncmp(line, L"\u9F20\u6807\u79FB\u52A8\u5230:", 6) == 0) {
            s.cmd = MC_MOUSE_MOVE; swscanf(line + 6, L"%d,%d", &s.x, &s.y); cmds[count++] = s;
        } else if (wcscmp(line, L"\u9F20\u6807\u5DE6\u952E\u6309\u4E0B") == 0) { s.cmd = MC_MOUSE_LDOWN; cmds[count++] = s; }
          else if (wcscmp(line, L"\u9F20\u6807\u5DE6\u952E\u5F39\u8D77") == 0) { s.cmd = MC_MOUSE_LUP; cmds[count++] = s; }
          else if (wcscmp(line, L"\u9F20\u6807\u53F3\u952E\u6309\u4E0B") == 0) { s.cmd = MC_MOUSE_RDOWN; cmds[count++] = s; }
          else if (wcscmp(line, L"\u9F20\u6807\u53F3\u952E\u5F33\u8D77") == 0 || wcscmp(line, L"\u9F20\u6807\u53F3\u952E\u5F39\u8D77") == 0) { s.cmd = MC_MOUSE_RUP; cmds[count++] = s; }
          else if (wcscmp(line, L"\u9F20\u6807\u4E2D\u952E\u6309\u4E0B") == 0) { s.cmd = MC_MOUSE_MDOWN; cmds[count++] = s; }
          else if (wcscmp(line, L"\u9F20\u6807\u4E2D\u952E\u5F39\u8D77") == 0) { s.cmd = MC_MOUSE_MUP; cmds[count++] = s; }
    }
    return count;
}

/* ================================================================== */
/*                     PLAYBACK                                        */
/* ================================================================== */

static InterceptionDevice find_keyboard_device(void) {
    if (!g_ctx) return 0;
    wchar_t hwid[256];
    for (int i = 0; i < INTERCEPTION_MAX_KEYBOARD; i++) {
        InterceptionDevice dev = INTERCEPTION_KEYBOARD(i);
        if (pfn_get_hardware_id(g_ctx, dev, hwid, sizeof(hwid)) > 0) return dev;
    }
    return 0;
}

static InterceptionDevice find_mouse_device(void) {
    if (!g_ctx) return 0;
    wchar_t hwid[256];
    for (int i = 0; i < INTERCEPTION_MAX_MOUSE; i++) {
        InterceptionDevice dev = INTERCEPTION_MOUSE(i);
        if (pfn_get_hardware_id(g_ctx, dev, hwid, sizeof(hwid)) > 0) return dev;
    }
    return 0;
}

static void play_sleep_ms(int ms) {
    if (ms <= 0 || !s_macro_playing) return;
    LARGE_INTEGER start, now;
    QueryPerformanceCounter(&start);
    long long target = (long long)ms * g_qpc_freq.QuadPart / 1000;
    int remaining = ms - 1;
    while (remaining > 0 && s_macro_playing) {
        int chunk = (remaining > 2) ? 2 : remaining;
        Sleep(chunk); remaining -= chunk;
        if (s_press_poll_vk > 0 && !(GetAsyncKeyState(s_press_poll_vk) & 0x8000)) {
            s_macro_playing = FALSE;
            return;
        }
        QueryPerformanceCounter(&now);
        if (now.QuadPart - start.QuadPart >= target) return;
    }
    do { YieldProcessor(); QueryPerformanceCounter(&now); }
    while (now.QuadPart - start.QuadPart < target && s_macro_playing);
}

static DWORD WINAPI play_thread_proc(LPVOID p) {
    (void)p;
    InterceptionDevice kb_dev = find_keyboard_device();
    InterceptionDevice ms_dev = find_mouse_device();
    int loops = s_play_loop_total, loop_idx = 0;
    while (s_macro_playing && (loops == 0 || loop_idx < loops)) {
        s_play_loop_cur = loop_idx + 1;
        for (int i = 0; i < s_play_cmd_count && s_macro_playing; i++) {
            MacroStep *cmd = &s_play_cmds[i];
            switch (cmd->cmd) {
            case MC_KEY_DOWN: {
                InterceptionKeyStroke ks; ks.code = cmd->scan; ks.state = INTERCEPTION_KEY_DOWN; ks.information = 0;
                if (kb_dev && g_ctx) pfn_send(g_ctx, kb_dev, (InterceptionStroke*)&ks, 1);
                break; }
            case MC_KEY_UP: {
                InterceptionKeyStroke ks; ks.code = cmd->scan; ks.state = INTERCEPTION_KEY_UP; ks.information = 0;
                if (kb_dev && g_ctx) pfn_send(g_ctx, kb_dev, (InterceptionStroke*)&ks, 1);
                break; }
            case MC_DELAY: play_sleep_ms(cmd->delay_ms); break;
            case MC_MOUSE_MOVE: {
                int sx = GetSystemMetrics(SM_CXSCREEN), sy = GetSystemMetrics(SM_CYSCREEN);
                InterceptionMouseStroke ms; memset(&ms, 0, sizeof(ms));
                ms.flags = INTERCEPTION_MOUSE_MOVE_ABSOLUTE;
                ms.x = (int)((long long)cmd->x * 65535 / sx);
                ms.y = (int)((long long)cmd->y * 65535 / sy);
                if (ms_dev && g_ctx) pfn_send(g_ctx, ms_dev, (InterceptionStroke*)&ms, 1);
                break; }
            case MC_MOUSE_LDOWN: { InterceptionMouseStroke ms; memset(&ms,0,sizeof(ms)); ms.state=INTERCEPTION_MOUSE_LEFT_BUTTON_DOWN; if(ms_dev&&g_ctx) pfn_send(g_ctx,ms_dev,(InterceptionStroke*)&ms,1); break; }
            case MC_MOUSE_LUP:   { InterceptionMouseStroke ms; memset(&ms,0,sizeof(ms)); ms.state=INTERCEPTION_MOUSE_LEFT_BUTTON_UP;   if(ms_dev&&g_ctx) pfn_send(g_ctx,ms_dev,(InterceptionStroke*)&ms,1); break; }
            case MC_MOUSE_RDOWN: { InterceptionMouseStroke ms; memset(&ms,0,sizeof(ms)); ms.state=INTERCEPTION_MOUSE_RIGHT_BUTTON_DOWN;if(ms_dev&&g_ctx) pfn_send(g_ctx,ms_dev,(InterceptionStroke*)&ms,1); break; }
            case MC_MOUSE_RUP:   { InterceptionMouseStroke ms; memset(&ms,0,sizeof(ms)); ms.state=INTERCEPTION_MOUSE_RIGHT_BUTTON_UP;  if(ms_dev&&g_ctx) pfn_send(g_ctx,ms_dev,(InterceptionStroke*)&ms,1); break; }
            case MC_MOUSE_MDOWN: { InterceptionMouseStroke ms; memset(&ms,0,sizeof(ms)); ms.state=INTERCEPTION_MOUSE_MIDDLE_BUTTON_DOWN;if(ms_dev&&g_ctx) pfn_send(g_ctx,ms_dev,(InterceptionStroke*)&ms,1); break; }
            case MC_MOUSE_MUP:   { InterceptionMouseStroke ms; memset(&ms,0,sizeof(ms)); ms.state=INTERCEPTION_MOUSE_MIDDLE_BUTTON_UP;  if(ms_dev&&g_ctx) pfn_send(g_ctx,ms_dev,(InterceptionStroke*)&ms,1); break; }
            }
        }
        Sleep(1);
        loop_idx++;
    }
    s_macro_playing = FALSE;
    return 0;
}

static void stop_playback(void) {
    if (!s_macro_playing) return;
    s_macro_playing = FALSE;
    if (s_play_thread) { WaitForSingleObject(s_play_thread, 2000); CloseHandle(s_play_thread); s_play_thread = NULL; }
    s_playing_slot = -1;
    if (s_lbl_status) SetWindowTextW(s_lbl_status, L"\u72B6\u6001: \u5DF2\u505C\u6B62");
}

static void start_playback(void) {
    if (s_macro_playing || g_macro_recording) return;
    if (!g_drv_ok || !g_ctx) {
        if (s_macro_hwnd) MessageBoxW(s_macro_hwnd, L"Interception \u9A71\u52A8\u672A\u5C31\u7EEA", L"\u9519\u8BEF", MB_ICONERROR);
        return;
    }
    if (!s_edit) return;
    int text_len = GetWindowTextLengthW(s_edit);
    if (text_len <= 0) return;
    WCHAR *text = (WCHAR*)malloc((text_len + 2) * sizeof(WCHAR));
    if (!text) return;
    GetWindowTextW(s_edit, text, text_len + 1);
    s_play_cmd_count = parse_script(text, s_play_cmds, MAX_MACRO_CMDS);
    free(text);
    if (s_play_cmd_count <= 0) return;
    int loops = 1;
    BOOL infinite = FALSE;
    if (s_chk_infinite) infinite = (SendMessageW(s_chk_infinite, BM_GETCHECK, 0, 0) == BST_CHECKED);
    if (!infinite && s_edit_loops) { WCHAR buf[32]; GetWindowTextW(s_edit_loops, buf, 32); loops = _wtoi(buf); if (loops < 1) loops = 1; }
    s_play_loop_total = infinite ? 0 : loops;
    s_play_loop_cur = 0;
    s_playing_slot = s_current_slot;
    s_macro_playing = TRUE;
    s_play_thread = CreateThread(NULL, 0, play_thread_proc, NULL, 0, NULL);
    if (s_lbl_status) SetWindowTextW(s_lbl_status, L"\u72B6\u6001: \u6B63\u5728\u6267\u884C...");
}

static void start_playback_slot(int slot, BOOL force_infinite) {
    if (s_macro_playing || g_macro_recording) return;
    if (!g_drv_ok || !g_ctx) return;
    if (slot < 0 || slot >= g_macro_slot_count) return;
    MacroSlot *ms = &g_macro_slots[slot];
    if (!ms->script || ms->script[0] == 0) return;
    s_play_cmd_count = parse_script(ms->script, s_play_cmds, MAX_MACRO_CMDS);
    if (s_play_cmd_count <= 0) return;
    if (force_infinite)
        s_play_loop_total = 0;
    else
        s_play_loop_total = ms->infinite ? 0 : (ms->loops < 1 ? 1 : ms->loops);
    s_play_loop_cur = 0;
    s_playing_slot = slot;
    s_macro_playing = TRUE;
    s_play_thread = CreateThread(NULL, 0, play_thread_proc, NULL, 0, NULL);
}

/* ================================================================== */
/*                     FILE SAVE / LOAD (.macro)                       */
/* ================================================================== */

static void do_save(void) {
    WCHAR path[MAX_PATH] = {0};
    OPENFILENAMEW ofn; memset(&ofn, 0, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn); ofn.hwndOwner = s_macro_hwnd;
    ofn.lpstrFilter = L"\u5B8F\u811A\u672C (*.macro)\0*.macro\0\u6240\u6709\u6587\u4EF6 (*.*)\0*.*\0";
    ofn.lpstrFile = path; ofn.nMaxFile = MAX_PATH; ofn.lpstrDefExt = L"macro"; ofn.Flags = OFN_OVERWRITEPROMPT;
    if (!GetSaveFileNameW(&ofn)) return;
    int text_len = GetWindowTextLengthW(s_edit);
    WCHAR *text = (WCHAR*)malloc((text_len + 2) * sizeof(WCHAR));
    if (!text) return;
    GetWindowTextW(s_edit, text, text_len + 1);
    int utf8_len = WideCharToMultiByte(CP_UTF8, 0, text, -1, NULL, 0, NULL, NULL);
    char *utf8 = (char*)malloc(utf8_len + 1);
    if (utf8) {
        WideCharToMultiByte(CP_UTF8, 0, text, -1, utf8, utf8_len, NULL, NULL);
        HANDLE hf = CreateFileW(path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hf != INVALID_HANDLE_VALUE) { DWORD written; WriteFile(hf, utf8, utf8_len - 1, &written, NULL); CloseHandle(hf); }
        free(utf8);
    }
    free(text);
    if (s_lbl_status) SetWindowTextW(s_lbl_status, L"\u72B6\u6001: \u5DF2\u4FDD\u5B58");
}

static void do_load(void) {
    WCHAR path[MAX_PATH] = {0};
    OPENFILENAMEW ofn; memset(&ofn, 0, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn); ofn.hwndOwner = s_macro_hwnd;
    ofn.lpstrFilter = L"\u5B8F\u811A\u672C (*.macro)\0*.macro\0\u6240\u6709\u6587\u4EF6 (*.*)\0*.*\0";
    ofn.lpstrFile = path; ofn.nMaxFile = MAX_PATH; ofn.Flags = OFN_FILEMUSTEXIST;
    if (!GetOpenFileNameW(&ofn)) return;
    HANDLE hf = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, NULL);
    if (hf == INVALID_HANDLE_VALUE) return;
    DWORD size = GetFileSize(hf, NULL);
    if (size > 1024*1024) { CloseHandle(hf); return; }
    char *utf8 = (char*)malloc(size + 1);
    if (!utf8) { CloseHandle(hf); return; }
    DWORD bytesRead; ReadFile(hf, utf8, size, &bytesRead, NULL); CloseHandle(hf);
    utf8[bytesRead] = 0;
    int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
    WCHAR *text = (WCHAR*)malloc(wlen * sizeof(WCHAR));
    if (text) { MultiByteToWideChar(CP_UTF8, 0, utf8, -1, text, wlen); SetWindowTextW(s_edit, text); free(text); }
    free(utf8);
    if (s_lbl_status) SetWindowTextW(s_lbl_status, L"\u72B6\u6001: \u5DF2\u52A0\u8F7D");
}

/* ================================================================== */
/*                     SLOT MANAGEMENT                                 */
/* ================================================================== */

static void save_current_slot_text(void) {
    int slot = s_current_slot;
    if (slot < 0 || slot >= g_macro_slot_count || !s_edit) return;
    MacroSlot *ms = &g_macro_slots[slot];
    if (ms->script) { free(ms->script); ms->script = NULL; }
    if (!is_edit_hint()) {
        int len = GetWindowTextLengthW(s_edit);
        if (len > 0) {
            ms->script = (WCHAR*)malloc((len + 2) * sizeof(WCHAR));
            if (ms->script) GetWindowTextW(s_edit, ms->script, len + 1);
        }
    }
    if (s_chk_infinite) ms->infinite = (SendMessageW(s_chk_infinite, BM_GETCHECK, 0, 0) == BST_CHECKED);
    if (s_chk_mouse_track) ms->mouse_track = (SendMessageW(s_chk_mouse_track, BM_GETCHECK, 0, 0) == BST_CHECKED);
    if (s_chk_enabled) ms->enabled = (SendMessageW(s_chk_enabled, BM_GETCHECK, 0, 0) == BST_CHECKED);
    if (s_edit_loops) { WCHAR buf[32]; GetWindowTextW(s_edit_loops, buf, 32); ms->loops = _wtoi(buf); if (ms->loops < 1) ms->loops = 1; }
}

static void load_slot_to_ui(int slot) {
    if (slot < 0 || slot >= g_macro_slot_count) return;
    MacroSlot *ms = &g_macro_slots[slot];
    if (s_edit) {
        BOOL empty = (!ms->script || ms->script[0] == 0);
        if (ms->enabled && empty)
            SetWindowTextW(s_edit, s_edit_hint);
        else
            SetWindowTextW(s_edit, ms->script ? ms->script : L"");
    }
    if (s_chk_infinite) SendMessageW(s_chk_infinite, BM_SETCHECK, ms->infinite ? BST_CHECKED : BST_UNCHECKED, 0);
    if (s_chk_mouse_track) SendMessageW(s_chk_mouse_track, BM_SETCHECK, ms->mouse_track ? BST_CHECKED : BST_UNCHECKED, 0);
    if (s_chk_enabled) SendMessageW(s_chk_enabled, BM_SETCHECK, ms->enabled ? BST_CHECKED : BST_UNCHECKED, 0);
    if (s_edit) SendMessageW(s_edit, EM_SETREADONLY, ms->enabled ? TRUE : FALSE, 0);
    if (s_edit_loops) {
        WCHAR buf[32]; wsprintfW(buf, L"%d", ms->loops > 0 ? ms->loops : 1); SetWindowTextW(s_edit_loops, buf);
        EnableWindow(s_edit_loops, !ms->infinite);
    }
}

/* ================================================================== */
/*                     HOTKEY UI LABELS                                */
/* ================================================================== */

static void get_vk_name(int vk, WCHAR *buf, int sz) {
    vk_to_script_name(vk, buf, sz);
}

static void update_hk_labels(void) {
    WCHAR name[64];
    if (s_lbl_rec_hk) {
        if (s_macro_setting_hk == 1) lstrcpyW(name, L"\u8BF7\u6309\u952E...");
        else if (g_macro_hk_rec_vk > 0) get_vk_name(g_macro_hk_rec_vk, name, 64);
        else lstrcpyW(name, L"\u65E0");
        SetWindowTextW(s_lbl_rec_hk, name);
    }
    if (s_lbl_stop_hk) {
        if (s_macro_setting_hk == 2) lstrcpyW(name, L"\u8BF7\u6309\u952E...");
        else if (g_macro_hk_stop_vk > 0) get_vk_name(g_macro_hk_stop_vk, name, 64);
        else lstrcpyW(name, L"\u65E0");
        SetWindowTextW(s_lbl_stop_hk, name);
    }
    if (s_lbl_macro_hk) {
        if (s_macro_setting_hk == 3) lstrcpyW(name, L"\u8BF7\u6309\u952E...");
        else if (s_current_slot >= 0 && s_current_slot < g_macro_slot_count && g_macro_slots[s_current_slot].hk_play_vk > 0)
            get_vk_name(g_macro_slots[s_current_slot].hk_play_vk, name, 64);
        else lstrcpyW(name, L"\u65E0");
        SetWindowTextW(s_lbl_macro_hk, name);
    }
}

static BOOL is_vk_conflicting(int vk, int setting_type) {
    if (vk == g_hk_toggle_vk || vk == g_hk_game_vk) return TRUE;
    if (setting_type != 1 && vk == g_macro_hk_rec_vk) return TRUE;
    if (setting_type != 2 && g_macro_hk_stop_vk > 0 && vk == g_macro_hk_stop_vk) return TRUE;
    for (int i = 0; i < g_macro_slot_count; i++) {
        if (setting_type == 3 && i == s_current_slot) continue;
        if (g_macro_slots[i].hk_play_vk == vk) return TRUE;
    }
    return FALSE;
}

/* ================================================================== */
/*                     HOOKS FOR HOTKEY SETTING                        */
/* ================================================================== */

static void unhook_setting_hooks(void) {
    if (s_macro_kbhook)    { UnhookWindowsHookEx(s_macro_kbhook);    s_macro_kbhook = NULL; }
    if (s_macro_mousehook) { UnhookWindowsHookEx(s_macro_mousehook); s_macro_mousehook = NULL; }
}

static void apply_set_hotkey(int vk) {
    if (is_vk_conflicting(vk, s_macro_setting_hk)) {
        if (s_lbl_status) SetWindowTextW(s_lbl_status, L"\u72B6\u6001: \u8BE5\u952E\u5DF2\u88AB\u5360\u7528");
        return;
    }
    switch (s_macro_setting_hk) {
    case 1:
        g_macro_hk_rec_vk = vk;
        break;
    case 2:
        g_macro_hk_stop_vk = vk;
        break;
    case 3:
        if (s_current_slot >= 0 && s_current_slot < g_macro_slot_count) {
            g_macro_slots[s_current_slot].hk_play_vk = vk;
        }
        break;
    }
    s_macro_setting_hk = 0;
    update_hk_labels();
    save_config();
    unhook_setting_hooks();
}

static LRESULT CALLBACK macro_kb_hook_proc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode == HC_ACTION && s_macro_setting_hk && wParam == WM_KEYDOWN) {
        KBDLLHOOKSTRUCT *p = (KBDLLHOOKSTRUCT *)lParam;
        int vk = (int)p->vkCode;
        if (vk == VK_ESCAPE) {
            switch (s_macro_setting_hk) {
            case 1:
                if (!macro_is_mouse_vk(g_macro_hk_rec_vk)) UnregisterHotKey(s_macro_hwnd, ID_M_HK_REC);
                g_macro_hk_rec_vk = 0;
                break;
            case 2:
                if (!macro_is_mouse_vk(g_macro_hk_stop_vk)) UnregisterHotKey(s_macro_hwnd, ID_M_HK_STOP);
                g_macro_hk_stop_vk = 0;
                break;
            case 3:
                if (s_current_slot >= 0 && s_current_slot < g_macro_slot_count) {
                    if (!macro_is_mouse_vk(g_macro_slots[s_current_slot].hk_play_vk))
                        UnregisterHotKey(s_macro_hwnd, ID_HK_MACRO_BASE + s_current_slot);
                    g_macro_slots[s_current_slot].hk_play_vk = 0;
                }
                break;
            }
            s_macro_setting_hk = 0;
            update_hk_labels();
            save_config();
            unhook_setting_hooks();
            return 1;
        }
        apply_set_hotkey(vk);
        return 1;
    }
    return CallNextHookEx(s_macro_kbhook, nCode, wParam, lParam);
}

static LRESULT CALLBACK macro_mouse_hook_proc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode == HC_ACTION && s_macro_setting_hk) {
        int vk = 0;
        if (wParam == WM_MBUTTONDOWN)  vk = VK_MBUTTON;
        else if (wParam == WM_XBUTTONDOWN) {
            MSLLHOOKSTRUCT *p = (MSLLHOOKSTRUCT *)lParam;
            WORD xbtn = HIWORD(p->mouseData);
            vk = (xbtn == XBUTTON1) ? VK_XBUTTON1 : VK_XBUTTON2;
        }
        if (vk) { apply_set_hotkey(vk); return 1; }
    }
    return CallNextHookEx(s_macro_mousehook, nCode, wParam, lParam);
}

static int macro_name_cmp(const void *a, const void *b) {
    return StrCmpLogicalW(((const MacroSlot*)a)->name, ((const MacroSlot*)b)->name);
}

static void sort_macro_slots(void) {
    if (g_macro_slot_count > 1) {
        WCHAR cur_name[64] = {0};
        if (s_current_slot >= 0 && s_current_slot < g_macro_slot_count)
            lstrcpynW(cur_name, g_macro_slots[s_current_slot].name, 64);
        qsort(g_macro_slots, g_macro_slot_count, sizeof(MacroSlot), macro_name_cmp);
        if (cur_name[0]) {
            for (int i = 0; i < g_macro_slot_count; i++)
                if (lstrcmpW(g_macro_slots[i].name, cur_name) == 0) { s_current_slot = i; break; }
        }
    }
    if (s_combo) {
        SendMessageW(s_combo, CB_RESETCONTENT, 0, 0);
        for (int i = 0; i < g_macro_slot_count; i++)
            SendMessageW(s_combo, CB_ADDSTRING, 0, (LPARAM)g_macro_slots[i].name);
        SendMessageW(s_combo, CB_SETCURSEL, s_current_slot, 0);
    }
}

static void finish_rename(BOOL apply) {
    if (!s_renaming) return;
    s_renaming = FALSE;
    if (apply && s_rename_edit && s_combo) {
        WCHAR buf[64];
        GetWindowTextW(s_rename_edit, buf, 64);
        if (buf[0] && s_current_slot >= 0 && s_current_slot < g_macro_slot_count) {
            lstrcpynW(g_macro_slots[s_current_slot].name, buf, 64);
            sort_macro_slots();
            save_config();
        }
    }
    if (s_rename_edit) ShowWindow(s_rename_edit, SW_HIDE);
    if (s_combo) {
        ShowWindow(s_combo, SW_SHOW);
        SendMessageW(s_combo, CB_SETCURSEL, s_current_slot, 0);
    }
    if (s_btn_new) SetWindowTextW(s_btn_new, L"\u65B0\u5EFA");
}

static LRESULT CALLBACK rename_edit_subproc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    if (msg == WM_KEYDOWN) {
        if (wp == VK_RETURN) { finish_rename(TRUE); return 0; }
        if (wp == VK_ESCAPE) { finish_rename(FALSE); return 0; }
    }
    if (msg == WM_KILLFOCUS && s_renaming && (HWND)wp != s_btn_new)
        finish_rename(TRUE);
    return CallWindowProcW(s_rename_origproc, hwnd, msg, wp, lp);
}

static void begin_set_hotkey(int which) {
    unhook_setting_hooks();
    s_macro_setting_hk = which;
    HINSTANCE hinst = GetModuleHandleW(NULL);
    s_macro_kbhook = SetWindowsHookExW(WH_KEYBOARD_LL, macro_kb_hook_proc, hinst, 0);
    s_macro_mousehook = SetWindowsHookExW(WH_MOUSE_LL, macro_mouse_hook_proc, hinst, 0);
    update_hk_labels();
    if (s_lbl_status) SetWindowTextW(s_lbl_status, L"\u72B6\u6001: \u8BF7\u6309\u4E0B\u65B0\u70ED\u952E\u6216\u9F20\u6807\u952E (Esc\u53D6\u6D88)");
}

/* ================================================================== */
/*                     JSON CONFIG HELPERS                              */
/* ================================================================== */

static void scfg_json_wstr(char *buf, int *pos, int limit, const WCHAR *ws) {
    if (!ws) { int n = snprintf(buf+*pos, limit-*pos, "\"\""); if(n>0)*pos+=n; return; }
    char utf8[131072];
    int len = WideCharToMultiByte(CP_UTF8, 0, ws, -1, utf8, sizeof(utf8)-1, NULL, NULL);
    if (len <= 0) { int n = snprintf(buf+*pos, limit-*pos, "\"\""); if(n>0)*pos+=n; return; }
    utf8[len] = 0;
    if (*pos < limit) buf[(*pos)++] = '"';
    for (int i = 0; utf8[i] && *pos < limit - 6; i++) {
        char c = utf8[i];
        if (c == '"')       { buf[(*pos)++]='\\'; buf[(*pos)++]='"'; }
        else if (c == '\\') { buf[(*pos)++]='\\'; buf[(*pos)++]='\\'; }
        else if (c == '\n') { buf[(*pos)++]='\\'; buf[(*pos)++]='n'; }
        else if (c == '\r') { buf[(*pos)++]='\\'; buf[(*pos)++]='r'; }
        else buf[(*pos)++] = c;
    }
    if (*pos < limit) buf[(*pos)++] = '"';
}

static WCHAR *read_json_wstr(const char *json, const char *key) {
    char search[128];
    snprintf(search, sizeof(search), "\"%s\"", key);
    const char *p = strstr(json, search);
    if (!p) return NULL;
    p += strlen(search);
    while (*p == ' ' || *p == ':' || *p == '\t') p++;
    if (*p != '"') return NULL;
    p++;
    char utf8[131072];
    int ui = 0;
    while (*p && *p != '"' && ui < (int)sizeof(utf8)-4) {
        if (*p == '\\') {
            p++;
            if (*p == 'n') { utf8[ui++]='\r'; utf8[ui++]='\n'; }
            else if (*p == 'r') { /* skip, \n already adds \r\n */ }
            else if (*p == '"') utf8[ui++]='"';
            else if (*p == '\\') utf8[ui++]='\\';
            else { utf8[ui++]='\\'; utf8[ui++]=*p; }
            p++;
        } else {
            utf8[ui++] = *p++;
        }
    }
    utf8[ui] = 0;
    int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
    WCHAR *ws = (WCHAR*)malloc(wlen * sizeof(WCHAR));
    if (ws) MultiByteToWideChar(CP_UTF8, 0, utf8, -1, ws, wlen);
    return ws;
}

static int find_json_int(const char *json, const char *key, int def) {
    char search[128];
    snprintf(search, sizeof(search), "\"%s\"", key);
    const char *p = strstr(json, search);
    if (!p) return def;
    p += strlen(search);
    while (*p == ' ' || *p == ':' || *p == '\t') p++;
    return atoi(p);
}

void macro_save_config_to(char *buf, int *pos, int rem) {
    #define MCFG(fmt, ...) do { int n = snprintf(buf+*pos, rem, fmt, __VA_ARGS__); if (n>0){*pos+=n; rem-=n;} } while(0)
    MCFG("  \"macro_hk_rec\": %d,\n", g_macro_hk_rec_vk);
    MCFG("  \"macro_hk_stop\": %d,\n", g_macro_hk_stop_vk);
    MCFG("  \"macro_press_mode\": %d,\n", g_macro_press_mode ? 1 : 0);
    MCFG("  \"macro_count\": %d,\n", g_macro_slot_count);
    for (int i = 0; i < g_macro_slot_count && rem > 256; i++) {
        MacroSlot *s = &g_macro_slots[i];
        char key[64];
        snprintf(key, sizeof(key), "macro_%d_name", i);
        MCFG("  \"%s\": ", key); scfg_json_wstr(buf, pos, *pos + rem, s->name); MCFG(",\n%s", "");
        MCFG("  \"macro_%d_loops\": %d,\n", i, s->loops);
        MCFG("  \"macro_%d_infinite\": %d,\n", i, s->infinite ? 1 : 0);
        MCFG("  \"macro_%d_mouse_track\": %d,\n", i, s->mouse_track ? 1 : 0);
        MCFG("  \"macro_%d_hk_play\": %d,\n", i, s->hk_play_vk);
        MCFG("  \"macro_%d_enabled\": %d,\n", i, s->enabled ? 1 : 0);
        snprintf(key, sizeof(key), "macro_%d_script", i);
        MCFG("  \"%s\": ", key); scfg_json_wstr(buf, pos, *pos + rem, s->script); MCFG(",\n%s", "");
    }
    #undef MCFG
}

void macro_load_config_from(const char *json) {
    g_macro_hk_rec_vk  = find_json_int(json, "macro_hk_rec",  VK_F11);
    g_macro_hk_stop_vk = find_json_int(json, "macro_hk_stop", 0);
    g_macro_press_mode = find_json_int(json, "macro_press_mode", 0) != 0;
    int count = find_json_int(json, "macro_count", 0);
    if (count < 0) count = 0;
    if (count > MAX_MACRO_SLOTS) count = MAX_MACRO_SLOTS;
    for (int i = 0; i < count; i++) {
        MacroSlot *s = &g_macro_slots[i];
        memset(s, 0, sizeof(MacroSlot));
        s->loops = 1;
        char key[64];
        snprintf(key, sizeof(key), "macro_%d_name", i);
        WCHAR *name = read_json_wstr(json, key);
        if (name) { lstrcpynW(s->name, name, 64); free(name); }
        else wsprintfW(s->name, L"\u5B8F%d", i + 1);
        snprintf(key, sizeof(key), "macro_%d_loops", i);
        s->loops = find_json_int(json, key, 1); if (s->loops < 1) s->loops = 1;
        snprintf(key, sizeof(key), "macro_%d_infinite", i);
        s->infinite = find_json_int(json, key, 0) != 0;
        snprintf(key, sizeof(key), "macro_%d_mouse_track", i);
        s->mouse_track = find_json_int(json, key, 0) != 0;
        snprintf(key, sizeof(key), "macro_%d_hk_play", i);
        s->hk_play_vk = find_json_int(json, key, 0);
        snprintf(key, sizeof(key), "macro_%d_enabled", i);
        s->enabled = find_json_int(json, key, 1) != 0;
        snprintf(key, sizeof(key), "macro_%d_script", i);
        s->script = read_json_wstr(json, key);
    }
    g_macro_slot_count = count;
    if (g_macro_slot_count < 1) macro_init_defaults();
    else qsort(g_macro_slots, g_macro_slot_count, sizeof(MacroSlot), macro_name_cmp);
}

/* ================================================================== */
/*                     STATUS TIMER UPDATE                             */
/* ================================================================== */

static void update_macro_status(void) {
    if (!s_lbl_status) return;
    if (g_macro_recording) {
        WCHAR t[256];
        if (s_last_evt_desc[0])
            wsprintfW(t, L"\u72B6\u6001: \u5F55\u5236\u4E2D... %s (%d \u6761)", s_last_evt_desc, s_rec_count);
        else
            wsprintfW(t, L"\u72B6\u6001: \u5F55\u5236\u4E2D... (%d \u6761)", s_rec_count);
        SetWindowTextW(s_lbl_status, t);
    } else if (s_macro_playing) {
        WCHAR t[256];
        WCHAR slot_name[64] = {0};
        if (s_playing_slot >= 0 && s_playing_slot < g_macro_slot_count)
            lstrcpynW(slot_name, g_macro_slots[s_playing_slot].name, 64);
        if (s_play_loop_total == 0)
            wsprintfW(t, L"\u72B6\u6001: \u6267\u884C [%s] \u7B2C %d \u8F6E (\u65E0\u9650)", slot_name, s_play_loop_cur);
        else
            wsprintfW(t, L"\u72B6\u6001: \u6267\u884C [%s] \u7B2C %d/%d \u8F6E", slot_name, s_play_loop_cur, s_play_loop_total);
        SetWindowTextW(s_lbl_status, t);
    }
}

/* ================================================================== */
/*                     WINDOW PROCEDURE                                */
/* ================================================================== */

static LRESULT CALLBACK macro_wndproc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_CREATE: {
        HFONT hf = g_font_ui;
        int y = 10;
        HWND lbl;

        /* Row 0: macro selector & management */
        lbl = CreateWindowW(L"STATIC", L"\u5F53\u524D\u5B8F:", WS_CHILD|WS_VISIBLE, 10, y+4, 55, 20, hwnd, NULL, NULL, NULL);
        SendMessageW(lbl, WM_SETFONT, (WPARAM)hf, TRUE);
        s_combo = CreateWindowW(L"COMBOBOX", NULL, WS_CHILD|WS_VISIBLE|CBS_DROPDOWNLIST|WS_VSCROLL,
            68, y, 150, 300, hwnd, (HMENU)(INT_PTR)IDC_M_COMBO_SLOT, NULL, NULL);
        SendMessageW(s_combo, WM_SETFONT, (WPARAM)hf, TRUE);
        s_rename_edit = CreateWindowW(L"EDIT", L"", WS_CHILD|WS_BORDER|ES_AUTOHSCROLL,
            68, y+2, 150, 24, hwnd, (HMENU)(INT_PTR)IDC_M_RENAME_EDIT, NULL, NULL);
        SendMessageW(s_rename_edit, WM_SETFONT, (WPARAM)hf, TRUE);
        s_rename_origproc = (WNDPROC)SetWindowLongPtrW(s_rename_edit, GWLP_WNDPROC, (LONG_PTR)rename_edit_subproc);
        for (int i = 0; i < g_macro_slot_count; i++)
            SendMessageW(s_combo, CB_ADDSTRING, 0, (LPARAM)g_macro_slots[i].name);
        SendMessageW(s_combo, CB_SETCURSEL, s_current_slot, 0);
        s_btn_new = CreateWindowW(L"BUTTON", L"\u65B0\u5EFA", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, 232, y, 62, 28, hwnd, (HMENU)(INT_PTR)IDC_M_BTN_NEW, NULL, NULL);
        CreateWindowW(L"BUTTON", L"\u5220\u9664", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, 300, y, 62, 28, hwnd, (HMENU)(INT_PTR)IDC_M_BTN_DEL, NULL, NULL);
        CreateWindowW(L"BUTTON", L"\u6539\u540D", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, 368, y, 62, 28, hwnd, (HMENU)(INT_PTR)IDC_M_BTN_RENAME, NULL, NULL);
        s_chk_enabled = CreateWindowW(L"BUTTON", L"\u751F\u6548", WS_CHILD|WS_VISIBLE|BS_AUTOCHECKBOX, 446, y+4, 55, 20, hwnd, (HMENU)(INT_PTR)IDC_M_CHK_ENABLED, NULL, NULL);
        SendMessageW(s_chk_enabled, WM_SETFONT, (WPARAM)hf, TRUE);

        lbl = CreateWindowW(L"STATIC", L"\u70ED\u952E\u8BBE\u7F6E", WS_CHILD|WS_VISIBLE, 10, y+40, 80, 20, hwnd, NULL, NULL, NULL);
        SendMessageW(lbl, WM_SETFONT, (WPARAM)hf, TRUE);
        CreateWindowW(L"STATIC", NULL, WS_CHILD|WS_VISIBLE|SS_ETCHEDHORZ, 92, y+48, MACRO_WND_W-112, 2, hwnd, NULL, NULL, NULL);

        /* Row 1: hotkeys */
        y += 62;
        lbl = CreateWindowW(L"STATIC", L"\u5F53\u524D\u5B8F\u542F\u505C\u952E:", WS_CHILD|WS_VISIBLE, 10, y+5, 96, 20, hwnd, NULL, NULL, NULL);
        SendMessageW(lbl, WM_SETFONT, (WPARAM)hf, TRUE);
        s_lbl_macro_hk = CreateWindowW(L"STATIC", L"", WS_CHILD|WS_VISIBLE|SS_CENTER|WS_BORDER, 108, y+2, 65, 22, hwnd, (HMENU)(INT_PTR)IDC_M_LBL_MACRO_HK, NULL, NULL);
        SendMessageW(s_lbl_macro_hk, WM_SETFONT, (WPARAM)hf, TRUE);
        CreateWindowW(L"BUTTON", L"\u6539", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, 176, y+1, 30, 24, hwnd, (HMENU)(INT_PTR)IDC_M_BTN_SET_MACRO_HK, NULL, NULL);

        s_chk_press_mode = CreateWindowW(L"BUTTON", L"\u6309\u538B\u6A21\u5F0F", WS_CHILD|WS_VISIBLE|BS_AUTOCHECKBOX, 214, y+4, 88, 20, hwnd, (HMENU)(INT_PTR)IDC_M_CHK_PRESS_MODE, NULL, NULL);
        SendMessageW(s_chk_press_mode, WM_SETFONT, (WPARAM)hf, TRUE);

        lbl = CreateWindowW(L"STATIC", L"\u5F55\u5236\u70ED\u952E:", WS_CHILD|WS_VISIBLE, 310, y+5, 68, 20, hwnd, NULL, NULL, NULL);
        SendMessageW(lbl, WM_SETFONT, (WPARAM)hf, TRUE);
        s_lbl_rec_hk = CreateWindowW(L"STATIC", L"", WS_CHILD|WS_VISIBLE|SS_CENTER|WS_BORDER, 380, y+2, 65, 22, hwnd, (HMENU)(INT_PTR)IDC_M_LBL_REC_HK, NULL, NULL);
        SendMessageW(s_lbl_rec_hk, WM_SETFONT, (WPARAM)hf, TRUE);
        CreateWindowW(L"BUTTON", L"\u6539", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, 448, y+1, 30, 24, hwnd, (HMENU)(INT_PTR)IDC_M_BTN_SET_REC_HK, NULL, NULL);

        lbl = CreateWindowW(L"STATIC", L"\u5168\u5C40\u505C\u6B62\u952E:", WS_CHILD|WS_VISIBLE, 486, y+5, 80, 20, hwnd, NULL, NULL, NULL);
        SendMessageW(lbl, WM_SETFONT, (WPARAM)hf, TRUE);
        s_lbl_stop_hk = CreateWindowW(L"STATIC", L"", WS_CHILD|WS_VISIBLE|SS_CENTER|WS_BORDER, 568, y+2, 65, 22, hwnd, (HMENU)(INT_PTR)IDC_M_LBL_STOP_HK, NULL, NULL);
        SendMessageW(s_lbl_stop_hk, WM_SETFONT, (WPARAM)hf, TRUE);
        CreateWindowW(L"BUTTON", L"\u6539", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, 636, y+1, 30, 24, hwnd, (HMENU)(INT_PTR)IDC_M_BTN_SET_STOP_HK, NULL, NULL);

        lbl = CreateWindowW(L"STATIC", L"\u5F55\u5236\u4E0E\u4FDD\u5B58", WS_CHILD|WS_VISIBLE, 10, y+34, 90, 20, hwnd, NULL, NULL, NULL);
        SendMessageW(lbl, WM_SETFONT, (WPARAM)hf, TRUE);
        CreateWindowW(L"STATIC", NULL, WS_CHILD|WS_VISIBLE|SS_ETCHEDHORZ, 104, y+42, MACRO_WND_W-124, 2, hwnd, NULL, NULL, NULL);

        /* Row 2: recording controls */
        y += 54;
        CreateWindowW(L"BUTTON", L"\u5F00\u59CB\u5F55\u5236", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, 10, y, 90, 28, hwnd, (HMENU)(INT_PTR)IDC_M_BTN_REC, NULL, NULL);
        s_chk_mouse_track = CreateWindowW(L"BUTTON", L"\u5F55\u5236\u9F20\u6807\u8F68\u8FF9", WS_CHILD|WS_VISIBLE|BS_AUTOCHECKBOX, 108, y+4, 96, 20, hwnd, (HMENU)(INT_PTR)IDC_M_CHK_MOUSE_TRACK, NULL, NULL);
        SendMessageW(s_chk_mouse_track, WM_SETFONT, (WPARAM)hf, TRUE);
        CreateWindowW(L"BUTTON", L"\u505C\u6B62\u5F55\u5236", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, 210, y, 90, 28, hwnd, (HMENU)(INT_PTR)IDC_M_BTN_STOPREC, NULL, NULL);
        lbl = CreateWindowW(L"STATIC", L"\u5FAA\u73AF\u6B21\u6570:", WS_CHILD|WS_VISIBLE, 316, y+5, 65, 20, hwnd, NULL, NULL, NULL);
        SendMessageW(lbl, WM_SETFONT, (WPARAM)hf, TRUE);
        s_edit_loops = CreateWindowW(L"EDIT", L"1", WS_CHILD|WS_VISIBLE|WS_BORDER|ES_NUMBER|ES_CENTER, 384, y+2, 45, 24, hwnd, (HMENU)(INT_PTR)IDC_M_EDIT_LOOPS, NULL, NULL);
        SendMessageW(s_edit_loops, WM_SETFONT, (WPARAM)hf, TRUE);
        s_chk_infinite = CreateWindowW(L"BUTTON", L"\u65E0\u9650", WS_CHILD|WS_VISIBLE|BS_AUTOCHECKBOX, 438, y+4, 55, 20, hwnd, (HMENU)(INT_PTR)IDC_M_CHK_INFINITE, NULL, NULL);
        SendMessageW(s_chk_infinite, WM_SETFONT, (WPARAM)hf, TRUE);
        CreateWindowW(L"BUTTON", L"\u4FDD\u5B58", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, 524, y, 74, 28, hwnd, (HMENU)(INT_PTR)IDC_M_BTN_SAVE, NULL, NULL);
        CreateWindowW(L"BUTTON", L"\u52A0\u8F7D", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, 606, y, 74, 28, hwnd, (HMENU)(INT_PTR)IDC_M_BTN_LOAD, NULL, NULL);

        /* Set font for all child controls */
        for (HWND ch = GetWindow(hwnd, GW_CHILD); ch; ch = GetWindow(ch, GW_HWNDNEXT))
            SendMessageW(ch, WM_SETFONT, (WPARAM)hf, TRUE);

        /* Script edit */
        lbl = CreateWindowW(L"STATIC", L"\u5B8F\u811A\u672C", WS_CHILD|WS_VISIBLE, 10, y+36, 70, 20, hwnd, NULL, NULL, NULL);
        SendMessageW(lbl, WM_SETFONT, (WPARAM)hf, TRUE);
        CreateWindowW(L"STATIC", NULL, WS_CHILD|WS_VISIBLE|SS_ETCHEDHORZ, 82, y+44, MACRO_WND_W-102, 2, hwnd, NULL, NULL, NULL);
        y += 60;
        s_edit = CreateWindowW(L"EDIT", L"",
            WS_CHILD|WS_VISIBLE|WS_BORDER|WS_VSCROLL|WS_HSCROLL|ES_MULTILINE|ES_AUTOVSCROLL|ES_AUTOHSCROLL|ES_WANTRETURN,
            10, y, MACRO_WND_W - 30, MACRO_WND_H - y - 50, hwnd, (HMENU)(INT_PTR)IDC_M_EDIT, NULL, NULL);
        SendMessageW(s_edit, WM_SETFONT, (WPARAM)hf, TRUE);

        /* Status bar */
        s_lbl_status = CreateWindowW(L"STATIC", L"\u72B6\u6001: \u5C31\u7EEA",
            WS_CHILD|WS_VISIBLE, 10, MACRO_WND_H - 42, MACRO_WND_W - 30, 20, hwnd, (HMENU)(INT_PTR)IDC_M_LABEL_STATUS, NULL, NULL);
        SendMessageW(s_lbl_status, WM_SETFONT, (WPARAM)hf, TRUE);

        /* Load current slot into UI */
        load_slot_to_ui(s_current_slot);
        if (s_chk_press_mode) SendMessageW(s_chk_press_mode, BM_SETCHECK, g_macro_press_mode ? BST_CHECKED : BST_UNCHECKED, 0);
        update_hk_labels();

        /* 热键统一走 main.c 的 Interception 转发，避免与 RegisterHotKey 双触发 */
        SetTimer(hwnd, IDT_M_STATUS, 200, NULL);
        return 0;
    }

    case WM_TIMER:
        if (wp == IDT_M_STATUS) update_macro_status();
        else if (wp == IDT_MACRO_PRESS_POLL) { macro_check_press_release(); }
        else if (wp == IDT_M_MOUSE_POLL && g_macro_recording) {
            POINT pt; GetCursorPos(&pt);
            if (pt.x != s_last_mouse_x || pt.y != s_last_mouse_y) {
                MacroEvent me;
                me.type = MACRO_EVT_MOUSE_MOVE;
                me.vk = 0; me.scan = 0; me.flags = 0; me.mouse_state = 0;
                me.mouse_x = pt.x; me.mouse_y = pt.y;
                QueryPerformanceCounter(&me.timestamp);
                macro_push_event(&me);
                s_last_mouse_x = pt.x; s_last_mouse_y = pt.y;
            }
        }
        return 0;

    case WM_HOTKEY:
        if (wp == ID_M_HK_REC) {
            if (g_macro_recording) stop_recording(); else start_recording();
        } else if (wp == ID_M_HK_STOP) {
            macro_handle_play_hotkey(-1);
        } else if (wp >= ID_HK_MACRO_BASE && wp < ID_HK_MACRO_BASE + MAX_MACRO_SLOTS) {
            macro_handle_play_hotkey((int)(wp - ID_HK_MACRO_BASE));
        }
        return 0;

    case WM_MACRO_MOUSE_HK: {
        int mvk = (int)wp;
        if (mvk == g_macro_hk_rec_vk) {
            if (g_macro_recording) stop_recording(); else start_recording();
        }
        if (mvk == g_macro_hk_stop_vk) {
            macro_handle_play_hotkey(-1);
            if (g_macro_recording) stop_recording();
        }
        for (int mi = 0; mi < g_macro_slot_count; mi++) {
            if (g_macro_slots[mi].hk_play_vk == mvk && g_macro_slots[mi].enabled) {
                macro_handle_play_hotkey(mi);
                break;
            }
        }
        return 0;
    }

    case WM_MACRO_KBD_HK: {
        int kvk = (int)wp;
        if (kvk == g_macro_hk_rec_vk) {
            if (g_macro_recording) stop_recording(); else start_recording();
        }
        if (kvk == g_macro_hk_stop_vk) {
            macro_handle_play_hotkey(-1);
            if (g_macro_recording) stop_recording();
        }
        for (int mi = 0; mi < g_macro_slot_count; mi++) {
            if (g_macro_slots[mi].hk_play_vk == kvk && g_macro_slots[mi].enabled) {
                macro_handle_play_hotkey(mi);
                break;
            }
        }
        return 0;
    }

    case WM_COMMAND:
        switch (LOWORD(wp)) {
        case IDC_M_COMBO_SLOT:
            if (HIWORD(wp) == CBN_SELCHANGE) {
                int ns = (int)SendMessageW(s_combo, CB_GETCURSEL, 0, 0);
                if (ns != s_current_slot && ns >= 0 && ns < g_macro_slot_count) {
                    save_current_slot_text();
                    s_current_slot = ns;
                    load_slot_to_ui(ns);
                    update_hk_labels();
                }
            }
            break;
        case IDC_M_BTN_NEW:
            if (s_renaming) {
                finish_rename(TRUE);
            } else if (g_macro_slot_count < MAX_MACRO_SLOTS) {
                save_current_slot_text();
                int idx = g_macro_slot_count;
                memset(&g_macro_slots[idx], 0, sizeof(MacroSlot));
                { int num = 1;
                  for (;;) { WCHAR trial[64]; wsprintfW(trial, L"\u5B8F%d", num);
                    BOOL dup = FALSE;
                    for (int j = 0; j < g_macro_slot_count; j++) if (lstrcmpW(g_macro_slots[j].name, trial)==0) { dup=TRUE; break; }
                    if (!dup) { lstrcpyW(g_macro_slots[idx].name, trial); break; } num++; } }
                g_macro_slots[idx].loops = 1;
                g_macro_slots[idx].infinite = TRUE;
                g_macro_slots[idx].enabled = TRUE;
                g_macro_slot_count++;
                s_current_slot = idx;
                sort_macro_slots();
                load_slot_to_ui(s_current_slot);
                update_hk_labels();
                save_config();
            }
            break;
        case IDC_M_BTN_DEL:
            if (g_macro_slot_count > 1) {
                int del = s_current_slot;
                if (g_macro_slots[del].script) free(g_macro_slots[del].script);
                for (int i = del; i < g_macro_slot_count - 1; i++)
                    g_macro_slots[i] = g_macro_slots[i+1];
                g_macro_slot_count--;
                memset(&g_macro_slots[g_macro_slot_count], 0, sizeof(MacroSlot));
                s_current_slot = (del >= g_macro_slot_count) ? g_macro_slot_count - 1 : del;
                sort_macro_slots();
                load_slot_to_ui(s_current_slot);
                update_hk_labels();
                save_config();
            }
            break;
        case IDC_M_BTN_RENAME:
            if (s_renaming) {
                finish_rename(TRUE);
            } else if (s_current_slot >= 0 && s_current_slot < g_macro_slot_count && s_rename_edit && s_combo) {
                SetWindowTextW(s_rename_edit, g_macro_slots[s_current_slot].name);
                ShowWindow(s_combo, SW_HIDE);
                ShowWindow(s_rename_edit, SW_SHOW);
                SetFocus(s_rename_edit);
                SendMessageW(s_rename_edit, EM_SETSEL, 0, -1);
                s_renaming = TRUE;
                if (s_btn_new) SetWindowTextW(s_btn_new, L"\u786E\u8BA4");
            }
            break;
        case IDC_M_BTN_SET_MACRO_HK: begin_set_hotkey(3); break;
        case IDC_M_BTN_REC:          start_recording(); break;
        case IDC_M_BTN_STOPREC:      stop_recording(); break;
        case IDC_M_BTN_SET_REC_HK:   begin_set_hotkey(1); break;
        case IDC_M_BTN_SET_STOP_HK:  begin_set_hotkey(2); break;
        case IDC_M_CHK_PRESS_MODE:
            g_macro_press_mode = (SendMessageW(s_chk_press_mode, BM_GETCHECK, 0, 0) == BST_CHECKED);
            save_config(); break;
        case IDC_M_CHK_ENABLED:
            if (s_current_slot >= 0 && s_current_slot < g_macro_slot_count) {
                BOOL en = (SendMessageW(s_chk_enabled, BM_GETCHECK, 0, 0) == BST_CHECKED);
                if (en && s_edit) {
                    int len = GetWindowTextLengthW(s_edit);
                    if (len <= 0) SetWindowTextW(s_edit, s_edit_hint);
                } else if (!en && s_edit && is_edit_hint()) {
                    SetWindowTextW(s_edit, L"");
                }
                g_macro_slots[s_current_slot].enabled = en;
                if (s_edit) SendMessageW(s_edit, EM_SETREADONLY, en ? TRUE : FALSE, 0);
            }
            save_config(); break;
        case IDC_M_CHK_INFINITE:
            if (s_edit_loops) {
                BOOL inf = (SendMessageW(s_chk_infinite, BM_GETCHECK, 0, 0) == BST_CHECKED);
                EnableWindow(s_edit_loops, !inf);
            }
            break;
        case IDC_M_BTN_SAVE:         do_save(); break;
        case IDC_M_BTN_LOAD:         do_load(); break;
        }
        return 0;

    case WM_CTLCOLORSTATIC: {
        HDC hdcCtl = (HDC)wp;
        SetBkColor(hdcCtl, GetSysColor(COLOR_BTNFACE));
        SetBkMode(hdcCtl, TRANSPARENT);
        return (LRESULT)GetSysColorBrush(COLOR_BTNFACE);
    }

    case WM_CLOSE:
        if (s_renaming) finish_rename(TRUE);
        if (g_macro_recording) stop_recording();
        if (s_macro_playing) stop_playback();
        save_current_slot_text();
        save_config();
        KillTimer(hwnd, IDT_M_STATUS);
        KillTimer(hwnd, IDT_M_MOUSE_POLL);
        unhook_setting_hooks(); s_macro_setting_hk = 0;
        UnregisterHotKey(hwnd, ID_M_HK_REC);
        UnregisterHotKey(hwnd, ID_M_HK_STOP);
        for (int ui = 0; ui < MAX_MACRO_SLOTS; ui++) UnregisterHotKey(hwnd, ID_HK_MACRO_BASE + ui);
        KillTimer(hwnd, IDT_MACRO_PRESS_POLL);
        DestroyWindow(hwnd);
        return 0;

    case WM_DESTROY:
        s_macro_hwnd = NULL;
        s_edit = NULL; s_lbl_status = NULL; s_edit_loops = NULL;
        s_chk_infinite = NULL; s_chk_mouse_track = NULL; s_combo = NULL; s_rename_edit = NULL; s_btn_new = NULL; s_renaming = FALSE;
        s_lbl_macro_hk = NULL; s_lbl_rec_hk = NULL; s_lbl_stop_hk = NULL; s_chk_press_mode = NULL; s_chk_enabled = NULL;
        PostMessage(g_hwnd, WM_MACRO_CLOSED, 0, 0);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

/* ================================================================== */
/*                     PUBLIC FUNCTIONS                                 */
/* ================================================================== */

void macro_init_defaults(void) {
    if (!s_cs_initialized) { InitializeCriticalSection(&s_cs_rec); s_cs_initialized = TRUE; }
    for (int i = 0; i < g_macro_slot_count; i++)
        if (g_macro_slots[i].script) { free(g_macro_slots[i].script); g_macro_slots[i].script = NULL; }
    memset(g_macro_slots, 0, sizeof(g_macro_slots));
    g_macro_slot_count = 1;
    wsprintfW(g_macro_slots[0].name, L"\u5B8F1");
    g_macro_slots[0].loops = 1;
    g_macro_slots[0].enabled = TRUE;
    g_macro_hk_rec_vk  = VK_F11;
    g_macro_hk_stop_vk = 0;
    g_macro_press_mode = FALSE;
}

void macro_show_window(HWND parent) {
    if (!s_cs_initialized) { InitializeCriticalSection(&s_cs_rec); s_cs_initialized = TRUE; }
    if (s_macro_hwnd) { SetForegroundWindow(s_macro_hwnd); return; }

    WNDCLASSEXW wc = {0};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = macro_wndproc;
    wc.hInstance = GetModuleHandleW(NULL);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = L"MacroWndClass";
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClassExW(&wc);

    DWORD style = WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU|WS_MINIMIZEBOX;
    RECT adj = {0, 0, MACRO_WND_W, MACRO_WND_H};
    AdjustWindowRectEx(&adj, style, FALSE, 0);
    s_macro_hwnd = CreateWindowExW(WS_EX_APPWINDOW, L"MacroWndClass", L"\u5B8F\u5F55\u5236",
        style, 0, 0, adj.right - adj.left, adj.bottom - adj.top,
        parent, NULL, GetModuleHandleW(NULL), NULL);

    /* Center on screen */
    RECT rc; GetWindowRect(s_macro_hwnd, &rc);
    int sw = GetSystemMetrics(SM_CXSCREEN), sh = GetSystemMetrics(SM_CYSCREEN);
    int ww = rc.right - rc.left, wh = rc.bottom - rc.top;
    SetWindowPos(s_macro_hwnd, NULL, (sw-ww)/2, (sh-wh)/2, 0, 0, SWP_NOSIZE|SWP_NOZORDER);

    ShowWindow(s_macro_hwnd, SW_SHOW);
    UpdateWindow(s_macro_hwnd);
}

BOOL macro_is_bound_vk(int vk) {
    if (vk == g_macro_hk_rec_vk) return TRUE;
    if (vk == g_macro_hk_stop_vk) return TRUE;
    for (int i = 0; i < g_macro_slot_count; i++)
        if (g_macro_slots[i].hk_play_vk == vk) return TRUE;
    return FALSE;
}

BOOL macro_is_mouse_vk(int vk) {
    return vk == VK_MBUTTON || vk == VK_XBUTTON1 || vk == VK_XBUTTON2;
}

HWND macro_get_hwnd(void) { return s_macro_hwnd; }

void macro_register_play_hotkeys(HWND hwnd) {
    (void)hwnd;
    /* 主界面不再注册宏键盘热键，避免吞掉原始按键输入。
       键盘宏热键触发统一走 main.c 的 Interception 拦截路径。 */
}

void macro_unregister_play_hotkeys(HWND hwnd) {
    (void)hwnd;
}

void macro_handle_play_hotkey(int slot_index) {
    BOOL was_playing = s_macro_playing;
    if (slot_index < 0) {
        stop_playback();
    } else if (slot_index >= g_macro_slot_count) {
        /* invalid slot */
    } else if (!g_macro_slots[slot_index].enabled) {
        /* disabled slot */
    } else {
        if (s_macro_hwnd && s_edit && slot_index == s_current_slot)
            save_current_slot_text();
        if (g_macro_press_mode) {
            if (!(s_macro_playing && s_playing_slot == slot_index)) {
                if (s_macro_playing) stop_playback();
                start_playback_slot(slot_index, TRUE);
                s_press_poll_vk = g_macro_slots[slot_index].hk_play_vk;
                HWND timer_hwnd = s_macro_hwnd ? s_macro_hwnd : g_hwnd;
                SetTimer(timer_hwnd, IDT_MACRO_PRESS_POLL, 10, NULL);
            }
        } else {
            if (s_macro_playing) {
                if (s_playing_slot == slot_index) stop_playback();
                else {
                    stop_playback();
                    start_playback_slot(slot_index, FALSE);
                }
            } else {
                start_playback_slot(slot_index, FALSE);
            }
        }
    }
    if (!g_macro_press_mode && was_playing != s_macro_playing && g_hwnd)
        PostMessageW(g_hwnd, WM_PLAY_TOGGLE_SOUND, (WPARAM)(s_macro_playing ? TRUE : FALSE), 0);
}

void macro_check_press_release(void) {
    if (s_press_poll_vk > 0 && !(GetAsyncKeyState(s_press_poll_vk) & 0x8000)) {
        s_press_poll_vk = 0;
        if (s_macro_hwnd) KillTimer(s_macro_hwnd, IDT_MACRO_PRESS_POLL);
        KillTimer(g_hwnd, IDT_MACRO_PRESS_POLL);
        /* 按压模式松开时静默停止：不触发开/关音效消息 */
        stop_playback();
    }
}
