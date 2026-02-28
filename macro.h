#ifndef MACRO_H
#define MACRO_H

#include <windows.h>

#define INTERCEPTION_STATIC
#include "interception.h"

#define MACRO_EVT_KEY_DOWN    1
#define MACRO_EVT_KEY_UP      2
#define MACRO_EVT_MOUSE_BTN   3
#define MACRO_EVT_MOUSE_MOVE  4

typedef struct {
    int type;
    int vk;
    unsigned short scan;
    unsigned short flags;
    unsigned short mouse_state;
    int mouse_x, mouse_y;
    LARGE_INTEGER timestamp;
} MacroEvent;

#define MAX_MACRO_SLOTS   20
#define ID_HK_MACRO_BASE  100
#define WM_MACRO_CLOSED   (WM_USER + 100)
#define WM_MACRO_MOUSE_HK (WM_USER + 101)
#define WM_MACRO_KBD_HK   (WM_USER + 102)

typedef struct {
    WCHAR name[64];
    WCHAR *script;
    int loops;
    BOOL infinite;
    BOOL mouse_track;
    int hk_play_vk;
    BOOL enabled;
} MacroSlot;

extern MacroSlot g_macro_slots[MAX_MACRO_SLOTS];
extern int g_macro_slot_count;
extern int g_macro_hk_rec_vk;
extern int g_macro_hk_stop_vk;
extern BOOL g_macro_press_mode;
extern volatile BOOL g_macro_recording;

#define IDT_MACRO_PRESS_POLL 352

void macro_init_defaults(void);
void macro_show_window(HWND parent);
void macro_push_event(const MacroEvent *evt);
BOOL macro_is_bound_vk(int vk);
BOOL macro_is_mouse_vk(int vk);
HWND macro_get_hwnd(void);
void macro_register_play_hotkeys(HWND hwnd);
void macro_unregister_play_hotkeys(HWND hwnd);
void macro_handle_play_hotkey(int slot_index);
void macro_check_press_release(void);
void macro_save_config_to(char *buf, int *pos, int rem);
void macro_load_config_from(const char *json);

#endif
