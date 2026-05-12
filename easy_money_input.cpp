#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#include <cwchar>

namespace {

constexpr size_t kMaxTextUtf16Units = 512;
constexpr size_t kMaxInputEvents = 3 + 4 + (kMaxTextUtf16Units * 2) + 8;

LONG clamp_absolute_coord(double value) {
    if (value < 0.0) {
        return 0;
    }
    if (value > 65535.0) {
        return 65535;
    }
    return static_cast<LONG>(value + 0.5);
}

int int_max(int a, int b) {
    return a > b ? a : b;
}

void absolute_mouse_coords(int x, int y, LONG& dx, LONG& dy) {
    const int left = GetSystemMetrics(SM_XVIRTUALSCREEN);
    const int top = GetSystemMetrics(SM_YVIRTUALSCREEN);
    const int width = int_max(1, GetSystemMetrics(SM_CXVIRTUALSCREEN));
    const int height = int_max(1, GetSystemMetrics(SM_CYVIRTUALSCREEN));
    const int denom_x = int_max(width - 1, 1);
    const int denom_y = int_max(height - 1, 1);
    dx = clamp_absolute_coord((static_cast<double>(x - left) * 65535.0) / denom_x);
    dy = clamp_absolute_coord((static_cast<double>(y - top) * 65535.0) / denom_y);
}

void add_mouse(INPUT* inputs, size_t& index, LONG dx, LONG dy, DWORD flags) {
    INPUT& input = inputs[index++];
    input.type = INPUT_MOUSE;
    input.mi.dx = dx;
    input.mi.dy = dy;
    input.mi.mouseData = 0;
    input.mi.dwFlags = flags;
    input.mi.time = 0;
    input.mi.dwExtraInfo = 0;
}

void add_key(INPUT* inputs, size_t& index, WORD vk) {
    INPUT& down = inputs[index++];
    down.type = INPUT_KEYBOARD;
    down.ki.wVk = vk;
    down.ki.wScan = 0;
    down.ki.dwFlags = 0;
    down.ki.time = 0;
    down.ki.dwExtraInfo = 0;

    INPUT& up = inputs[index++];
    up.type = INPUT_KEYBOARD;
    up.ki.wVk = vk;
    up.ki.wScan = 0;
    up.ki.dwFlags = KEYEVENTF_KEYUP;
    up.ki.time = 0;
    up.ki.dwExtraInfo = 0;
}

void add_unicode(INPUT* inputs, size_t& index, wchar_t unit) {
    INPUT& down = inputs[index++];
    down.type = INPUT_KEYBOARD;
    down.ki.wVk = 0;
    down.ki.wScan = static_cast<WORD>(unit);
    down.ki.dwFlags = KEYEVENTF_UNICODE;
    down.ki.time = 0;
    down.ki.dwExtraInfo = 0;

    INPUT& up = inputs[index++];
    up.type = INPUT_KEYBOARD;
    up.ki.wVk = 0;
    up.ki.wScan = static_cast<WORD>(unit);
    up.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
    up.ki.time = 0;
    up.ki.dwExtraInfo = 0;
}

bool valid_text(const wchar_t* text, size_t len) {
    if (text == nullptr || len == 0 || len > kMaxTextUtf16Units) {
        return false;
    }
    for (size_t i = 0; i < len; ++i) {
        const wchar_t ch = text[i];
        if (ch == L'\r' || ch == L'\n' || ch == L'\t' || ch < 32) {
            return false;
        }
    }
    return true;
}

} // namespace

extern "C" __declspec(dllexport) int easy_money_send_comment_default(
    int x,
    int y,
    const wchar_t* text
) noexcept {
    const size_t text_len = text == nullptr ? 0 : std::wcslen(text);
    if (!valid_text(text, text_len)) {
        return 1;
    }

    INPUT inputs[kMaxInputEvents] = {};
    size_t index = 0;
    LONG dx = 0;
    LONG dy = 0;
    absolute_mouse_coords(x, y, dx, dy);

    add_mouse(inputs, index, dx, dy, MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK);
    add_mouse(inputs, index, 0, 0, MOUSEEVENTF_LEFTDOWN);
    add_mouse(inputs, index, 0, 0, MOUSEEVENTF_LEFTUP);

    add_key(inputs, index, VK_TAB);
    add_key(inputs, index, VK_RETURN);

    for (size_t i = 0; i < text_len; ++i) {
        add_unicode(inputs, index, text[i]);
    }

    add_key(inputs, index, VK_TAB);
    add_key(inputs, index, VK_TAB);
    add_key(inputs, index, VK_TAB);
    add_key(inputs, index, VK_RETURN);

    const UINT sent = SendInput(static_cast<UINT>(index), inputs, sizeof(INPUT));
    if (sent != index) {
        const DWORD error = GetLastError();
        return error == 0 ? 2 : static_cast<int>(error);
    }
    return 0;
}

extern "C" __declspec(dllexport) int easy_money_send_comment_setpos(
    int x,
    int y,
    const wchar_t* text
) noexcept {
    const size_t text_len = text == nullptr ? 0 : std::wcslen(text);
    if (!valid_text(text, text_len)) {
        return 1;
    }

    if (!SetCursorPos(x, y)) {
        const DWORD error = GetLastError();
        return error == 0 ? 3 : static_cast<int>(error);
    }

    INPUT inputs[kMaxInputEvents] = {};
    size_t index = 0;

    add_mouse(inputs, index, 0, 0, MOUSEEVENTF_LEFTDOWN);
    add_mouse(inputs, index, 0, 0, MOUSEEVENTF_LEFTUP);

    add_key(inputs, index, VK_TAB);
    add_key(inputs, index, VK_RETURN);

    for (size_t i = 0; i < text_len; ++i) {
        add_unicode(inputs, index, text[i]);
    }

    add_key(inputs, index, VK_TAB);
    add_key(inputs, index, VK_TAB);
    add_key(inputs, index, VK_TAB);
    add_key(inputs, index, VK_RETURN);

    const UINT sent = SendInput(static_cast<UINT>(index), inputs, sizeof(INPUT));
    if (sent != index) {
        const DWORD error = GetLastError();
        return error == 0 ? 2 : static_cast<int>(error);
    }
    return 0;
}
