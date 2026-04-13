import koffi from "koffi";

// ── Types ──────────────────────────────────────────────────────────────────
const HWND = koffi.pointer("HWND", koffi.opaque());
const HDC = koffi.pointer("HDC", koffi.opaque());
const HBITMAP = koffi.pointer("HBITMAP", koffi.opaque());
const HGDIOBJ = koffi.pointer("HGDIOBJ", koffi.opaque());

const RECT = koffi.struct("RECT", {
  left: "long",
  top: "long",
  right: "long",
  bottom: "long",
});

const POINT = koffi.struct("POINT", { x: "long", y: "long" });

const BITMAPINFOHEADER = koffi.struct("BITMAPINFOHEADER", {
  biSize: "uint32",
  biWidth: "int32",
  biHeight: "int32",
  biPlanes: "uint16",
  biBitCount: "uint16",
  biCompression: "uint32",
  biSizeImage: "uint32",
  biXPelsPerMeter: "int32",
  biYPelsPerMeter: "int32",
  biClrUsed: "uint32",
  biClrImportant: "uint32",
});

const BITMAPINFO = koffi.struct("BITMAPINFO", {
  bmiHeader: BITMAPINFOHEADER,
  bmiColors: koffi.array("uint32", 1), // dummy, not used for 32-bit
});

// INPUT struct layout on x64: type(4) + pad(4) + union(32) = 40 bytes.
// The union starts at offset 8 due to ULONG_PTR alignment.
// We add explicit _pad0 to match this layout.
const InputMouse = koffi.struct("InputMouse", {
  type: "uint32",
  _pad0: "uint32",    // alignment padding before union
  dx: "int32",
  dy: "int32",
  mouseData: "uint32",
  dwFlags: "uint32",
  time: "uint32",
  _pad1: "uint32",    // padding before dwExtraInfo (8-byte alignment)
  dwExtraInfo: "uintptr",
});

const InputKbd = koffi.struct("InputKbd", {
  type: "uint32",
  _pad0: "uint32",    // alignment padding before union
  wVk: "uint16",
  wScan: "uint16",
  dwFlags: "uint32",
  time: "uint32",
  _pad1: "uint32",    // padding before dwExtraInfo (8-byte alignment)
  dwExtraInfo: "uintptr",
  _pad2: koffi.array("uint8", 8), // fill to union size (32 bytes)
});

// ── Libraries ──────────────────────────────────────────────────────────────
const user32 = koffi.load("user32.dll");
const gdi32 = koffi.load("gdi32.dll");

// ── user32 functions ───────────────────────────────────────────────────────
const FindWindowW = user32.func("FindWindowW", HWND, ["str16", "str16"]);
const FindWindowExW = user32.func("FindWindowExW", HWND, [HWND, HWND, "str16", "str16"]);
const GetForegroundWindow = user32.func("GetForegroundWindow", HWND, []);
const SetForegroundWindow = user32.func("SetForegroundWindow", "int", [HWND]);
const GetWindowRect = user32.func("GetWindowRect", "int", [HWND, koffi.out(koffi.pointer(RECT))]);
const GetClientRect = user32.func("GetClientRect", "int", [HWND, koffi.out(koffi.pointer(RECT))]);
const MoveWindow = user32.func("MoveWindow", "int", [HWND, "int", "int", "int", "int", "int"]);
const GetDC = user32.func("GetDC", HDC, [HWND]);
const ReleaseDC = user32.func("ReleaseDC", "int", [HWND, HDC]);
const SendInput = user32.func("SendInput", "uint32", ["uint32", "void *", "int"]);
const SetCursorPos = user32.func("SetCursorPos", "int", ["int", "int"]);
const IsWindow = user32.func("IsWindow", "int", [HWND]);
const GetWindowTextW = user32.func("GetWindowTextW", "int", [HWND, koffi.out("str16"), "int"]);
const GetWindowTextLengthW = user32.func("GetWindowTextLengthW", "int", [HWND]);
const IsWindowVisible = user32.func("IsWindowVisible", "int", [HWND]);
const PrintWindow = user32.func("PrintWindow", "int", [HWND, HDC, "uint32"]);
const VkKeyScanW = user32.func("VkKeyScanW", "short", ["uint16"]);
const MapVirtualKeyW = user32.func("MapVirtualKeyW", "uint32", ["uint32", "uint32"]);
const PW_RENDERFULLCONTENT = 0x00000002;

// ── gdi32 functions ────────────────────────────────────────────────────────
const CreateCompatibleDC = gdi32.func("CreateCompatibleDC", HDC, [HDC]);
const CreateCompatibleBitmap = gdi32.func("CreateCompatibleBitmap", HBITMAP, [HDC, "int", "int"]);
const SelectObject = gdi32.func("SelectObject", "void *", [HDC, "void *"]);
const BitBlt = gdi32.func("BitBlt", "int", [HDC, "int", "int", "int", "int", HDC, "int", "int", "uint32"]);
const GetDIBits = gdi32.func("GetDIBits", "int", [HDC, HBITMAP, "uint32", "uint32", "void *", koffi.inout(koffi.pointer(BITMAPINFO)), "uint32"]);
const DeleteDC = gdi32.func("DeleteDC", "int", [HDC]);
const DeleteObject = gdi32.func("DeleteObject", "int", ["void *"]);
const GetPixel = gdi32.func("GetPixel", "uint32", [HDC, "int", "int"]);

const SRCCOPY = 0x00cc0020;
const KEYEVENTF_KEYUP = 0x0002;
const KEYEVENTF_SCANCODE = 0x0008;
const INPUT_TYPE_KEYBOARD = 1;
const INPUT_TYPE_MOUSE = 0;
const VK_CONTROL = 0x11;
const VK_SHIFT = 0x10;
const VK_ALT = 0x12;
const VK_RETURN = 0x0d;
const VK_A = 0x41;
const VK_V = 0x56;
const MOUSEEVENTF_LEFTDOWN = 0x0002;
const MOUSEEVENTF_LEFTUP = 0x0004;
const MAPVK_VK_TO_VSC = 0;

// ── Clipboard ──────────────────────────────────────────────────────────────
const OpenClipboard = user32.func("OpenClipboard", "int", [HWND]);
const CloseClipboard = user32.func("CloseClipboard", "int", []);
const EmptyClipboard = user32.func("EmptyClipboard", "int", []);
const SetClipboardData = user32.func("SetClipboardData", "void *", ["uint32", "void *"]);
const kernel32 = koffi.load("kernel32.dll");
const GlobalAlloc = kernel32.func("GlobalAlloc", "void *", ["uint32", "uintptr"]);
const GlobalLock = kernel32.func("GlobalLock", "void *", ["void *"]);
const GlobalUnlock = kernel32.func("GlobalUnlock", "int", ["void *"]);
const RtlMoveMemory = kernel32.func("RtlMoveMemory", "void", ["void *", "void *", "uintptr"]);
const CF_UNICODETEXT = 13;
const GMEM_MOVEABLE = 0x0002;

// ── Public API ─────────────────────────────────────────────────────────────

/** Get the title of a window. */
function getWindowText(hwnd) {
  const len = GetWindowTextLengthW(hwnd);
  if (len <= 0) return "";
  const buf = Buffer.alloc((len + 2) * 2);
  GetWindowTextW(hwnd, buf, len + 1);
  return buf.toString("utf16le").replace(/\0+$/, "");
}

/**
 * Find a top-level window whose title *contains* the given substring
 * (like AHK's WinExist which does substring matching).
 * Returns HWND or null.
 */
export function findWindow(titlePart) {
  let hwnd = null;
  // Enumerate top-level windows via FindWindowExW(NULL, prev, NULL, NULL)
  for (;;) {
    hwnd = FindWindowExW(null, hwnd, null, null);
    if (!hwnd) break;
    if (!IsWindowVisible(hwnd)) continue;
    const text = getWindowText(hwnd);
    if (text && text.includes(titlePart)) return hwnd;
  }
  return null;
}

/** Check if an HWND is still a valid window. */
export function isWindow(hwnd) {
  return IsWindow(hwnd) !== 0;
}

/**
 * Bring window to foreground and wait until it's actually in the foreground.
 * Returns true if successful within timeout, false otherwise.
 */
export function setForeground(hwnd, timeoutMs = 5000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    SetForegroundWindow(hwnd);
    const fg = GetForegroundWindow();
    // Compare pointer addresses
    if (fg && koffi.address(fg) === koffi.address(hwnd)) return true;
    // Busy-wait a short interval
    const end = Date.now() + 100;
    while (Date.now() < end) { /* spin */ }
  }
  return false;
}

/** Get window position & size {x, y, w, h}. */
export function getWindowRect(hwnd) {
  const r = {};
  GetWindowRect(hwnd, r);
  return { x: r.left, y: r.top, w: r.right - r.left, h: r.bottom - r.top };
}

/** Move/resize a window. */
export function moveWindow(hwnd, x, y, w, h) {
  MoveWindow(hwnd, x, y, w, h, 1);
}

/**
 * Capture the full window (including non-client/title bar area) via PrintWindow.
 * Uses PW_RENDERFULLCONTENT for DWM-composited windows.
 * Returns { width, height, data: Buffer } with BGRA pixel data.
 */
function captureFullWindow(hwnd) {
  const rect = getWindowRect(hwnd);
  const ww = rect.w;
  const wh = rect.h;

  const hdcScreen = GetDC(null);
  const hdcMem = CreateCompatibleDC(hdcScreen);
  const hBmp = CreateCompatibleBitmap(hdcScreen, ww, wh);
  const hOld = SelectObject(hdcMem, hBmp);

  PrintWindow(hwnd, hdcMem, PW_RENDERFULLCONTENT);

  const bi = {
    bmiHeader: {
      biSize: 40,
      biWidth: ww,
      biHeight: -wh, // top-down
      biPlanes: 1,
      biBitCount: 32,
      biCompression: 0,
      biSizeImage: 0,
      biXPelsPerMeter: 0,
      biYPelsPerMeter: 0,
      biClrUsed: 0,
      biClrImportant: 0,
    },
    bmiColors: [0],
  };

  const buf = Buffer.alloc(ww * wh * 4);
  GetDIBits(hdcMem, hBmp, 0, wh, buf, bi, 0);

  SelectObject(hdcMem, hOld);
  DeleteObject(hBmp);
  DeleteDC(hdcMem);
  ReleaseDC(null, hdcScreen);

  return { width: ww, height: wh, data: buf };
}

/**
 * Read a single pixel from window-relative coordinates (including non-client area).
 * Returns {r, g, b} or null.
 */
export function getPixel(hwnd, x, y) {
  const snap = captureFullWindow(hwnd);
  if (x < 0 || y < 0 || x >= snap.width || y >= snap.height) return null;
  const i = (y * snap.width + x) * 4;
  // BGRA layout
  return { r: snap.data[i + 2], g: snap.data[i + 1], b: snap.data[i] };
}

/**
 * Capture a region of the full window (including non-client area) as raw BGRA buffer.
 * Coordinates are window-relative (same as AHK CoordMode, Pixel, Window).
 * Returns { width, height, data: Buffer }.
 */
export function captureRegion(hwnd, sx, sy, w, h) {
  const snap = captureFullWindow(hwnd);
  // Guard against empty or too-small captures
  if (snap.width <= 0 || snap.height <= 0 || snap.data.length === 0) return null;
  if (sx + w > snap.width || sy + h > snap.height) return null;
  // Extract the sub-region
  const buf = Buffer.alloc(w * h * 4);
  for (let row = 0; row < h; row++) {
    const srcOff = ((sy + row) * snap.width + sx) * 4;
    const dstOff = row * w * 4;
    snap.data.copy(buf, dstOff, srcOff, srcOff + w * 4);
  }
  return { width: w, height: h, data: buf };
}

/**
 * Search for a template image inside a window region.
 * @param {HWND} hwnd
 * @param {{ data: Buffer, width: number, height: number }} template - raw RGBA pixel data
 * @param {number} sx search region left (window coords)
 * @param {number} sy search region top
 * @param {number} sw search region width
 * @param {number} sh search region height
 * @param {number} tolerance per-channel difference (like AHK *15)
 * @returns {{ x: number, y: number } | null} position of best match (window coords), or null
 */
export function imageSearch(hwnd, template, sx, sy, sw, sh, tolerance = 15) {
  const haystack = captureRegion(hwnd, sx, sy, sw, sh);
  if (!haystack) return null; // window not ready or region out of bounds
  const tw = template.width;
  const th = template.height;

  if (tw > sw || th > sh) return null;

  const hData = haystack.data; // BGRA
  const tData = template.data; // RGBA

  for (let y = 0; y <= sh - th; y++) {
    for (let x = 0; x <= sw - tw; x++) {
      let match = true;
      // Sample pixels to speed up — check corners first, then full
      outer:
      for (let ty = 0; ty < th; ty++) {
        for (let tx = 0; tx < tw; tx++) {
          const hi = (y + ty) * sw * 4 + (x + tx) * 4;
          const ti = (ty * tw + tx) * 4;
          const tAlpha = tData[ti + 3];
          if (tAlpha < 128) continue; // skip transparent pixels
          const dr = Math.abs(hData[hi + 2] - tData[ti + 0]); // B vs R
          const dg = Math.abs(hData[hi + 1] - tData[ti + 1]); // G vs G
          const db = Math.abs(hData[hi + 0] - tData[ti + 2]); // R vs B
          if (dr > tolerance || dg > tolerance || db > tolerance) {
            match = false;
            break outer;
          }
        }
      }
      if (match) {
        // Return center of match in window coordinates
        return { x: sx + x + Math.floor(tw / 2), y: sy + y + Math.floor(th / 2) };
      }
    }
  }

  return null;
}

/** Check if a pixel at window coords matches a color within tolerance. */
export function pixelMatch(hwnd, x, y, targetR, targetG, targetB, tolerance = 15) {
  const p = getPixel(hwnd, x, y);
  if (!p) return false;
  return (
    Math.abs(p.r - targetR) <= tolerance &&
    Math.abs(p.g - targetG) <= tolerance &&
    Math.abs(p.b - targetB) <= tolerance
  );
}

/** Click at screen coordinates. */
export function clickAt(screenX, screenY) {
  SetCursorPos(screenX, screenY);
  // Small delay between move and click
  const down = Buffer.alloc(koffi.sizeof(InputMouse));
  koffi.encode(down, 0, InputMouse, {
    type: INPUT_TYPE_MOUSE, _pad0: 0, dx: 0, dy: 0, mouseData: 0,
    dwFlags: MOUSEEVENTF_LEFTDOWN, time: 0, _pad1: 0, dwExtraInfo: 0,
  });
  const up = Buffer.alloc(koffi.sizeof(InputMouse));
  koffi.encode(up, 0, InputMouse, {
    type: INPUT_TYPE_MOUSE, _pad0: 0, dx: 0, dy: 0, mouseData: 0,
    dwFlags: MOUSEEVENTF_LEFTUP, time: 0, _pad1: 0, dwExtraInfo: 0,
  });
  SendInput(1, down, koffi.sizeof(InputMouse));
  SendInput(1, up, koffi.sizeof(InputMouse));
}

/** Send Ctrl+A. */
export function sendCtrlA() {
  const size = koffi.sizeof(InputKbd);
  function kbdBuf(vk, flags) {
    const scan = MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    const b = Buffer.alloc(size);
    koffi.encode(b, 0, InputKbd, {
      type: INPUT_TYPE_KEYBOARD, _pad0: 0, wVk: vk, wScan: scan, dwFlags: flags,
      time: 0, _pad1: 0, dwExtraInfo: 0, _pad2: new Array(8).fill(0),
    });
    return b;
  }
  SendInput(1, kbdBuf(VK_CONTROL, 0), size);
  SendInput(1, kbdBuf(VK_A, 0), size);
  SendInput(1, kbdBuf(VK_A, KEYEVENTF_KEYUP), size);
  SendInput(1, kbdBuf(VK_CONTROL, KEYEVENTF_KEYUP), size);
}

/** Send Enter key. */
export function sendEnter() {
  const size = koffi.sizeof(InputKbd);
  const scanEnter = MapVirtualKeyW(VK_RETURN, MAPVK_VK_TO_VSC);
  const down = Buffer.alloc(size);
  koffi.encode(down, 0, InputKbd, {
    type: INPUT_TYPE_KEYBOARD, _pad0: 0, wVk: VK_RETURN, wScan: scanEnter, dwFlags: 0,
    time: 0, _pad1: 0, dwExtraInfo: 0, _pad2: new Array(8).fill(0),
  });
  const up = Buffer.alloc(size);
  koffi.encode(up, 0, InputKbd, {
    type: INPUT_TYPE_KEYBOARD, _pad0: 0, wVk: VK_RETURN, wScan: scanEnter, dwFlags: KEYEVENTF_KEYUP,
    time: 0, _pad1: 0, dwExtraInfo: 0, _pad2: new Array(8).fill(0),
  });
  SendInput(1, down, size);
  SendInput(1, up, size);
}

/**
 * Type a string using raw hardware scan codes (KEYEVENTF_SCANCODE).
 * This is what Citrix Desktop Viewer actually forwards to the remote session.
 * Uses VkKeyScanW → MapVirtualKeyW to get the scan code for each character,
 * then sends Shift/Ctrl/Alt via scan codes as well.
 */
export function sendString(text) {
  const size = koffi.sizeof(InputKbd);
  const SCAN_SHIFT = MapVirtualKeyW(VK_SHIFT, MAPVK_VK_TO_VSC);
  const SCAN_CTRL = MapVirtualKeyW(VK_CONTROL, MAPVK_VK_TO_VSC);
  const SCAN_ALT = MapVirtualKeyW(VK_ALT, MAPVK_VK_TO_VSC);

  function scanBuf(scan, flags) {
    const b = Buffer.alloc(size);
    koffi.encode(b, 0, InputKbd, {
      type: INPUT_TYPE_KEYBOARD, _pad0: 0, wVk: 0, wScan: scan,
      dwFlags: KEYEVENTF_SCANCODE | flags,
      time: 0, _pad1: 0, dwExtraInfo: 0, _pad2: new Array(8).fill(0),
    });
    return b;
  }

  for (const ch of text) {
    const vkScan = VkKeyScanW(ch.charCodeAt(0)) & 0xffff;
    if (vkScan === 0xffff) continue; // unmappable
    const vk = vkScan & 0xff;
    const mods = (vkScan >> 8) & 0x07;
    const needShift = (mods & 0x01) !== 0;
    const needCtrl  = (mods & 0x02) !== 0;
    const needAlt   = (mods & 0x04) !== 0;
    const scan = MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    if (!scan) continue; // no scan code available

    // Press modifiers (scan code only)
    if (needCtrl)  SendInput(1, scanBuf(SCAN_CTRL, 0), size);
    if (needAlt)   SendInput(1, scanBuf(SCAN_ALT, 0), size);
    if (needShift) SendInput(1, scanBuf(SCAN_SHIFT, 0), size);
    // Key press + release
    SendInput(1, scanBuf(scan, 0), size);
    SendInput(1, scanBuf(scan, KEYEVENTF_KEYUP), size);
    // Release modifiers
    if (needShift) SendInput(1, scanBuf(SCAN_SHIFT, KEYEVENTF_KEYUP), size);
    if (needAlt)   SendInput(1, scanBuf(SCAN_ALT, KEYEVENTF_KEYUP), size);
    if (needCtrl)  SendInput(1, scanBuf(SCAN_CTRL, KEYEVENTF_KEYUP), size);
  }
}

/** Copy text to the Windows clipboard as Unicode. */
function setClipboard(text) {
  const utf16 = Buffer.from(text + "\0", "utf16le");
  const hMem = GlobalAlloc(GMEM_MOVEABLE, utf16.length);
  const ptr = GlobalLock(hMem);
  RtlMoveMemory(ptr, utf16, utf16.length);
  GlobalUnlock(hMem);
  OpenClipboard(null);
  EmptyClipboard();
  SetClipboardData(CF_UNICODETEXT, hMem);
  CloseClipboard();
}

/** Paste text via clipboard (Ctrl+V). Most reliable for Citrix/RDP. */
export function sendPaste(text) {
  setClipboard(text);
  // Send Ctrl+V
  const size = koffi.sizeof(InputKbd);
  function kbdBuf(vk, flags) {
    const scan = MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    const b = Buffer.alloc(size);
    koffi.encode(b, 0, InputKbd, {
      type: INPUT_TYPE_KEYBOARD, _pad0: 0, wVk: vk, wScan: scan, dwFlags: flags,
      time: 0, _pad1: 0, dwExtraInfo: 0, _pad2: new Array(8).fill(0),
    });
    return b;
  }
  SendInput(1, kbdBuf(VK_CONTROL, 0), size);
  SendInput(1, kbdBuf(VK_V, 0), size);
  SendInput(1, kbdBuf(VK_V, KEYEVENTF_KEYUP), size);
  SendInput(1, kbdBuf(VK_CONTROL, KEYEVENTF_KEYUP), size);
}
