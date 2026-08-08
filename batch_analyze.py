#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WinRHIZO 半自动批量分析工具
============================
- 按样本顺序列出当天图片
- OpenCV 自动检测根系包围盒(ROI)，生成预览图
- 在 WinRHIZO 中自动打开图片、应用 ROI、触发/等待分析
- 数据自动追加到结果 TXT（全新数据文件，绝不覆盖已有文件）
- 汇总生成 _汇总.csv / .xlsx（F/G + 样地 + 重复 分组，多次分析全部保留）

依赖: numpy, opencv-python, pillow, openpyxl (见 requirements.txt)
运行: python batch_analyze.py  或双击 启动.bat
"""

import ctypes
import json
import os
import re
import sys
import time
import threading
import queue
import traceback
from ctypes import (c_int, c_uint, c_ulong, c_long, c_ubyte, c_void_p,
                    c_wchar, c_wchar_p, c_ssize_t, c_size_t, POINTER, Structure,
                    byref, WINFUNCTYPE)

import numpy as np

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

# ============================================================
# Win32 绑定
# ============================================================
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

EnumProc = WINFUNCTYPE(c_int, c_void_p, c_void_p)


class RECT(Structure):
    _fields_ = [("left", c_int), ("top", c_int), ("right", c_int), ("bottom", c_int)]


class POINT(Structure):
    _fields_ = [("x", c_int), ("y", c_int)]


class MSG(Structure):
    _fields_ = [("hwnd", c_void_p), ("message", c_uint), ("wParam", c_ulong),
                ("lParam", c_long), ("time", c_ulong), ("pt_x", c_int), ("pt_y", c_int)]


class PROCESSENTRY32W(Structure):
    _fields_ = [("dwSize", c_ulong), ("cntUsage", c_ulong), ("th32ProcessID", c_ulong),
                ("th32DefaultHeapID", c_void_p), ("th32ModuleID", c_ulong),
                ("cntThreads", c_ulong), ("th32ParentProcessID", c_ulong),
                ("pcPriClassBase", c_long), ("dwFlags", c_ulong),
                ("szExeFile", c_wchar * 260)]


user32.EnumWindows.argtypes = [EnumProc, c_void_p]
user32.EnumWindows.restype = c_int
user32.EnumChildWindows.argtypes = [c_void_p, EnumProc, c_void_p]
user32.EnumChildWindows.restype = c_int
user32.GetWindowThreadProcessId.argtypes = [c_void_p, POINTER(c_uint)]
user32.GetWindowThreadProcessId.restype = c_uint
user32.GetWindowTextW.argtypes = [c_void_p, c_wchar_p, c_int]
user32.GetWindowTextW.restype = c_int
user32.GetClassNameW.argtypes = [c_void_p, c_wchar_p, c_int]
user32.GetClassNameW.restype = c_int
user32.IsWindowVisible.argtypes = [c_void_p]
user32.IsWindowVisible.restype = c_int
user32.IsWindow.argtypes = [c_void_p]
user32.IsWindow.restype = c_int
user32.GetDlgCtrlID.argtypes = [c_void_p]
user32.GetDlgCtrlID.restype = c_int
user32.GetParent.argtypes = [c_void_p]
user32.GetParent.restype = c_void_p
user32.PostMessageW.argtypes = [c_void_p, c_uint, c_void_p, c_void_p]
user32.PostMessageW.restype = c_int
user32.SendMessageW.argtypes = [c_void_p, c_uint, c_void_p, c_void_p]
user32.SendMessageW.restype = c_void_p
user32.GetDlgItem.argtypes = [c_void_p, c_int]
user32.GetDlgItem.restype = c_void_p
user32.GetWindowRect.argtypes = [c_void_p, POINTER(RECT)]
user32.GetWindowRect.restype = c_int
user32.GetClientRect.argtypes = [c_void_p, POINTER(RECT)]
user32.GetClientRect.restype = c_int
user32.ClientToScreen.argtypes = [c_void_p, POINTER(POINT)]
user32.ClientToScreen.restype = c_int
user32.SetForegroundWindow.argtypes = [c_void_p]
user32.SetForegroundWindow.restype = c_int
user32.SetFocus.argtypes = [c_void_p]
user32.SetFocus.restype = c_void_p
user32.SetWindowPos.argtypes = [c_void_p, c_void_p, c_int, c_int, c_int, c_int, c_uint]
user32.SetWindowPos.restype = c_int
user32.RegisterHotKey.argtypes = [c_void_p, c_int, c_uint, c_uint]
user32.RegisterHotKey.restype = c_int
user32.UnregisterHotKey.argtypes = [c_void_p, c_int]
user32.UnregisterHotKey.restype = c_int
user32.GetMessageW.argtypes = [POINTER(MSG), c_void_p, c_uint, c_uint]
user32.GetMessageW.restype = c_int
user32.TranslateMessage.argtypes = [POINTER(MSG)]
user32.TranslateMessage.restype = c_int
user32.DispatchMessageW.argtypes = [POINTER(MSG)]
user32.DispatchMessageW.restype = c_void_p
user32.OpenClipboard.argtypes = [c_void_p]
user32.OpenClipboard.restype = c_int
user32.EmptyClipboard.argtypes = []
user32.EmptyClipboard.restype = c_int
user32.SetClipboardData.argtypes = [c_uint, c_void_p]
user32.SetClipboardData.restype = c_void_p
user32.CloseClipboard.argtypes = []
user32.CloseClipboard.restype = c_int
user32.SetCursorPos.argtypes = [c_int, c_int]
user32.SetCursorPos.restype = c_int
user32.mouse_event.argtypes = [c_uint, c_uint, c_uint, c_uint, c_void_p]
user32.mouse_event.restype = None
user32.keybd_event.argtypes = [c_ubyte, c_ubyte, c_ulong, c_void_p]
user32.keybd_event.restype = None

kernel32.GlobalAlloc.argtypes = [c_uint, c_size_t]
kernel32.GlobalAlloc.restype = c_void_p
kernel32.GlobalLock.argtypes = [c_void_p]
kernel32.GlobalLock.restype = c_void_p
kernel32.GlobalUnlock.argtypes = [c_void_p]
kernel32.GlobalUnlock.restype = c_int
kernel32.GlobalFree.argtypes = [c_void_p]
kernel32.GlobalFree.restype = c_void_p
kernel32.CreateToolhelp32Snapshot.argtypes = [c_uint, c_ulong]
kernel32.CreateToolhelp32Snapshot.restype = c_void_p
kernel32.Process32FirstW.argtypes = [c_void_p, POINTER(PROCESSENTRY32W)]
kernel32.Process32FirstW.restype = c_int
kernel32.Process32NextW.argtypes = [c_void_p, POINTER(PROCESSENTRY32W)]
kernel32.Process32NextW.restype = c_int
kernel32.CloseHandle.argtypes = [c_void_p]
kernel32.CloseHandle.restype = c_int

WM_COMMAND = 0x0111
WM_CLOSE = 0x0010
WM_SETTEXT = 0x000C
WM_GETTEXT = 0x000D
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
BM_CLICK = 0x00F5
WM_PASTE = 0x0302
WM_CLEAR = 0x0303
EM_SETSEL = 0x00B1
WM_HOTKEY = 0x0312
VK_RETURN = 0x0D
VK_CONTROL = 0x11
VK_A = 0x41
VK_V = 0x56
VK_F12 = 0x7B
KEYEVENTF_KEYUP = 0x0002
MOD_NOREPEAT = 0x4000
CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002
GMEM_ZEROINIT = 0x0040
TH32CS_SNAPPROCESS = 0x00000002
HWND_TOP = 0
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_SHOWWINDOW = 0x0040


# ============================================================
# 工具函数
# ============================================================
def log(msg):
    ts = time.strftime("%H:%M:%S")
    print("[%s] %s" % (ts, msg), flush=True)


def find_pid_by_name(name):
    """按进程名找 PID（不区分大小写），返回列表。"""
    results = []
    snapshot = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if snapshot == 0 or snapshot == c_void_p(-1).value:
        return results
    try:
        pe = PROCESSENTRY32W()
        pe.dwSize = ctypes.sizeof(PROCESSENTRY32W)
        ok = kernel32.Process32FirstW(snapshot, byref(pe))
        while ok:
            if pe.szExeFile.lower() == name.lower():
                results.append(pe.th32ProcessID)
            ok = kernel32.Process32NextW(snapshot, byref(pe))
    finally:
        kernel32.CloseHandle(snapshot)
    return results


def windows_of_pid(pid, visible_only=False):
    """返回该 PID 的所有顶层窗口 (hwnd, class, title)。"""
    out = []
    def cb(h, l):
        p = c_uint()
        user32.GetWindowThreadProcessId(h, byref(p))
        if p.value == pid and (not visible_only or user32.IsWindowVisible(h)):
            cls = ctypes.create_unicode_buffer(128)
            title = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(h, cls, 128)
            user32.GetWindowTextW(h, title, 256)
            out.append((h, cls.value, title.value))
        return 1
    user32.EnumWindows(EnumProc(cb), 0)
    return out


def children_of(hwnd):
    out = []
    def cb(h, l):
        cls = ctypes.create_unicode_buffer(128)
        title = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(h, cls, 128)
        user32.GetWindowTextW(h, title, 256)
        out.append((h, cls.value, title.value, user32.GetDlgCtrlID(h), user32.IsWindowVisible(h)))
        return 1
    user32.EnumChildWindows(hwnd, EnumProc(cb), 0)
    return out


def get_window_text(hwnd):
    b = ctypes.create_unicode_buffer(512)
    user32.SendMessageW(hwnd, WM_GETTEXT, 512, b)
    return b.value


def set_window_text(hwnd, text):
    """仅用于简单控件（如 Edit）；不要用于现代文件对话框。"""
    user32.SendMessageW(hwnd, WM_SETTEXT, 0, ctypes.cast(ctypes.c_wchar_p(text), c_void_p))


def click_button(dialog_hwnd, btn_id):
    """在对话框内点击指定 ID 的按钮。"""
    btn = user32.GetDlgItem(dialog_hwnd, btn_id)
    if not btn:
        return False
    user32.SendMessageW(btn, BM_CLICK, 0, 0)
    return True


def set_clipboard_text(text):
    """把 Unicode 文本放进系统剪贴板。"""
    user32.OpenClipboard(0)
    user32.EmptyClipboard()
    data = (text + "\x00").encode("utf-16-le")
    hmem = kernel32.GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, len(data))
    if not hmem:
        user32.CloseClipboard()
        return False
    ptr = kernel32.GlobalLock(hmem)
    if not ptr:
        kernel32.GlobalFree(hmem)
        user32.CloseClipboard()
        return False
    ctypes.memmove(ptr, data, len(data))
    kernel32.GlobalUnlock(hmem)
    user32.SetClipboardData(CF_UNICODETEXT, hmem)
    user32.CloseClipboard()
    return True


def focus_and_foreground(hwnd):
    user32.SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_SHOWWINDOW)
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.15)


def key_press(hwnd, vk, with_ctrl=False):
    """向窗口发虚拟键（尽力而为）。"""
    if with_ctrl:
        user32.PostMessageW(hwnd, WM_KEYDOWN, VK_CONTROL, 0)
        time.sleep(0.05)
    user32.PostMessageW(hwnd, WM_KEYDOWN, vk, 0)
    user32.PostMessageW(hwnd, WM_KEYUP, vk, 0)
    if with_ctrl:
        time.sleep(0.05)
        user32.PostMessageW(hwnd, WM_KEYUP, VK_CONTROL, 0)


# ============================================================
# 配置
# ============================================================
def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    return cfg


def resolve(base, path):
    if os.path.isabs(path):
        return path
    return os.path.join(base, path)


# ============================================================
# ROIDetector —— 自动检测根系包围盒
# ============================================================
class ROIDetector:
    def __init__(self, cfg_roi):
        self.edge_margin_px = int(cfg_roi.get("edge_margin_px", 50))
        self.open_kernel = int(cfg_roi.get("open_kernel", 5))
        self.merge_dilate = int(cfg_roi.get("merge_dilate", 15))
        self.max_components = int(cfg_roi.get("max_components", 4))
        self.comp_area_ratio = float(cfg_roi.get("comp_area_ratio", 0.02))
        self.bbox_margin_ratio = float(cfg_roi.get("bbox_margin_ratio", 0.02))
        self.margin_top = float(cfg_roi.get("margin_top", self.bbox_margin_ratio))
        self.margin_left = float(cfg_roi.get("margin_left", self.bbox_margin_ratio))
        self.margin_right = float(cfg_roi.get("margin_right", self.bbox_margin_ratio))
        self.margin_bottom = float(cfg_roi.get("margin_bottom", self.bbox_margin_ratio))

    @staticmethod
    def load_image(path):
        """读取灰度图，并应用 EXIF 旋转，使坐标与显示一致（OpenCV imdecode 不应用 EXIF）。"""
        import cv2
        data = np.fromfile(path, dtype=np.uint8)
        img = cv2.imdecode(data, cv2.IMREAD_GRAYSCALE)
        if img is None:
            return None
        try:
            from PIL import Image
            im = Image.open(path)
            orient = im.getexif().get(0x0112)
            if orient == 3:
                img = cv2.rotate(img, cv2.ROTATE_180)
            elif orient == 6:
                img = cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
            elif orient == 8:
                img = cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
        except Exception:
            pass
        return img

    def detect(self, path):
        """返回 ((L,T,R,B), img) ；检测失败返回 (None, img)。"""
        import cv2
        img = self.load_image(path)
        if img is None:
            return None, None
        thr = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
        thr = cv2.morphologyEx(thr, cv2.MORPH_OPEN,
                               np.ones((self.open_kernel, self.open_kernel), np.uint8))
        if self.merge_dilate > 1:
            thr = cv2.dilate(thr, np.ones((self.merge_dilate, self.merge_dilate), np.uint8))
        em = self.edge_margin_px
        if em > 0:
            thr[:em, :] = 0
            thr[-em:, :] = 0
            thr[:, :em] = 0
            thr[:, -em:] = 0
        n, _, stats, _ = cv2.connectedComponentsWithStats(thr)
        if n <= 1:
            return None, img
        sizes = stats[1:, cv2.CC_STAT_AREA]
        order = np.argsort(-sizes)
        keep = []
        for i in order[:self.max_components]:
            if sizes[i] >= self.comp_area_ratio * sizes[order[0]]:
                keep.append(1 + int(i))
        if not keep:
            keep = [1 + int(order[0])]
        ys, xs = np.where(np.isin(cv2.connectedComponentsWithStats(thr)[1], keep))
        if len(xs) == 0:
            return None, img
        L, T, R, B = int(xs.min()), int(ys.min()), int(xs.max()), int(ys.max())
        mL = max(0, int((R - L) * self.margin_left))
        mT = max(0, int((B - T) * self.margin_top))
        mR = max(0, int((R - L) * self.margin_right))
        mB = max(0, int((B - T) * self.margin_bottom))
        L = max(0, L - mL)
        T = max(0, T - mT)
        R = min(img.shape[1] - 1, R + mR)
        B = min(img.shape[0] - 1, B + mB)
        return (L, T, R, B), img

    def save_preview(self, path, roi, preview_dir, label=""):
        """把检测框画到缩小图上存预览，不动原图。"""
        import cv2
        img = self.load_image(path)
        if img is None:
            return
        img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
        scale = 1.0
        maxw = 1200
        if img.shape[1] > maxw:
            scale = maxw / img.shape[1]
            img = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
        if roi:
            L, T, R, B = [int(v * scale) for v in roi]
            cv2.rectangle(img, (L, T), (R, B), (0, 0, 255), max(2, int(3 * scale)))
            cv2.putText(img, label, (L, max(20, T - 8)), cv2.FONT_HERSHEY_SIMPLEX,
                        0.8 * scale + 0.3, (0, 0, 255), 2)
        os.makedirs(preview_dir, exist_ok=True)
        out = os.path.join(preview_dir, os.path.basename(path).rsplit(".", 1)[0] + "_roi.jpg")
        cv2.imencode(".jpg", img)[1].tofile(out)
        return out


# ============================================================
# WinRhizo 驱动
# ============================================================
class WinRhizo:
    def __init__(self, cfg, log_fn=log):
        self.cfg = cfg
        self.wc = cfg.get("winrhizo", {})
        self.log = log_fn
        self.main = None
        self.pid = 0

    # ---------- 定位 ----------
    def find_main(self):
        pids = find_pid_by_name(self.wc.get("process_name", "WinRHIZO.exe"))
        if not pids:
            return None
        for pid in pids:
            for h, cls, title in windows_of_pid(pid, visible_only=True):
                if self.wc.get("title_contains", "WinRHIZO") in title:
                    self.main = h
                    self.pid = pid
                    return h
        return None

    def ensure_main(self):
        if self.main is None or not user32.IsWindow(self.main):
            self.find_main()
        return self.main

    # ---------- 菜单命令 ----------
    def menu(self, cmd_id):
        h = self.ensure_main()
        if not h:
            return False
        user32.PostMessageW(h, WM_COMMAND, cmd_id, 0)
        return True

    # ---------- 对话框 ----------
    def _find_new_dialog(self, before, wait_s, cls="#32770"):
        """等一个新对话框出现。"""
        deadline = time.time() + wait_s
        while time.time() < deadline:
            def cb(h, l):
                if h not in before:
                    p = c_uint()
                    user32.GetWindowThreadProcessId(h, byref(p))
                    if p.value == self._pid():
                        c = ctypes.create_unicode_buffer(64)
                        user32.GetClassNameW(h, c, 64)
                        if c.value == cls:
                            dialogs.append(h)
                            return 0
                return 1
            dialogs = []
            user32.EnumWindows(EnumProc(cb), 0)
            if dialogs:
                return dialogs[0]
            time.sleep(0.15)
        return None

    def _pid(self):
        if self.pid:
            return self.pid
        for pid in find_pid_by_name(self.wc.get("process_name", "WinRHIZO.exe")):
            return pid
        return 0

    def _find_filename_edit(self, dlg, wait_s=2.0):
        """找文件对话框的文件名编辑框。不同对话框 ID 不同（1148=打开/新建，1001=导出等），
        逐个尝试并轮询等待；最后兜底枚举所有 Edit。"""
        deadline = time.time() + wait_s
        last_try = 0
        while time.time() < deadline:
            # 1) id=1148 (ComboBoxEx32 → 内部 Edit)
            for ctrl_id in (1148, 1001):
                h = user32.GetDlgItem(dlg, ctrl_id)
                if h:
                    inner = [0]
                    def cb(hw, l):
                        c = ctypes.create_unicode_buffer(32)
                        user32.GetClassNameW(hw, c, 32)
                        if c.value == "Edit":
                            inner[0] = hw
                            return 0
                        return 1
                    user32.EnumChildWindows(h, EnumProc(cb), 0)
                    if inner[0]:
                        return inner[0]
                    # 该控件本身是 Edit 也可用
                    c = ctypes.create_unicode_buffer(32)
                    user32.GetClassNameW(h, c, 32)
                    if c.value == "Edit":
                        return h
            # 2) 兜底：枚举所有后代，收集 Edit 控件（跳过地址栏/搜索框特征类）
            last_try += 1
            if last_try >= 3:
                candidates = []
                def cb2(hw, l):
                    c = ctypes.create_unicode_buffer(64)
                    p = ctypes.create_unicode_buffer(64)
                    user32.GetClassNameW(hw, c, 64)
                    user32.GetClassNameW(user32.GetParent(hw), p, 64)
                    if c.value == "Edit" and p.value not in ("SearchEditBoxWrapperClass", "ComboLBox"):
                        candidates.append(hw)
                    return 1
                user32.EnumChildWindows(dlg, EnumProc(cb2), 0)
                if candidates:
                    # 取矩形最靠下的 Edit（文件名框通常位于底部）
                    best = None
                    best_bottom = -1
                    for hw in candidates:
                        rc = RECT()
                        user32.GetWindowRect(hw, byref(rc))
                        if rc.bottom > best_bottom:
                            best_bottom = rc.bottom
                            best = hw
                    if best:
                        return best
            time.sleep(0.2)
        return None

    def close_dialogs(self, wait_s=1.0):
        """关闭该进程所有可见 #32770 对话框（点取消/确定）。"""
        pid = self._pid()
        if not pid:
            return
        def cb(h, l):
            p = c_uint()
            user32.GetWindowThreadProcessId(h, byref(p))
            if p.value == pid and user32.IsWindowVisible(h):
                c = ctypes.create_unicode_buffer(64)
                user32.GetClassNameW(h, c, 64)
                if c.value == "#32770":
                    click_button(h, 2) or click_button(h, 1)
            return 1
        user32.EnumWindows(EnumProc(cb), 0)
        time.sleep(wait_s)

    def wait_closed(self, hwnd, wait_s=3.0):
        deadline = time.time() + wait_s
        while time.time() < deadline:
            if not user32.IsWindow(hwnd):
                return True
            time.sleep(0.15)
        return False

    def _real_key(self, vk):
        """真实键盘按下+抬起（走系统输入队列，现代对话框可正常识别）。"""
        user32.keybd_event(vk, 0, 0, 0)
        time.sleep(0.05)
        user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.05)

    def _real_ctrl(self, vk):
        """真实键盘 Ctrl+按键。"""
        user32.keybd_event(VK_CONTROL, 0, 0, 0)
        time.sleep(0.05)
        user32.keybd_event(vk, 0, 0, 0)
        time.sleep(0.05)
        user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)
        user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.05)

    # ---------- 打开文件（真实键盘驱动，避免弄坏现代文件对话框） ----------
    def open_file(self, path, is_save=False, wait_dlg=None, cmd=None):
        """
        发文件命令 → 等对话框 → 文件名框用【真实键盘】Ctrl+A/Ctrl+V 粘贴完整路径 → 回车。
        返回 True=对话框已关闭（视为成功），False=失败/超时。
        注意：绝不用 WM_SETTEXT/BM_CLICK 操作现代文件对话框（会损坏 IFileDialog 实例）。
        cmd 可显式指定（如 Regions>Export=9488）。
        """
        if cmd is None:
            cmd = self.wc.get("cmd_new_file" if is_save else "cmd_open_file", 4025 if is_save else 4030)
        before = set()
        def snap(h, l):
            before.add(h)
            return 1
        user32.EnumWindows(EnumProc(snap), 0)
        self.menu(cmd)
        dlg = self._find_new_dialog(before, wait_dlg or float(self.wc.get("file_dialog_wait_s", 3.0)))
        if not dlg:
            self.log("  [open_file] 文件对话框未出现")
            return False
        focus_and_foreground(dlg)
        edit = self._find_filename_edit(dlg)
        if not edit:
            self.log("  [open_file] 未找到文件名编辑框")
            click_button(dlg, 2)
            return False
        user32.SetFocus(edit)
        time.sleep(0.3)
        set_clipboard_text(path)
        # 真实键盘：Ctrl+A 全选 → Ctrl+V 粘贴 → 读回验证
        self._real_ctrl(VK_A)
        self._real_ctrl(VK_V)
        time.sleep(0.4)
        typed = get_window_text(edit)
        if os.path.basename(path) not in typed:
            # 重试一次
            self.log("  [open_file] 粘贴验证失败(%r)，重试一次" % typed)
            self._real_ctrl(VK_A)
            self._real_ctrl(VK_V)
            time.sleep(0.4)
            typed = get_window_text(edit)
        self.log("  [open_file] 文件名框=[%s]" % typed)
        if os.path.basename(path) not in typed:
            self.log("  [open_file] 粘贴始终未生效，取消（不做危险兜底）")
            click_button(dlg, 2)
            return False
        # 回车确认
        self._real_key(VK_RETURN)
        closed = self.wait_closed(dlg, float(self.wc.get("file_dialog_wait_s", 3.0)))
        self.log("  [open_file] %s -> %s" % ("保存" if is_save else "打开", "OK" if closed else "FAIL"))
        return closed

    # ---------- 获取图片（Image > Acquire Image） ----------
    def acquire_image(self, path):
        """
        用 Image>Acquire Image(6400) 打开图片。文件名框自动填入 '*jpg'（或上次名字），
        需替换为图片原始文件名（不含后缀）。返回 True=成功。
        """
        import uiautomation as auto
        base = os.path.splitext(os.path.basename(path))[0]
        before = set()
        def snap(hw, l):
            before.add(hw)
            return 1
        user32.EnumWindows(EnumProc(snap), 0)
        self.menu(6400)
        dlg = self._find_new_dialog(before, 4.0)
        if not dlg:
            self.log("  [acquire] 对话框未出现")
            return False
        time.sleep(1.5)
        root = auto.ControlFromHandle(dlg)
        edit = None
        for aid in ("1148", "1001"):
            e = root.EditControl(automationId=aid, searchDepth=8)
            if e.Exists(0.5, 0.1):
                edit = e
                break
        if edit is None:
            self.log("  [acquire] 找不到文件名框")
            click_button(dlg, 2)
            return False
        ok = False
        for i in range(5):
            try:
                edit.GetLegacyIAccessiblePattern().SetValue(base)
                ok = True
                break
            except Exception:
                time.sleep(0.6)
        if not ok:
            self.log("  [acquire] 设置文件名失败")
            click_button(dlg, 2)
            return False
        time.sleep(0.4)
        focus_and_foreground(dlg)
        time.sleep(0.5)
        self._real_key(VK_RETURN)
        time.sleep(3.0)
        closed = self.wait_closed(dlg, 2.0)
        self.log("  [acquire] %s -> %s" % (base, "OK" if closed else "未关闭"))
        return closed

    # ---------- 缩放 ----------
    def _main_rect(self):
        h = self.ensure_main()
        if not h:
            return None
        rc = RECT()
        user32.GetWindowRect(h, byref(rc))
        return (rc.left, rc.top, rc.right, rc.bottom)

    def _cmd_button_pos(self, btn_id):
        out = []
        def cb(w, l):
            c = ctypes.create_unicode_buffer(80)
            user32.GetClassNameW(w, c, 80)
            if c.value == "Button" and user32.GetDlgCtrlID(w) == btn_id:
                rc = RECT()
                user32.GetWindowRect(w, byref(rc))
                out.append(((rc.left + rc.right) // 2, (rc.top + rc.bottom) // 2))
            return 1
        user32.EnumChildWindows(self.ensure_main(), EnumProc(cb), 0)
        return out[0] if out else None

    def _real_click(self, x, y):
        user32.SetCursorPos(x, y)
        time.sleep(0.15)
        user32.mouse_event(0x0002, 0, 0, 0, 0)
        time.sleep(0.08)
        user32.mouse_event(0x0004, 0, 0, 0, 0)
        time.sleep(0.4)

    def get_zoom(self):
        """读缩放显示值（如 '1/8.0'）。返回字符串。"""
        out = []
        def cb(w, l):
            t = ctypes.create_unicode_buffer(32)
            user32.GetWindowTextW(w, t, 32)
            if user32.GetDlgCtrlID(w) == 615:
                out.append(t.value.strip())
            return 1
        user32.EnumChildWindows(self.ensure_main(), EnumProc(cb), 0)
        return out[0] if out else ""

    def zoom_to(self, target="1/8.0", max_clicks=30):
        """
        缩小到目标缩放值（默认 1/8.0）。按钮：604=放大, 605=缩小。
        用真实鼠标点击。返回达到的缩放值。
        """
        self.log("  [zoom] 目标 %s ..." % target)
        focus_and_foreground(self.ensure_main())
        time.sleep(0.3)
        cur = self.get_zoom()
        for _ in range(max_clicks):
            if cur == target:
                break
            if cur == "":
                break
            # 解析当前值为数字
            try:
                if "/" in cur:
                    num, den = cur.split("/")
                    curv = float(num) / float(den)
                else:
                    curv = float(cur)
                tv = 0.125 if target == "1/8.0" else float(target)
            except Exception:
                break
            if curv > tv:
                pos = self._cmd_button_pos(605)  # 缩小
            else:
                pos = self._cmd_button_pos(604)  # 放大
            if not pos:
                break
            self._real_click(*pos)
            time.sleep(0.5)
            cur = self.get_zoom()
        self.log("  [zoom] 当前 %s" % cur)
        return cur

    # ---------- 分析后对话框处理 ----------
    def handle_analysis_dialogs(self, data_file):
        """
        处理分析后可能出现的一系列对话框：
          - Sample Identification -> OK
          - "No data file is open" -> Open one -> 数据文件选择器 -> 填 testauto -> 回车
        在 20 秒内轮询处理。返回 True=全部处理完。
        """
        import uiautomation as auto
        deadline = time.time() + 25
        last = None
        while time.time() < deadline:
            # 找当前可见、且不是刚处理过的对话框
            dlg = None
            for w in windows_of_pid(self._pid()):
                if w[1] == "#32770" and user32.IsWindowVisible(w[0]) and w[0] != last:
                    dlg = w[0]
                    break
            if dlg is None:
                # 没有新对话框：可能正在等待选择器出现，继续轮询
                time.sleep(0.3)
                continue
            title = ctypes.create_unicode_buffer(128)
            user32.GetWindowTextW(dlg, title, 128)
            t = title.value
            if "Sample" in t or "Identific" in t:
                self.log("  [dlg] Sample Identification -> OK")
                click_button(dlg, 1)
                last = dlg
            elif t == "" or "data" in t.lower():
                self.log("  [dlg] 数据未保存 -> Open one")
                self._click_dlg_button_id(dlg, 101)
                last = dlg
            elif ("Open" in t or "打开" in t or "另存" in t):
                time.sleep(0.8)
                root = auto.ControlFromHandle(dlg)
                edit = None
                for aid in ("1148", "1001"):
                    e = root.EditControl(automationId=aid, searchDepth=8)
                    if e.Exists(0.5, 0.1):
                        edit = e
                        break
                if edit:
                    base = os.path.splitext(os.path.basename(data_file))[0]
                    for i in range(4):
                        try:
                            edit.GetLegacyIAccessiblePattern().SetValue(base)
                            break
                        except Exception:
                            time.sleep(0.5)
                    time.sleep(0.4)
                    focus_and_foreground(dlg)
                    time.sleep(0.4)
                    self._real_key(VK_RETURN)
                    self.log("  [dlg] 数据文件选择 -> %s" % base)
                last = dlg
            else:
                # 未知对话框，点取消
                self.log("  [dlg] 未知对话框关闭: %r" % t)
                click_button(dlg, 2)
                last = dlg
            time.sleep(0.5)
        # 清理可能残留的对话框
        self.close_dialogs(0.3)
        return True

    def _click_dlg_button_id(self, dlg, btn_id):
        btn = user32.GetDlgItem(dlg, btn_id)
        if btn:
            user32.SendMessageW(btn, BM_CLICK, 0, 0)
            return True
        return False

    # ---------- 图像是否显示（像素方差校验） ----------
    def image_displayed(self):
        """
        截图主窗口图像区，返回 (mean, std)。
        图像已打开时 std 较高（暗根系 vs 浅背景）；空窗口 std 低。失败返回 (-1, -1)。
        """
        try:
            from PIL import ImageGrab
        except Exception:
            return -1, -1
        h = self.ensure_main()
        if not h:
            return -1, -1
        focus_and_foreground(h)
        rc = RECT()
        user32.GetWindowRect(h, byref(rc))
        img = np.asarray(ImageGrab.grab(bbox=(rc.left, rc.top, rc.right, rc.bottom)).convert("L"))
        H, W = img.shape
        # 图像区 = 命令栏右侧、顶部条下方
        area = img[int(H * 0.12):int(H * 0.95), int(W * 0.12):int(W * 0.97)]
        return float(area.mean()), float(area.std())

    # ---------- 应用 ROI ----------
    def _view_rect(self):
        """
        估算图像显示区（主窗口内命令栏右侧、顶部条下方），用 MDI 子窗口客户区换算屏幕坐标。
        返回 (left, top, right, bottom) 屏幕坐标或 None。
        """
        h = self.ensure_main()
        if not h:
            return None
        # 找可见的 MDI 子框架
        mdi = [0]
        def cb(w, l):
            c = ctypes.create_unicode_buffer(128)
            user32.GetClassNameW(w, c, 128)
            if ":b:" in c.value and user32.IsWindowVisible(w):
                mdi[0] = w
                return 0
            return 1
        user32.EnumChildWindows(h, EnumProc(cb), 0)
        if not mdi[0]:
            return None
        rc = RECT()
        if not user32.GetClientRect(mdi[0], byref(rc)):
            return None
        pt = POINT(0, 0)
        user32.ClientToScreen(mdi[0], byref(pt))
        # 命令栏在左（约 110px），顶部有条状控制条（约 55px）
        left = pt.x + 115
        top = pt.y + 55
        return (left, top, pt.x + rc.right - 5, pt.y + rc.bottom - 15)

    def apply_roi(self, roi, img_size):
        """
        应用 ROI。返回 True/False。
        方法 B：模拟鼠标拖框。映射用当前缩放值（如 1/8.0 -> scale=0.125），图像在视图内居中。
        """
        if not roi or not img_size:
            return False
        L, T, R, B = roi
        vw, vh, vx, vy = self._view_size()
        if vw is None:
            self.log("  [apply_roi] 找不到图像显示区")
            return False
        img_w, img_h = img_size
        # 从缩放显示解析 scale（如 "1/8.0" -> 0.125）
        zoom = self.get_zoom()
        scale = None
        try:
            if "/" in zoom:
                num, den = zoom.split("/")
                scale = float(num) / float(den)
            else:
                scale = float(zoom)
        except Exception:
            scale = None
        if scale is None or scale <= 0:
            # 兜底：fit-to-window
            scale = min(vw / img_w, vh / img_h)
        # 先恢复窗口到校准尺寸/位置，保证映射一致
        h = self.ensure_main()
        if h:
            win_w = int(float(self.wc.get("calib_win_w", 1643)))
            win_h = int(float(self.wc.get("calib_win_h", 1003)))
            win_x = int(float(self.wc.get("calib_win_x", 40)))
            win_y = int(float(self.wc.get("calib_win_y", 40)))
            user32.SetWindowPos(h, 0, win_x, win_y, win_w, win_h, 0x40)
            time.sleep(0.6)
        # 恢复后再读窗口位置，算图像左上角屏幕坐标（相对窗口原点）
        win = self._main_rect()
        wox = float(self.wc.get("view_ox_off", vx - (win[0] if win else 0)))
        woy = float(self.wc.get("view_oy_off", vy - (win[1] if win else 0)))
        ox = (win[0] if win else 0) + wox
        oy = (win[1] if win else 0) + woy
        x1, y1 = ox + L * scale, oy + T * scale
        x2, y2 = ox + R * scale, oy + B * scale
        focus_and_foreground(h)
        self._mouse_drag(int(x1), int(y1), int(x2), int(y2))
        self.log("  [apply_roi] 拖框 (%d,%d)->(%d,%d) zoom=%s scale=%.4f" % (int(x1), int(y1), int(x2), int(y2), zoom, scale))
        return True

    def _view_size(self):
        vr = self._view_rect()
        if not vr:
            return None, None, None, None
        x, y, rx, ry = vr
        return rx - x, ry - y, x, y

    def _mouse_drag(self, x1, y1, x2, y2):
        """用 SendInput 模拟一次鼠标拖拽。"""
        time.sleep(0.2)
        ctypes.windll.user32.SetCursorPos(x1, y1)
        time.sleep(0.15)
        self._mouse_event_down()
        steps = 12
        for i in range(1, steps + 1):
            ctypes.windll.user32.SetCursorPos(
                int(x1 + (x2 - x1) * i / steps),
                int(y1 + (y2 - y1) * i / steps))
            time.sleep(0.03)
        time.sleep(0.15)
        self._mouse_event_up()

    def _mouse_event_down(self):
        user32.mouse_event(0x0002, 0, 0, 0, 0)  # LEFTDOWN

    def _mouse_event_up(self):
        user32.mouse_event(0x0004, 0, 0, 0, 0)  # LEFTUP

    # ---------- 等待分析结果 ----------
    def wait_for_data(self, data_file, filename, before_rows, timeout_s):
        """等 data_file 出现 filename 的新数据行。返回新增行数。"""
        deadline = time.time() + timeout_s
        last_count = before_rows
        while time.time() < deadline:
            try:
                rows = DataFileReader.read_data_rows(data_file)
            except Exception:
                rows = []
            new_rows = [r for r in rows[last_count:] if r.get("filename") == filename]
            if new_rows:
                return len(new_rows)
            time.sleep(0.5)
        return 0


# ============================================================
# 数据文件解析
# ============================================================
COL_INDEX = {
    "datetime": 2, "filename": 3,
    "area": 5, "width": 6, "height": 7, "roi": 8,
    "Length": 15, "ProjArea": 17, "SurfArea": 19, "AvgDiam": 21,
    "RootVolume": 25, "Tips": 28,
}


class DataFileReader:
    @staticmethod
    def read_all(path):
        """读取 WinRHIZO 结果 TXT。返回 {'header': [...], 'rows': [{...}]}。"""
        if not os.path.exists(path):
            return {"header": [], "rows": []}
        with open(path, "rb") as f:
            raw = f.read()
        text = raw.decode("gbk", errors="replace")
        lines = [l for l in text.split("\r\n") if l.strip()]
        header = lines[0].split("\t") if lines else []
        rows = []
        for ln in lines[1:]:
            if not ln.strip() or ln.startswith("SampleId") or ln.startswith("RHIZO"):
                continue
            cells = ln.split("\t")
            if len(cells) <= 8:
                continue
            r = {}
            for key, idx in COL_INDEX.items():
                r[key] = cells[idx] if idx < len(cells) else ""
            rows.append(r)
        return {"header": header, "rows": rows}

    @staticmethod
    def read_data_rows(path):
        return DataFileReader.read_all(path)["rows"]

    @staticmethod
    def count_rows(path):
        return len(DataFileReader.read_data_rows(path))

    @staticmethod
    def count_for_file(path, filename):
        rows = DataFileReader.read_data_rows(path)
        return len([r for r in rows if r.get("filename") == filename])


class DataWatcher:
    def __init__(self, data_file):
        self.data_file = data_file
        self.last_rows = 0

    def snapshot(self):
        self.last_rows = DataFileReader.count_rows(self.data_file)

    def new_rows_for(self, filename, timeout_s, poll=0.5):
        deadline = time.time() + timeout_s
        start = self.last_rows
        while time.time() < deadline:
            rows = DataFileReader.read_data_rows(self.data_file)
            new = [r for r in rows[start:] if r.get("filename") == filename]
            if new:
                self.last_rows = len(rows)
                return new
            time.sleep(poll)
        return []


# ============================================================
# 汇总
# ============================================================
def parse_sample_key(filename):
    """F120-001.jpg -> ('F', '120', '001', 'F120-001')；解析不了返回 None。"""
    m = re.match(r"^\s*([A-Za-z])\s*(\d+)\s*[-_]?\s*(\d{2,3})", filename)
    if not m:
        return None
    prefix, plot, rep = m.group(1).upper(), m.group(2), m.group(3)
    return prefix, plot, rep, "%s%s-%s" % (prefix, plot, rep)


class Aggregator:
    def __init__(self, cfg, log_fn=log):
        self.cfg = cfg
        self.labels = cfg.get("labels", {})
        self.log_fn = log_fn

    def build(self, data_file, out_csv, out_xlsx=None):
        """解析 TXT，全部保留多次分析，按样本分组，输出 CSV/XLSX。"""
        res = DataFileReader.read_all(data_file)
        rows = res["rows"]
        records = []
        for r in rows:
            fn = r.get("filename", "")
            key = parse_sample_key(fn)
            if not key:
                continue
            prefix, plot, rep, sample = key
            roi_str = r.get("roi", "")
            parts = roi_str.split()
            roi = " / ".join(parts[:4]) if len(parts) >= 4 else roi_str
            records.append({
                "样本": sample,
                "处理": self.labels.get(prefix, prefix),
                "样地": plot,
                "重复": rep,
                "文件名": fn,
                "分析时间": r.get("datetime", ""),
                "ROI(左 上 右 下)": roi,
                "区域面积cm2": r.get("area", ""),
                "总根长cm": r.get("Length", ""),
                "投影面积cm2": r.get("ProjArea", ""),
                "表面积cm2": r.get("SurfArea", ""),
                "平均直径mm": r.get("AvgDiam", ""),
                "体积cm3": r.get("RootVolume", ""),
                "根尖数Tips": r.get("Tips", ""),
            })
        # 每样本内按时间排序 + 分析序号
        order = {}
        for rec in records:
            key = rec["样本"]
            order.setdefault(key, 0)
            order[key] += 1
            rec["分析序号"] = order[key]
        records.sort(key=lambda x: (x["样本"], x["分析序号"]))
        columns = ["样本", "处理", "样地", "重复", "分析序号", "文件名", "分析时间",
                   "ROI(左 上 右 下)", "区域面积cm2", "总根长cm", "投影面积cm2",
                   "表面积cm2", "平均直径mm", "体积cm3", "根尖数Tips"]
        # CSV（UTF-8 BOM，Excel 友好）
        with open(out_csv, "w", encoding="utf-8-sig", newline="") as f:
            import csv
            w = csv.DictWriter(f, fieldnames=columns)
            w.writeheader()
            for rec in records:
                w.writerow({k: rec.get(k, "") for k in columns})
        # XLSX
        if out_xlsx:
            try:
                import openpyxl
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "汇总"
                ws.append(columns)
                for rec in records:
                    ws.append([rec.get(k, "") for k in columns])
                wb.save(out_xlsx)
            except Exception as e:
                self.log_fn and self.log_fn("XLSX 输出失败: %s" % e)
        return len(records)


# ============================================================
# GUI + 主流程
# ============================================================
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox


class App:
    def __init__(self, cfg):
        self.cfg = cfg
        self.image_base = resolve(BASE_DIR, cfg.get("image_dir", "图片"))
        self.data_dir = resolve(BASE_DIR, cfg.get("data_dir", "数据"))
        self.date = str(cfg.get("date", ""))
        self.images_dir = os.path.join(self.image_base, str(cfg.get("image_date", self.date)))
        self.preview_dir = os.path.join(self.images_dir, cfg.get("roi_preview_dir", "_ROI预览"))
        if cfg.get("data_file"):
            self.default_data_file = resolve(BASE_DIR, cfg["data_file"])
        else:
            self.default_data_file = os.path.join(self.data_dir, "%s.TXT" % self.date)
        self.data_file = self.default_data_file
        self.roi_detector = ROIDetector(cfg.get("roi_detect", {}))
        self.wr = WinRhizo(cfg, log_fn=self._log)
        self.aggregator = Aggregator(cfg, log_fn=self._log)

        self.events = queue.Queue()
        self.ctl = {
            "pause": threading.Event(),
            "abort": threading.Event(),
            "skip": threading.Event(),
            "redo": threading.Event(),
            "user_done": threading.Event(),
        }
        self.worker = None
        self.awaiting = False
        self.running = False
        self.images = []

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(120, self._poll)

    # ---------- UI ----------
    def _build_ui(self):
        self.root = tk.Tk()
        self.root.title("WinRHIZO 半自动批量分析  (%s)" % self.date)
        self.root.geometry("720x560")

        top = ttk.Frame(self.root, padding=(8, 6))
        top.pack(fill="x")
        ttk.Label(top, text="待处理图片：").pack(side="left")
        self.lbl_images = ttk.Label(top, text="")
        self.lbl_images.pack(side="left", padx=6)
        ttk.Label(top, text="进度：").pack(side="left", padx=(20, 0))
        self.lbl_progress = ttk.Label(top, text="0/0")
        self.lbl_progress.pack(side="left", padx=6)

        mid = ttk.Frame(self.root, padding=(8, 0))
        mid.pack(fill="both", expand=True)
        self.listbox = tk.Listbox(mid, width=30)
        self.listbox.pack(side="left", fill="y", padx=(0, 6))
        info = ttk.Frame(mid)
        info.pack(side="left", fill="both", expand=True)
        ttk.Label(info, text="当前图片：").pack(anchor="w")
        self.lbl_current = ttk.Label(info, text="（未开始）", font=("", 10, "bold"))
        self.lbl_current.pack(anchor="w", pady=(0, 6))
        ttk.Label(info, text="检测 ROI：").pack(anchor="w")
        self.lbl_roi = ttk.Label(info, text="")
        self.lbl_roi.pack(anchor="w", pady=(0, 6))
        ttk.Label(info, text="状态：").pack(anchor="w")
        self.lbl_status = ttk.Label(info, text="就绪", foreground="#1a6bd6")
        self.lbl_status.pack(anchor="w", pady=(0, 6))
        ttk.Label(info, text="数据文件：").pack(anchor="w")
        self.lbl_data = ttk.Label(info, text=self.default_data_file, foreground="#666")
        self.lbl_data.pack(anchor="w")

        logframe = ttk.LabelFrame(self.root, text="日志", padding=(6, 4))
        logframe.pack(fill="both", expand=True, padx=8, pady=(6, 0))
        self.txt = scrolledtext.ScrolledText(logframe, height=12, state="disabled")
        self.txt.pack(fill="both", expand=True)

        btns = ttk.Frame(self.root, padding=(8, 6))
        btns.pack(fill="x")
        self.btn_start = ttk.Button(btns, text="开始", command=self._on_start)
        self.btn_start.pack(side="left")
        self.btn_pause = ttk.Button(btns, text="暂停", command=self._on_pause, state="disabled")
        self.btn_pause.pack(side="left", padx=4)
        self.btn_skip = ttk.Button(btns, text="跳过", command=self._on_skip, state="disabled")
        self.btn_skip.pack(side="left", padx=4)
        self.btn_redo = ttk.Button(btns, text="重做", command=self._on_redo, state="disabled")
        self.btn_redo.pack(side="left", padx=4)
        self.btn_done = ttk.Button(btns, text="完成", command=self._on_done, state="disabled")
        self.btn_done.pack(side="left", padx=4)
        self.btn_abort = ttk.Button(btns, text="中止", command=self._on_abort, state="disabled")
        self.btn_abort.pack(side="left", padx=4)
        ttk.Label(btns, text="  F12 = 本图已人工完成", foreground="#888").pack(side="left", padx=12)

    def _log(self, msg):
        self.events.put(("log", msg))

    # ---------- 按钮事件 ----------
    def _on_start(self):
        if self.running:
            return
        # 数据文件策略
        if os.path.exists(self.default_data_file):
            use_existing = messagebox.askyesno(
                "数据文件已存在",
                "已有 %s\n\n是否继续使用它（追加，支持断点续跑）？\n选【否】则新建一个全新数据文件（绝不覆盖）。" % os.path.basename(self.default_data_file))
            if use_existing:
                self.data_file = self.default_data_file
            else:
                self.data_file = self._fresh_path(self.default_data_file)
        else:
            self.data_file = self.default_data_file
        self.lbl_data.config(text=self.data_file)
        self.ctl["abort"].clear()
        self.ctl["pause"].clear()
        self.running = True
        self.btn_start.config(state="disabled")
        for b in (self.btn_pause, self.btn_skip, self.btn_redo, self.btn_done, self.btn_abort):
            b.config(state="normal")
        self.worker = threading.Thread(target=self._run_batch, daemon=True)
        self.worker.start()

    def _on_pause(self):
        if self.ctl["pause"].is_set():
            self.ctl["pause"].clear()
            self.btn_pause.config(text="暂停")
            self._log("已继续")
        else:
            self.ctl["pause"].set()
            self.btn_pause.config(text="继续")
            self._log("已暂停（当前图片完成后再暂停）")

    def _on_skip(self):
        self.ctl["skip"].set()
        self._log("请求跳过当前图片")

    def _on_redo(self):
        self.ctl["redo"].set()
        self.ctl["user_done"].set()
        self._log("请求重做当前图片")

    def _on_done(self):
        self.ctl["user_done"].set()
        self._log("收到【完成】信号")

    def _on_abort(self):
        self.ctl["abort"].set()
        self.ctl["user_done"].set()
        self._log("已请求中止")

    # ---------- 工具 ----------
    def _fresh_path(self, path):
        base, ext = os.path.splitext(path)
        i = 1
        while os.path.exists("%s_新%d%s" % (base, i, ext)):
            i += 1
        return "%s_新%d%s" % (base, i, ext)

    def _enumerate_images(self):
        if not os.path.isdir(self.images_dir):
            return []
        files = [f for f in os.listdir(self.images_dir)
                 if f.lower().endswith((".jpg", ".jpeg", ".tif", ".tiff"))]
        files.sort()
        return [(f, os.path.join(self.images_dir, f)) for f in files]

    def _await_user(self):
        """等待用户在 WinRHIZO 中人工完成当前图（按 F12 或点【完成】）。"""
        self.awaiting = True
        self.ctl["user_done"].clear()
        self.events.put(("prompt", True))
        while not self.ctl["user_done"].is_set():
            if self.ctl["abort"].is_set():
                break
            if self.ctl["skip"].is_set():
                self.ctl["skip"].clear()
                break
            time.sleep(0.2)
        self.ctl["user_done"].clear()
        self.awaiting = False
        self.events.put(("prompt", False))

    def _wait_if_paused(self):
        while self.ctl["pause"].is_set() and not self.ctl["abort"].is_set():
            time.sleep(0.3)

    # ---------- 主流程（工作线程） ----------
    def _create_data_file(self):
        if os.path.exists(self.data_file):
            self._log("使用已有数据文件：%s" % self.data_file)
            return True
        self._log("创建数据文件：%s" % self.data_file)
        self.wr.close_dialogs()
        ok = self.wr.open_file(self.data_file, is_save=True)
        time.sleep(1.2)
        if not os.path.exists(self.data_file):
            self._log("数据文件未生成（可能弹出了覆盖/格式确认框），请人工检查后按 F12 继续")
            self.wr.close_dialogs()
            self._await_user()
            return os.path.exists(self.data_file)
        self._log("数据文件就绪")
        return True

    def _run_batch(self):
        try:
            if not self.wr.find_main():
                self.events.put(("alert", "未找到 WinRHIZO 窗口。请先启动 WinRHIZO 并确认 Sentinel 加密狗已插入。"))
                return
            self._log("WinRHIZO 主窗口已定位")
            self.wr.close_dialogs()
            if not self._create_data_file():
                self.events.put(("alert", "数据文件创建失败，已中止。"))
                return
            images = self._enumerate_images()
            self.images = images
            if not images:
                self.events.put(("alert", "图片目录为空：%s" % self.images_dir))
                return
            total = len(images)
            self._log("共 %d 张图片" % total)
            self.events.put(("queue", [f for f, _ in images]))
            done = 0
            for idx, (img, path) in enumerate(images):
                if self.ctl["abort"].is_set():
                    break
                self._wait_if_paused()
                self.events.put(("image", img, idx))
                # 断点续跑：数据文件中已有该图则跳过
                if DataFileReader.count_for_file(self.data_file, img) > 0:
                    self._log("跳过 %s（数据文件中已存在）" % img)
                    done += 1
                    self.events.put(("progress", done, total))
                    continue
                if self.ctl["skip"].is_set():
                    self.ctl["skip"].clear()
                    self._log("跳过 %s" % img)
                    done += 1
                    self.events.put(("progress", done, total))
                    continue
                # ---- 1. 检测 ROI ----
                self._log("检测 ROI：%s" % img)
                roi, gimg = self.roi_detector.detect(path)
                if roi:
                    self.events.put(("roi", "L=%d  T=%d  R=%d  B=%d" % roi))
                    prev = self.roi_detector.save_preview(path, roi, self.preview_dir, img)
                    if prev:
                        self._log("ROI 预览：%s" % prev)
                else:
                    self.events.put(("roi", "（未检测到根系）"))
                img_size = (gimg.shape[1], gimg.shape[0]) if gimg is not None else None
                # ---- 2. 打开图片（手动：程序化打开现代对话框不可靠）----
                self.events.put(("status", "请用 WinRHIZO 获取图片：%s（获取后按 F12 继续）" % img))
                self._await_user()
                time.sleep(1.0)
                # ---- 2.5 缩小到能看到完整根系 ----
                self.events.put(("status", "缩小视图 ..."))
                target_zoom = self.cfg.get("winrhizo", {}).get("target_zoom", "1/8.0")
                self.wr.zoom_to(target_zoom)
                time.sleep(0.5)
                cur_zoom = self.wr.get_zoom()
                if cur_zoom != target_zoom:
                    self._log("警告：缩放未达目标 %s（当前 %s），画框可能不准" % (target_zoom, cur_zoom))
                # ---- 3. 应用 ROI（画框即自动分析）----
                if roi and img_size:
                    if self.wr.apply_roi(roi, img_size):
                        self._log("已绘制自动 ROI")
                    else:
                        self.events.put(("status", "自动 ROI 失败，请人工框选后按 F12"))
                        self._await_user()
                else:
                    self.events.put(("status", "未检测到根系，请人工框选 ROI 后按 F12"))
                    self._await_user()
                # ---- 4. 处理对话框 + 等待分析结果 ----
                self.events.put(("status", "等待分析结果 ..."))
                self.wr.handle_analysis_dialogs(self.data_file)
                watch = DataWatcher(self.data_file)
                watch.snapshot()
                new = watch.new_rows_for(img, float(self.cfg["winrhizo"].get("analysis_wait_s", 8)))
                if not new:
                    self._log("未检测到新数据。若弹出了保存对话框，请手动处理：选择 Open one → 选 %s → 确定" % os.path.basename(self.data_file))
                    self.events.put(("status", "请在 WinRHIZO 确认分析并保存数据到 %s，完成后按 F11" % os.path.basename(self.data_file)))
                    self._await_user()
                    new = watch.new_rows_for(img, float(self.cfg["winrhizo"].get("analysis_wait_s", 8)))
                if new:
                    self._log("✓ %s 已获得 %d 条结果" % (img, len(new)))
                    done += 1
                else:
                    self._log("✗ %s 未获得结果（请人工检查）" % img)
                self.events.put(("progress", done, total))
            # ---- 5. 收尾 + 汇总 ----
            self.events.put(("status", "批量完成，关闭数据文件并汇总 ..."))
            self.wr.menu(self.cfg["winrhizo"].get("cmd_close_file", 4035))
            time.sleep(1.0)
            self.wr.close_dialogs()
            base = os.path.basename(self.data_file).rsplit(".", 1)[0]
            out_csv = os.path.join(self.data_dir, base + "_汇总.csv")
            out_xlsx = os.path.join(self.data_dir, base + "_汇总.xlsx")
            n = self.aggregator.build(self.data_file, out_csv, out_xlsx)
            self._log("汇总完成：%s（%d 条记录）" % (out_csv, n))
            self.events.put(("status", "全部完成"))
            self.events.put(("done", None))
        except Exception as e:
            self.events.put(("alert", "运行出错：%s\n%s" % (e, traceback.format_exc())))

    # ---------- 事件轮询 ----------
    def _poll(self):
        try:
            while True:
                item = self.events.get_nowait()
                self._handle(item)
        except queue.Empty:
            pass
        self.root.after(120, self._poll)

    def _handle(self, item):
        kind = item[0]
        if kind == "log":
            self._append_log(item[1])
        elif kind == "status":
            self.lbl_status.config(text=item[1])
        elif kind == "queue":
            self.listbox.delete(0, "end")
            for f in item[1]:
                self.listbox.insert("end", f)
            self.lbl_images.config(text="%d 张" % len(item[1]))
        elif kind == "image":
            img, idx = item[1], item[2]
            self.lbl_current.config(text=img)
            self.listbox.selection_clear(0, "end")
            self.listbox.selection_set(idx)
            self.listbox.see(idx)
        elif kind == "roi":
            self.lbl_roi.config(text=item[1])
        elif kind == "progress":
            self.lbl_progress.config(text="%d/%d" % (item[1], item[2]))
        elif kind == "prompt":
            self.btn_done.config(state="normal" if item[1] else "disabled")
        elif kind == "hotkey":
            if self.awaiting:
                self.ctl["user_done"].set()
                self._append_log("F12 已按下")
        elif kind == "alert":
            messagebox.showerror("提示", item[1])
            self._reset_ui()
        elif kind == "done":
            self._reset_ui()

    def _append_log(self, msg):
        self.txt.config(state="normal")
        self.txt.insert("end", "[%s] %s\n" % (time.strftime("%H:%M:%S"), msg))
        self.txt.see("end")
        self.txt.config(state="disabled")

    def _reset_ui(self):
        self.running = False
        self.awaiting = False
        self.btn_start.config(state="normal")
        for b in (self.btn_pause, self.btn_skip, self.btn_redo, self.btn_done, self.btn_abort):
            b.config(state="disabled")
        self.btn_pause.config(text="暂停")

    # ---------- 热键线程 ----------
    def _hotkey_loop(self):
        self.hotkey_id = None
        for hid, vk, name in ((1, VK_F12, "F12"), (2, 0x7A, "F11"), (3, 0x77, "F10")):
            ok = user32.RegisterHotKey(None, hid, 0, vk)
            if ok:
                self.hotkey_id = hid
                self._log("%s 热键注册成功（按 %s = 本图已完成）" % (name, name))
                break
        if self.hotkey_id is None:
            self._log("F12/F11/F10 热键都注册失败，请用窗口【完成】按钮")
        msg = MSG()
        while True:
            r = user32.GetMessageW(byref(msg), 0, 0, 0)
            if r <= 0:
                break
            if msg.message == WM_HOTKEY and (msg.wParam == 1 or msg.wParam == 2 or msg.wParam == 3):
                self.events.put(("hotkey", None))
            else:
                user32.TranslateMessage(byref(msg))
                user32.DispatchMessageW(byref(msg))

    def _on_close(self):
        self.ctl["abort"].set()
        if getattr(self, "hotkey_id", None):
            user32.UnregisterHotKey(None, self.hotkey_id)
        for hid in (1, 2, 3):
            user32.UnregisterHotKey(None, hid)
        try:
            self.root.destroy()
        except Exception:
            pass

    def run(self):
        threading.Thread(target=self._hotkey_loop, daemon=True).start()
        self.root.mainloop()


def main():
    cfg = load_config()
    App(cfg).run()


if __name__ == "__main__":
    main()
