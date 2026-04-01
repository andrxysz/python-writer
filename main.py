from __future__ import annotations

import ctypes
from ctypes import wintypes
from dataclasses import dataclass
import sys

from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)


if sys.platform != "win32":
    raise OSError("Python-Writer currently supports only Windows.")


user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

WM_HOTKEY = 0x0312
PM_REMOVE = 0x0001

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
MOD_NOREPEAT = 0x4000

INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004

VK_BACK = 0x08
VK_TAB = 0x09
VK_RETURN = 0x0D
VK_SHIFT = 0x10
VK_CONTROL = 0x11
VK_MENU = 0x12
VK_SPACE = 0x20

KEY_CHOICES = [*list("ABCDEFGHIJKLMNOPQRSTUVWXYZ"), *list("0123456789")]
KEY_CHOICES += [f"F{i}" for i in range(1, 13)]
KEY_CHOICES += ["Space", "Enter", "Tab"]

KEY_NAME_TO_VK = {letter: ord(letter) for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ"}
KEY_NAME_TO_VK.update({digit: ord(digit) for digit in "0123456789"})
KEY_NAME_TO_VK.update({f"F{i}": 0x6F + i for i in range(1, 13)})
KEY_NAME_TO_VK.update(
    {
        "Space": VK_SPACE,
        "Enter": VK_RETURN,
        "Tab": VK_TAB,
    }
)


user32.RegisterHotKey.argtypes = [
    wintypes.HWND,
    ctypes.c_int,
    wintypes.UINT,
    wintypes.UINT,
]
user32.RegisterHotKey.restype = wintypes.BOOL

user32.UnregisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int]
user32.UnregisterHotKey.restype = wintypes.BOOL

user32.PeekMessageW.argtypes = [
    ctypes.POINTER(wintypes.MSG),
    wintypes.HWND,
    wintypes.UINT,
    wintypes.UINT,
    wintypes.UINT,
]
user32.PeekMessageW.restype = wintypes.BOOL

user32.MapVirtualKeyW.argtypes = [wintypes.UINT, wintypes.UINT]
user32.MapVirtualKeyW.restype = wintypes.UINT

user32.VkKeyScanW.argtypes = [ctypes.c_wchar]
user32.VkKeyScanW.restype = ctypes.c_short

user32.GetForegroundWindow.argtypes = []
user32.GetForegroundWindow.restype = wintypes.HWND

user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD

kernel32.GetCurrentProcessId.argtypes = []
kernel32.GetCurrentProcessId.restype = wintypes.DWORD

ULONG_PTR = wintypes.WPARAM


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", wintypes.DWORD),
        ("wParamL", wintypes.WORD),
        ("wParamH", wintypes.WORD),
    ]


class INPUT_UNION(ctypes.Union):
    _fields_ = [
        ("ki", KEYBDINPUT),
        ("mi", MOUSEINPUT),
        ("hi", HARDWAREINPUT),
    ]


class INPUT(ctypes.Structure):
    _fields_ = [
        ("type", wintypes.DWORD),
        ("union", INPUT_UNION),
    ]


user32.SendInput.argtypes = [wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int]
user32.SendInput.restype = wintypes.UINT


@dataclass(frozen=True)
class HotkeySpec:
    ctrl: bool
    alt: bool
    shift: bool
    win: bool
    key_name: str

    def modifier_mask(self) -> int:
        mask = MOD_NOREPEAT
        if self.ctrl:
            mask |= MOD_CONTROL
        if self.alt:
            mask |= MOD_ALT
        if self.shift:
            mask |= MOD_SHIFT
        if self.win:
            mask |= MOD_WIN
        return mask

    def virtual_key(self) -> int:
        return KEY_NAME_TO_VK[self.key_name]

    def display_label(self) -> str:
        parts = []
        if self.ctrl:
            parts.append("Ctrl")
        if self.alt:
            parts.append("Alt")
        if self.shift:
            parts.append("Shift")
        if self.win:
            parts.append("Win")
        parts.append(self.key_name)
        return " + ".join(parts)

    def has_modifier(self) -> bool:
        return self.ctrl or self.alt or self.shift or self.win


DEFAULT_HOTKEY = HotkeySpec(ctrl=True, alt=True, shift=False, win=False, key_name="W")
DEFAULT_DELAY_MS = 45
SAMPLE_TEXT = (
    "Escrevendo com o Python-Writer.\n"
    "Escolha a hotkey, deixe o cursor no campo certo e acompanhe a digitacao."
)


class GlobalHotkeyManager:
    def __init__(self, hwnd_provider):
        self._hwnd_provider = hwnd_provider
        self._hotkey_id = 1
        self._registered = False

    def register(self, hotkey: HotkeySpec) -> None:
        self.unregister()
        result = user32.RegisterHotKey(
            self._hwnd(),
            self._hotkey_id,
            hotkey.modifier_mask(),
            hotkey.virtual_key(),
        )
        if not result:
            raise ctypes.WinError(ctypes.get_last_error())
        self._registered = True

    def unregister(self) -> None:
        if self._registered:
            user32.UnregisterHotKey(self._hwnd(), self._hotkey_id)
            self._registered = False

    def is_hotkey_message(self, message: wintypes.MSG) -> bool:
        return message.message == WM_HOTKEY and message.wParam == self._hotkey_id

    def _hwnd(self) -> int:
        return int(self._hwnd_provider())


def _send_keyboard_input(w_vk: int, w_scan: int, flags: int) -> None:
    event = INPUT(
        type=INPUT_KEYBOARD,
        union=INPUT_UNION(
            ki=KEYBDINPUT(
                wVk=w_vk,
                wScan=w_scan,
                dwFlags=flags,
                time=0,
                dwExtraInfo=0,
            )
        ),
    )
    sent = user32.SendInput(1, ctypes.byref(event), ctypes.sizeof(INPUT))
    if sent != 1:
        raise ctypes.WinError(ctypes.get_last_error())


def _send_virtual_key(vk_code: int) -> None:
    scan_code = user32.MapVirtualKeyW(vk_code, 0)
    _send_keyboard_input(vk_code, scan_code, 0)
    _send_keyboard_input(vk_code, scan_code, KEYEVENTF_KEYUP)


def _send_key_down(vk_code: int) -> None:
    scan_code = user32.MapVirtualKeyW(vk_code, 0)
    _send_keyboard_input(vk_code, scan_code, 0)


def _send_key_up(vk_code: int) -> None:
    scan_code = user32.MapVirtualKeyW(vk_code, 0)
    _send_keyboard_input(vk_code, scan_code, KEYEVENTF_KEYUP)


def _send_layout_character(char: str) -> bool:
    vk_combo = user32.VkKeyScanW(char)
    if vk_combo == -1:
        return False

    vk_code = vk_combo & 0xFF
    modifiers = (vk_combo >> 8) & 0xFF

    if modifiers & (2 | 4):
        return False

    if modifiers & 1:
        _send_key_down(VK_SHIFT)

    _send_key_down(vk_code)
    _send_key_up(vk_code)

    if modifiers & 1:
        _send_key_up(VK_SHIFT)

    return True


def _foreground_belongs_to_current_process() -> bool:
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return False

    process_id = wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
    return process_id.value == kernel32.GetCurrentProcessId()


def _utf16_units(char: str) -> list[int]:
    encoded = char.encode("utf-16-le")
    return [
        int.from_bytes(encoded[index : index + 2], "little")
        for index in range(0, len(encoded), 2)
    ]


def send_character(char: str) -> None:
    if char == "\n":
        _send_virtual_key(VK_RETURN)
        return
    if char == "\t":
        _send_virtual_key(VK_TAB)
        return
    if char == "\b":
        _send_virtual_key(VK_BACK)
        return

    if not _foreground_belongs_to_current_process() and _send_layout_character(char):
        return

    for code_unit in _utf16_units(char):
        _send_keyboard_input(0, code_unit, KEYEVENTF_UNICODE)
        _send_keyboard_input(0, code_unit, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP)


class TypewriterController:
    START_GRACE_MS = 140

    def __init__(self, notify, parent: QWidget):
        self._notify = notify
        self._timer = QTimer(parent)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self._type_next_character)
        self._text = ""
        self._index = 0
        self._delay_ms = 0

    def is_running(self) -> bool:
        return self._timer.isActive()

    def start(self, text: str, delay_ms: int) -> bool:
        if self.is_running():
            return False

        normalized = text.replace("\r\n", "\n").replace("\r", "\n")
        if normalized == "":
            return False

        self._text = normalized
        self._index = 0
        self._delay_ms = max(delay_ms, 0)

        self._notify("started", "")
        self._timer.start(self.START_GRACE_MS)
        return True

    def stop(self, announce: bool = True) -> None:
        was_running = self.is_running()
        self._timer.stop()
        self._text = ""
        self._index = 0
        if was_running and announce:
            self._notify("stopped", "")

    def _type_next_character(self) -> None:
        if self._index >= len(self._text):
            self._timer.stop()
            self._notify("finished", "")
            return

        try:
            send_character(self._text[self._index])
        except Exception as exc:  # pragma: no cover - Win32 runtime branch
            self._timer.stop()
            self._notify("error", str(exc))
            return

        self._index += 1

        if self._index >= len(self._text):
            self._timer.stop()
            self._notify("finished", "")
            return

        self._timer.start(self._delay_ms)


class PythonWriterWindow(QWidget):
    BG = "#F3EEE6"
    PANEL = "#FFFBF4"
    FIELD = "#F7F1E7"
    BORDER = "#D9CCBC"
    TEXT = "#17212C"
    MUTED = "#5B6874"
    ACCENT = "#1E6C7A"
    ACCENT_ACTIVE = "#164F59"
    HERO_TOP = "#2A6578"
    HERO_BOTTOM = "#18384C"
    SOFT = "#E5EFF1"
    SOFT_ALT = "#EEF4EA"
    WARN = "#B45309"
    WARN_ACTIVE = "#8B3E12"

    def __init__(self) -> None:
        super().__init__()
        self._current_hotkey_label = ""
        self._active_hotkey: HotkeySpec | None = None

        self.setWindowTitle("Python-Writer")
        self.resize(560, 760)
        self.setMinimumSize(520, 720)
        self.setObjectName("window")

        self.hotkeys = GlobalHotkeyManager(self.winId)
        self.typewriter = TypewriterController(self._handle_typewriter_event, self)

        self._build_ui()
        self._wire_ui_signals()
        self._apply_styles()
        self._refresh_hotkey_details()
        self._update_text_metrics()
        QTimer.singleShot(0, self.apply_hotkey)

    def _build_ui(self) -> None:
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(20, 20, 20, 20)
        root_layout.setSpacing(14)

        root_layout.addWidget(self._build_hero_card())

        root_layout.addWidget(
            self._build_info_card(
                "Fluxo rapido",
                "Cole o texto, deixe o cursor no campo desejado e pressione a hotkey ativa. "
                "A mesma hotkey interrompe a digitacao e o botao Parar fica disponivel durante o envio.",
            )
        )

        root_layout.addWidget(self._build_config_card())
        root_layout.addWidget(self._build_text_card(), stretch=1)
        root_layout.addWidget(self._build_status_card())

    def _build_hero_card(self) -> QFrame:
        card = QFrame()
        card.setProperty("hero", True)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        title = QLabel("Python-Writer")
        title.setProperty("role", "heroTitle")

        subtitle = QLabel(
            "Automatize textos com uma hotkey global, ritmo configuravel e feedback visual em tempo real."
        )
        subtitle.setWordWrap(True)
        subtitle.setProperty("role", "heroSubtitle")

        badges = QHBoxLayout()
        badges.setSpacing(8)
        self.active_hotkey_badge = QLabel("Hotkey ativa: preparando...")
        self.active_hotkey_badge.setProperty("role", "heroBadge")
        self.delay_badge = QLabel("")
        self.delay_badge.setProperty("role", "heroBadgeMuted")

        badges.addWidget(self.active_hotkey_badge, 0, Qt.AlignmentFlag.AlignLeft)
        badges.addWidget(self.delay_badge, 0, Qt.AlignmentFlag.AlignLeft)
        badges.addStretch(1)

        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addLayout(badges)
        return card

    def _build_info_card(self, title_text: str, body_text: str) -> QFrame:
        card = self._make_card()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(4)

        title = QLabel(title_text)
        title.setProperty("role", "cardTitle")
        body = QLabel(body_text)
        body.setWordWrap(True)
        body.setProperty("role", "cardBody")

        layout.addWidget(title)
        layout.addWidget(body)
        return card

    def _build_config_card(self) -> QFrame:
        card = self._make_card()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        title = QLabel("Hotkey e ritmo")
        title.setProperty("role", "cardTitle")
        body = QLabel(
            "Escolha a combinacao global, revise a selecao atual e aplique quando estiver satisfeito."
        )
        body.setProperty("role", "cardBody")
        body.setWordWrap(True)

        layout.addWidget(title)
        layout.addWidget(body)

        modifiers_label = QLabel("Modificadores")
        modifiers_label.setProperty("role", "fieldLabel")
        layout.addWidget(modifiers_label)

        modifiers_row = QHBoxLayout()
        modifiers_row.setSpacing(10)
        self.ctrl_checkbox = QCheckBox("Ctrl")
        self.ctrl_checkbox.setChecked(DEFAULT_HOTKEY.ctrl)
        self.alt_checkbox = QCheckBox("Alt")
        self.alt_checkbox.setChecked(DEFAULT_HOTKEY.alt)
        self.shift_checkbox = QCheckBox("Shift")
        self.shift_checkbox.setChecked(DEFAULT_HOTKEY.shift)
        self.win_checkbox = QCheckBox("Win")
        self.win_checkbox.setChecked(DEFAULT_HOTKEY.win)

        modifiers_row.addWidget(self.ctrl_checkbox)
        modifiers_row.addWidget(self.alt_checkbox)
        modifiers_row.addWidget(self.shift_checkbox)
        modifiers_row.addWidget(self.win_checkbox)
        modifiers_row.addStretch(1)
        layout.addLayout(modifiers_row)

        controls_row = QHBoxLayout()
        controls_row.setSpacing(12)

        key_group = QVBoxLayout()
        key_group.setSpacing(6)
        key_label = QLabel("Tecla")
        key_label.setProperty("role", "fieldLabel")
        self.key_combo = QComboBox()
        self.key_combo.addItems(KEY_CHOICES)
        self.key_combo.setCurrentText(DEFAULT_HOTKEY.key_name)
        self.key_combo.setMinimumHeight(38)
        key_group.addWidget(key_label)
        key_group.addWidget(self.key_combo)

        delay_group = QVBoxLayout()
        delay_group.setSpacing(6)
        delay_label = QLabel("Delay")
        delay_label.setProperty("role", "fieldLabel")
        self.delay_spinbox = QSpinBox()
        self.delay_spinbox.setRange(0, 2000)
        self.delay_spinbox.setSingleStep(5)
        self.delay_spinbox.setValue(DEFAULT_DELAY_MS)
        self.delay_spinbox.setSuffix(" ms")
        self.delay_spinbox.setMinimumHeight(38)
        delay_group.addWidget(delay_label)
        delay_group.addWidget(self.delay_spinbox)

        controls_row.addLayout(key_group, 1)
        controls_row.addLayout(delay_group, 1)
        layout.addLayout(controls_row)

        preview_panel = QFrame()
        preview_panel.setProperty("subtlePanel", True)
        preview_layout = QVBoxLayout(preview_panel)
        preview_layout.setContentsMargins(14, 12, 14, 12)
        preview_layout.setSpacing(4)

        self.hotkey_preview = QLabel("")
        self.hotkey_preview.setProperty("role", "fieldHint")
        self.hotkey_preview.setWordWrap(True)
        self.active_hotkey_hint = QLabel("")
        self.active_hotkey_hint.setProperty("role", "cardBody")
        self.active_hotkey_hint.setWordWrap(True)

        preview_layout.addWidget(self.hotkey_preview)
        preview_layout.addWidget(self.active_hotkey_hint)
        layout.addWidget(preview_panel)

        actions_row = QHBoxLayout()
        actions_row.setSpacing(10)
        self.restore_defaults_button = QPushButton("Restaurar padrao")
        self.restore_defaults_button.setObjectName("ghostButton")
        self.restore_defaults_button.setMinimumHeight(40)
        self.restore_defaults_button.clicked.connect(self.restore_defaults)

        self.apply_button = QPushButton("Aplicar hotkey")
        self.apply_button.setObjectName("accentButton")
        self.apply_button.setMinimumHeight(40)
        self.apply_button.clicked.connect(self.apply_hotkey)

        actions_row.addWidget(self.restore_defaults_button)
        actions_row.addStretch(1)
        actions_row.addWidget(self.apply_button)
        layout.addLayout(actions_row)
        return card

    def _build_text_card(self) -> QFrame:
        card = self._make_card()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        title = QLabel("Texto para digitar")
        title.setProperty("role", "cardTitle")
        body = QLabel(
            "Espacos, quebras de linha e paragrafos serao enviados do mesmo jeito. "
            "Use as metricas para conferir tamanho e duracao estimada."
        )
        body.setProperty("role", "cardBody")
        body.setWordWrap(True)

        metrics_grid = QGridLayout()
        metrics_grid.setHorizontalSpacing(8)
        metrics_grid.setVerticalSpacing(8)

        characters_card, self.characters_value = self._build_metric_tile("Caracteres")
        words_card, self.words_value = self._build_metric_tile("Palavras")
        lines_card, self.lines_value = self._build_metric_tile("Linhas")
        duration_card, self.duration_value = self._build_metric_tile("Duracao")

        metrics_grid.addWidget(characters_card, 0, 0)
        metrics_grid.addWidget(words_card, 0, 1)
        metrics_grid.addWidget(lines_card, 1, 0)
        metrics_grid.addWidget(duration_card, 1, 1)

        self.textbox = QPlainTextEdit()
        self.textbox.setPlaceholderText("Cole ou escreva o texto que deseja enviar...")

        footer = QHBoxLayout()
        footer.setSpacing(10)
        self.counter_label = QLabel("0 caracteres")
        self.counter_label.setProperty("role", "cardBody")

        self.sample_button = QPushButton("Usar exemplo")
        self.sample_button.setObjectName("ghostButton")
        self.sample_button.clicked.connect(self.insert_sample_text)

        footer.addWidget(self.counter_label)
        footer.addStretch(1)

        self.clear_button = QPushButton("Limpar")
        self.clear_button.setObjectName("ghostButton")
        self.clear_button.clicked.connect(self.clear_text)

        self.stop_button = QPushButton("Parar")
        self.stop_button.setObjectName("warnButton")
        self.stop_button.setEnabled(False)
        self.stop_button.clicked.connect(self.stop_typing)

        footer.addWidget(self.sample_button)
        footer.addWidget(self.clear_button)
        footer.addWidget(self.stop_button)

        layout.addWidget(title)
        layout.addWidget(body)
        layout.addLayout(metrics_grid)
        layout.addWidget(self.textbox, stretch=1)
        layout.addLayout(footer)
        return card

    def _build_status_card(self) -> QFrame:
        card = self._make_card()
        layout = QGridLayout(card)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setHorizontalSpacing(10)
        layout.setVerticalSpacing(6)

        self.status_dot = QLabel()
        self.status_dot.setFixedSize(12, 12)

        self.status_label = QLabel("Preparando hotkey...")
        self.status_label.setWordWrap(True)
        self.status_label.setProperty("role", "status")

        note = QLabel(
            "Dica: use pelo menos um modificador para evitar conflitos. Durante a digitacao, os campos ficam bloqueados para evitar mudancas acidentais."
        )
        note.setWordWrap(True)
        note.setProperty("role", "cardBody")

        layout.addWidget(self.status_dot, 0, 0, alignment=Qt.AlignmentFlag.AlignTop)
        layout.addWidget(self.status_label, 0, 1)
        layout.addWidget(note, 1, 1)
        layout.setColumnStretch(1, 1)
        return card

    def _build_metric_tile(self, title_text: str) -> tuple[QFrame, QLabel]:
        card = QFrame()
        card.setProperty("metric", True)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(2)

        value = QLabel("0")
        value.setProperty("role", "metricValue")
        caption = QLabel(title_text)
        caption.setProperty("role", "metricLabel")

        layout.addWidget(value)
        layout.addWidget(caption)
        return card, value

    def _make_card(self) -> QFrame:
        card = QFrame()
        card.setProperty("card", True)
        return card

    def _wire_ui_signals(self) -> None:
        for checkbox in (
            self.ctrl_checkbox,
            self.alt_checkbox,
            self.shift_checkbox,
            self.win_checkbox,
        ):
            checkbox.toggled.connect(self._refresh_hotkey_details)

        self.key_combo.currentTextChanged.connect(self._refresh_hotkey_details)
        self.delay_spinbox.valueChanged.connect(self._refresh_hotkey_details)
        self.delay_spinbox.valueChanged.connect(self._update_text_metrics)
        self.textbox.textChanged.connect(self._update_text_metrics)

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            f"""
            QWidget#window {{
                background: {self.BG};
                color: {self.TEXT};
                font-family: "Segoe UI";
            }}
            QFrame[hero="true"] {{
                background: qlineargradient(
                    x1: 0,
                    y1: 0,
                    x2: 1,
                    y2: 1,
                    stop: 0 {self.HERO_TOP},
                    stop: 1 {self.HERO_BOTTOM}
                );
                border: 1px solid #244A5C;
                border-radius: 24px;
            }}
            QFrame[card="true"] {{
                background: {self.PANEL};
                border: 1px solid #E7DDCF;
                border-radius: 18px;
            }}
            QFrame[subtlePanel="true"] {{
                background: {self.FIELD};
                border: 1px solid #E7DDCF;
                border-radius: 14px;
            }}
            QFrame[metric="true"] {{
                background: #FBF7F0;
                border: 1px solid #E5D9CB;
                border-radius: 14px;
            }}
            QLabel[role="heroTitle"] {{
                color: #FFFDF8;
                font-family: "Bahnschrift";
                font-size: 25px;
                font-weight: 700;
            }}
            QLabel[role="heroSubtitle"] {{
                color: #DCEAF0;
                font-size: 10pt;
            }}
            QLabel[role="heroBadge"],
            QLabel[role="heroBadgeMuted"] {{
                padding: 6px 12px;
                border-radius: 14px;
                font-size: 9pt;
                font-weight: 600;
                background: transparent;
                border: none;
            }}
            QLabel[role="heroBadge"] {{
                background: rgba(255, 255, 255, 0.14);
                color: white;
            }}
            QLabel[role="heroBadgeMuted"] {{
                background: rgba(255, 255, 255, 0.08);
                color: #D7E5EA;
            }}
            QLabel[role="cardTitle"] {{
                color: {self.TEXT};
                font-size: 11pt;
                font-weight: 600;
                background: transparent;
                border: none;
            }}
            QLabel[role="cardBody"] {{
                color: {self.MUTED};
                font-size: 9pt;
                background: transparent;
                border: none;
            }}
            QLabel[role="fieldLabel"] {{
                color: {self.TEXT};
                font-size: 9pt;
                font-weight: 600;
                background: transparent;
                border: none;
            }}
            QLabel[role="fieldHint"] {{
                color: {self.TEXT};
                font-size: 9pt;
                font-weight: 600;
                background: transparent;
                border: none;
            }}
            QLabel[role="metricValue"] {{
                color: {self.TEXT};
                font-family: "Bahnschrift";
                font-size: 17px;
                font-weight: 700;
                background: transparent;
                border: none;
            }}
            QLabel[role="metricLabel"] {{
                color: {self.MUTED};
                font-size: 8.5pt;
                background: transparent;
                border: none;
            }}
            QLabel[role="status"] {{
                color: {self.TEXT};
                font-size: 10pt;
                font-weight: 600;
                background: transparent;
                border: none;
            }}
            QPlainTextEdit,
            QComboBox,
            QSpinBox {{
                background: {self.FIELD};
                color: {self.TEXT};
                border: 1px solid {self.BORDER};
                border-radius: 12px;
                min-height: 38px;
                padding: 0 10px;
                selection-background-color: #BFE4E0;
            }}
            QPlainTextEdit {{
                min-height: 0;
                padding: 12px;
            }}
            QPlainTextEdit:focus,
            QComboBox:focus,
            QSpinBox:focus {{
                border: 1px solid {self.ACCENT};
            }}
            QComboBox::drop-down,
            QSpinBox::up-button,
            QSpinBox::down-button {{
                border: none;
                width: 22px;
                background: transparent;
            }}
            QPlainTextEdit:disabled,
            QComboBox:disabled,
            QSpinBox:disabled {{
                background: #ECE3D6;
                color: #7B8791;
            }}
            QCheckBox {{
                color: {self.TEXT};
                spacing: 8px;
            }}
            QCheckBox:disabled {{
                color: #8A939A;
            }}
            QCheckBox::indicator {{
                width: 16px;
                height: 16px;
                border: 1px solid {self.BORDER};
                border-radius: 5px;
                background: {self.FIELD};
            }}
            QCheckBox::indicator:checked {{
                background: {self.ACCENT};
                border: 1px solid {self.ACCENT};
            }}
            QPushButton {{
                border-radius: 12px;
                min-height: 40px;
                padding: 8px 14px;
                font-weight: 600;
                border: none;
            }}
            QPushButton#accentButton {{
                background: {self.ACCENT};
                color: white;
            }}
            QPushButton#accentButton:hover {{
                background: {self.ACCENT_ACTIVE};
            }}
            QPushButton#ghostButton {{
                background: {self.FIELD};
                color: {self.TEXT};
                border: 1px solid transparent;
            }}
            QPushButton#ghostButton:hover {{
                background: #ECE5D7;
            }}
            QPushButton#warnButton {{
                background: #F9E7DB;
                color: {self.WARN_ACTIVE};
                border: 1px solid #F0C8AA;
            }}
            QPushButton#warnButton:hover {{
                background: #F4D4BF;
            }}
            QPushButton:disabled {{
                background: #DED4C6;
                color: #7B8791;
                border: none;
            }}
            """
        )
        self._set_status_color(self.ACCENT)

    def _set_status_color(self, color: str) -> None:
        self.status_dot.setStyleSheet(
            f"background: {color}; border-radius: 6px; min-width: 12px; min-height: 12px;"
        )

    def _selected_hotkey(self) -> HotkeySpec:
        return HotkeySpec(
            ctrl=self.ctrl_checkbox.isChecked(),
            alt=self.alt_checkbox.isChecked(),
            shift=self.shift_checkbox.isChecked(),
            win=self.win_checkbox.isChecked(),
            key_name=self.key_combo.currentText(),
        )

    def _refresh_hotkey_details(self) -> None:
        selected = self._selected_hotkey()
        if selected.has_modifier():
            self.hotkey_preview.setText(f"Selecao atual: {selected.display_label()}")
        else:
            self.hotkey_preview.setText(
                "Selecao atual: adicione ao menos um modificador para registrar a hotkey."
            )

        active_text = self._current_hotkey_label or "nenhuma combinacao aplicada"
        self.active_hotkey_hint.setText(f"Hotkey ativa agora: {active_text}")
        badge_text = self._current_hotkey_label or "nenhuma"
        self.active_hotkey_badge.setText(f"Hotkey ativa: {badge_text}")
        self.delay_badge.setText(f"{self.get_delay_ms()} ms por tecla")
        self._refresh_action_states()

    def _refresh_action_states(self) -> None:
        is_typing = self.typewriter.is_running()
        has_modifier = self._selected_hotkey().has_modifier()
        has_text = bool(self.get_text())

        for widget in (
            self.ctrl_checkbox,
            self.alt_checkbox,
            self.shift_checkbox,
            self.win_checkbox,
            self.key_combo,
            self.delay_spinbox,
            self.textbox,
        ):
            widget.setEnabled(not is_typing)

        self.apply_button.setEnabled(not is_typing and has_modifier)
        self.restore_defaults_button.setEnabled(not is_typing)
        self.sample_button.setEnabled(not is_typing)
        self.clear_button.setEnabled(not is_typing and has_text)
        self.stop_button.setEnabled(is_typing)

    def _update_text_metrics(self) -> None:
        text = self.get_text()
        normalized = text.replace("\r\n", "\n").replace("\r", "\n")
        text_length = len(text)
        word_count = len(text.split())
        line_count = 0 if text == "" else normalized.count("\n") + 1

        label = "caractere" if text_length == 1 else "caracteres"
        self.counter_label.setText(f"{text_length} {label}")
        self.characters_value.setText(str(text_length))
        self.words_value.setText(str(word_count))
        self.lines_value.setText(str(line_count))
        self.duration_value.setText(self._format_duration(normalized))
        self._refresh_action_states()

    def _format_duration(self, text: str) -> str:
        if text == "":
            return "0 s"

        total_ms = self.typewriter.START_GRACE_MS + (
            max(len(text) - 1, 0) * self.get_delay_ms()
        )
        if total_ms < 1000:
            return f"{total_ms} ms"

        total_seconds = total_ms / 1000
        if total_seconds < 60:
            return f"{total_seconds:.1f} s"

        minutes = int(total_seconds // 60)
        seconds = int(round(total_seconds % 60))
        if seconds == 60:
            minutes += 1
            seconds = 0
        return f"{minutes}m {seconds:02d}s"

    def get_hotkey_spec(self) -> HotkeySpec:
        hotkey = self._selected_hotkey()
        if not hotkey.has_modifier():
            raise ValueError("Selecione ao menos um modificador para a hotkey.")
        return hotkey

    def get_delay_ms(self) -> int:
        return self.delay_spinbox.value()

    def get_text(self) -> str:
        return self.textbox.toPlainText()

    def insert_sample_text(self) -> None:
        self.textbox.setPlainText(SAMPLE_TEXT)
        self.set_status("Texto de exemplo inserido.", self.ACCENT)

    def clear_text(self) -> None:
        if self.get_text() == "":
            self.set_status("O campo de texto ja esta vazio.", self.ACCENT)
            return
        self.textbox.clear()
        self.set_status("Texto limpo.", self.ACCENT)

    def restore_defaults(self) -> None:
        self.ctrl_checkbox.setChecked(DEFAULT_HOTKEY.ctrl)
        self.alt_checkbox.setChecked(DEFAULT_HOTKEY.alt)
        self.shift_checkbox.setChecked(DEFAULT_HOTKEY.shift)
        self.win_checkbox.setChecked(DEFAULT_HOTKEY.win)
        self.key_combo.setCurrentText(DEFAULT_HOTKEY.key_name)
        self.delay_spinbox.setValue(DEFAULT_DELAY_MS)
        self.apply_hotkey()

    def stop_typing(self) -> None:
        if not self.typewriter.is_running():
            self.set_status("Nenhuma digitacao esta em andamento.", self.ACCENT)
            return
        self.typewriter.stop(announce=True)

    def apply_hotkey(self) -> None:
        previous_hotkey = self._active_hotkey
        previous_label = self._current_hotkey_label
        try:
            hotkey = self.get_hotkey_spec()
        except ValueError as exc:
            QMessageBox.critical(self, "Hotkey invalida", str(exc))
            self.set_status(str(exc), self.WARN)
            self._refresh_hotkey_details()
            return

        try:
            self.hotkeys.register(hotkey)
        except OSError:
            message = (
                "Nao foi possivel registrar essa hotkey. "
                "Ela pode ja estar em uso por outro aplicativo."
            )
            restored = False
            if previous_hotkey is not None:
                try:
                    self.hotkeys.register(previous_hotkey)
                    restored = True
                except OSError:
                    self._active_hotkey = None
                    self._current_hotkey_label = ""

            if restored and previous_label:
                message += f" A hotkey anterior ({previous_label}) foi mantida."
            elif previous_hotkey is None:
                self._active_hotkey = None
                self._current_hotkey_label = ""
                message += " Nenhuma hotkey ficou ativa."

            QMessageBox.critical(self, "Hotkey em uso", message)
            self.set_status(message, self.WARN)
            self._refresh_hotkey_details()
            return

        self._active_hotkey = hotkey
        self._current_hotkey_label = hotkey.display_label()
        self._refresh_hotkey_details()
        self.set_status(
            f"Pronto para digitar com {self._current_hotkey_label}.",
            self.ACCENT,
        )

    def _handle_hotkey_pressed(self) -> None:
        if self.typewriter.is_running():
            self.set_status("Interrompendo digitacao...", self.WARN)
            self.typewriter.stop(announce=True)
            return

        text = self.get_text()
        if text == "":
            self.set_status("Adicione algum texto antes de acionar a hotkey.", self.WARN)
            return

        if not self.typewriter.start(text, self.get_delay_ms()):
            self.set_status("A digitacao ja esta em andamento.", self.WARN)

    def _handle_typewriter_event(self, state: str, detail: str) -> None:
        self._refresh_action_states()
        if state == "started":
            self.set_status(
                "Digitando no app ativo. Pressione a mesma hotkey para interromper.",
                self.ACCENT,
            )
        elif state == "stopped":
            self.set_status("Digitacao interrompida.", self.WARN)
        elif state == "finished":
            label = self._current_hotkey_label or "a hotkey ativa"
            self.set_status(f"Texto enviado. Use {label} para digitar novamente.", self.ACCENT)
        elif state == "error":
            message = detail or "Falha ao enviar as teclas para o app ativo."
            self.set_status(message, self.WARN)

    def set_status(self, message: str, color: str) -> None:
        self.status_label.setText(message)
        self._set_status_color(color)

    def nativeEvent(self, eventType, message):  # noqa: N802
        if isinstance(eventType, str):
            event_name = eventType
        elif hasattr(eventType, "data"):
            event_name = bytes(eventType.data()).decode()
        else:
            event_name = bytes(eventType).decode()

        if event_name in {"windows_generic_MSG", "windows_dispatcher_MSG"}:
            message_ptr = int(message)
            if message_ptr:
                msg = ctypes.cast(message_ptr, ctypes.POINTER(wintypes.MSG)).contents
                if self.hotkeys.is_hotkey_message(msg):
                    self._handle_hotkey_pressed()
                    return True, 0
        return False, 0

    def closeEvent(self, event) -> None:  # noqa: N802
        self.typewriter.stop(announce=False)
        self.hotkeys.unregister()
        super().closeEvent(event)


def main() -> None:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = PythonWriterWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
