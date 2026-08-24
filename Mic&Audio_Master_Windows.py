#!/usr/bin/env python3
"""
Global Audio Device Control for Windows (PySide6 + pycaw + PowerShell)

Operates at the operating-system level, so it affects ALL applications
using the device at once - including Microsoft Teams, Zoom, Discord,
browsers, etc.

Two tabs, identical in behavior, one per device type:
- "Microphone" (input / capture devices)   -> default tab
- "Audio"      (output / render devices, i.e. speakers/headphones)

Features (per tab):
- "Mute" button
    -> sets a global mute on the selected (or all) device(s)
       (Windows Core Audio API / IAudioEndpointVolume.SetMute).
       Takes effect immediately, does NOT require administrator rights.

- "Disable" button
    -> fully disables the device in Device Manager
       (equivalent to right-click -> Disable on an audio device).
       REQUIRES administrator rights, because it uses the PowerShell
       commands Disable-PnpDevice / Enable-PnpDevice.

- "Sensitivity / volume" slider
    -> sets the global level of the device in Windows
       (IAudioEndpointVolume.SetMasterVolumeLevelScalar).

- "Set for all microphones" / "Set for all audio devices" checkbox
    -> when checked (default), the device dropdown is disabled and every
       action (mute, volume, disable) is applied to ALL devices of that
       type at once, instead of just the one selected in the dropdown.

Requirements (Windows only):
    pip install PySide6 pycaw comtypes

Running with the "Disable" feature:
    Launch the terminal / IDE as Administrator, then run:
    python mic_control_windows.py
"""

import sys
import ctypes
import subprocess

from PySide6.QtCore import Qt, QByteArray
from PySide6.QtGui import QIcon, QPixmap, QPainter, QColor
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QSlider, QComboBox, QFrame, QMessageBox,
    QGraphicsDropShadowEffect, QCheckBox, QTabWidget,
)

try:
    import comtypes
    from ctypes import POINTER, cast
    from pycaw.pycaw import (
        AudioUtilities,
        IAudioEndpointVolume,
        IMMDeviceEnumerator,
    )
    # Newer pycaw releases dropped this constant from the module, so we
    # define it directly here - it is a fixed Windows COM GUID that never
    # changes across OS or library versions.
    CLSID_MMDeviceEnumerator = comtypes.GUID("{BCDE0395-E52F-467C-8E3D-C4579291692E}")
    PYCAW_AVAILABLE = True
    _IMPORT_ERROR_DETAIL = ""
except ImportError as _e:
    PYCAW_AVAILABLE = False
    _IMPORT_ERROR_DETAIL = str(_e)

# --- Windows Core Audio API constants (independent of pycaw version) ---
eRender, eCapture, eAll = 0, 1, 2
eConsole, eMultimedia, eCommunications = 0, 1, 2
DEVICE_STATE_ACTIVE = 0x1


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin():
    """Relaunches this same script with administrator rights."""
    params = " ".join([f'"{a}"' for a in sys.argv])
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, f'"{sys.argv[0]}" {params}', None, 1
    )


# ---------------------------------------------------------------------
# Windows Core Audio access layer (mute / volume - no admin needed)
# Works for both capture (microphones) and render (speakers) devices,
# depending on the "flow" passed in (eCapture or eRender).
# ---------------------------------------------------------------------
class WindowsAudioDevices:
    def __init__(self, flow: int):
        self.flow = flow
        self._enumerator = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER,
        )

    def list_devices(self):
        """Returns a list of (friendly_name, IMMDevice) for active devices."""
        devices = []
        collection = self._enumerator.EnumAudioEndpoints(self.flow, DEVICE_STATE_ACTIVE)
        count = collection.GetCount()
        for i in range(count):
            raw_dev = collection.Item(i)
            dev = AudioUtilities.CreateDevice(raw_dev)
            devices.append((dev.FriendlyName, raw_dev))
        return devices

    def get_default_device(self):
        return self._enumerator.GetDefaultAudioEndpoint(self.flow, eMultimedia)

    @staticmethod
    def get_volume_interface(imm_device):
        interface = imm_device.Activate(IAudioEndpointVolume._iid_, comtypes.CLSCTX_ALL, None)
        return cast(interface, POINTER(IAudioEndpointVolume))

    def set_mute(self, imm_device, muted: bool):
        vol = self.get_volume_interface(imm_device)
        vol.SetMute(1 if muted else 0, None)

    def get_mute(self, imm_device) -> bool:
        vol = self.get_volume_interface(imm_device)
        return bool(vol.GetMute())

    def set_gain(self, imm_device, gain_0_to_1: float):
        vol = self.get_volume_interface(imm_device)
        vol.SetMasterVolumeLevelScalar(max(0.0, min(1.0, gain_0_to_1)), None)

    def get_gain(self, imm_device) -> float:
        vol = self.get_volume_interface(imm_device)
        return vol.GetMasterVolumeLevelScalar()


# ---------------------------------------------------------------------
# Device Manager enable/disable layer (admin required)
# Works for any AudioEndpoint device (input or output).
# ---------------------------------------------------------------------
class PnpDeviceControl:
    @staticmethod
    def _run_ps(command: str) -> str:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True, text=True
        )
        if result.returncode != 0 and result.stderr.strip():
            raise RuntimeError(result.stderr.strip())
        return result.stdout.strip()

    @classmethod
    def find_instance_id(cls, friendly_name: str) -> str | None:
        # Looks for the matching entry in the "AudioEndpoint" class
        escaped = friendly_name.replace("'", "''")
        cmd = (
            f"Get-PnpDevice -Class AudioEndpoint | "
            f"Where-Object {{ $_.FriendlyName -like '*{escaped}*' }} | "
            f"Select-Object -First 1 -ExpandProperty InstanceId"
        )
        out = cls._run_ps(cmd)
        return out or None

    @classmethod
    def disable(cls, instance_id: str):
        cls._run_ps(f"Disable-PnpDevice -InstanceId '{instance_id}' -Confirm:$false")

    @classmethod
    def enable(cls, instance_id: str):
        cls._run_ps(f"Enable-PnpDevice -InstanceId '{instance_id}' -Confirm:$false")


# ---------------------------------------------------------------------
# SVG icons (inline, colored via a {color} placeholder replaced at render time)
# ---------------------------------------------------------------------
_MIC_ON_SVG = """
<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
  <path d="M12 15a3.5 3.5 0 0 0 3.5-3.5V6a3.5 3.5 0 0 0-7 0v5.5A3.5 3.5 0 0 0 12 15Z"
        fill="{color}"/>
  <path d="M19 11.5a1 1 0 0 0-2 0 5 5 0 0 1-10 0 1 1 0 0 0-2 0 7 7 0 0 0 6 6.93V21h-2a1 1 0 0 0 0 2h6a1 1 0 0 0 0-2h-2v-2.57a7 7 0 0 0 6-6.93Z"
        fill="{color}"/>
</svg>
""".strip()

_MIC_OFF_SVG = """
<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
  <path d="M12 15a3.5 3.5 0 0 0 3.5-3.5V6a3.5 3.5 0 0 0-6.53-1.76l3.03 3.03V6a1 1 0 0 1 2 0v3.5c0 .3-.04.58-.11.85l1.47 1.47c.27-.55.44-1.16.44-1.82Z"
        fill="{color}"/>
  <path d="M19 11.5a1 1 0 0 0-2 0 4.98 4.98 0 0 1-1.28 3.34l1.42 1.42A6.97 6.97 0 0 0 19 11.5Z"
        fill="{color}"/>
  <path d="M4.7 3.29a1 1 0 0 0-1.41 1.42l4.02 4.02V11.5A4.5 4.5 0 0 0 12 16a4.46 4.46 0 0 0 1.6-.3l1.6 1.6A6.96 6.96 0 0 1 7 11.5a1 1 0 0 0-2 0 8.97 8.97 0 0 0 6 8.46V21H9a1 1 0 0 0 0 2h6a1 1 0 0 0 0-2h-2v-2.04a8.94 8.94 0 0 0 3.05-1.23l2.24 2.25a1 1 0 0 0 1.42-1.42Z"
        fill="{color}"/>
</svg>
""".strip()

_SPEAKER_ON_SVG = """
<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
  <path d="M11 4.7 6.9 8H4a1 1 0 0 0-1 1v6a1 1 0 0 0 1 1h2.9L11 19.3a1 1 0 0 0 1.7-.72V5.42A1 1 0 0 0 11 4.7Z" fill="{color}"/>
  <path d="M15.5 8.5a1 1 0 0 1 1.41 0 5 5 0 0 1 0 7 1 1 0 1 1-1.41-1.41 3 3 0 0 0 0-4.18 1 1 0 0 1 0-1.41Z" fill="{color}"/>
  <path d="M18.1 5.9a1 1 0 0 1 1.41 0 9.98 9.98 0 0 1 0 14.14 1 1 0 1 1-1.41-1.41 7.98 7.98 0 0 0 0-11.32 1 1 0 0 1 0-1.41Z" fill="{color}"/>
</svg>
""".strip()

_SPEAKER_OFF_SVG = """
<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
  <path d="M11 4.7 6.9 8H4a1 1 0 0 0-1 1v6a1 1 0 0 0 1 1h2.9L11 19.3a1 1 0 0 0 1.7-.72V5.42A1 1 0 0 0 11 4.7Z" fill="{color}"/>
  <path d="M15.9 9.4a1 1 0 0 1 1.41 1.41L16.41 12l.9 1.19a1 1 0 0 1-1.41 1.41L15 13.6l-1.19.9a1 1 0 0 1-1.41-1.41l.9-1.19-.9-1.19A1 1 0 0 1 13.8 8.1l1.2.9Z" fill="{color}"/>
</svg>
""".strip()

_POWER_SVG = """
<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
  <path d="M12 2a1 1 0 0 1 1 1v9a1 1 0 0 1-2 0V3a1 1 0 0 1 1-1Z" fill="{color}"/>
  <path d="M6.35 5.64A1 1 0 0 0 4.9 7.03 8 8 0 1 0 19.1 7a1 1 0 0 0-1.45-1.38 6 6 0 1 1-11.3.02Z" fill="{color}"/>
</svg>
""".strip()


def svg_icon(svg_template: str, color: str, size: int = 28) -> QIcon:
    """Renders an inline SVG string (with a {color} placeholder) into a QIcon."""
    svg_data = svg_template.format(color=color)
    renderer = QSvgRenderer(QByteArray(svg_data.encode("utf-8")))
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    renderer.render(painter)
    painter.end()
    return QIcon(pixmap)


# ---------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------
BG = "#0f1826"
CARD = "#152238"
ACCENT = "#2dd4bf"
ACCENT_DARK = "#12b3a3"
DANGER = "#ef4444"
TEXT = "#e6edf5"
SUBTEXT = "#8ba0b8"

STYLESHEET = f"""
QWidget#root {{
    background-color: {BG};
}}
QLabel {{
    color: {TEXT};
}}
QLabel#title {{
    font-size: 19px;
    font-weight: 700;
    letter-spacing: 0.3px;
}}
QLabel#subtext {{
    color: {SUBTEXT};
    font-size: 11px;
}}
QLabel#warning {{
    color: #f59e0b;
    font-size: 12px;
}}
QFrame#card {{
    background-color: {CARD};
    border-radius: 16px;
}}
QComboBox {{
    background-color: #1c2c47;
    color: {TEXT};
    border: 1px solid #253854;
    border-radius: 8px;
    padding: 6px 10px;
}}
QComboBox QAbstractItemView {{
    background-color: #1c2c47;
    color: {TEXT};
    selection-background-color: {ACCENT_DARK};
}}
QComboBox:disabled {{
    color: {SUBTEXT};
    background-color: #14203a;
}}
QSlider::groove:horizontal {{
    height: 6px;
    background: #223351;
    border-radius: 3px;
}}
QSlider::sub-page:horizontal {{
    background: {ACCENT};
    border-radius: 3px;
}}
QSlider::handle:horizontal {{
    background: {TEXT};
    width: 16px;
    height: 16px;
    margin: -6px 0;
    border-radius: 8px;
}}
QPushButton#relaunch {{
    background-color: #1c2c47;
    color: {TEXT};
    border: 1px solid #253854;
    border-radius: 8px;
    padding: 8px;
}}
QPushButton#relaunch:hover {{
    background-color: #253854;
}}
QPushButton#disableBtn {{
    background-color: #1c2c47;
    color: {TEXT};
    border: 1px solid #253854;
    border-radius: 10px;
    padding: 10px;
    font-weight: 600;
}}
QPushButton#disableBtn:hover {{
    background-color: #253854;
}}
QPushButton#disableBtn:checked {{
    background-color: {DANGER};
    border: 1px solid {DANGER};
}}
QCheckBox {{
    color: {SUBTEXT};
    font-size: 12px;
    spacing: 8px;
}}
QCheckBox::indicator {{
    width: 16px;
    height: 16px;
    border-radius: 4px;
    border: 1px solid #35496b;
    background-color: #1c2c47;
}}
QCheckBox::indicator:checked {{
    background-color: {ACCENT};
    border: 1px solid {ACCENT};
}}
QTabWidget::pane {{
    border: none;
    background-color: {BG};
}}
QTabBar::tab {{
    background-color: transparent;
    color: {SUBTEXT};
    padding: 10px 18px;
    margin-right: 4px;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    font-weight: 600;
    font-size: 13px;
}}
QTabBar::tab:selected {{
    color: {TEXT};
    background-color: {CARD};
}}
QTabBar::tab:hover:!selected {{
    color: {TEXT};
}}
"""


# ---------------------------------------------------------------------
# Reusable control panel - used once for microphones (capture) and once
# for audio output devices (render). Identical look and behavior.
# ---------------------------------------------------------------------
class AudioControlPanel(QWidget):
    def __init__(self, flow: int, device_word: str, all_label: str,
                 icon_on: str, icon_off: str, parent=None):
        """
        flow: eCapture or eRender
        device_word: "microphone" / "audio device" - used in dialog text
        all_label: text for the "apply to all" checkbox
        icon_on / icon_off: SVG templates for the big toggle button
        """
        super().__init__(parent)
        self._device_word = device_word
        self._all_label = all_label
        self._icon_on = icon_on
        self._icon_off = icon_off

        if not PYCAW_AVAILABLE:
            layout = QVBoxLayout(self)
            detail = f"\n\nError details:\n{_IMPORT_ERROR_DETAIL}" if _IMPORT_ERROR_DETAIL else ""
            label = QLabel(
                "Could not import pycaw / comtypes.\n\n"
                "If the libraries are installed, this may be caused by an "
                "incompatible pycaw version.\n\n"
                "Try:\npip install --upgrade pycaw comtypes"
                + detail
            )
            label.setWordWrap(True)
            layout.addWidget(label)
            return

        self.audio = WindowsAudioDevices(flow)
        self.current_device = None       # IMMDevice currently controlled
        self.current_name = None
        self.pnp_instance_id = None      # used for Disable/Enable-PnpDevice
        self.admin_mode = is_admin()
        self.all_devices = []            # list of (name, imm_dev) for "all" mode
        self.pnp_instance_ids = {}       # name -> instance_id cache for "all" mode

        self._build_ui()
        self._populate_devices()

    # ---------- UI ----------
    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(20, 20, 20, 20)
        outer.setSpacing(16)

        if not self.admin_mode:
            admin_warning = QLabel(
                "⚠ The app is NOT running as Administrator.\n"
                "Mute and volume will work normally, but the "
                "\"Disable\" button requires admin rights."
            )
            admin_warning.setObjectName("warning")
            admin_warning.setWordWrap(True)
            outer.addWidget(admin_warning)

            relaunch_btn = QPushButton("Restart as Administrator")
            relaunch_btn.setObjectName("relaunch")
            relaunch_btn.setCursor(Qt.PointingHandCursor)
            relaunch_btn.clicked.connect(self._relaunch_admin)
            outer.addWidget(relaunch_btn)

        # ---- Card ----
        card = QFrame()
        card.setObjectName("card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(22, 22, 22, 22)
        card_layout.setSpacing(16)

        device_label = QLabel("Device")
        device_label.setObjectName("subtext")
        card_layout.addWidget(device_label)

        self.device_combo = QComboBox()
        self.device_combo.currentIndexChanged.connect(self._on_device_changed)
        card_layout.addWidget(self.device_combo)

        self.all_checkbox = QCheckBox(self._all_label)
        self.all_checkbox.setCursor(Qt.PointingHandCursor)
        self.all_checkbox.setChecked(True)
        self.all_checkbox.toggled.connect(self._on_all_toggled)
        card_layout.addWidget(self.all_checkbox)

        # ---- Big circular mute button ----
        btn_wrap = QHBoxLayout()
        btn_wrap.addStretch()

        self.mute_btn = QPushButton()
        self.mute_btn.setCheckable(True)
        self.mute_btn.setFixedSize(96, 96)
        self.mute_btn.setCursor(Qt.PointingHandCursor)
        self.mute_btn.setIconSize(self.mute_btn.size() * 0.4)
        self.mute_btn.clicked.connect(self._on_mute_clicked)
        shadow = QGraphicsDropShadowEffect(blurRadius=30, xOffset=0, yOffset=0)
        shadow.setColor(QColor(ACCENT))
        self.mute_btn.setGraphicsEffect(shadow)
        btn_wrap.addWidget(self.mute_btn)
        btn_wrap.addStretch()
        card_layout.addLayout(btn_wrap)

        self.status_label = QLabel("Status: —")
        self.status_label.setAlignment(Qt.AlignCenter)
        card_layout.addWidget(self.status_label)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("background-color: #223351; max-height: 1px; border: none;")
        card_layout.addWidget(line)

        gain_col = QVBoxLayout()
        gain_col.setSpacing(8)
        self.gain_label = QLabel("Sensitivity / volume: 100%")
        self.gain_label.setObjectName("subtext")
        gain_col.addWidget(self.gain_label)

        self.gain_slider = QSlider(Qt.Horizontal)
        self.gain_slider.setRange(0, 100)
        self.gain_slider.setValue(100)
        self.gain_slider.valueChanged.connect(self._on_gain_changed)
        gain_col.addWidget(self.gain_slider)
        card_layout.addLayout(gain_col)

        self.disable_btn = QPushButton()
        self.disable_btn.setObjectName("disableBtn")
        self.disable_btn.setCheckable(True)
        self.disable_btn.setCursor(Qt.PointingHandCursor)
        self.disable_btn.setIcon(svg_icon(_POWER_SVG, TEXT, 18))
        self.disable_btn.setText("  Disable")
        self.disable_btn.clicked.connect(self._on_disable_clicked)
        if not self.admin_mode:
            self.disable_btn.setToolTip("Requires running as Administrator")
        card_layout.addWidget(self.disable_btn)

        outer.addWidget(card)

        note = QLabel(
            f"Mute and volume apply instantly to ALL applications "
            f"(Teams, Zoom, browsers, etc.), because they change "
            f"Windows system settings rather than this app alone."
        )
        note.setObjectName("subtext")
        note.setWordWrap(True)
        outer.addWidget(note)
        outer.addStretch()

        self._update_mute_button_style(muted=False)

    def _populate_devices(self):
        self.device_combo.blockSignals(True)
        self.device_combo.clear()
        try:
            devices = self.audio.list_devices()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not retrieve the device list:\n{e}")
            devices = []

        self.all_devices = devices
        for name, imm_dev in devices:
            self.device_combo.addItem(name, userData=imm_dev)
        self.device_combo.blockSignals(False)

        if devices:
            self.device_combo.setCurrentIndex(0)
            self._on_device_changed(0)

        self.device_combo.setEnabled(not self.all_checkbox.isChecked())

    # ---------- Slots ----------
    def _relaunch_admin(self):
        relaunch_as_admin()
        sys.exit(0)

    def _on_all_toggled(self, checked):
        self.device_combo.setEnabled(not checked)

    def _on_device_changed(self, index):
        imm_dev = self.device_combo.itemData(index)
        if imm_dev is None:
            return
        self.current_device = imm_dev
        self.current_name = self.device_combo.currentText()
        self.pnp_instance_id = None  # resolved lazily, only when needed

        try:
            muted = self.audio.get_mute(imm_dev)
            gain = self.audio.get_gain(imm_dev)
        except Exception as e:
            QMessageBox.warning(self, "Error", str(e))
            return

        self.mute_btn.blockSignals(True)
        self.mute_btn.setChecked(muted)
        self.mute_btn.blockSignals(False)
        self._update_mute_button_style(muted)

        self.gain_slider.blockSignals(True)
        self.gain_slider.setValue(int(gain * 100))
        self.gain_label.setText(f"Sensitivity / volume: {int(gain * 100)}%")
        self.gain_slider.blockSignals(False)

        self.disable_btn.blockSignals(True)
        self.disable_btn.setChecked(False)
        self.disable_btn.blockSignals(False)

        self._refresh_status()

    def _on_mute_clicked(self, checked):
        self._update_mute_button_style(checked)

        if self.all_checkbox.isChecked():
            errors = []
            for name, imm_dev in self.all_devices:
                try:
                    self.audio.set_mute(imm_dev, checked)
                except Exception as e:
                    errors.append(f"{name}: {e}")
            if errors:
                QMessageBox.warning(self, "Error", "\n".join(errors))
        elif self.current_device is not None:
            try:
                self.audio.set_mute(self.current_device, checked)
            except Exception as e:
                QMessageBox.warning(self, "Error", str(e))

        self._refresh_status()

    def _update_mute_button_style(self, muted: bool):
        """Swaps the circular button's color and icon between on / off."""
        if muted:
            bg, bg_hover = DANGER, "#dc2626"
            icon = svg_icon(self._icon_off, "#ffffff", 34)
        else:
            bg, bg_hover = ACCENT, ACCENT_DARK
            icon = svg_icon(self._icon_on, "#08201c", 34)
        self.mute_btn.setIcon(icon)
        self.mute_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {bg};
                border: none;
                border-radius: 48px;
            }}
            QPushButton:hover {{
                background-color: {bg_hover};
            }}
        """)
        effect = self.mute_btn.graphicsEffect()
        if effect is not None:
            effect.setColor(QColor(bg))

    def _on_gain_changed(self, value):
        self.gain_label.setText(f"Sensitivity / volume: {value}%")

        if self.all_checkbox.isChecked():
            for name, imm_dev in self.all_devices:
                try:
                    self.audio.set_gain(imm_dev, value / 100.0)
                except Exception:
                    pass  # avoid spamming warnings while dragging the slider
        elif self.current_device is not None:
            try:
                self.audio.set_gain(self.current_device, value / 100.0)
            except Exception as e:
                QMessageBox.warning(self, "Error", str(e))

    def _on_disable_clicked(self, checked):
        if not self.admin_mode:
            QMessageBox.information(
                self, "Administrator rights required",
                "This feature requires running the app as Administrator.\n"
                "Use the \"Restart as Administrator\" button at the top of the window."
            )
            self.disable_btn.setChecked(False)
            return

        if self.all_checkbox.isChecked():
            errors = []
            for name, _imm_dev in self.all_devices:
                try:
                    instance_id = self.pnp_instance_ids.get(name)
                    if instance_id is None:
                        instance_id = PnpDeviceControl.find_instance_id(name)
                        if instance_id is None:
                            raise RuntimeError(
                                "Could not find a matching device in Device "
                                "Manager (AudioEndpoint class)."
                            )
                        self.pnp_instance_ids[name] = instance_id

                    if checked:
                        PnpDeviceControl.disable(instance_id)
                    else:
                        PnpDeviceControl.enable(instance_id)
                except Exception as e:
                    errors.append(f"{name}: {e}")
            if errors:
                QMessageBox.critical(self, "Error", "\n".join(errors))
                self.disable_btn.setChecked(not checked)
        else:
            try:
                if self.pnp_instance_id is None:
                    self.pnp_instance_id = PnpDeviceControl.find_instance_id(self.current_name)
                    if self.pnp_instance_id is None:
                        raise RuntimeError(
                            "Could not find a matching device in Device Manager "
                            "(AudioEndpoint class)."
                        )

                if checked:
                    PnpDeviceControl.disable(self.pnp_instance_id)
                else:
                    PnpDeviceControl.enable(self.pnp_instance_id)
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))
                self.disable_btn.setChecked(not checked)

        self._refresh_status()

    def _refresh_status(self):
        if self.disable_btn.isChecked():
            self.disable_btn.setText("  Enable")
            self.status_label.setText(f"{self._device_word.capitalize()} COMPLETELY DISABLED (system)")
            self.status_label.setStyleSheet(f"color: {DANGER}; font-size: 13px; font-weight: 600;")
        elif self.mute_btn.isChecked():
            self.disable_btn.setText("  Disable")
            self.status_label.setText(f"{self._device_word.capitalize()} active, but MUTED globally")
            self.status_label.setStyleSheet("color: #f59e0b; font-size: 13px; font-weight: 600;")
        else:
            self.disable_btn.setText("  Disable")
            self.status_label.setText(f"{self._device_word.capitalize()} active")
            self.status_label.setStyleSheet(f"color: {ACCENT}; font-size: 13px; font-weight: 600;")


# ---------------------------------------------------------------------
# Main window - hosts the two tabs
# ---------------------------------------------------------------------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Global Audio Device Control (Windows)")
        self.setFixedWidth(420)
        self.setObjectName("root")
        self.setStyleSheet(STYLESHEET)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        title = QLabel("Global Audio Device Control")
        title.setObjectName("title")
        title.setContentsMargins(20, 20, 20, 4)
        outer.addWidget(title)

        tabs = QTabWidget()
        tabs.setDocumentMode(True)

        mic_panel = AudioControlPanel(
            flow=eCapture,
            device_word="microphone",
            all_label="Set for all microphones",
            icon_on=_MIC_ON_SVG,
            icon_off=_MIC_OFF_SVG,
        )
        audio_panel = AudioControlPanel(
            flow=eRender,
            device_word="audio device",
            all_label="Set for all audio devices",
            icon_on=_SPEAKER_ON_SVG,
            icon_off=_SPEAKER_OFF_SVG,
        )

        tabs.addTab(mic_panel, "Microphone")
        tabs.addTab(audio_panel, "Audio")
        tabs.setCurrentIndex(0)  # Microphone is the default tab

        outer.addWidget(tabs)


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
