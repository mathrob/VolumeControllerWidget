# Copyright 2024 Mathieu R.
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files 
# (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, 
# publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, 
# subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE 
# WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT 
# HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, 
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


import sys
import json
from pathlib import Path
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QSlider, QLabel, QMenu, QSystemTrayIcon, QPushButton,
                             QColorDialog, QSpinBox)
from PyQt6.QtCore import Qt, QPoint, QTimer
from PyQt6.QtGui import QAction, QIcon, QColor, QPixmap, QPainter
import winreg
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


class AudioWidget(QWidget):
    """
    A customizable audio control widget that allows adjusting master volume, microphone volume, and muting.

    The widget also provides options to customize the appearance, set the widget to autostart at system boot, and access a system tray icon.
    """
    def __init__(self):
        super().__init__()
        self.settings = self.load_settings()  # Load user settings from a JSON file
        self.is_mic_muted = False
        self.init_ui()  # Initialize the user interface
        self.setup_tray()  # Set up the system tray icon
        self.setup_auto_start()  # Set up the autostart functionality

        # Restore the saved window position
        saved_position = self.settings.get('position')
        if saved_position:
            self.move(QPoint(*saved_position))

        # Start a timer to update the volume levels
        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self.update_volumes)
        self.update_timer.start(200)

        # Initialize the mute states
        self.is_master_muted = False
        self.is_mic_muted = False

    def init_ui(self):
        """
        Initialize the user interface of the audio control widget.

        This method sets up the layout, widgets, and event handlers for the widget.
        """
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        layout = QVBoxLayout()
        layout.setSpacing(20)

        self.background_style = QPushButton(self)
        self.background_style.setGeometry(self.rect())
        self.update_background_style()
        self.background_style.setFlat(True)
        self.background_style.mousePressEvent = self.mousePressEvent
        self.background_style.mouseMoveEvent = self.mouseMoveEvent
        
        self.mic_label = QLabel()
        self.update_mic_icon()
        self.mic_label.mousePressEvent = self.toggle_mute_mic
        
        self.master_volume = QSlider(Qt.Orientation.Vertical)
        self.master_volume.setRange(0, 100)
        self.master_volume.valueChanged.connect(self.change_master_volume)
        self.master_volume.setFixedHeight(200)
        
        self.mic_volume = QSlider(Qt.Orientation.Vertical)
        self.mic_volume.setRange(0, 100)
        self.mic_volume.valueChanged.connect(self.change_mic_volume)
        self.mic_volume.setFixedHeight(200)
        
        self.speaker_label = QLabel("🔊")
        self.speaker_label.setFixedSize(96, 96)
        self.speaker_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.speaker_label.mousePressEvent = self.toggle_mute_master

        self.mic_label = QLabel("🎤")
        self.mic_label.setFixedSize(96, 96)
        self.mic_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.mic_label.mousePressEvent = self.toggle_mute_mic
        
        options_button = QPushButton("⋮")
        options_button.setFixedSize(48, 96)
        options_button.setStyleSheet("color: white; border: none; font-size: 18px;")
        options_button.clicked.connect(self.show_options_menu)

        self.apply_style()
        
        controls_layout = QHBoxLayout()
        controls_layout.setSpacing(30)
        controls_layout.addWidget(self.speaker_label)
        controls_layout.addWidget(self.master_volume)
        controls_layout.addWidget(self.mic_label)
        controls_layout.addWidget(self.mic_volume)
        controls_layout.addWidget(options_button)
        
        layout.addLayout(controls_layout)
        self.setLayout(layout)
        
        self.init_volumes()

    def apply_style(self):
        """
        Apply the custom styles to the audio control widget based on user settings.

        This method updates the background color, button color, and opacity of the widget.
        """
        buttons_color = self.settings.get('buttons_color', '#2E2E2E')
        buttons_opacity = self.settings.get('buttons_opacity', 0.9)
        
        style = f"""
        QWidget {{
            background-color: rgba{QColor(buttons_color).getRgb()[:-1] + (int(buttons_opacity * 255),)};
            border-radius: 10px;
            padding: 20px;
        }}
        QSlider::groove:vertical {{
            background: #4A4A4A;
            width: 40px;
            border-radius: 10px;
        }}
        QSlider::handle:vertical {{
            background: #007AFF;
            height: 40px;
            width: 40px;
            margin: 0 -8px;
            border-radius: 20px;
        }}
        QLabel {{
            color: white;
            font-size: 48px;
            text-align: center;
            padding: 10px;
        }}
        QPushButton {{
            color: white;
            font-size: 48px;
            border: none;
            text-align: center;
        }}
        """
        self.setStyleSheet(style)

    def update_background_style(self):
        """
        Update the background style of the audio control widget based on user settings.

        This method updates the background color and opacity.
        """
        background_color = self.settings.get('background_color', '#0000FF')
        background_opacity = self.settings.get('background_opacity', 0.5)

        style = f"background: rgba({QColor(background_color).red()}, {QColor(background_color).green()}, {QColor(background_color).blue()}, {int(background_opacity * 255)});"
        self.background_style.setStyleSheet(style)

    def toggle_mute_master(self, event):
        """
        Toggle the mute state of the master (speaker) volume.

        This method interacts with the system audio APIs to mute or unmute the master volume.
        """
        self.is_master_muted = not self.is_master_muted

        try:
            sessions = AudioUtilities.GetSpeakers()
            interface = sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))

            volume.SetMute(self.is_master_muted, None)
        except Exception as e:
            print(f"Erreur lors de l'initialisation du volume maître : {e}")

        self.speaker_label.setText("🔇" if self.is_master_muted else "🔊")

    def update_mic_icon(self):
        """
        Update the microphone icon to indicate the mute state.

        This method creates a custom icon with a red strike-through when the microphone is muted.
        """
        mic_icon = QPixmap(64, 64)
        mic_icon.fill(Qt.GlobalColor.transparent)

        painter = QPainter(mic_icon)
        
        painter.setPen(Qt.GlobalColor.black)
        painter.drawText(mic_icon.rect(), Qt.AlignmentFlag.AlignCenter, "🎤")
        
        if self.is_mic_muted:
            painter.setPen(QColor("red"))
            painter.setBrush(QColor("red"))
            painter.drawLine(5, 5, mic_icon.width() - 5, mic_icon.height() - 5) 
            painter.drawLine(mic_icon.width() - 5, 5, 5, mic_icon.height() - 5) 

        painter.end()
        
        self.mic_label.setPixmap(mic_icon)

    def toggle_mute_mic(self, event):
        """
        Toggle the mute state of the microphone.

        This method interacts with the system audio APIs to mute or unmute the microphone.
        """
        self.is_mic_muted = not self.is_mic_muted
        
        try: 
            sessions = AudioUtilities.GetMicrophone() 
            interface = sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))

            volume.SetMute(self.is_mic_muted, None)
        except Exception as e:
            print(f"Error while managing the microphone: {e}")

        self.update_mic_icon()

    def update_volumes(self):
        """
        Periodically update the master and microphone volume levels.

        This method retrieves the current volume levels from the system audio APIs and updates the UI accordingly.
        """
        try:
            sessions = AudioUtilities.GetSpeakers()
            interface = sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            current_volume = int(volume.GetMasterVolumeLevelScalar() * 100)
            
            if self.master_volume.value() != current_volume:
                self.master_volume.blockSignals(True)
                self.master_volume.setValue(current_volume)
                self.master_volume.blockSignals(False)

            self.is_master_muted = volume.GetMute()
            self.speaker_label.setText("🔇" if self.is_master_muted else "🔊")
        except Exception as e:
            print(f"Error while managing the speaker volume: {e}")

        try:
            mic_sessions = AudioUtilities.GetMicrophone()
            mic_interface = mic_sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            mic_volume = cast(mic_interface, POINTER(IAudioEndpointVolume))
            mic_current_volume = int(mic_volume.GetMasterVolumeLevelScalar() * 100)

            if self.mic_volume.value() != mic_current_volume:
                self.mic_volume.blockSignals(True)
                self.mic_volume.setValue(mic_current_volume)
                self.mic_volume.blockSignals(False)

            self.is_mic_muted = mic_volume.GetMute()
        except Exception as e:
            print(f"Error while managing the microphone volume: {e}")

        self.update_mic_icon()

    def closeEvent(self, event):
        """
        Save the current window position when the widget is closed.

        This method is called when the user closes the audio control widget.
        """
        self.settings['position'] = (self.x(), self.y())
        self.save_settings()
        super().closeEvent(event)

    def show_options_menu(self):
        """
        Show the options menu for the audio control widget.

        The options menu allows the user to customize the widget, toggle autostart, and quit the application.
        """
        menu = QMenu(self)
        
        customize_action = QAction("Customize", self)
        customize_action.triggered.connect(self.show_customize_dialog)
        menu.addAction(customize_action)
        
        autostart_action = QAction("Autostart", self)
        autostart_action.setCheckable(True)
        autostart_action.setChecked(self.settings.get('autostart', False))
        autostart_action.triggered.connect(self.toggle_autostart)
        menu.addAction(autostart_action)
        
        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(QApplication.quit)
        menu.addAction(quit_action)
        
        menu.exec(self.mapToGlobal(QPoint(self.width() - 10, 30)))

    def setup_tray(self):
        """
        Set up the system tray icon for the audio control widget.

        The system tray icon allows the user to access the options menu and quit the application.
        """
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon("audio_icon.ico"))
        
        tray_menu = QMenu()
        
        customize_action = QAction("Customize", self)
        customize_action.triggered.connect(self.show_customize_dialog)
        
        self.autostart_action = QAction("Autostart", self)
        self.autostart_action.setCheckable(True)
        self.autostart_action.setChecked(self.settings.get('autostart', False))
        self.autostart_action.triggered.connect(self.toggle_autostart)
        
        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(QApplication.quit)
        
        tray_menu.addAction(customize_action)
        tray_menu.addAction(self.autostart_action)
        tray_menu.addSeparator()
        tray_menu.addAction(quit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

    def show_customize_dialog(self):
        """
        Show the customization dialog for the audio control widget.

        The customization dialog allows the user to change the background color, button color, and opacity.
        """
        dialog = QWidget()
        dialog.setWindowTitle("Customization")
        layout = QVBoxLayout()

        background_color = QPushButton("Change background color")
        background_color.clicked.connect(lambda: self.change_background_color(dialog))

        background_opacity_layout = QHBoxLayout()
        background_opacity_label = QLabel("Background opacity:")
        background_opacity_spin = QSpinBox()
        background_opacity_spin.setRange(0, 100)
        background_opacity_spin.setValue(int(self.settings.get('background_opacity', 0.5) * 100))
        background_opacity_spin.valueChanged.connect(lambda v: self.change_background_opacity(v / 100))

        background_opacity_layout.addWidget(background_opacity_label)
        background_opacity_layout.addWidget(background_opacity_spin)
        
        color_buttons = QPushButton("Change button color")
        color_buttons.clicked.connect(lambda: self.change_buttons_color(dialog))
        
        buttons_opacity_layout = QHBoxLayout()
        buttons_opacity_label = QLabel("Button opacity:")
        buttons_opacity_spin = QSpinBox()
        buttons_opacity_spin.setRange(20, 100)
        buttons_opacity_spin.setValue(int(self.settings.get('buttons_opacity', 0.9) * 100))
        buttons_opacity_spin.valueChanged.connect(lambda v: self.change_buttons_opacity(v / 100))
        
        buttons_opacity_layout.addWidget(buttons_opacity_label)
        buttons_opacity_layout.addWidget(buttons_opacity_spin)
        
        layout.addWidget(color_buttons)
        layout.addWidget(background_color)
        layout.addLayout(buttons_opacity_layout)
        layout.addLayout(background_opacity_layout)
        dialog.setLayout(layout)
        dialog.show()

    def change_background_color(self, dialog):
        """
        Change the background color of the audio control widget.

        This method opens a color dialog and updates the settings and style accordingly.
        """
        color = QColorDialog.getColor()
        if color.isValid():
            self.settings['background_color'] = color.name()
            self.save_settings()
            self.update_background_style() 

    def change_background_opacity(self, opacity):
        """
        Change the background opacity of the audio control widget.

        This method updates the settings and style based on the new opacity value.
        """
        self.settings['background_opacity'] = opacity
        self.save_settings()
        self.update_background_style() 

    def change_buttons_color(self, dialog):
        """
        Change the button color of the audio control widget.

        This method opens a color dialog and updates the settings and style accordingly.
        """
        color = QColorDialog.getColor()
        if color.isValid():
            self.settings['buttons_color'] = color.name()
            self.save_settings()
            self.apply_style()

    def change_buttons_opacity(self, opacity):
        """
        Change the button opacity of the audio control widget.

        This method updates the settings and style based on the new opacity value.
        """
        self.settings['buttons_opacity'] = opacity
        self.save_settings()
        self.apply_style()

    def load_settings(self):
        """
        Load the user settings from a JSON file.

        If the settings file doesn't exist, it creates a default settings dictionary.
        """
        try:
            with open('widget_settings.json', 'r') as f:
                return json.load(f)
        except:
            return {
                'background_color': '#2E2E2E',
                'background_opacity': 0.6,
                'buttons_color': '#2E2E2E',
                'buttons_opacity': 0.9,
                'autostart': False
            }

    def save_settings(self):
        """
        Save the current user settings to a JSON file.
        """
        with open('widget_settings.json', 'w') as f:
            json.dump(self.settings, f)

    def setup_auto_start(self):
        """
        Set up the autostart functionality for the audio control widget.

        This method checks the user settings and registers or unregisters the application to start automatically at system boot.
        """
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, 
                                winreg.KEY_ALL_ACCESS)
            if self.settings.get('autostart', False):
                winreg.SetValueEx(key, "AudioWidget", 0, winreg.REG_SZ, sys.argv[0])
            else:
                try:
                    winreg.DeleteValue(key, "AudioWidget")
                except:
                    pass
            key.Close()
        except:
            pass

    def toggle_autostart(self, checked):
        """
        Toggle the autostart functionality for the audio control widget.

        This method updates the user settings and sets up the autostart accordingly.
        """
        self.settings['autostart'] = checked
        self.save_settings()
        self.setup_auto_start()

    def init_volumes(self):
        """
        Initialize the master and microphone volume levels.

        This method retrieves the current volume levels from the system audio APIs and sets the corresponding sliders.
        """
        try:
            sessions = AudioUtilities.GetSpeakers()
            interface = sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            current_volume = volume.GetMasterVolumeLevelScalar()
            self.master_volume.setValue(int(current_volume * 100))
        except Exception as e:
            print(f"Error while initializing the speaker volume: {e}")

        try:
            mic_sessions = AudioUtilities.GetMicrophone() 
            mic_interface = mic_sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            mic_volume = cast(mic_interface, POINTER(IAudioEndpointVolume))
            mic_current_volume = mic_volume.GetMasterVolumeLevelScalar()
            self.mic_volume.setValue(int(mic_current_volume * 100)) 
        except Exception as e:
            print(f"Error while initializing the microphone volume: {e}")

    def change_master_volume(self, value):
        """
        Change the master (speaker) volume.

        This method interacts with the system audio APIs to update the master volume level.
        """
        try:
            sessions = AudioUtilities.GetSpeakers()
            interface = sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            
            volume.SetMasterVolumeLevelScalar(value / 100, None)
        except Exception as e:
            print(f"Error while changing the speaker volume: {e}")

    def change_mic_volume(self, value):
        """
        Change the microphone volume.

        This method interacts with the system audio APIs to update the microphone volume level.
        """
        try:
            sessions = AudioUtilities.GetMicrophone() 
            interface = sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            
            volume.SetMasterVolumeLevelScalar(value / 100, None)
        except Exception as e:
            print(f"Error while changing the microphone volume: {e}")

    def mousePressEvent(self, event):
        """
        Handle the mouse press event to enable window dragging.

        This method is called when the user presses the mouse button on the widget.
        """
        if event.button() == Qt.MouseButton.LeftButton: 
            self.oldPos = event.globalPosition().toPoint()
            event.accept() 

    def mouseMoveEvent(self, event):
        """
        Handle the mouse move event to enable window dragging.

        This method is called when the user moves the mouse while holding the left button.
        """
        if event.buttons() == Qt.MouseButton.LeftButton and hasattr(self, 'oldPos'):
            delta = event.globalPosition().toPoint() - self.oldPos
            self.move(self.x() + delta.x(), self.y() + delta.y())
            self.oldPos = event.globalPosition().toPoint() 
            event.accept()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    widget = AudioWidget()
    widget.show()
    sys.exit(app.exec())