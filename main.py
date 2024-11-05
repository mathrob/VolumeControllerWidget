#import os
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
    def __init__(self):
        super().__init__()
        self.settings = self.load_settings()
        self.is_mic_muted = False
        self.init_ui()
        self.setup_tray()
        self.setup_auto_start()

        # Restore saved position if it exists
        saved_position = self.settings.get('position')
        if saved_position:
            self.move(QPoint(*saved_position))

        # Timer pour vérifier le volume toutes les secondes
        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self.update_volumes)
        self.update_timer.start(200)  # Vérifie toutes les 200 ms

        # Initialize mute states
        self.is_master_muted = False
        self.is_mic_muted = False

        # # Setting file location
        # folder_path = Path("C:/Program Files/Volume Controller Widget")
        # os.makedirs(folder_path, exist_ok=True)
        # self.file_path = os.path.join(folder_path, 'widget_settings.json')
        
    def init_ui(self):
        # Configuration de la fenêtre principale
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        # Layout principal
        layout = QVBoxLayout()

        # Ajouter un bouton invisible pour capturer les clics
        self.move_button = QPushButton(self)
        self.move_button.setGeometry(self.rect())  # Occupe toute la zone du widget
        self.update_background_style()  # Met à jour le style du bouton au démarrage
        self.move_button.setFlat(True)  # Pas de bordure
        self.move_button.mousePressEvent = self.mousePressEvent
        self.move_button.mouseMoveEvent = self.mouseMoveEvent
        
        self.mic_label = QLabel()
        self.update_mic_icon()  # Set the initial icon
        self.mic_label.mousePressEvent = self.toggle_mute_mic
        
        # Contrôle du volume principal
        self.master_volume = QSlider(Qt.Orientation.Vertical)
        self.master_volume.setRange(0, 100)
        self.master_volume.valueChanged.connect(self.change_master_volume)
        
        # Contrôle du volume du microphone
        self.mic_volume = QSlider(Qt.Orientation.Vertical)
        self.mic_volume.setRange(0, 100)
        self.mic_volume.valueChanged.connect(self.change_mic_volume)
        
        # Labels pour le speaker et le micro
        self.speaker_label = QLabel("🔊")
        self.speaker_label.setFixedSize(48, 48)
        self.speaker_label.mousePressEvent = self.toggle_mute_master  # Connect click to toggle mute

        self.mic_label = QLabel("🎤")
        self.mic_label.setFixedSize(48, 48)
        self.mic_label.mousePressEvent = self.toggle_mute_mic  # Connect click to toggle mute
        
        # Bouton d'options avec une icône de roue dentée
        options_button = QPushButton("⋮")
        options_button.setFixedSize(24, 48)
        options_button.setStyleSheet("color: white; border: none; font-size: 18px;")
        options_button.clicked.connect(self.show_options_menu)

        # Application du style
        self.apply_style()
        
        # Ajout des widgets au layout
        controls_layout = QHBoxLayout()
        controls_layout.addWidget(self.speaker_label)
        controls_layout.addWidget(self.master_volume)
        controls_layout.addWidget(self.mic_label)
        controls_layout.addWidget(self.mic_volume)
        controls_layout.addWidget(options_button)
        
        layout.addLayout(controls_layout)
        self.setLayout(layout)
        
        # Initialisation des volumes
        self.init_volumes()
        
    def apply_style(self):
        buttons_color = self.settings.get('buttons_color', '#2E2E2E')
        buttons_opacity = self.settings.get('buttons_opacity', 0.9)
        
        style = f"""
        QWidget {{
            background-color: rgba{QColor(buttons_color).getRgb()[:-1] + (int(buttons_opacity * 255),)};
            border-radius: 10px;
            padding: 10px;
        }}
        QSlider::groove:vertical {{
            background: #4A4A4A;
            width: 20px;
            border-radius: 5px;
        }}
        QSlider::handle:vertical {{
            background: #007AFF;
            height: 20px;
            width: 20px;
            margin: 0 -4px;
            border-radius: 10px;
        }}
        QLabel {{
            color: white;
            font-size: 16px;
        }}
        """
        self.setStyleSheet(style)
    
    def update_background_style(self):
        # Mise à jour de la couleur et de l'opacité du bouton move_button
        background_color = self.settings.get('background_color', '#0000FF')  # Couleur par défaut (bleu)
        background_opacity = self.settings.get('background_opacity', 0.5)  # Opacité par défaut (50%)

        # Appliquer le style
        style = f"background: rgba({QColor(background_color).red()}, {QColor(background_color).green()}, {QColor(background_color).blue()}, {int(background_opacity * 255)});"
        self.move_button.setStyleSheet(style)

    def toggle_mute_master(self, event):
        # Toggle mute state for the speaker
        sessions = AudioUtilities.GetSpeakers()
        interface = sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))

        # Check and toggle mute
        self.is_master_muted = not self.is_master_muted
        volume.SetMute(self.is_master_muted, None)

        # Update the label to indicate mute/unmute
        self.speaker_label.setText("🔇" if self.is_master_muted else "🔊")

    def update_mic_icon(self):
        # Set base microphone icon
        mic_icon = QPixmap(24, 24)
        mic_icon.fill(Qt.GlobalColor.transparent)  # Transparent background

        painter = QPainter(mic_icon)
        
        # Draw microphone icon (simple representation)
        painter.setPen(Qt.GlobalColor.black)
        painter.drawText(mic_icon.rect(), Qt.AlignmentFlag.AlignCenter, "🎤")
        
        # If muted, overlay a red "X" on top of the icon
        if self.is_mic_muted:
            painter.setPen(QColor("red"))
            painter.setBrush(QColor("red"))
            painter.drawLine(5, 5, mic_icon.width() - 5, mic_icon.height() - 5)  # Diagonal line
            painter.drawLine(mic_icon.width() - 5, 5, 5, mic_icon.height() - 5)  # Cross line

        painter.end()
        
        # Set the resulting icon on the QLabel
        self.mic_label.setPixmap(mic_icon)

    def toggle_mute_mic(self, event):
        # Toggle mute state for the microphone
        self.is_mic_muted = not self.is_mic_muted
        
        # Obtenir le microphone par défaut
        sessions = AudioUtilities.GetMicrophone()  # Cela devrait vous donner le microphone par défaut
        interface = sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))

        # Mute ou rétablir le microphone
        volume.SetMute(self.is_mic_muted, None)

        # Update the microphone icon to show or hide the red "X"
        self.update_mic_icon()

    def update_volumes(self):
        # Récupère le volume actuel du haut-parleur
        sessions = AudioUtilities.GetSpeakers()
        interface = sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        current_volume = int(volume.GetMasterVolumeLevelScalar() * 100)
        
        # Met à jour le slider de volume principal si nécessaire
        if self.master_volume.value() != current_volume:
            self.master_volume.blockSignals(True)
            self.master_volume.setValue(current_volume)
            self.master_volume.blockSignals(False)

        # Vérifier si le haut-parleur est muet
        self.is_master_muted = volume.GetMute()
        self.speaker_label.setText("🔇" if self.is_master_muted else "🔊")

        # Récupère le volume actuel du microphone
        mic_sessions = AudioUtilities.GetMicrophone()  # Cela devrait vous donner le microphone par défaut
        mic_interface = mic_sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        mic_volume = cast(mic_interface, POINTER(IAudioEndpointVolume))
        mic_current_volume = int(mic_volume.GetMasterVolumeLevelScalar() * 100)

        # Met à jour le slider de volume du microphone si nécessaire
        if self.mic_volume.value() != mic_current_volume:
            self.mic_volume.blockSignals(True)
            self.mic_volume.setValue(mic_current_volume)
            self.mic_volume.blockSignals(False)

        # Vérifier si le microphone est muet
        self.is_mic_muted = mic_volume.GetMute()
        self.update_mic_icon()  # Met à jour l'icône du microphone en fonction de l'état de sourdine

    def closeEvent(self, event):
        # Sauvegarde la position actuelle avant de fermer
        self.settings['position'] = (self.x(), self.y())
        self.save_settings()
        super().closeEvent(event)

    def show_options_menu(self):
        # Créez le menu d'options
        menu = QMenu(self)
        
        # Action de personnalisation
        customize_action = QAction("Personnaliser", self)
        customize_action.triggered.connect(self.show_customize_dialog)
        menu.addAction(customize_action)
        
        # Action pour démarrage automatique
        autostart_action = QAction("Démarrage automatique", self)
        autostart_action.setCheckable(True)
        autostart_action.setChecked(self.settings.get('autostart', False))
        autostart_action.triggered.connect(self.toggle_autostart)
        menu.addAction(autostart_action)
        
        # Option de quitter
        quit_action = QAction("Quitter", self)
        quit_action.triggered.connect(QApplication.quit)
        menu.addAction(quit_action)
        
        # Afficher le menu sous le bouton
        menu.exec(self.mapToGlobal(QPoint(self.width() - 10, 30)))
    
    def setup_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon("audio_icon.ico"))  # Remplacer par votre icône
        
        # Menu contextuel
        tray_menu = QMenu()
        
        # Options de personnalisation
        customize_action = QAction("Personnaliser", self)
        customize_action.triggered.connect(self.show_customize_dialog)
        
        # Option de démarrage automatique
        self.autostart_action = QAction("Démarrage automatique", self)
        self.autostart_action.setCheckable(True)
        self.autostart_action.setChecked(self.settings.get('autostart', False))
        self.autostart_action.triggered.connect(self.toggle_autostart)
        
        # Quitter
        quit_action = QAction("Quitter", self)
        quit_action.triggered.connect(QApplication.quit)
        
        # Ajout des actions au menu
        tray_menu.addAction(customize_action)
        tray_menu.addAction(self.autostart_action)
        tray_menu.addSeparator()
        tray_menu.addAction(quit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
    
    def show_customize_dialog(self):
        dialog = QWidget()
        dialog.setWindowTitle("Personnalisation")
        layout = QVBoxLayout()

        # Sélecteur de couleur du background
        background_color = QPushButton("Changer la couleur du background")
        background_color.clicked.connect(lambda: self.change_background_color(dialog))

        # Contrôle de l'opacité du bouton de déplacement
        background_opacity_layout = QHBoxLayout()
        background_opacity_label = QLabel("Opacité du background:")
        background_opacity_spin = QSpinBox()
        background_opacity_spin.setRange(0, 100)
        background_opacity_spin.setValue(int(self.settings.get('background_opacity', 0.5) * 100))
        background_opacity_spin.valueChanged.connect(lambda v: self.change_background_opacity(v / 100))

        background_opacity_layout.addWidget(background_opacity_label)
        background_opacity_layout.addWidget(background_opacity_spin)
        
        # Sélecteur de couleur
        color_buttons = QPushButton("Changer la couleur des boutons")
        color_buttons.clicked.connect(lambda: self.change_buttons_color(dialog))
        
        # Contrôle de l'opacité
        buttons_opacity_layout = QHBoxLayout()
        buttons_opacity_label = QLabel("Opacité des boutons:")
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
        color = QColorDialog.getColor()
        if color.isValid():
            self.settings['background_color'] = color.name()
            self.save_settings()
            self.update_background_style()  # Mettre à jour le style du bouton
    
    def change_background_opacity(self, opacity):
        self.settings['background_opacity'] = opacity
        self.save_settings()
        self.update_background_style()  # Mettre à jour le style du bouton

    def change_buttons_color(self, dialog):
        color = QColorDialog.getColor()
        if color.isValid():
            self.settings['buttons_color'] = color.name()
            self.save_settings()
            self.apply_style()
    
    def change_buttons_opacity(self, opacity):
        self.settings['buttons_opacity'] = opacity
        self.save_settings()
        self.apply_style()
    
    def load_settings(self):
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
        
        with open('widget_settings.json', 'w') as f:
            json.dump(self.settings, f)
    
    def setup_auto_start(self):
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
        self.settings['autostart'] = checked
        self.save_settings()
        self.setup_auto_start()
    
    def init_volumes(self):
        # Initialisation du volume principal
        sessions = AudioUtilities.GetSpeakers()
        interface = sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        current_volume = volume.GetMasterVolumeLevelScalar()
        self.master_volume.setValue(int(current_volume * 100))

        # Initialisation du volume du microphone
        mic_sessions = AudioUtilities.GetMicrophone()  # Cela devrait vous donner le microphone par défaut
        mic_interface = mic_sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        mic_volume = cast(mic_interface, POINTER(IAudioEndpointVolume))
        mic_current_volume = mic_volume.GetMasterVolumeLevelScalar()
        self.mic_volume.setValue(int(mic_current_volume * 100))  # Met à jour le slider du microphone
    
    def change_master_volume(self, value):
        sessions = AudioUtilities.GetSpeakers()
        interface = sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(value / 100, None)
    
    def change_mic_volume(self, value):
        # Récupérer le microphone par défaut
        sessions = AudioUtilities.GetMicrophone()  # Cela devrait vous donner le microphone par défaut
        interface = sessions.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        
        # Change le volume du microphone
        volume.SetMasterVolumeLevelScalar(value / 100, None)

    
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:  # Répondre seulement au bouton gauche
            self.oldPos = event.globalPosition().toPoint()
            event.accept()  # Accepter l'événement pour indiquer qu'il a été traité

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.MouseButton.LeftButton and hasattr(self, 'oldPos'):
            delta = event.globalPosition().toPoint() - self.oldPos
            self.move(self.x() + delta.x(), self.y() + delta.y())
            self.oldPos = event.globalPosition().toPoint()  # Mettre à jour la position ancienne
            event.accept()  # Accepter l'événement pour indiquer qu'il a été traité



if __name__ == '__main__':
    app = QApplication(sys.argv)
    widget = AudioWidget()
    widget.show()
    sys.exit(app.exec())