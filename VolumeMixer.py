import sys
import os
import json
import re
import winreg as reg
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QSlider, QLabel, QHBoxLayout,
                             QProgressBar, QPushButton, QMenu, QAction, QSystemTrayIcon, QSizePolicy)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QIcon
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume, IAudioMeterInformation

class VolumeMixer(QWidget):
    def __init__(self):
        super().__init__()
        self.settings_file = 'settings.json'
        self.initUI()
        self.load_settings()  # Load settings when initializing
        self.add_to_startup()
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon('icon.png'))
        self.tray_icon.setToolTip('Volume Mixer')
        self.tray_icon.show()

        self.tray_menu = QMenu()
        self.restore_action = QAction('Restore')
        self.restore_action.triggered.connect(self.show)
        self.quit_action = QAction('Quit')
        self.quit_action.triggered.connect(self.close)

        self.tray_menu.addAction(self.restore_action)
        self.tray_menu.addAction(self.quit_action)
        self.tray_icon.setContextMenu(self.tray_menu)

        self.programs = {}
        self.sliders = {}
        self.level_bars = {}
        self.mute_buttons = {}
        self.process_ids = set()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_levels)
        self.timer.start(1000)  # Update every second
        self.drag_start_position = None

    def initUI(self):
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)

        # Create the main layout
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)  # Remove margins around the main layout
        layout.setSpacing(0)  # Remove spacing between widgets in the main layout
        self.setLayout(layout)

        # Create the program list layout
        self.program_list_layout = QVBoxLayout()
        self.program_list_layout.setSpacing(0)  # No spacing between items in the program list
        self.program_list_layout.setContentsMargins(0, 0, 0, 0)  # No margins for compact layout
        layout.addLayout(self.program_list_layout)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_start_position = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if self.drag_start_position:
            self.move(event.globalPos() - self.drag_start_position)
            event.accept()

    def mouseReleaseEvent(self, event):
        self.drag_start_position = None

    def set_volume(self, program_name, value):
        try:
            if program_name in self.programs:
                volume = self.programs[program_name]['volume']
                volume.SetMasterVolume(value / 100, None)
                if value > 0:
                    self.mute_buttons[program_name].setStyleSheet("background-color: #555555; color: white; border-radius: 15px;")
        except Exception as e:
            print(f"Error setting volume for {program_name}: {e}")

    def toggle_mute(self, program_name):
        try:
            if program_name in self.programs:
                volume = self.programs[program_name]['volume']
                current_volume = volume.GetMasterVolume()
                if current_volume > 0:
                    self.programs[program_name]['last_volume'] = current_volume
                    volume.SetMasterVolume(0, None)
                    self.mute_buttons[program_name].setStyleSheet("background-color: red; color: white; border-radius: 15px;")
                else:
                    last_volume = self.programs[program_name].get('last_volume', 1)
                    volume.SetMasterVolume(last_volume, None)
                    self.mute_buttons[program_name].setStyleSheet("background-color: #555555; color: white; border-radius: 15px;")
        except Exception as e:
            print(f"Error toggling mute for {program_name}: {e}")

    def update_levels(self):
        try:
            sessions = AudioUtilities.GetAllSessions()
            current_process_ids = set()
            new_programs = []
            exclude_processes = [
                'audiodg.exe', 'explorer.exe', 'SndVol.exe', 'SearchUI.exe',
                'svchost.exe', 'avastUI.exe', 'avgui.exe', 'OneDrive.exe',
                'Dropbox.exe', 'googledrivesync.exe','SteelSeriesSonar.exe',
                'SteelSeriesEngine.exe', 'SteelSeriesPrism.exe', 'steelseriessonar.exe'
            ]
            media_players = {
                'vlc.exe': 'VLC',
                'wmplayer.exe': 'Windows Media Player',
                'iTunes.exe': 'iTunes',
                'potplayer.exe': 'PotPlayer',
                'mpc-hc.exe': 'Media Player',
                'zoom.exe': 'Zoom',
                'Zoom.exe': 'Zoom',
                'zoommeeting.exe': 'Zoom',
                'MicrosoftEdge.exe': 'Edge',
                'Teams.exe': 'Microsoft Teams',
                'Word.exe': 'Microsoft Word',
                'Excel.exe': 'Microsoft Excel',
                'PowerPoint.exe': 'PowerPoint',
                'whatsapp.exe': 'WhatsApp',
                'whatsappvoip.exe': 'WhatsApp',
                'steamwebhelper.exe': 'Steam',
                'discord.exe': 'Discord',
                'opera.exe': 'Opera',
                'zWebview2Agent': 'Zoom'
            }

            process_map = {}

            for session in sessions:
                if session.Process:
                    process_name = session.Process.name().lower()
                    process_id = session.Process.pid

                    # Skip excluded processes
                    if process_name in exclude_processes:
                        continue

                    main_program_name = media_players.get(process_name, process_name)
                    cleaned_name = re.split(r'[.,()]', main_program_name)[0].strip()

                    if cleaned_name not in process_map:
                        process_map[cleaned_name] = set()

                    process_map[cleaned_name].add(process_id)
                    current_process_ids.add(process_id)

                    if cleaned_name not in self.programs:
                        display_name = media_players.get(cleaned_name, cleaned_name)
                        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                        audio_meter = session._ctl.QueryInterface(IAudioMeterInformation)
                        self.programs[display_name] = {'volume': volume, 'meter': audio_meter}

                        hbox = QHBoxLayout()
                        hbox.setSpacing(0)
                        hbox.setContentsMargins(5, 5, 5, 5)

                        label = QLabel(display_name)
                        label.setFixedSize(120, 30)
                        label.setWordWrap(True)
                        label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                        label.setStyleSheet("""
                            QLabel {
                                background-color: #333333;
                                color: #FFFFFF;
                                padding: 5px;
                                border-radius: 5px;
                                font-weight: bold;
                                font-size: 15px;
                            }
                        """)
                        hbox.addWidget(label)

                        mute_button = QPushButton("M")
                        mute_button.setFixedSize(30, 30)
                        mute_button.setStyleSheet("background-color: #666666; color: white; border-radius: 15px;")
                        mute_button.clicked.connect(lambda value, name=display_name: self.toggle_mute(name))
                        self.mute_buttons[display_name] = mute_button

                        hbox.addWidget(mute_button)

                        slider = QSlider(Qt.Horizontal)
                        slider.setMinimum(0)
                        slider.setMaximum(100)
                        slider.setValue(int(volume.GetMasterVolume() * 100))
                        slider.valueChanged.connect(lambda value, name=display_name: self.set_volume(name, value))
                        self.sliders[display_name] = slider

                        slider.setFixedSize(200, 30)
                        slider.setStyleSheet("""
                            QSlider::groove:horizontal {
                                border: 1px solid #999999;
                                height: 8px;
                                background: #B0C4DE;
                                margin: 2px 0;
                            }
                            QSlider::handle:horizontal {
                                background: #FFFFFF;
                                border: 1px solid #FFFFFF;
                                width: 15px;
                                margin: -2px 0;
                                border-radius: 15px;
                            }
                        """)
                        hbox.addWidget(slider)

                        level_bar = QProgressBar()
                        level_bar.setMinimum(0)
                        level_bar.setMaximum(100)
                        level_bar.setTextVisible(False)
                        level_bar.setFixedSize(100, 30)
                        level_bar.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

                        level_bar.setStyleSheet("""
                            QProgressBar {
                                border: 1px solid #999999;
                                border-radius: 3px;
                                background-color: rgba(224, 224, 224, 150);
                                margin-left: 10px;
                                margin-right: 10px;
                            }
                            QProgressBar::chunk {
                                background-color: #66CDAA;
                                width: 10px;
                            }
                        """)
                        hbox.addWidget(level_bar)
                        self.level_bars[display_name] = level_bar
                        new_programs.append(hbox)

            programs_to_remove = [program_name for program_name in self.programs if program_name not in process_map]

            for program_name in programs_to_remove:
                self.remove_program_from_ui(program_name)

            for hbox in new_programs:
                self.program_list_layout.addLayout(hbox)

            self.process_ids = current_process_ids

            for program_name in self.programs:
                process_ids = process_map.get(program_name, set())
                program_volume = self.programs[program_name]['volume']
                program_meter = self.programs[program_name]['meter']

                if not process_ids:
                    continue

                process_ids = list(process_ids)
                session_volume = program_volume.GetMasterVolume()
                session_meter = program_meter.GetPeakValue() * 100

                self.sliders[program_name].setValue(int(session_volume * 100))
                self.level_bars[program_name].setValue(int(session_meter))
        except Exception as e:
            print(f"Error updating levels: {e}")

    def remove_program_from_ui(self, program_name):
        try:
            # Iterate over the items in the program_list_layout
            for i in range(self.program_list_layout.count()):
                item = self.program_list_layout.itemAt(i)
                if item:
                    layout = item.layout()
                    if layout and layout.count() > 0:
                        label = layout.itemAt(0).widget()
                        if label and label.text() == program_name:
                            # Remove widgets in the layout
                            label.deleteLater()
                            slider = layout.itemAt(1).widget()
                            if slider:
                                slider.deleteLater()
                            mute_button = layout.itemAt(2).widget()
                            if mute_button:
                                mute_button.deleteLater()
                            level_bar = layout.itemAt(3).widget()
                            if level_bar:
                                level_bar.deleteLater()
                            layout.deleteLater()
                            if program_name in self.programs:
                                del self.programs[program_name]
                            if program_name in self.sliders:
                                del self.sliders[program_name]
                            if program_name in self.mute_buttons:
                                del self.mute_buttons[program_name]
                            if program_name in self.level_bars:
                                del self.level_bars[program_name]
                            break
        except Exception as e:
            print(f"Error removing program from UI: {e}")


    def add_to_startup(self):
        try:
            key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
            value = 'VolumeMixer'
            reg_key = reg.OpenKey(reg.HKEY_CURRENT_USER, key, 0, reg.KEY_SET_VALUE)
            reg.SetValueEx(reg_key, value, 0, reg.REG_SZ, sys.executable)
            reg.CloseKey(reg_key)
        except Exception as e:
            print(f"Error adding to startup: {e}")

    def closeEvent(self, event):
        self.save_settings()  # Save settings before closing
        super().closeEvent(event)

    def save_settings(self):
        try:
            settings = {
                'geometry': self.saveGeometry().data().decode('latin1')  # Convert QByteArray to string
            }
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f)
            print("Settings saved:", settings)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def load_settings(self):
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)
                    geometry = settings.get('geometry')
                    if geometry:
                        self.restoreGeometry(bytes(geometry, 'latin1'))  # Convert string to QByteArray
                        print("Settings loaded:", settings)
        except Exception as e:
            print(f"Error loading settings: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mixer = VolumeMixer()
    mixer.show()
    sys.exit(app.exec_())

