import sys
import os
import re
import winreg as reg
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QSlider, QLabel, QHBoxLayout, QProgressBar, QPushButton, \
    QMenu, QAction, QSystemTrayIcon, QSizePolicy
from PyQt5.QtCore import Qt, QTimer, QSettings
from PyQt5.QtGui import QIcon
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume, IAudioMeterInformation

class VolumeMixer(QWidget):
    def __init__(self):
        super().__init__()
        self.settings = QSettings('MyApp', 'VolumeMixer')
        self.initUI()
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
        self.restore_geometry()

    def initUI(self):
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        layout = QVBoxLayout()
        self.setGeometry(600, 600, 600, 800)  # Default geometry

        # Program list layout
        self.program_list_layout = QVBoxLayout()
        self.program_list_layout.setSpacing(3)  # Increased spacing for better visibility
        self.program_list_layout.setContentsMargins(10, 10, 10, 10)  # Margins around the layout
        layout.addLayout(self.program_list_layout)
        self.setLayout(layout)

    def set_volume(self, program_name, value):
        try:
            if program_name in self.programs:
                volume = self.programs[program_name]['volume']
                volume.SetMasterVolume(value / 100, None)
                # Update mute button color if volume is not 0
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
                    # Save the current volume before muting
                    self.programs[program_name]['last_volume'] = current_volume
                    volume.SetMasterVolume(0, None)
                    self.mute_buttons[program_name].setStyleSheet(
                        "background-color: red; color: white; border-radius: 15px;")
                else:
                    # Restore the last volume after unmuting
                    last_volume = self.programs[program_name].get('last_volume', 1)
                    volume.SetMasterVolume(last_volume, None)
                    self.mute_buttons[program_name].setStyleSheet(
                        "background-color: #555555; color: white; border-radius: 15px;")
        except Exception as e:
            print(f"Error toggling mute for {program_name}: {e}")

    def update_levels(self):
        try:
            sessions = AudioUtilities.GetAllSessions()
            current_process_ids = set()
            new_programs = []
            exclude_processes = ['audiodg.exe', 'explorer.exe', 'SndVol.exe', 'SearchUI.exe', 'svchost.exe',
                                 'chrome.exe', 'msedge.exe', 'firefox.exe', 'avastUI.exe', 'avgui.exe',
                                 'OneDrive.exe', 'Dropbox.exe', 'googledrivesync.exe']
            media_players = {
                'vlc.exe': 'VLC Media Player',
                'wmplayer.exe': 'Windows Media Player',
                'iTunes.exe': 'iTunes',
                'potplayer.exe': 'PotPlayer',
                'mpc-hc.exe': 'Media Player Classic',
                'Microsoft.Media.Player.exe': 'Microsoft Media Player',
                'MicrosoftEdge.exe': 'Microsoft Edge',
                'Teams.exe': 'Microsoft Teams',
                'Word.exe': 'Microsoft Word',
                'Excel.exe': 'Microsoft Excel',
                'PowerPoint.exe': 'Microsoft PowerPoint'
            }
            process_map = {}

            for session in sessions:
                if session.Process:
                    process_name = session.Process.name()
                    process_id = session.Process.pid

                    cleaned_name = re.split(r'[.,()]', process_name)[0].strip()

                    if cleaned_name not in process_map:
                        process_map[cleaned_name] = set()

                    process_map[cleaned_name].add(process_id)
                    current_process_ids.add(process_id)

                    if cleaned_name not in self.programs:
                        # If the process name matches one of the media players, use the corresponding display name
                        display_name = media_players.get(cleaned_name, cleaned_name)
                        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                        audio_meter = session._ctl.QueryInterface(IAudioMeterInformation)
                        self.programs[display_name] = {'volume': volume, 'meter': audio_meter}

                        hbox = QHBoxLayout()
                        hbox.setSpacing(5)
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

                        mute_button = QPushButton("M")
                        mute_button.setFixedSize(30, 30)
                        mute_button.setStyleSheet("background-color: #555555; color: white; border-radius: 15px;")
                        mute_button.clicked.connect(lambda _, name=display_name: self.toggle_mute(name))
                        self.mute_buttons[display_name] = mute_button

                        hbox.addWidget(mute_button)

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

            programs_to_remove = [program_name for program_name in self.programs.keys() if
                                  program_name not in process_map]
            for program_name in programs_to_remove:
                self.remove_program_from_ui(program_name)

            for hbox in new_programs:
                self.program_list_layout.addLayout(hbox)

            self.process_ids = current_process_ids

            for program_name, program_data in self.programs.items():
                meter = program_data['meter']
                peak_value = meter.GetPeakValue()
                if program_name in self.level_bars:
                    self.level_bars[program_name].setValue(int(peak_value * 100))

            self.adjust_size()
        except Exception as e:
            print(f"Error updating levels: {e}")

    def adjust_size(self):
        num_programs = len(self.programs)
        height = max(0, 50 * num_programs + 50)  # Ensure height remains at least 800
        self.setFixedHeight(height)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_start_position = event.pos()

    def mouseMoveEvent(self, event):
        if self.drag_start_position:
            self.move(self.pos() + event.pos() - self.drag_start_position)

    def mouseReleaseEvent(self, event):
        self.drag_start_position = None

    def closeEvent(self, event):
        self.save_geometry()
        self.tray_icon.hide()  # Hide tray icon when closing
        super().closeEvent(event)

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

    def save_geometry(self):
        self.settings.setValue('geometry', self.saveGeometry())

    def restore_geometry(self):
        geometry = self.settings.value('geometry')
        if geometry:
            self.restoreGeometry(geometry)

    def add_to_startup(self):
        try:
            # Get the path to the executable
            exe_path = sys.executable

            # If running as a script, change the path to the script
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
            else:
                exe_path = os.path.abspath(__file__)

            # Key for the Windows startup folder
            key = reg.HKEY_CURRENT_USER
            key_value = r"Software\Microsoft\Windows\CurrentVersion\Run"
            key_to_open = reg.OpenKey(key, key_value, 0, reg.KEY_ALL_ACCESS)
            reg.SetValueEx(key_to_open, "VolumeMixer", 0, reg.REG_SZ, exe_path)
            reg.CloseKey(key_to_open)
            print("Successfully added to startup")
        except Exception as e:
            print(f"Failed to add to startup: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mixer = VolumeMixer()
    mixer.show()
    sys.exit(app.exec_())
