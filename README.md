# Volume Mixer

Volume Mixer is a PyQt-based application that allows users to control the volume of different programs on their Windows machine. The application leverages the PyCAW library to interact with the Windows Core Audio API.

## Features

- **Real-time Volume Control:** Adjust the volume of individual applications in real-time using a sleek, transparent GUI.
- **Mute/Unmute Programs:** Easily mute or unmute applications with a single click.
- **Audio Level Visualization:** View the audio levels for each program using a progress bar.
- **Minimize to System Tray:** The app can be minimized to the system tray, allowing easy access without cluttering your taskbar.
- **Drag-and-Drop Window Positioning:** Click and drag anywhere on the window to reposition it on your screen.

## Installation

### Prerequisites

Ensure you have Python 3.x installed on your system. The required packages can be installed using `pip`.

1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/volume-mixer.git
   cd volume-mixer

2. **Create and activate a virtual environment**:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`

3. **Install the required packages**:
   ```bash
   pip install -r requirements.txt

4. **Install PyCAW**:
   ```bash
   pip install pycaw

Running the Application
After installing the required packages, you can run the application using:
   ```bash
   python volume_mixer.py

