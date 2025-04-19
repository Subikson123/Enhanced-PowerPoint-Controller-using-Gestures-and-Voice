# 🖐️🎤 Enhanced PowerPoint Control System

An intelligent PowerPoint controller using **hand gestures** and **voice commands** to navigate slides during presentations. This project combines computer vision and speech recognition to create a hands-free, intuitive presentation experience.

---

## 📌 Features

- **🎯 Gesture-Based Control**
  - **Next Slide**: Show **thumbs up**
  - **Previous Slide**: Raise **index + middle fingers**
  - Gesture must be held for a configurable duration (default: 1.2s)

- **🎙️ Voice-Activated Commands**
  - Use the wake word `"control"` followed by commands:
    - `control next`
    - `control previous` or `control back`
    - `control stop` or `control exit`
  - Configurable sensitivity and cooldown settings

- **📊 PowerPoint Automation**
  - Opens `.pptx` files and starts the slideshow automatically
  - Uses COM interface to switch slides or fallbacks to `pyautogui` keystrokes

- **📷 Real-Time Feedback**
  - Visual overlay for gesture progress
  - On-screen voice activation status
  - Live webcam feed with hand tracking using MediaPipe

---

## 🛠️ Tech Stack

| Module              | Purpose                            |
|---------------------|------------------------------------|
| OpenCV              | Webcam integration & UI rendering  |
| MediaPipe           | Real-time hand tracking            |
| SpeechRecognition   | Voice input capture & parsing      |
| comtypes / pythoncom| PowerPoint COM automation          |
| pyautogui           | Fallback slide control             |
| win32gui / win32con | Always-on-top camera window        |

---

## 🚀 Getting Started

### 🔧 Prerequisites

- Python 3.8+
- Windows OS (PowerPoint COM automation is Windows-only)
- PowerPoint installed (2013 or newer recommended)
- A working webcam and microphone

### 📦 Install Requirements

```bash
pip install opencv-python mediapipe pyautogui comtypes pypiwin32 SpeechRecognition
Optional (for first-time microphone setup):

bash
Copy
Edit
pip install PyAudio
🧾 File Structure
graphql
Copy
Edit
📁 Project
│
├── presentation.pptx   # PowerPoint file to control
├── controller.py       # Main program file
└── README.md           # Documentation
Make sure the .pptx file is in the same directory as the script or provide the full path on execution.

🧪 How It Works
🖐️ Gesture Detection Logic
MediaPipe tracks hand landmarks in real-time.

Logic checks for specific hand shapes:

Thumb Up: Next slide

Index + Middle: Previous slide

A gesture is triggered only when held steadily for a defined time.

🎙️ Voice Command Flow
Listens for input using Google’s speech recognition engine.

Commands are queued only if prefixed by the wake word.

Each command type is throttled to prevent rapid repeat actions.

🧠 Slide Control
If PowerPoint COM is available:

Uses SlideShowWindow.View.Next() and Previous()

Else:

Uses pyautogui.press("right") or "left" as fallback

🖥️ Usage Guide
Run the Program

bash
Copy
Edit
python controller.py
Interface

A webcam window will open.

Visual feedback appears for gestures and voice activity.

Start Presenting

Use gestures or voice to navigate.

Press ESC or say "control stop" to exit.

⚙️ Configuration
You can modify these values in the script:

python
Copy
Edit
GESTURE_HOLD_TIME = 1.2       # seconds to hold a gesture
GESTURE_COOLDOWN = 1.8        # delay between slide changes
VOICE_COOLDOWN = 2.0          # delay between voice commands
WAKE_WORD = "control"         # wake word to activate voice commands
MICROPHONE_SENSITIVITY = 3500 # adjust based on noise
CAMERA_WIDTH = 640
CAMERA_HEIGHT = 480
📈 Use Cases
Teachers & Educators: Navigate slides while facing students.

Public Speakers: Move slides without needing a clicker.

Remote Presentations: Reduce dependency on physical hardware.

🧩 Troubleshooting

Issue	Solution
Microphone not working	Check audio input settings or re-install PyAudio
No gesture detected	Ensure adequate lighting and visible hand
Presentation not loading	Provide correct .pptx path or try .ppt
🧹 Future Enhancements
✅ Laser-pointer-like finger tracking

🔄 Slide drawing with fingertip

🔊 Voice feedback via TTS

🧠 AI-powered smart commands (e.g., "skip to conclusion")

🧑‍💻 Author
Developed by: Your Name Here
If you use or modify this project, feel free to ⭐️ the repo and share your improvements!

📄 License
This project is licensed under the MIT License - see the LICENSE file for details.