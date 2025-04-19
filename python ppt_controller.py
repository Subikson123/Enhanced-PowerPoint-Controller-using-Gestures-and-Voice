import cv2
import time
import os
import sys
import pythoncom
import comtypes.client
import mediapipe as mp
import pyautogui
import speech_recognition as sr
import threading
from queue import Queue
import win32gui
import win32con

# ===== Configuration =====
PPT_FILENAME = "presentation.pptx"
GESTURE_HOLD_TIME = 1.2
GESTURE_COOLDOWN = 1.8
VOICE_COOLDOWN = 2.0
CAMERA_WIDTH = 640
CAMERA_HEIGHT = 480
WAKE_WORD = "control"
MICROPHONE_SENSITIVITY = 3500

# ===== Utility Functions =====
def find_presentation():
    for file in os.listdir():
        if file.endswith(".pptx") or file.endswith(".ppt"):
            return file
    return None

def init_camera():
    cap = cv2.VideoCapture(0)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, CAMERA_WIDTH)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, CAMERA_HEIGHT)
    return cap if cap.isOpened() else None

def set_window_topmost(window_name):
    hwnd = win32gui.FindWindow(None, window_name)
    if hwnd:
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                              win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

# ===== PowerPoint Control =====
class PowerPointController:
    def __init__(self):
        self.ppt = None
        self.presentation = None
        self.last_slide_change = 0
        self.presentation_loaded = False

    def start(self, filepath):
        try:
            pythoncom.CoInitialize()
            self.ppt = comtypes.client.CreateObject("PowerPoint.Application")
            self.ppt.Visible = True

            for attempt in range(3):
                try:
                    abs_path = os.path.abspath(filepath)
                    print(f"Attempting to open: {abs_path}")
                    self.presentation = self.ppt.Presentations.Open(abs_path)
                    self.presentation_loaded = True
                    break
                except Exception as e:
                    if attempt == 2:
                        raise
                    time.sleep(1.5)

            if self.presentation_loaded:
                self.presentation.SlideShowSettings.Run()
                time.sleep(1.5)
                return True
        except Exception as e:
            print(f"PowerPoint Error: {str(e)}")
            try:
                os.startfile(filepath)
                print("Opened presentation in default viewer as fallback")
                return True
            except:
                return False

    def can_change_slide(self):
        return (time.time() - self.last_slide_change) > GESTURE_COOLDOWN

    def next_slide(self):
        if self.can_change_slide():
            try:
                if self.presentation and hasattr(self.presentation, 'SlideShowWindow'):
                    self.presentation.SlideShowWindow.View.Next()
                else:
                    pyautogui.press("right")
                self.last_slide_change = time.time()
                print("Action: Next slide")
            except Exception as e:
                print(f"Slide change error: {e}")
                pyautogui.press("right")

    def prev_slide(self):
        if self.can_change_slide():
            try:
                if self.presentation and hasattr(self.presentation, 'SlideShowWindow'):
                    self.presentation.SlideShowWindow.View.Previous()
                else:
                    pyautogui.press("left")
                self.last_slide_change = time.time()
                print("Action: Previous slide")
            except Exception as e:
                print(f"Slide change error: {e}")
                pyautogui.press("left")

    def close(self):
        try:
            if self.presentation_loaded and self.presentation:
                try:
                    if hasattr(self.presentation, 'SlideShowWindow'):
                        self.presentation.SlideShowWindow.View.Exit()
                except Exception as e:
                    print(f"Error exiting slideshow view: {e}")
                try:
                    self.presentation.Close()
                except Exception as e:
                    print(f"Error closing presentation: {e}")
            if self.ppt:
                self.ppt.Quit()
        finally:
            pythoncom.CoUninitialize()

# ===== Gesture Detector =====
class GestureDetector:
    def __init__(self):
        self.hands = mp.solutions.hands.Hands(
            max_num_hands=1,
            min_detection_confidence=0.75,
            min_tracking_confidence=0.6,
            static_image_mode=False
        )
        self.current_gesture = None
        self.gesture_start_time = 0
        self.gesture_triggered = False

    def detect(self, frame):
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        results = self.hands.process(frame_rgb)
        gesture = None

        if results.multi_hand_landmarks:
            for hand_landmarks in results.multi_hand_landmarks:
                mp.solutions.drawing_utils.draw_landmarks(
                    frame, hand_landmarks, mp.solutions.hands.HAND_CONNECTIONS,
                    mp.solutions.drawing_styles.get_default_hand_landmarks_style(),
                    mp.solutions.drawing_styles.get_default_hand_connections_style())

                landmarks = hand_landmarks.landmark
                fingers_open = [
                    landmarks[8].y < landmarks[6].y,
                    landmarks[12].y < landmarks[10].y,
                    landmarks[16].y < landmarks[14].y,
                    landmarks[20].y < landmarks[18].y
                ]
                thumb_up = landmarks[4].y < landmarks[3].y and landmarks[4].x < landmarks[3].x

                if thumb_up and not any(fingers_open):
                    gesture = "next"
                elif fingers_open[0] and fingers_open[1] and not any(fingers_open[2:]) and not thumb_up:
                    gesture = "previous"

        if gesture != self.current_gesture:
            self.current_gesture = gesture
            self.gesture_start_time = time.time()
            self.gesture_triggered = False

        if self.current_gesture:
            hold_duration = time.time() - self.gesture_start_time
            progress = min(int((hold_duration / GESTURE_HOLD_TIME) * 100), 100)
            color = (0, int(255 * (progress/100)), int(255 * (1 - progress/100)))
            thickness = 2 + int(progress/50)

            cv2.putText(frame, f"{self.current_gesture.upper()} {progress}%", 
                       (50, 80), cv2.FONT_HERSHEY_SIMPLEX, 1, color, thickness)

            if hold_duration >= GESTURE_HOLD_TIME and not self.gesture_triggered:
                cv2.putText(frame, "READY", (200, 80), 
                           cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 255), 3)
                self.gesture_triggered = True
                return frame, self.current_gesture, True

        return frame, None, False

# ===== Voice Controller =====
class VoiceController:
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        self.last_command_time = 0
        self.listening = False
        self.command_queue = Queue()

    def listen_in_background(self):
        self.listening = True
        thread = threading.Thread(target=self._listen_loop, daemon=True)
        thread.start()

    def _listen_loop(self):
        with self.microphone as source:
            self.recognizer.adjust_for_ambient_noise(source, duration=1.5)
            print("Ambient noise adjusted")

        while self.listening:
            try:
                with self.microphone as source:
                    print("Listening for voice input...")
                    audio = self.recognizer.listen(source, timeout=3, phrase_time_limit=4)

                command_text = self.recognizer.recognize_google(audio).lower()
                print("Heard:", command_text)

                if WAKE_WORD in command_text:
                    command_text = command_text.replace(WAKE_WORD, "").strip()

                    if command_text in ["next", "previous", "back", "stop", "exit"]:
                        self.command_queue.put(command_text)
                        self.last_command_time = time.time()
                        print("Wake word detected, command:", command_text)

            except sr.UnknownValueError:
                print("Could not understand audio")
            except sr.WaitTimeoutError:
                continue
            except Exception as e:
                print(f"Voice recognition error: {e}")

    def get_command(self):
        if not self.command_queue.empty():
            return self.command_queue.get()
        return None

    def process_command(self, command):
        if "next" in command:
            return "next"
        elif "previous" in command or "back" in command:
            return "previous"
        elif "stop" in command or "exit" in command:
            return "exit"
        return None

    def stop(self):
        self.listening = False

# ===== Main Loop =====
def main():
    print("\n=== Enhanced PowerPoint Control System ===")
    print("Initializing components...")

    ppt_path = find_presentation()
    if not ppt_path:
        ppt_path = input("Enter full path to PowerPoint file: ").strip('"')
        if not os.path.exists(ppt_path):
            print("Error: File not found")
            return

    ppt_controller = PowerPointController()
    if not ppt_controller.start(ppt_path):
        print("Failed to initialize PowerPoint control")
        return

    cap = init_camera()
    if not cap:
        ppt_controller.close()
        return

    gesture_detector = GestureDetector()
    voice_controller = VoiceController()
    voice_controller.listen_in_background()

    window_name = "PowerPoint Control"
    cv2.namedWindow(window_name, cv2.WINDOW_NORMAL)
    set_window_topmost(window_name)

    print("\nSystem ready! Control options:")
    print(f"- Gestures: Thumb up = Next | Index+Middle = Previous (Hold {GESTURE_HOLD_TIME}s)")
    if WAKE_WORD:
        print(f"- Voice: Say '{WAKE_WORD} [next/previous/stop]' (e.g., '{WAKE_WORD} next')")
    else:
        print("- Voice: Say 'next', 'previous', or 'stop'")
    print("- Press ESC to exit\n")

    try:
        while True:
            ret, frame = cap.read()
            if not ret:
                print("Camera error - attempting to reconnect...")
                cap.release()
                time.sleep(1)
                cap = init_camera()
                if not cap:
                    break
                continue

            frame = cv2.flip(frame, 1)
            frame, gesture, gesture_ready = gesture_detector.detect(frame)
            if gesture_ready:
                if gesture == "next":
                    ppt_controller.next_slide()
                elif gesture == "previous":
                    ppt_controller.prev_slide()

            command = voice_controller.get_command()
            if command:
                print(f"Processing voice command: {command}")
                action = voice_controller.process_command(command)

                if action == "next":
                    ppt_controller.next_slide()
                elif action == "previous":
                    ppt_controller.prev_slide()
                elif action == "exit":
                    print("Voice command received: Exit")
                    break

            cv2.putText(frame, f"Hold gesture for {GESTURE_HOLD_TIME}s", (10, 30),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2)
            cv2.putText(frame, "Press ESC to exit", (10, CAMERA_HEIGHT - 20),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2)
            mic_status = "Voice: Active" if voice_controller.listening else "Voice: Off"
            cv2.putText(frame, mic_status, (CAMERA_WIDTH-150, CAMERA_HEIGHT-20),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 0), 1)

            cv2.imshow(window_name, frame)
            set_window_topmost(window_name)

            if cv2.waitKey(1) == 27:
                print("ESC key pressed: Exit")
                break

    except KeyboardInterrupt:
        print("\nKeyboardInterrupt: Stopping system...")
    except Exception as e:
        print(f"Unexpected error: {e}")
    finally:
        print("\nShutting down components...")
        if cap:
            cap.release()
        cv2.destroyAllWindows()
        voice_controller.stop()
        ppt_controller.close()
        print("System shutdown complete.")

if __name__ == "__main__":
    main()
