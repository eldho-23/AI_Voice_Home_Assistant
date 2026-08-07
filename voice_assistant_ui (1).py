"""
AI Voice-Controlled Home Assistant - FINAL UI VERSION
Group 12 Assignment

Features:
- Modern dark UI (CustomTkinter)
- Real microphone listening
- Natural-sounding female voice (edge-tts)
- Live status of all appliances
- Chat-like conversation history
"""

import os
import sys
import threading
import tempfile
import asyncio
import speech_recognition as sr
from typing import Dict, Optional
from dataclasses import dataclass
from enum import Enum

# UI
import customtkinter as ctk
from tkinter import messagebox

# Better voice
try:
    import edge_tts
    import pygame
    HAS_EDGE = True
except ImportError:
    HAS_EDGE = False

# Fallback TTS
try:
    import win32com.client
    HAS_SAPI = True
except ImportError:
    HAS_SAPI = False


# ============================================================
# CONFIG
# ============================================================
VOICE = "en-US-JennyNeural"   # Natural female voice (very good)
# Other good options: en-US-AriaNeural, en-GB-SoniaNeural, en-US-MichelleNeural

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


# ============================================================
# SPEAKING
# ============================================================
def speak_edge(text: str):
    """High quality neural voice using edge-tts"""
    async def _speak():
        communicate = edge_tts.Communicate(text, VOICE)
        temp_file = os.path.join(tempfile.gettempdir(), "assistant_reply.mp3")
        await communicate.save(temp_file)
        pygame.mixer.init()
        pygame.mixer.music.load(temp_file)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            await asyncio.sleep(0.1)
        pygame.mixer.quit()
        try:
            os.remove(temp_file)
        except:
            pass

    try:
        asyncio.run(_speak())
    except Exception as e:
        print("edge-tts error:", e)
        speak_fallback(text)


def speak_fallback(text: str):
    """Windows SAPI fallback"""
    if HAS_SAPI:
        try:
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            # Try to select a female voice
            voices = speaker.GetVoices()
            for i in range(voices.Count):
                v = voices.Item(i)
                if "zira" in v.GetDescription().lower() or "female" in v.GetDescription().lower():
                    speaker.Voice = v
                    break
            speaker.Speak(text)
            return
        except:
            pass
    print(f"[SPEAK] {text}")


def speak(text: str, ui_callback=None):
    if ui_callback:
        ui_callback(text)
    if HAS_EDGE:
        speak_edge(text)
    else:
        speak_fallback(text)


# ============================================================
# HOME MODEL
# ============================================================
class ApplianceType(Enum):
    LIGHT = "light"
    FAN = "fan"
    AC = "ac"
    TV = "tv"
    PLUG = "plug"


@dataclass
class Appliance:
    name: str
    room: str
    type: ApplianceType
    is_on: bool = False
    value: Optional[int] = None

    def nice(self):
        return self.name.replace("_", " ").title()


class Home:
    def __init__(self):
        self.devices: Dict[str, Appliance] = {}
        data = [
            ("living_room_light", "Living Room", ApplianceType.LIGHT),
            ("bedroom_light", "Bedroom", ApplianceType.LIGHT),
            ("kitchen_light", "Kitchen", ApplianceType.LIGHT),
            ("living_room_fan", "Living Room", ApplianceType.FAN),
            ("bedroom_fan", "Bedroom", ApplianceType.FAN),
            ("living_room_ac", "Living Room", ApplianceType.AC),
            ("bedroom_ac", "Bedroom", ApplianceType.AC),
            ("living_room_tv", "Living Room", ApplianceType.TV),
            ("smart_plug", "Kitchen", ApplianceType.PLUG),
        ]
        for n, r, t in data:
            a = Appliance(n, r, t)
            if t == ApplianceType.AC:
                a.value = 24
            elif t == ApplianceType.LIGHT:
                a.value = 70
            elif t == ApplianceType.TV:
                a.value = 30
            self.devices[n] = a

    def status_text(self):
        parts = []
        for a in self.devices.values():
            state = "ON" if a.is_on else "OFF"
            if a.value is not None and a.is_on:
                parts.append(f"{a.nice()} is {state} ({a.value})")
            else:
                parts.append(f"{a.nice()} is {state}")
        return "Current status: " + ". ".join(parts)


# ============================================================
# AGENT
# ============================================================
class Agent:
    def __init__(self, home: Home):
        self.home = home
        self.last_device = None
        self.last_room = None

    def handle(self, text: str) -> str:
        cmd = text.lower().strip()
        print(f"[DEBUG] Raw command: '{cmd}'")

        if any(w in cmd for w in ["exit", "quit", "goodbye", "bye", "stop"]):
            return "EXIT"

        if any(w in cmd for w in ["status", "what is the status", "report", "list devices", "device status"]):
            return self.home.status_text()

        # ----- Detect ACTION -----
        action = None
        if any(w in cmd for w in ["turn on", "switch on", "switch on the", "on the", "start the"]):
            action = "on"
        elif any(w in cmd for w in ["turn off", "switch off", "switch off the", "off the", "stop the"]):
            action = "off"
        elif any(w in cmd for w in ["set", "change to", "make it"]):
            action = "set"
        elif any(w in cmd for w in ["increase", "raise", "up", "higher", "brighter", "louder"]):
            action = "increase"
        elif any(w in cmd for w in ["decrease", "lower", "down", "dim", "quieter", "reduce"]):
            action = "decrease"

        # ----- Detect ROOM (more flexible) -----
        room = None
        if "living room" in cmd or "livingroom" in cmd or "hall" in cmd:
            room = "Living Room"
        elif "bed room" in cmd or "bedroom" in cmd or "bed" in cmd:
            room = "Bedroom"
        elif "kitchen" in cmd:
            room = "Kitchen"

        # ----- Detect APPLIANCE TYPE (order matters – more specific first) -----
        atype = None
        if any(w in cmd for w in ["ac", "a c", "air conditioner", "aircon", "air condition"]):
            atype = ApplianceType.AC
        elif "fan" in cmd:
            atype = ApplianceType.FAN
        elif "tv" in cmd or "television" in cmd or "t v" in cmd:
            atype = ApplianceType.TV
        elif "plug" in cmd:
            atype = ApplianceType.PLUG
        elif "light" in cmd or "lamp" in cmd or "lights" in cmd:
            atype = ApplianceType.LIGHT

        # ----- Detect VALUE -----
        import re
        nums = re.findall(r"\b(\d+)\b", cmd)
        value = int(nums[0]) if nums else None

        print(f"[DEBUG] Detected → action={action}, room={room}, type={atype}, value={value}")

        # ----- FIND THE CORRECT DEVICE -----
        target = None

        # Priority 1: both room and type are known → exact match
        if atype is not None and room is not None:
            for d in self.home.devices.values():
                if d.type == atype and d.room == room:
                    target = d
                    break

        # Priority 2: only type known → use last room if available
        elif atype is not None and self.last_room is not None:
            for d in self.home.devices.values():
                if d.type == atype and d.room == self.last_room:
                    target = d
                    break

        # Priority 3: incomplete command (“turn it off”) → use last device
        elif action in ["on", "off", "increase", "decrease", "set"] and self.last_device:
            target = self.home.devices.get(self.last_device)

        if target is None:
            return "I could not identify the device. Please clearly say the room and the device, for example: turn on the bedroom light."

        print(f"[DEBUG] Matched device: {target.name}")

        # ----- EXECUTE ACTION -----
        if action == "on":
            if target.is_on:
                reply = f"The {target.nice()} is already on."
            else:
                target.is_on = True
                reply = f"Okay, turning on the {target.nice()}."
        elif action == "off":
            if not target.is_on:
                reply = f"The {target.nice()} is already off."
            else:
                target.is_on = False
                reply = f"Okay, turning off the {target.nice()}."
        elif action == "set" and value is not None:
            target.value = value
            target.is_on = True
            reply = f"Setting {target.nice()} to {value}."
        elif action == "increase" and target.value is not None:
            maxv = 30 if target.type == ApplianceType.AC else 100
            target.value = min(maxv, target.value + 5)
            target.is_on = True
            reply = f"Increased {target.nice()} to {target.value}."
        elif action == "decrease" and target.value is not None:
            minv = 16 if target.type == ApplianceType.AC else 0
            target.value = max(minv, target.value - 5)
            reply = f"Decreased {target.nice()} to {target.value}."
        else:
            reply = "I did not understand that command. Please try again."

        # Update context only after a successful match
        self.last_device = target.name
        self.last_room = target.room
        return reply


# ============================================================
# UI APPLICATION
# ============================================================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("AI Voice Home Assistant  |  Group 12")
        self.geometry("900x650")
        self.minsize(800, 580)

        self.home = Home()
        self.agent = Agent(self.home)
        self.recognizer = sr.Recognizer()
        self.listening = False
        self.stop_flag = False

        self._build_ui()
        self.after(800, self.welcome)

    def _build_ui(self):
        # Left panel - Conversation
        left = ctk.CTkFrame(self, corner_radius=12)
        left.pack(side="left", fill="both", expand=True, padx=12, pady=12)

        title = ctk.CTkLabel(left, text="AI Voice Home Assistant",
                             font=ctk.CTkFont(size=22, weight="bold"))
        title.pack(pady=(15, 5))

        subtitle = ctk.CTkLabel(left, text="Speak naturally • I listen & respond",
                                font=ctk.CTkFont(size=13), text_color="gray")
        subtitle.pack(pady=(0, 10))

        self.chat_box = ctk.CTkTextbox(left, font=ctk.CTkFont(size=14),
                                       wrap="word", state="disabled",
                                       corner_radius=10)
        self.chat_box.pack(fill="both", expand=True, padx=10, pady=5)

        # Status indicator
        self.status_label = ctk.CTkLabel(left, text="● Ready",
                                         font=ctk.CTkFont(size=13),
                                         text_color="#4ade80")
        self.status_label.pack(pady=5)

        # Buttons
        btn_frame = ctk.CTkFrame(left, fg_color="transparent")
        btn_frame.pack(pady=10)

        self.listen_btn = ctk.CTkButton(btn_frame, text="🎤  Start Listening",
                                        width=180, height=40,
                                        font=ctk.CTkFont(size=14, weight="bold"),
                                        command=self.toggle_listen)
        self.listen_btn.pack(side="left", padx=6)

        self.status_btn = ctk.CTkButton(btn_frame, text="📋 Status",
                                        width=110, height=40,
                                        command=self.show_status)
        self.status_btn.pack(side="left", padx=6)

        # Right panel - Devices
        right = ctk.CTkFrame(self, width=280, corner_radius=12)
        right.pack(side="right", fill="y", padx=(0, 12), pady=12)
        right.pack_propagate(False)

        ctk.CTkLabel(right, text="Home Devices",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(15, 10))

        self.device_frame = ctk.CTkScrollableFrame(right, corner_radius=8)
        self.device_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.device_labels = {}
        for name, app in self.home.devices.items():
            f = ctk.CTkFrame(self.device_frame, corner_radius=8)
            f.pack(fill="x", pady=4, padx=4)
            lbl = ctk.CTkLabel(f, text=self._device_text(app),
                               font=ctk.CTkFont(size=12), anchor="w")
            lbl.pack(fill="x", padx=8, pady=6)
            self.device_labels[name] = lbl

        ctk.CTkLabel(right, text="Group 12 Assignment",
                     font=ctk.CTkFont(size=11), text_color="gray").pack(pady=10)

    def _device_text(self, app: Appliance) -> str:
        icon = "🟢" if app.is_on else "⚫"
        state = "ON" if app.is_on else "OFF"
        if app.value is not None and app.is_on:
            return f"{icon}  {app.nice()}\n     {state}  •  {app.value}"
        return f"{icon}  {app.nice()}\n     {state}"

    def refresh_devices(self):
        for name, app in self.home.devices.items():
            self.device_labels[name].configure(text=self._device_text(app))

    def add_chat(self, who: str, text: str):
        self.chat_box.configure(state="normal")
        if who == "You":
            self.chat_box.insert("end", f"You:  {text}\n\n")
        else:
            self.chat_box.insert("end", f"Assistant:  {text}\n\n")
        self.chat_box.see("end")
        self.chat_box.configure(state="disabled")

    def set_status(self, text: str, color: str = "#4ade80"):
        self.status_label.configure(text=text, text_color=color)

    def welcome(self):
        msg = "Hello! I am your AI home assistant. How can I help you today?"
        self.add_chat("Assistant", msg)
        threading.Thread(target=lambda: speak(msg), daemon=True).start()

    def show_status(self):
        reply = self.home.status_text()
        self.add_chat("Assistant", reply)
        threading.Thread(target=lambda: speak(reply), daemon=True).start()
        self.refresh_devices()

    def toggle_listen(self):
        if not self.listening:
            self.listening = True
            self.stop_flag = False
            self.listen_btn.configure(text="⏹  Stop Listening", fg_color="#dc2626")
            self.set_status("● Listening...", "#facc15")
            threading.Thread(target=self.listen_loop, daemon=True).start()
        else:
            self.stop_flag = True
            self.listening = False
            self.listen_btn.configure(text="🎤  Start Listening",
                                      fg_color=["#3B8ED0", "#1F6AA5"])
            self.set_status("● Ready", "#4ade80")

    def listen_loop(self):
        mic = sr.Microphone()
        while self.listening and not self.stop_flag:
            try:
                with mic as source:
                    self.recognizer.adjust_for_ambient_noise(source, duration=0.3)
                    self.set_status("● Listening... Speak now", "#facc15")
                    audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=6)

                self.set_status("● Recognizing...", "#60a5fa")
                text = self.recognizer.recognize_google(audio)
                self.after(0, lambda t=text: self.process_command(t))

            except sr.WaitTimeoutError:
                continue
            except sr.UnknownValueError:
                self.after(0, lambda: self.set_status("● Could not understand", "#f87171"))
                continue
            except Exception as e:
                print("Listen error:", e)
                self.after(0, lambda: self.set_status("● Error", "#f87171"))
                break

        self.after(0, lambda: self.set_status("● Ready", "#4ade80"))

    def process_command(self, text: str):
        self.add_chat("You", text)
        reply = self.agent.handle(text)

        if reply == "EXIT":
            self.add_chat("Assistant", "Goodbye! Have a nice day.")
            threading.Thread(target=lambda: speak("Goodbye! Have a nice day."), daemon=True).start()
            self.stop_flag = True
            self.listening = False
            self.listen_btn.configure(text="🎤  Start Listening",
                                      fg_color=["#3B8ED0", "#1F6AA5"])
            return

        self.add_chat("Assistant", reply)
        self.refresh_devices()
        threading.Thread(target=lambda: speak(reply), daemon=True).start()
        self.set_status("● Ready", "#4ade80")


# ============================================================
# MAIN
# ============================================================
if __name__ == "__main__":
    # Check critical packages
    missing = []
    try:
        import customtkinter
    except ImportError:
        missing.append("customtkinter")
    try:
        import speech_recognition
    except ImportError:
        missing.append("SpeechRecognition")

    if missing:
        print("Please install missing packages:")
        print("  pip install " + " ".join(missing))
        print("  pip install edge-tts pygame pywin32 pyaudio")
        sys.exit(1)

    app = App()
    app.mainloop()
