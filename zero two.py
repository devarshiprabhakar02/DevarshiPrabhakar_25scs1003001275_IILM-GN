import speech_recognition as sr
import pyttsx3
import webbrowser
import os
import subprocess
from datetime import datetime
import threading
import re
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.scrollview import ScrollView
from kivy.uix.textinput import TextInput
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.uix.gridlayout import GridLayout


class Athena:
    def __init__(self, gui_callback=None):
        self.recognizer = sr.Recognizer()
        self.engine = pyttsx3.init()
        self.setup_voice()
        self.desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        self.running = True
        self.listening_active = True
        self.gui_callback = gui_callback

    def setup_voice(self):
        voices = self.engine.getProperty('voices')
        for voice in voices:
            if 'female' in voice.name.lower() or 'zira' in voice.name.lower():
                self.engine.setProperty('voice', voice.id)
                break
        self.engine.setProperty('rate', 200)
        self.engine.setProperty('volume', 1.0)
        try:
            self.engine.setProperty('pitch', 1.5)
        except:
            pass

    def speak(self, text):
        print(f"Athena: {text}")
        if self.gui_callback:
            self.gui_callback(f"Athena: {text}")
        self.engine.say(text)
        self.engine.runAndWait()

    def listen(self):
        with sr.Microphone() as source:
            if self.gui_callback:
                self.gui_callback("🎤 Listening...")
            print("Listening...")
            self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
            try:
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=10)
                text = self.recognizer.recognize_google(audio)
                print(f"You said: {text}")
                if self.gui_callback:
                    self.gui_callback(f"You: {text}")
                self.speak("Got it")
                return text.lower()
            except sr.WaitTimeoutError:
                return ""
            except sr.UnknownValueError:
                self.speak("Sorry, I didn't catch that. Could you repeat please?")
                return ""
            except sr.RequestError:
                self.speak("Sorry, I'm having trouble connecting to the speech service")
                return ""

    def ask_copilot(self, question):
        """Open Windows Copilot with the query"""
        try:
            # Method 1: Using Windows 11 Copilot URI
            copilot_uri = f"microsoft-edge:///?ux=copilot&q={question.replace(' ', '+')}"
            os.startfile(copilot_uri)
            self.speak(f"Opening Copilot to answer your question about {question}")
            return True
        except Exception as e:
            try:
                # Method 2: Using keyboard shortcut simulation
                subprocess.Popen(['powershell', '-Command',
                                  '(New-Object -ComObject WScript.Shell).SendKeys("^+{ESC}")'])
                self.speak("Opening Windows Copilot for you")
                return True
            except:
                self.speak("I couldn't open Copilot. Let me search that for you instead.")
                self.google_search(question)
                return False

    def extract_query(self, command, keywords):
        query = command
        for keyword in keywords:
            query = query.replace(keyword, '')
        query = query.replace('for', '').replace('about', '').replace('the', '').strip()
        return query

    def google_search(self, query):
        search_url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
        try:
            edge_path = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
            if os.path.exists(edge_path):
                subprocess.Popen([edge_path, search_url])
            else:
                webbrowser.open(search_url)
            self.speak(f"Searching Google for {query}! Here you go~")
        except Exception as e:
            self.speak("Oh no! I couldn't open the browser. Sorry~")

    def youtube_search(self, query):
        search_url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
        try:
            webbrowser.open(search_url)
            self.speak(f"Searching YouTube for {query}! Enjoy~")
        except Exception as e:
            self.speak("Oh no! I couldn't open YouTube. Sorry~")

    def open_website(self, site):
        sites = {
            'youtube': 'https://www.youtube.com',
            'gmail': 'https://mail.google.com',
            'github': 'https://github.com',
            'reddit': 'https://www.reddit.com',
            'twitter': 'https://twitter.com',
            'facebook': 'https://www.facebook.com',
            'instagram': 'https://www.instagram.com',
            'linkedin': 'https://www.linkedin.com',
            'amazon': 'https://www.amazon.com',
            'netflix': 'https://www.netflix.com'
        }

        if site in sites:
            try:
                webbrowser.open(sites[site])
                self.speak(f"Opening {site} for you! Here we go~")
            except Exception as e:
                self.speak(f"Oh no! I couldn't open {site}. Sorry~")
        else:
            if not site.startswith('http'):
                site = 'https://' + site
            try:
                webbrowser.open(site)
                self.speak(f"Opening the website! Here we go~")
            except Exception as e:
                self.speak("Hmm, I couldn't open that website. Sorry~")

    def get_time(self):
        now = datetime.now()
        time_str = now.strftime("%I:%M %p")
        self.speak(f"It's {time_str} right now!")

    def get_date(self):
        now = datetime.now()
        date_str = now.strftime("%B %d, %Y")
        self.speak(f"Today is {date_str}!")

    def create_text_file(self, filename):
        if not filename.endswith('.txt'):
            filename += '.txt'
        filepath = os.path.join(self.desktop_path, filename)
        try:
            with open(filepath, 'w') as f:
                f.write(f"File created by Athena on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            self.speak(f"Yay! Created {filename} on your desktop~")
        except Exception as e:
            self.speak(f"Oopsie! I couldn't create the file. Sorry~")

    def delete_text_file(self, filename):
        if not filename.endswith('.txt'):
            filename += '.txt'
        filepath = os.path.join(self.desktop_path, filename)
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
                self.speak(f"Done! Deleted {filename} from your desktop~")
            else:
                self.speak(f"Hmm, I couldn't find {filename} on your desktop. Are you sure it exists?")
        except Exception as e:
            self.speak(f"Oh no! I couldn't delete the file. Sorry~")

    def list_desktop_files(self):
        try:
            files = [f for f in os.listdir(self.desktop_path) if os.path.isfile(os.path.join(self.desktop_path, f))]
            if files:
                file_count = len(files)
                self.speak(f"You have {file_count} files on your desktop. Let me list them for you!")
                for i, file in enumerate(files[:10], 1):
                    print(f"{i}. {file}")
                    if self.gui_callback:
                        self.gui_callback(f"{i}. {file}")
                if len(files) > 10:
                    self.speak(f"And {len(files) - 10} more files!")
            else:
                self.speak("Your desktop is empty! So clean~")
        except Exception as e:
            self.speak("Oh no! I couldn't read your desktop. Sorry~")

    def open_application(self, app_name):
        apps = {
            'paint': 'mspaint.exe',
            'calculator': 'calc.exe',
            'notepad': 'notepad.exe',
            'settings': 'ms-settings:',
            'file explorer': 'explorer.exe',
            'explorer': 'explorer.exe',
            'edge': 'msedge.exe',
            'microsoft edge': 'msedge.exe',
            'chrome': 'chrome.exe',
            'google chrome': 'chrome.exe',
            'word': 'winword.exe',
            'excel': 'excel.exe',
            'powerpoint': 'powerpnt.exe',
            'outlook': 'outlook.exe',
            'command prompt': 'cmd.exe',
            'terminal': 'cmd.exe',
            'task manager': 'taskmgr.exe',
            'copilot': 'microsoft-edge:///?ux=copilot'
        }

        if app_name in apps:
            try:
                if app_name in ['settings', 'copilot']:
                    os.startfile(apps[app_name])
                else:
                    subprocess.Popen(apps[app_name], shell=True)
                self.speak(f"Opening {app_name} for you! Here we go~")
            except Exception as e:
                self.speak(f"Oh no! I couldn't open {app_name}. Sorry~")
        else:
            self.speak(f"Hmm, I don't know how to open {app_name} yet. Maybe teach me later?")

    def shutdown_computer(self):
        self.speak("Are you sure you want to shut down? Say yes to confirm or no to cancel.")
        confirmation = self.listen()
        if confirmation and 'yes' in confirmation:
            self.speak("Okay, shutting down the computer. Goodbye!")
            os.system("shutdown /s /t 1")
        else:
            self.speak("Shutdown cancelled! We can keep working together~")

    def restart_computer(self):
        self.speak("Are you sure you want to restart? Say yes to confirm or no to cancel.")
        confirmation = self.listen()
        if confirmation and 'yes' in confirmation:
            self.speak("Okay, restarting the computer. See you in a moment!")
            os.system("shutdown /r /t 1")
        else:
            self.speak("Restart cancelled! Let's continue~")

    def set_reminder(self, reminder_text):
        self.speak("How many minutes should I remind you in?")
        time_response = self.listen()

        if time_response:
            numbers = re.findall(r'\d+', time_response)
            if numbers:
                minutes = int(numbers[0])
                self.speak(f"Okay! I'll remind you in {minutes} minutes about: {reminder_text}")

                def remind():
                    import time
                    time.sleep(minutes * 60)
                    self.speak(f"Reminder! {reminder_text}")

                threading.Thread(target=remind, daemon=True).start()
            else:
                self.speak("I didn't catch the time. Let's try again later~")

    def process_command(self, command):
        if not command:
            return

        if any(word in command for word in ['exit', 'quit', 'goodbye', 'bye', 'stop']):
            self.speak("Aww, goodbye! Have an amazing day! See you soon~")
            self.running = False
            return

        elif any(phrase in command for phrase in ['what time', 'current time', 'time is it']):
            self.get_time()

        elif any(phrase in command for phrase in ['what date', 'what day', 'today date', 'what is today']):
            self.get_date()

        elif 'youtube' in command:
            if 'open' in command and command.count(' ') <= 2:
                self.open_website('youtube')
            else:
                query = self.extract_query(command, ['youtube', 'search', 'on'])
                if query:
                    self.youtube_search(query)
                else:
                    self.open_website('youtube')

        elif any(word in command for word in ['search', 'google', 'look up', 'find']):
            query = self.extract_query(command, ['search', 'google', 'look up', 'find'])
            if query:
                self.google_search(query)
            else:
                self.speak("What would you like me to search for? I'm ready!")

        elif 'open' in command and any(word in command for word in ['website', 'site', '.com', 'www']):
            site = self.extract_query(command, ['open', 'website', 'site'])
            self.open_website(site)

        elif any(site in command for site in
                 ['gmail', 'github', 'reddit', 'twitter', 'facebook', 'instagram', 'linkedin', 'amazon', 'netflix']):
            for site in ['gmail', 'github', 'reddit', 'twitter', 'facebook', 'instagram', 'linkedin', 'amazon',
                         'netflix']:
                if site in command:
                    self.open_website(site)
                    break

        elif 'create' in command and 'file' in command:
            filename = self.extract_query(command, ['create', 'file', 'named', 'called'])
            if filename:
                self.create_text_file(filename)
            else:
                self.speak("Please tell me what to name the file~")

        elif 'delete' in command and 'file' in command:
            filename = self.extract_query(command, ['delete', 'file', 'named', 'called'])
            if filename:
                self.delete_text_file(filename)
            else:
                self.speak("Which file should I delete? Please tell me~")

        elif any(phrase in command for phrase in ['list files', 'show files', 'what files', 'files on desktop']):
            self.list_desktop_files()

        elif 'open copilot' in command or 'launch copilot' in command:
            self.open_application('copilot')

        elif 'open' in command:
            app = self.extract_query(command, ['open', 'launch', 'start'])
            self.open_application(app)

        elif any(phrase in command for phrase in ['shut down', 'shutdown', 'turn off computer']):
            self.shutdown_computer()

        elif any(phrase in command for phrase in ['restart', 'reboot']):
            self.restart_computer()

        elif 'remind me' in command or 'set reminder' in command:
            reminder = self.extract_query(command, ['remind me', 'set reminder', 'to', 'about'])
            if reminder:
                self.set_reminder(reminder)
            else:
                self.speak("What should I remind you about?")

        elif 'help' in command or 'what can you do' in command or 'commands' in command:
            self.speak("""Yay! Let me tell you what I can do! I can search Google or YouTube, 
                open websites like Gmail, GitHub, Reddit, and more! I can tell you the time and date,
                create and delete text files, list files on your desktop,
                open Paint, Calculator, Notepad, Settings, File Explorer, and other apps! 
                I can even set reminders, shut down or restart your computer!
                Plus, I can answer any questions using Windows Copilot! Just ask me anything!
                Just talk to me naturally! I'm here to help you!""")

        else:
            # Use Windows Copilot for all other queries
            self.speak("Let me ask Copilot about that...")
            self.ask_copilot(command)


class AthenaGUI(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = 10
        self.spacing = 10

        # Set window background color
        Window.clearcolor = (0.95, 0.95, 0.97, 1)

        # Title
        title = Label(
            text='ATHENA - AI Assistant',
            size_hint=(1, 0.1),
            font_size='24sp',
            bold=True,
            color=(0.2, 0.3, 0.5, 1)
        )
        self.add_widget(title)

        # Chat display
        self.chat_display = TextInput(
            text='Welcome to Athena! 🎀\n\n',
            size_hint=(1, 0.5),
            readonly=True,
            multiline=True,
            background_color=(1, 1, 1, 1),
            foreground_color=(0, 0, 0, 1),
            font_size='14sp'
        )
        self.add_widget(self.chat_display)

        # Text input for manual commands
        input_layout = BoxLayout(size_hint=(1, 0.1), spacing=5)
        self.text_input = TextInput(
            hint_text='Type your command here...',
            size_hint=(0.8, 1),
            multiline=False,
            font_size='14sp'
        )
        self.text_input.bind(on_text_validate=self.send_text_command)
        input_layout.add_widget(self.text_input)

        send_btn = Button(
            text='Send',
            size_hint=(0.2, 1),
            background_color=(0.3, 0.6, 0.9, 1),
            font_size='14sp'
        )
        send_btn.bind(on_press=self.send_text_command)
        input_layout.add_widget(send_btn)
        self.add_widget(input_layout)

        # Button panel
        button_layout = GridLayout(cols=3, size_hint=(1, 0.3), spacing=5)

        buttons = [
            ('🎤 Voice Command', self.start_voice_command),
            ('⏰ Time', self.get_time),
            ('📅 Date', self.get_date),
            ('🌐 Open Copilot', self.open_copilot),
            ('📂 List Files', self.list_files),
            ('❓ Help', self.show_help),
        ]

        for text, callback in buttons:
            btn = Button(
                text=text,
                background_color=(0.4, 0.7, 0.95, 1),
                font_size='13sp'
            )
            btn.bind(on_press=callback)
            button_layout.add_widget(btn)

        self.add_widget(button_layout)

        # Initialize Athena
        self.athena = Athena(gui_callback=self.update_chat)
        self.update_chat("Athena initialized! Ready to assist you! 🎀")

    def update_chat(self, message):
        def update(dt):
            self.chat_display.text += message + '\n'
            self.chat_display.cursor = (0, len(self.chat_display.text))

        Clock.schedule_once(update, 0)

    def send_text_command(self, instance):
        command = self.text_input.text.strip()
        if command:
            self.update_chat(f"\nYou: {command}")
            self.text_input.text = ''
            # Process command in a separate thread
            threading.Thread(target=self.athena.process_command, args=(command.lower(),), daemon=True).start()

    def start_voice_command(self, instance):
        self.update_chat("\n🎤 Starting voice command...")
        threading.Thread(target=self._voice_command_thread, daemon=True).start()

    def _voice_command_thread(self):
        command = self.athena.listen()
        if command:
            self.athena.process_command(command)

    def get_time(self, instance):
        threading.Thread(target=self.athena.get_time, daemon=True).start()

    def get_date(self, instance):
        threading.Thread(target=self.athena.get_date, daemon=True).start()

    def open_copilot(self, instance):
        threading.Thread(target=self.athena.open_application, args=('copilot',), daemon=True).start()

    def list_files(self, instance):
        threading.Thread(target=self.athena.list_desktop_files, daemon=True).start()

    def show_help(self, instance):
        threading.Thread(target=self.athena.process_command, args=('help',), daemon=True).start()


class AthenaApp(App):
    def build(self):
        Window.size = (800, 600)
        return AthenaGUI()


if __name__ == "__main__":
    print("=" * 50)
    print("ATHENA - AI Assistant with Windows Copilot")
    print("=" * 50)
    print("\nStarting GUI...")
    print("\nCommands you can use:")
    print("- Voice or text commands")
    print("- 'search [query]' - Search Google")
    print("- 'youtube [query]' - Search YouTube")
    print("- 'open [website]' - Open websites")
    print("- 'what time is it' - Get current time")
    print("- 'create file [name]' - Create text file")
    print("- 'open copilot' - Open Windows Copilot")
    print("- 'help' - Show all commands")
    print("- Any question will open Copilot!")
    print("\n" + "=" * 50 + "\n")

    AthenaApp().run()