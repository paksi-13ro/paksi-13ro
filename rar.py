import os
import tkinter as tk
from tkinter import PhotoImage, messagebox
from PIL import Image, ImageTk
import pygame
import pythoncom
import win32com.client


class MediaViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Media Viewer")

        # Инициализация Pygame
        pygame.init()

        # Переменные
        self.media_files = []
        self.current_index = 0

        # UI компоненты
        self.label = tk.Label(root)
        self.label.pack()

        self.play_button = tk.Button(root, text="Play", command=self.play_media)
        self.play_button.pack(side=tk.LEFT)

        self.pause_button = tk.Button(root, text="Pause", command=self.pause_media)
        self.pause_button.pack(side=tk.LEFT)

        self.stop_button = tk.Button(root, text="Stop", command=self.stop_media)
        self.stop_button.pack(side=tk.LEFT)

        self.prev_button = tk.Button(root, text="Previous", command=self.prev_media)
        self.prev_button.pack(side=tk.LEFT)

        self.next_button = tk.Button(root, text="Next", command=self.next_media)
        self.next_button.pack(side=tk.LEFT)

    def load_shortcuts(self, shortcut_paths):
        self.media_files = [self.resolve_shortcut(path) for path in shortcut_paths]
        self.show_media()

    def resolve_shortcut(self, shortcut_path):
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        return shortcut.TargetPath

    def show_media(self):
        if not self.media_files:
            messagebox.showerror("Error", "No media files loaded.")
            return

        media_file = self.media_files[self.current_index]
        ext = os.path.splitext(media_file)[1].lower()

        if ext in [".jpg", ".jpeg", ".png", ".gif"]:
            self.show_image(media_file)
        elif ext in [".mp4", ".avi", ".mov"]:
            self.show_video(media_file)
        elif ext in [".mp3", ".wav"]:
            self.show_audio(media_file)

    def show_image(self, filepath):
        img = Image.open(filepath)
        img = img.resize((800, 600), Image.ANTIALIAS)
        img_tk = ImageTk.PhotoImage(img)
        self.label.config(image=img_tk)
        self.label.image = img_tk

    def show_video(self, filepath):
        self.stop_media()  # Остановить предыдущее воспроизведение
        pygame.display.set_mode((800, 600))
        pygame.mixer.music.load(filepath)
        self.label.config(text="Playing video...")
        pygame.mixer.music.play()

    def show_audio(self, filepath):
        self.stop_media()  # Остановить предыдущее воспроизведение
        pygame.mixer.music.load(filepath)
        self.label.config(text="Playing audio...")
        pygame.mixer.music.play()

    def play_media(self):
        pygame.mixer.music.unpause()

    def pause_media(self):
        pygame.mixer.music.pause()

    def stop_media(self):
        pygame.mixer.music.stop()
        self.label.config(text="")

    def prev_media(self):
        self.current_index = (self.current_index - 1) % len(self.media_files)
        self.show_media()

    def next_media(self):
        self.current_index = (self.current_index + 1) % len(self.media_files)
        self.show_media()


if __name__ == "__main__":
    root = tk.Tk()
    viewer = MediaViewer(root)

    # Пример загрузки ярлыков
    shortcuts = ["C:\\path\\to\\your\\shortcut1.lnk", "C:\\path\\to\\your\\shortcut2.lnk"]
    viewer.load_shortcuts(shortcuts)

    root.mainloop()
