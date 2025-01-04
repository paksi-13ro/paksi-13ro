import os
import tkinter as tk
from tkinter import filedialog
import pygame
import win32com.client


class MediaPlayer:
    def __init__(self, master):
        self.master = master
        self.master.title("Media Player")
        self.master.geometry("400x200")

        self.playlist = []
        self.current_index = 0

        self.label = tk.Label(master, text="Media Player", font=("Arial", 16))
        self.label.pack()

        self.play_button = tk.Button(master, text="Play", command=self.play)
        self.play_button.pack(side=tk.LEFT)

        self.pause_button = tk.Button(master, text="Pause", command=self.pause)
        self.pause_button.pack(side=tk.LEFT)

        self.stop_button = tk.Button(master, text="Stop", command=self.stop)
        self.stop_button.pack(side=tk.LEFT)

        self.next_button = tk.Button(master, text="Next", command=self.next_media)
        self.next_button.pack(side=tk.LEFT)

        self.prev_button = tk.Button(master, text="Previous", command=self.prev_media)
        self.prev_button.pack(side=tk.LEFT)

        self.load_button = tk.Button(master, text="Load Shortcuts", command=self.load_shortcuts)
        self.load_button.pack(side=tk.LEFT)

        pygame.mixer.init()

    def load_shortcuts(self):
        shortcuts = filedialog.askopenfilenames(title="Select Shortcut Files", filetypes=[("Shortcut Files", "*.lnk")])
        self.playlist = [self.resolve_shortcut(lnk) for lnk in shortcuts]
        self.current_index = 0
        print("Loaded:", self.playlist)

    def resolve_shortcut(self, lnk_path):
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(lnk_path)
        return shortcut.Targetpath

    def play(self):
        if self.playlist:
            pygame.mixer.music.load(self.playlist[self.current_index])
            pygame.mixer.music.play()

    def pause(self):
        pygame.mixer.music.pause()

    def stop(self):
        pygame.mixer.music.stop()

    def next_media(self):
        self.current_index = (self.current_index + 1) % len(self.playlist)
        self.play()

    def prev_media(self):
        self.current_index = (self.current_index - 1) % len(self.playlist)
        self.play()


if __name__ == "__main__":
    root = tk.Tk()
    media_player = MediaPlayer(root)
    root.mainloop()
