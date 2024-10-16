import tkinter as tk
from tkinter import ttk
import subprocess

def launch_student_data():
    subprocess.Popen(['python', 'StudentData.py'])

def launch_activities():
    subprocess.Popen(['python', 'Activities.py'])

def launch_merit_demerit():
    subprocess.Popen(['python', 'MeritDemerit.py'])

def launch_rankings():
    subprocess.Popen(['python', 'rankings.py'])

def toggle_mode():
    # Toggle between light and dark themes
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
        root.configure(bg="#FFFFFF")  # Change root background to light
        btn_rankings.configure(style="TButton")
        mode_switch.configure(style="TCheckbutton")
    else:
        style.theme_use("forest-dark")
        root.configure(bg="#2E2E2E")  # Change root background to dark
        btn_rankings.configure(style="Dark.TButton")
        mode_switch.configure(style="Dark.TCheckbutton")

# Create the main window
root = tk.Tk()
root.title("House System!")
root.geometry("450x320")  # Adjusted height for additional button

# Set up style and theme switching
style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")  # Default to dark mode

# Create the main frame
frame = ttk.Frame(root)
frame.pack(padx=20, pady=20)

# Create buttons for each project
btn_student_data = ttk.Button(frame, text="Student Distributer", command=launch_student_data, cursor="hand2")
btn_student_data.pack(pady=10, fill="x")

btn_activities = ttk.Button(frame, text="Activities/House Points", command=launch_activities, cursor="hand2")
btn_activities.pack(pady=10, fill="x")

btn_merit_demerit = ttk.Button(frame, text="Merit/Demerit Points", command=launch_merit_demerit, cursor="hand2")
btn_merit_demerit.pack(pady=10, fill="x")

# New button for Rankings
btn_rankings = ttk.Button(frame, text="House Rankings", command=launch_rankings, cursor="hand2")
btn_rankings.pack(pady=10, fill="x")

# Separator for toggle switch
separator = ttk.Separator(frame)
separator.pack(fill="x", pady=10)

# Mode switch (Light/Dark)
mode_switch = ttk.Checkbutton(frame, text="Toggle Light/Dark Mode", style="Switch", command=toggle_mode)
mode_switch.pack(pady=10, fill="x")

# Run the main loop
root.mainloop()
