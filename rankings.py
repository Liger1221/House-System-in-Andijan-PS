import tkinter as tk
from tkinter import ttk
import openpyxl
from tkinter import messagebox

def load_house_points():
    path = "Houses.xlsx"
    
    try:
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active

        house_points = {
            "Blue House": sheet["A2"].value,
            "Red House": sheet["B2"].value,
            "Green House": sheet["C2"].value,
            "White House": sheet["D2"].value
        }

        return house_points
    except FileNotFoundError:
        messagebox.showerror("Error", f"File '{path}' not found.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def rank_houses(house_points):
    # Sort the houses by points in descending order
    ranked_houses = sorted(house_points.items(), key=lambda x: x[1], reverse=True)
    return ranked_houses

def display_rankings():
    house_points = load_house_points()

    if house_points:
        ranked_houses = rank_houses(house_points)
        
        # Clear the Treeview before inserting new data
        treeview.delete(*treeview.get_children())

        # Display the house rankings with points in quotes
        for i, (house, points) in enumerate(ranked_houses, start=1):
            treeview.insert('', tk.END, values=(i, f'{house} - "{points}"'))

def toggle_mode():
    # Toggle between light and dark themes
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
        root.configure(bg="#FFFFFF")  # Change root background to light
        show_rankings_button.configure(style="TButton")
        mode_switch.configure(style="TCheckbutton")
    else:
        style.theme_use("forest-dark")
        root.configure(bg="#2E2E2E")  # Change root background to dark
        show_rankings_button.configure(style="Dark.TButton")
        mode_switch.configure(style="Dark.TCheckbutton")

# Tkinter setup
root = tk.Tk()
root.title("House Rankings Podium")

# Set a fixed window size to prevent excessive resizing
root.geometry("400x350")
root.minsize(400, 350)

style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

# Set initial background color
root.configure(bg="#2E2E2E")  # Dark mode background

# Frame for displaying rankings
tree_frame = ttk.LabelFrame(root, text="House Podium Rankings", padding=10)
tree_frame.pack(padx=20, pady=20, fill="both", expand=True)

# Create a Treeview for ranking display
cols = ("Rank", "House and Points")
treeview = ttk.Treeview(tree_frame, columns=cols, show="headings", height=4)  # Set height to 4 to make it vertically smaller
treeview.pack(fill="both", expand=True)

# Configure column headings
for col in cols:
    treeview.heading(col, text=col)
    treeview.column(col, width=150)

# Button to display the rankings
show_rankings_button = ttk.Button(root, text="Show Rankings", command=display_rankings)
show_rankings_button.pack(pady=10)

# Toggle mode switch for dark/light mode
mode_switch = ttk.Checkbutton(root, text="Toggle Mode", style="Dark.TCheckbutton", command=toggle_mode)
mode_switch.pack(pady=5)

# Run the application
root.mainloop()
