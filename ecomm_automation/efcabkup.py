import tkinter as tk
from tkinter import Canvas, Button, PhotoImage
import os

# Define parent directory
parent_dir = 'ecomm_automation'

# Define subdirectories
frame0 = os.path.join(parent_dir, 'assets', 'frame0')

def create_interface():
    window = tk.Tk()
    window.geometry("700x550")
    window.configure(bg="#FFFFFF")

    canvas = Canvas(
        window,
        bg="#FFFFFF",
        height=550,
        width=700,
        bd=0,
        highlightthickness=0,
        relief="ridge"
    )
    canvas.place(x=0, y=0)

    # Load and display the background image
    image_image_1 = PhotoImage(file=os.path.join(frame0, "image_1.png"))
    image_1 = canvas.create_image(350.0, 275.0, image=image_image_1)

    # Header
    canvas.create_rectangle(0.0, 0.0, 700.0, 71.0, fill="#6EC8C3", outline="")
    canvas.create_text(103.0, 13.0, anchor="nw", text="ECOMMS AUTOMATION SYSTEM",
                       fill="#1D4916", font=("Inter Bold", 30 * -1))

    # Footer
    canvas.create_rectangle(0.0, 497.0, 700.0, 550.0, fill="#6EC8C3", outline="")

    # Buttons
    button_image_1 = PhotoImage(file=os.path.join(frame0, "button_1.png"))
    button_1 = Button(image=button_image_1, borderwidth=0, highlightthickness=0,
                      command=lambda: print("button_1 clicked"), relief="flat")
    button_1.place(x=609.0, y=504.0, width=62.0, height=40.0)

    button_image_2 = PhotoImage(file=os.path.join(frame0, "button_2.png"))
    button_2 = Button(image=button_image_2, borderwidth=0, highlightthickness=0,
                      command=lambda: print("button_2 clicked"), relief="flat")
    button_2.place(x=435.0, y=504.0, width=87.0, height=40.0)

    # Sidebar (you can add your sidebar widgets here)

    # Prevent window resizing
    window.resizable(False, False)
    window.mainloop()

create_interface()
