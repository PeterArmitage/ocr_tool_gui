import tkinter as tk

print("Creating window...")
root = tk.Tk()
root.title("Test Window")
root.geometry("300x200")

label = tk.Label(root, text="If you can see this, tkinter works!")
label.pack(pady=50)

print("Window created. Starting main loop...")
root.mainloop()
print("Window closed.")