import pandas as pd
import pywhatkit as kit
import time
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import webbrowser

stop_sending = False

def send_messages():
    global df, message_entry, stop_sending
    try:
        message = message_entry.get("1.0", tk.END).strip()
        phone_numbers = df['Phone Number'].astype(str)
        for phone_number in phone_numbers:
            if stop_sending:
                break
            sent = False
            attempts = 0
            while not sent and attempts < 3:
                if stop_sending:
                    break
                try:
                    kit.sendwhatmsg_instantly(f"+{phone_number}", message)
                    time.sleep(10)  # Wait to ensure message has time to be sent
                    sent = True
                except Exception as e:
                    attempts += 1
                    time.sleep(10)  # Wait before retrying
                    if attempts == 3:
                        messagebox.showerror("Error", f"Failed to send message to {phone_number}: {e}")
                        break
                time.sleep(5)  # Wait before sending the next message
        if not stop_sending:
            messagebox.showinfo("Success", "Messages sent successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def start_sending_messages():
    global stop_sending
    stop_sending = False
    thread = threading.Thread(target=send_messages)
    thread.start()

def select_file():
    global df
    file_path = filedialog.askopenfilename()
    if file_path:
        df = pd.read_excel(file_path)
        messagebox.showinfo("Success", "File loaded successfully!")

def on_closing():
    global stop_sending
    stop_sending = True
    if 'thread' in globals():
        thread.join()
    root.destroy()

def open_facebook_link(event):
    webbrowser.open_new("https://www.facebook.com/hamzahamedi123")

def open_whatsapp():
    kit.open_web()

# Create the GUI
root = tk.Tk()
root.title("WhatsApp Bulk Messenger")
root.geometry("500x500")
root.configure(bg="#ADD8E6")

# Message entry
message_label = tk.Label(root, text="Enter message:", font=('Arial', 14), bg="#ADD8E6")
message_label.pack(pady=10)
message_entry = tk.Text(root, width=50, height=10, font=('Arial', 14))
message_entry.pack(pady=10)

# Frame for buttons
button_frame = tk.Frame(root, bg="#ADD8E6")
button_frame.pack(pady=20)

# Load file button
load_button = tk.Button(button_frame, text="Load Excel File", command=select_file, font=('Arial', 12), bg="#4CAF50", fg="white", width=15)
load_button.pack(side=tk.LEFT, padx=10)

# Send messages button
send_button = tk.Button(button_frame, text="Send Messages", command=start_sending_messages, font=('Arial', 12), bg="#008CBA", fg="white", width=15)
send_button.pack(side=tk.LEFT, padx=10)

# Contact Information
contact_frame = tk.Frame(root, bg="#ADD8E6")
contact_frame.pack(pady=20)

contact_label = tk.Label(contact_frame, text="Contact Information", font=('Arial', 12, 'bold'), bg="#ADD8E6")
contact_label.pack(pady=5)

facebook_label = tk.Label(contact_frame, text="Facebook: https://www.facebook.com/hamzahamedi123", font=('Arial', 12), bg="#ADD8E6", fg="blue", cursor="hand2")
facebook_label.pack(pady=5)
facebook_label.bind("<Button-1>", open_facebook_link)

phone_label = tk.Label(contact_frame, text="Phone: +21694694875", font=('Arial', 12), bg="#ADD8E6")
phone_label.pack(pady=5)

# Set the closing protocol
root.protocol("WM_DELETE_WINDOW", on_closing)

# Open WhatsApp before starting the loop and wait for 5 seconds
open_whatsapp()
time.sleep(5)

root.mainloop()
