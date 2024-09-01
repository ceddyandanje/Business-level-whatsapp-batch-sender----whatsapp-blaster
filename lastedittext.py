import pandas as pd
import pywhatkit as kit
import time
import tkinter as tk
from tkinter import filedialog, simpledialog, scrolledtext, messagebox
import webbrowser
import random

# Function to open file explorer and select the Excel or CSV file
def get_file_path():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(title="Select the Excel or CSV file", filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
    return file_path

# Function to open a larger input box for custom messages
def get_custom_message():
    def submit_message():
        global custom_message
        custom_message = text_box.get("1.0", tk.END).strip()
        root.quit()  # Close the input window
    
    root = tk.Tk()
    root.title("Custom Message Input")
    
    label = tk.Label(root, text="Please input your custom message below:", padx=10, pady=10)
    label.pack()

    # Use a ScrolledText widget for a larger, scrollable input box
    text_box = scrolledtext.ScrolledText(root, width=60, height=15, wrap=tk.WORD)
    text_box.pack(padx=10, pady=10)

    # Add a button to submit the message
    submit_button = tk.Button(root, text="Submit", command=submit_message)
    submit_button.pack(pady=10)
    
    root.mainloop()
    
    return custom_message

# Function to choose between pasting a custom message or selecting a text file
def choose_message_input_method():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    input_method = messagebox.askquestion("Message Input Method", "Would you like to paste the message directly?", icon='question')

    if input_method == 'yes':
        return get_custom_message()
    else:
        file_path = filedialog.askopenfilename(title="Select a Text File", filetypes=[("Text files", "*.txt")])
        with open(file_path, 'r') as file:
            return file.read().strip()

# Main execution
if __name__ == "__main__":
    # Prompt user for file location
    print("Please select the Excel or CSV file containing the business names and phone numbers.")
    file_path = get_file_path()

    # Load the Excel/CSV file containing the business names and phone numbers
    if file_path.endswith('.xlsx'):
        df = pd.read_excel(file_path)
    else:
        df = pd.read_csv(file_path)

    # Get the custom message (either from direct input or from a
# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    business_name = row['Business Name']
    phone_number = str(row['Phone 1'])

    # Check if phone number starts with '7' and add '254' if necessary
    if phone_number.startswith('7'):
        phone_number = '254' + phone_number
    
    # Format the WhatsApp URL with the custom message
    custom_message = choose_message_input_method()
    personalized_message = f"Hello {business_name},\n{custom_message}"
    url = f"https://web.whatsapp.com/send?phone={phone_number}&text={personalized_message}"
    
    # Open the WhatsApp Web tab and wait for it to load
    webbrowser.open(url)
    time.sleep(random.uniform(10, 15))  # Random wait time between 10 to 15 seconds
    
    # Simulate pressing 'Enter' to send the message
    kit.sendwhatmsg_instantly(phone_number, personalized_message, 10, True, 3)

    # Random delay before sending the next message (15 to 35 seconds)
    time.sleep(random.uniform(15, 35))

    # Optional: Add a longer break after every 5 messages
    if (index + 1) % 5 == 0:
        time.sleep(random.uniform(300, 600))  # 5 to 10 minutes break

print("Messages sent successfully.")
