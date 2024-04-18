import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import tkinter.ttk as ttk
import logging
import ssl
from imapclient import IMAPClient
import mailbox

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def test_connection(server_address, email, password, log_area):
    try:
        # Create IMAPClient connection with SSL context
        context = ssl.create_default_context()
        with IMAPClient(server_address, ssl_context=context) as client:
            client.login(email, password)
            # If connection is successful, log it
            log_area.insert(tk.END, "IMAP connection successful.\n")
    except Exception as e:
        log_area.insert(tk.END, f"An error occurred: {e}\n")

def migrate_emails(mbox_file_path, dest_server, dest_email, dest_password, log_area, progress_bar):
    try:
        # Connect to destination server
        context = ssl.create_default_context()
        with IMAPClient(dest_server, ssl_context=context) as client:
            client.login(dest_email, dest_password)

            # Open the mbox file
            mbox = mailbox.mbox(mbox_file_path)

            # Initialize progress
            total_messages = len(mbox)
            progress = 0

            # Iterate over each email in the mbox file
            for message in mbox:
                # Append the message to the destination mailbox
                client.append("INBOX", message.as_bytes())
                
                # Update progress
                progress += 1
                progress_percent = (progress / total_messages) * 100
                progress_bar["value"] = progress_percent
                progress_bar.update()

            log_area.insert(tk.END, f"{total_messages} emails migrated successfully.\n")
    except Exception as e:
        logging.error("An error occurred: %s", str(e))
        log_area.insert(tk.END, f"An error occurred: {e}\n")

def browse_mbox_file(entry_widget, email_tree):
    mbox_file_path = filedialog.askopenfilename(filetypes=[("Mbox Files", "*.mbox")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, mbox_file_path)
    visualize_mbox_content(mbox_file_path, email_tree)

def visualize_mbox_content(mbox_file_path, email_tree):
    try:
        mbox = mailbox.mbox(mbox_file_path)
        email_tree.delete(*email_tree.get_children())  # Clear previous content

        folders = {}
        for message in mbox:
            from_address = message['From']
            folder_name = from_address.split('@')[0]  # Extract the folder name from the email address
            if folder_name not in folders:
                folders[folder_name] = []
            folders[folder_name].append(message)

        for folder_name, emails in folders.items():
            folder_id = email_tree.insert("", tk.END, text=folder_name, open=True)
            for email in emails:
                email_tree.insert(folder_id, tk.END, text=email['Subject'])

    except Exception as e:
        logging.error("An error occurred while visualizing mbox content: %s", str(e))

def create_gui():
    root = tk.Tk()
    root.title("Email Migration Tool | Lord Amdal")

    frame = tk.Frame(root)
    frame.pack(padx=20, pady=10)

    lbl_mbox = tk.Label(frame, text="Mbox File:")
    lbl_mbox.grid(row=0, column=0, sticky="e")

    entry_mbox = tk.Entry(frame, width=50)
    entry_mbox.grid(row=0, column=1, padx=5, pady=5)

    btn_browse = tk.Button(frame, text="Browse", command=lambda: browse_mbox_file(entry_mbox, email_tree))
    btn_browse.grid(row=0, column=2, padx=5, pady=5)

    lbl_server = tk.Label(frame, text="Destination Server IMAP:")
    lbl_server.grid(row=1, column=0, sticky="e")

    entry_server = tk.Entry(frame, width=50)
    entry_server.grid(row=1, column=1, padx=5, pady=5)

    lbl_email = tk.Label(frame, text="Destination Email:")
    lbl_email.grid(row=2, column=0, sticky="e")

    entry_email = tk.Entry(frame, width=50)
    entry_email.grid(row=2, column=1, padx=5, pady=5)

    lbl_password = tk.Label(frame, text="Destination Password:")
    lbl_password.grid(row=3, column=0, sticky="e")

    entry_password = tk.Entry(frame, width=50, show="*")
    entry_password.grid(row=3, column=1, padx=5, pady=5)

    btn_start_migration = tk.Button(frame, text="Start Migration", command=lambda: start_migration(entry_mbox.get(), entry_server.get(), entry_email.get(), entry_password.get(), log_area, progress_bar))
    btn_start_migration.grid(row=4, column=1, padx=5, pady=10)

    log_frame = tk.LabelFrame(root, text="Log")
    log_frame.pack(padx=20, pady=10, fill="both", expand=True)

    log_area = scrolledtext.ScrolledText(log_frame, width=60, height=10)
    log_area.pack(fill="both", expand=True)

    # Progress bar
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
    progress_bar.pack(pady=10)

    email_frame = tk.LabelFrame(root, text="Mbox Content")
    email_frame.pack(padx=20, pady=10, fill="both", expand=True)

    email_tree = ttk.Treeview(email_frame)
    email_tree.pack(fill="both", expand=True)

    btn_test_conn = tk.Button(root, text="Test IMAP Connection", command=lambda: test_connection(entry_server.get(), entry_email.get(), entry_password.get(), log_area))
    btn_test_conn.pack(pady=5)

    root.mainloop()

def start_migration(mbox_file_path, dest_server, dest_email, dest_password, log_area, progress_bar):
    if mbox_file_path == "":
        messagebox.showerror("Error", "Please select an mbox file.")
        return
    if dest_server == "" or dest_email == "" or dest_password == "":
        messagebox.showerror("Error", "Please enter destination server address, email address, and password.")
        return

    migrate_emails(mbox_file_path, dest_server, dest_email, dest_password, log_area, progress_bar)

if __name__ == "__main__":
    create_gui()
