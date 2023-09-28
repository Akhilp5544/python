import email
from email.header import decode_header
import os
import zipfile
import imaplib
import pandas as pd

class EmailProcessor:
    def __init__(self, email_address, password, imap_server, imap_port, subject_to_search):
        self.email_address = email_address
        self.password = password
        self.imap_server = imap_server
        self.imap_port = imap_port
        self.subject_to_search = subject_to_search

    def connect_to_email_account(self):
        # Connect to the IMAP server
        self.mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)

        if self.mail:
            # Login to the email account
            login_result, login_message = self.mail.login(self.email_address, self.password)

            if login_result == 'OK':
                print("Successfully logged in to the email account.")
                return True
            else:
                print(f"Error logging in to the email account: {login_result} - {login_message}")
                return False
        else:
            print(f"Error connecting to the IMAP server")
            return False

    def search_and_process_emails(self):
        if not self.connect_to_email_account():
            return

        # Select the mailbox (in this case, the inbox)
        mailbox = 'INBOX'
        self.mail.select(mailbox)

        # Define search criteria to find emails with a specific subject
        search_criteria = f'(SUBJECT "{self.subject_to_search}")'

        # Search for emails that match the criteria
        search_result, email_ids = self.mail.search(None, search_criteria)

        if search_result == 'OK':
            # Check if there is at least one email
            email_id_list = email_ids[0].split()
            if len(email_id_list) > 0:
                # Create a folder to save the zip attachments
                desktop_path = os.path.expanduser("~/Desktop")
                folder_name = "RazorPay_Settlement_attachments"
                folder_path = os.path.join(desktop_path, folder_name)

                # Create the folder if it doesn't exist
                if not os.path.exists(folder_path):
                    os.makedirs(folder_path)

                # Initialize an empty list to store DataFrames
                dataframes = []

                # Iterate through email IDs and process each email
                for email_id in email_id_list:
                    fetch_result, email_data = self.mail.fetch(email_id, '(RFC822)')

                    if fetch_result == 'OK':
                        # 'email_data' contains the email content
                        email_message = email_data[0][1]

                        # Parse the email message
                        msg = email.message_from_bytes(email_message)

                        # Decode the subject if it's encoded
                        subject, encoding = decode_header(msg['Subject'])[0]
                        if isinstance(subject, bytes):
                            subject = subject.decode(encoding or 'utf-8')
                            print(subject)

                        # Print the subject
                        print(f"Subject: {subject}")
                        zip_counter = 1  # Reset the counter for each email
                        zip_files = []  # Store the ZIP file paths

                        for part in msg.walk():
                            content_disposition = str(part.get("Content-Disposition"))

                            if "attachment" in content_disposition:
                                filename = part.get_filename()
                                if filename:
                                    # Check if the attachment is a zip file
                                    if filename.lower().endswith(".zip"):
                                        attachment_data = part.get_payload(decode=True)

                                        if attachment_data:
                                            # Generate a unique filename for each attachment
                                            unique_filename = f"email_{email_id}_attachment_{zip_counter}.zip"
                                            attachment_path = os.path.join(folder_path, unique_filename)

                                            with open(attachment_path, 'wb') as attachment_file:
                                                attachment_file.write(attachment_data)
                                            print(f"Saved ZIP attachment to folder: {attachment_path}")

                                            zip_files.append(attachment_path)
                                            zip_counter += 1  # Increment the counter for attachments in this email

                # Now, loop through the extracted ZIP files and process them one by one
                for zip_file in zip_files:
                    print(zip_file)

                    try:
                        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                            zipinfos = zip_ref.infolist()
                            for zipinfo in zipinfos:
                                filename = zipinfo.filename
                                if filename.lower().endswith(".csv"):
                                    # Read the CSV file into a DataFrame
                                    with zip_ref.open(filename) as csv_file:
                                        df = pd.read_csv(csv_file)
                                    dataframes.append(df)  # Append the DataFrame to the list
                                    print(f"Read CSV file '{filename}' from {zip_file}.")

                            # Extracted contents of the ZIP file and renamed files, no need to rename them again
                            print(f"Extracted contents of {zip_file} to folder.")
                    except Exception as e:
                        print(f"Error extracting {zip_file}: {str(e)}")

                # Check if any CSV data was found
                if dataframes:
                    # Concatenate all DataFrames into one
                    combined_df = pd.concat(dataframes, ignore_index=True)
                    # Store the combined DataFrame in the instance variable
                    self.final_dataframe = combined_df


                    # Define the path for the combined CSV file

                    combined_csv_path = os.path.join(folder_path, 'combined_data.csv')

                    # Save the concatenated DataFrame to a CSV file
                    combined_df.to_csv(combined_csv_path, index=False)

                    print(f"Combined all CSV data into '{combined_csv_path}'")
                else:
                    print("No CSV data found in the extracted ZIP files.")
            else:
                print(f"No emails found with the subject '{self.subject_to_search}' in the inbox.")
        else:
            print(f"Error searching for emails: {search_result}")

    def get_combined_dataframe(self):
        return self.final_dataframe


if __name__ == "__main__":
    # email_address = input('Enter the email_address: ')
    email_address = 'akhilp@algofusiontech.com'
    # password = input('Enter the password: ')
    password = 'foqf ospb dzwl jasn'
    # imap_server = input('Enter the imap_server_number: ')
    imap_server = 'imap.gmail.com'
    # imap_port = input('Enter the imap_port_number: ')
    imap_port = '993'
    # subject_to_search = input('Enter the subject_to_search: ')
    subject_to_search = 'Fwd: FW: RazorPay Settlement report'

    # Create an instance of the EmailProcessor class
    email_processor = EmailProcessor(email_address, password, imap_server, imap_port, subject_to_search)

    # Call the method to search and process emails
    email_processor.search_and_process_emails()







