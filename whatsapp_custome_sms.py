import pandas as pd
import pywhatkit as kit
import logging

# Configure logging
logging.basicConfig(
    filename="whatsapp_messages.log",
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def send_whatsapp_messages():
    # Load the Excel file and preserve the + sign in phone numbers
    excel_file = "contacts.xlsx"
    try:
        logging.info("Attempting to load Excel file.")
        contacts = pd.read_excel(excel_file, dtype={'PhoneNumber': str})
        logging.info("Excel file loaded successfully.")
    except Exception as e:
        logging.error(f"Error loading Excel file: {e}")
        return
    # Iterate through each contact
    for index, row in contacts.iterrows():
        phone_number = row['PhoneNumber']  # Already a string due to dtype
        message = row['Message']
        # Debugging: Check if phone number has + sign
        logging.debug(f"Read phone number: {phone_number}")

        # Ensure phone number is valid
        if not phone_number.startswith('+') or not phone_number[1:].isdigit():
            logging.warning(f"Invalid phone number: {phone_number}")
            continue
        try:
            # Send the WhatsApp message
            logging.info(f"Sending message to {phone_number}: {message}")
            kit.sendwhatmsg_instantly(phone_number, message, wait_time=15, tab_close=True)
            logging.info(f"Message sent to {phone_number}")
        except Exception as e:
            logging.error(f"Error sending message to {phone_number}: {e}")

if __name__ == "__main__":
    send_whatsapp_messages()