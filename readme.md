# Call Center MC Automation

This project automates the process of searching for company details on the SAFER website and sending emails to the companies. It uses Selenium for web automation, OpenPyXL for Excel integration, and PyYAML for reading email templates.

## Installation

1. Clone the repository to your local machine.
2. Navigate to the project directory.
3. Install the required Python packages using `requirements.txt`.

```pip install -r requirements.txt ```


> Configuration
1. Create a .env.local file in the project directory with the following content:

# MCMX is the Range for Company search on safer website
MCMX_START=1
MCMX_END=10

# Email Details:
ENABLE_EMAIL_SENDING="true"
EMAIL_LOGIN_URL="https://mail.hostinger.com/"
EMAIL_USERNAME="info@jslogisticsolutions.com"
EMAIL_PASSWORD="J&S.js$$$$03456"

2. Update the email_data.yaml file with your email template.


Usage
Run the safer_v1.1.py script to start the automation process.

python safer_v1.1.py