from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF
import os
import pandas as pd
import argparse
import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# ---------------------------------------------------------------
#                               CONSTANTS
# ---------------------------------------------------------------

# Email
EMAIL_HOST = "smtp.gmail.com"
EMAIL_PORT = 587
EMAIL_ADDRESS = "computersociety@itba.edu.ar"
EMAIL_FROM = "Computer Society ITBA<" + EMAIL_ADDRESS + ">"
EMAIL_PASS_VAR = "EMAIL_PASS"
EMAIL_SUBJECT = "Certificado de Asistencia | Curso de Web Backend"

# Paths
DIR_TEMP = "./temp"
DIR_CERTS = "./certificates"
INPUT_PATH = "./input/input.xlsx"
TEMPLATE_PATH = "./templates/template.svg"
TEMP_PATH = DIR_TEMP + "/output.svg"
CERT_PATH = DIR_CERTS + "/cert_" 

# Excel fields to keep
FIELD_NAME = "Nombre"
FIELD_SURNAME = "Apellido"
FIELD_EMAIL = "Email"
FIELD_FILE = "Filename"
FIELDS = [FIELD_NAME, FIELD_SURNAME, FIELD_EMAIL]

# Index to be used
INDEX = None

# ---------------------------------------------------------------
#                               FUNCTIONS
# ---------------------------------------------------------------

# Function to create directories
def create_dir(path):
    try:
        os.mkdir(path)
    except OSError as err:
        print ("Creation of the directory %s failed" % path, "-->", err)
    else:
        print ("Successfully created the directory %s " % path)

# Function to obtain the string of the template
def get_template_string(path):
    with open(path, "r") as f:
    	src = f.read()
    return src

# Start progress bar
def pre_progress_bar(width):
    print("PROGRESS (" + str(width), "certificates to process):")
    # setup toolbar
    sys.stdout.write("[%s]" % (" " * width))
    sys.stdout.flush()
    sys.stdout.write("\b" * (width+1)) # return to start of line, after '['

# Update progress bar
def progress_bar():
    # update the bar
    sys.stdout.write("-")
    sys.stdout.flush()

# Finish progress bar
def post_progress_bar():
    sys.stdout.write("]\n") # this ends the progress bar

# Gets a dataframe with all the information
def get_input_info(path):
    # Read excel
    data = pd.read_excel(path, index_col=INDEX)
    # Keep only email + name
    df = pd.DataFrame(data, columns= FIELDS)
    # Capitalize and trim names
    for i, row in df.iterrows():
        df.at[i, FIELD_NAME] = df.at[i, FIELD_NAME].title().strip()
        df.at[i, FIELD_SURNAME] = df.at[i, FIELD_SURNAME].title().strip()
    return df

# Generates a particular certificate
def generate_certificate(template, info):
    # Calculate name
    full_name = info[FIELD_NAME] + " " + info[FIELD_SURNAME]
    # Replace in template
    template = template.replace("%name", full_name)
    # Store result
    with open(TEMP_PATH, 'w') as svgout:
    	svgout.write(template)
    # Open result
    drawing = svg2rlg(TEMP_PATH)
    # Obtain url safe name
    full_name = full_name.replace(" ", "_")
    # Get filename
    filename = CERT_PATH + full_name + ".pdf"
    # Render
    renderPDF.drawToFile(drawing, filename)
    return filename

# Generates all the certificates
def generate_certificates():
    print("\nGENERATING CERTIFICATES")
    # Get template string
    template = get_template_string(TEMPLATE_PATH)
    # Get people info
    data = get_input_info(INPUT_PATH)
    # Init progress bar
    pre_progress_bar(len(data))
    # Generate certificates
    for i, row in data.iterrows():
        # Generate certificate
        filename = generate_certificate(template, row)
        # Store certificate direction
        data.at[i, FIELD_FILE] = filename
        # Print progress bar
        progress_bar()
    # Finish progress bar
    post_progress_bar()
    return data

# Setup for the program itself
def setup():
    # Create necessary directories
    print("\nCREATING NECESSARY DIRECTORIES")
    create_dir(DIR_TEMP)
    create_dir(DIR_CERTS)

# Setup for the email account
def setup_email():
    print("\nSENDING EMAILS")
    # set up the SMTP server
    s = smtplib.SMTP(host=EMAIL_HOST, port=EMAIL_PORT)
    # Set TLS
    s.starttls()
    # Get password
    password = os.environ.get(EMAIL_PASS_VAR)
    # Login to email
    s.login(EMAIL_ADDRESS, password)
    return s

# Close email connection
def finish_email(s):
    s.close()

# Send an email
def send_email(s, data):
    # Create message
    msg = MIMEMultipart()

    # Setup message params
    message = "¡Hola! \n\nTe adjuntamos tu Certificado de Asistencia del Curso de Web Backend. \n\n¡Saludos!"
    msg['From']=EMAIL_FROM
    msg['To']=data[FIELD_EMAIL]
    msg['Subject']=EMAIL_SUBJECT

    # Attach plaintext
    msg.attach(MIMEText(message, "plain"))
    # Attach the certificate
    with open(data[FIELD_FILE], "rb") as f:
        attach = MIMEApplication(f.read(), _subtype="pdf")
    attach.add_header('Content-Disposition', 'attachment', filename=str(data[FIELD_FILE]))
    msg.attach(attach)

    # Send the message
    s.send_message(msg)

# Send all emails
def send_emails(s, data):
    # Init progress bar
    pre_progress_bar(len(data))
    for i, row in data.iterrows():
        # Send email
        send_email(s, row)
        # Print progress bar
        progress_bar()
    # Finish progress bar
    post_progress_bar()


# ---------------------------------------------------------------
#                                 MAIN
# ---------------------------------------------------------------

# main() function
def main():
    # Command line args are in sys.argv[1], sys.argv[2] ..
    # sys.argv[0] is the script name itself and can be ignored
    # parse arguments
    parser = argparse.ArgumentParser(description="Certificate generator")

    # add arguments
    parser.add_argument('-s', '--send', action='store_true', dest='send', required=False, default=False, help="Enables sending of emails")
    args = parser.parse_args()

    # Setup for generator
    setup()

    # Generate the certificates
    data = generate_certificates()

    # Sending emails
    if args.send:
        s = setup_email()
        send_emails(s, data)
        finish_email(s)

# call main
if __name__ == '__main__':
    main()



