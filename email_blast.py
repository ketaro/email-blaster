#!/usr/bin/env python3

import click
import csv
import os
import smtplib
import subprocess

from email_validator import validate_email, EmailNotValidError
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from jinja2 import Environment, PackageLoader, select_autoescape
from InquirerPy import prompt, inquirer


SMTP_SETTINGS = {
    "smtp_host": os.environ.get("SMTP_HOST", ""),
    "smtp_user": os.environ.get("SMTP_USER", ""),
    "smtp_passwd": os.environ.get("SMTP_PASSWD", ""),
    "smtp_security": os.environ.get("SMTP_SECURITY", "tls"),
    "mail_from": os.environ.get("SMTP_MAIL_FROM", "noreply@hakaru.org")
}

jinja_env = Environment(
    loader=PackageLoader("email_blast"),
    autoescape=select_autoescape()
)


def get_file_lines(filename):
    """
    Use wc to count the number of lines in the csv file.
    """
    output = subprocess.check_output([
        "/usr/bin/wc", os.path.expanduser(filename)
    ])

    return int(output.split()[0])


def handle_smtp_error(error):
    if type(error) is smtplib.SMTPRecipientsRefused:
        if hasattr(error, "recipients"):
            for (key, value) in error.recipients.iteritems():
                print(value[1])

    elif hasattr(error, "smtp_error"):
        print(error.smtp_error)

    elif hasattr(error, "message"):
        print(error.message)

    else:
        print("Unknown Error (type: %s)" % type(error).__name__)

    exit(1)

def create_message(mail_to, mail_from, mail_from_name, subject, message_text):
    """Create the MIME Multipart message

    Args:
        mail_to (str): Email to
        mail_from (str): Email from
        subject (str): Email subject
        message_text (str): Plain text message

    Returns:
        _type_: _description_
    """
    msg = MIMEMultipart()
    msg["Subject"] = subject
    if mail_from_name:
        msg["From"] = f"{mail_from_name} <{mail_from}>"
    else:
        msg["From"] = mail_from
    msg["To"] = formataddr(("", mail_to))
    msg.preamble = "This is a multi-part message in MIME format"

    msgalt = MIMEMultipart("alternative")
    msg.attach(msgalt)

    msgalt.attach(MIMEText(message_text, "plain", "US-ASCII"))

    return msg

def send_email(mail_to, mail_from, mail_from_name, subject, message_text):
    """Send an email via smtp

    Args:
        mail_to (_type_): _description_
        mail_from (_type_): _description_
        subject (_type_): _description_
        message (_type_): _description_
    """
    
    # Start SMTP connection
    try:
        smtp = smtplib.SMTP(SMTP_SETTINGS["smtp_host"])
    except smtplib.SMTPException as err:
        handle_smtp_error(err)

    if not smtp:
        print(f"Unable to initialize SMTP to {SMTP_SETTINGS['smtp_host']}")
        exit(1)

    if SMTP_SETTINGS["smtp_security"] == "tls":
        try:
            smtp.starttls()
        except smtplib.SMTPException as err:
            handle_smtp_error(err)

    if SMTP_SETTINGS["smtp_user"] and SMTP_SETTINGS["smtp_passwd"]:
        try:
            smtp.login(SMTP_SETTINGS["smtp_user"], SMTP_SETTINGS["smtp_passwd"])
        except smtplib.SMTPException as err:
            handle_smtp_error(err)

    msg = create_message(mail_to, mail_from, mail_from_name, subject, message_text)

    try:
        smtp.sendmail(from_addr=mail_from, to_addrs=[mail_to], msg=msg.as_string())
    except smtplib.SMTPException as err:
        handle_smtp_error(err)


@click.command()
@click.option("--csv", prompt="CSV File", help="The CSV file containing data for mail merge.")
@click.option("--template", help="Email template to user for the mail merge.")
@click.option("--mail-from-name", help="Name of the email sender.")
@click.option("--dry-run", help="Send to this email address for testing")
def main(*args, **kwargs):
    csv_file = kwargs["csv"]
    email_column = "email"

    # The CSV data
    with open(csv_file, "r") as csvfile:
        csvreader = csv.reader(line.replace('\0', '') for line in csvfile)
        headers = None
        rows = []

        for row in csvreader:
            if not headers:
                # First row with more than 3 values in it is probably our header row
                if len(row) >= 3:
                    headers = row

            else:
                rows.append(row)

    csv_lines = len(rows)

    questions = []
    questions.append({
        "type": "list",
        "name": "email_column",
        "message": "Please select the column containing email addresses",
        "choices": headers,
        "default": next((h for h in headers if "email" in h.lower() or "e-mail" in h.lower()))
    })

    questions.append({
        "type": "list",
        "name": "template",
        "message": "Please select the email template to use",
        "choices": os.listdir("templates")
    })

    questions.append({
        "type": "input",
        "name": "subject",
        "message": "Please enter the email subject line:",
    })

    questions.append({
        "type": "input",
        "name": "mail_from",
        "message": "Send email from address: ",
        "default": SMTP_SETTINGS["mail_from"]
    })

    if not kwargs.get("mail_from_name"):
        questions.append({
            "type": "input",
            "name": "mail_from_name",
            "message": "Send email from name: ",
            "default": "",
        })

    # Check that headers include an email address column
    answers = prompt(questions)
    email_column = answers.get("email_column")
    template = jinja_env.get_template(answers.get("template"))
    subject = answers.get("subject")
    mail_from = answers.get("mail_from")
    mail_from_name = answers.get("mail_from_name", kwargs.get("mail_from_name"))

    print("\n--------------------------------------------------")
    print("%25s: %d data rows" % (os.path.split(csv_file)[-1], csv_lines))
    print(f"     Using Email Template: {template.name}")
    print(f"            Email Subject: {subject}")
    # print("            Email Backend: %s" % settings.EMAIL_BACKEND)
    if mail_from_name:
        print(f"          Send Email From: {mail_from_name} <{mail_from}>")
    else:
        print(f"          Send Email From: {mail_from}")
    if kwargs["dry_run"]:
        print(f"        DRY RUN Emails To: {kwargs['dry_run']}")
    else:
        print(f"     Send Email To Column: {email_column}")
    print("--------------------------------------------------")

    proceed = inquirer.confirm(message="Proceed?", default=False).execute()
    if not proceed:
        exit(0)

    for row in rows[:1]:
        data = dict(zip(headers, row))

        mail_to = data[email_column]

        # Verify the mail_to looks like an email address
        try:
            emailinfo = validate_email(mail_to, check_deliverability=False)
            mail_to = emailinfo.normalized

        except EmailNotValidError as err:
            print(f"Row contains invalid email address ({mail_to}), skipping")
            continue
 
        if kwargs["dry_run"]:
            mail_to = kwargs["dry_run"]
            print(f"mail to: {data[email_column]} (dry run to: {mail_to})")
        else:
            print(f"mail to: {mail_to}")

        # import pdb; pdb.set_trace()
        msg = template.render(data=data, **data)

        # print(msg)
        send_email(mail_to, mail_from, mail_from_name, subject, msg)


if __name__ == "__main__":
    main()
 