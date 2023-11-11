#!/usr/bin/env python3

import click
import csv
import os
import smtplib
import time

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

# Load the available templates (breaking out the file extensions)
TEMPLATES = {}
for file in os.listdir("templates"):
    fileparts = os.path.splitext(file)
    TEMPLATES.setdefault(fileparts[0], [])
    TEMPLATES[fileparts[0]].append(fileparts[1])


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

def create_message(mail_to, mail_from, subject, template_key, data):
    """Create the MIME Multipart message

    Args:
        mail_to (str): Email to
        mail_from (str): Email from
        subject (str): Email subject
        template (str): Template name
        data: Data to pass to template during render

    Returns:
        _type_: _description_
    """
    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = mail_from
    msg["To"] = formataddr(("", mail_to))
    msg.preamble = "This is a multi-part message in MIME format"

    msgalt = MIMEMultipart("alternative")
    msg.attach(msgalt)

    # Render each available template type and attach to multipart message
    for ext in TEMPLATES[template_key]:
        template = jinja_env.get_template(f"{template_key}{ext}")
        content = template.render(data=data, **data)

        # try to determine character set
        for charset in ["US-ASCII", "ISO-8859-1", "UTF-8"]:
            try:
                content.encode(charset)
            except UnicodeError:
                pass
            else:
                break

        if ext.lower() in [".html", ".htm"]:
            msgalt.attach(MIMEText(content.encode(charset), "html", charset))
        else:
            msgalt.attach(MIMEText(content, "plain", charset))

    return msg

def send_email(mail_to, mail_from, message):
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

    try:
        smtp.sendmail(
            from_addr=mail_from,
            to_addrs=[mail_to],
            msg=message.as_string()
        )
    except smtplib.SMTPException as err:
        handle_smtp_error(err)


@click.command()
@click.option("--csv", prompt="CSV File", help="The CSV file containing data for mail merge.")
@click.option("--template", help="Email template to user for the mail merge.")
@click.option("--mail-from-name", help="Name of the email sender.")
@click.option("--subject", help="Email Subject")
@click.option("--dry-run", help="Send to this email address for testing")
def main(*args, **kwargs):
    csv_file = kwargs["csv"]
    email_column = "email"
    subject = kwargs.get("subject")

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
        "choices": TEMPLATES.keys()
    })

    if not subject:
        questions.append({
            "type": "input",
            "name": "subject",
            "message": "Please enter the email subject line:",
        })

    questions.append({
        "type": "input",
        "name": "mail_from_addr",
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
    template = answers.get("template")
    subject = answers.get("subject", subject)
    mail_from_addr = answers.get("mail_from_addr")
    mail_from_name = answers.get("mail_from_name", kwargs.get("mail_from_name"))

    # Format the mail from address
    mail_from = mail_from_addr
    if mail_from_name:
        mail_from = f"{mail_from_name} <{mail_from_addr}>"

    print("\n--------------------------------------------------")
    print("%25s: %d data rows" % (os.path.split(csv_file)[-1], csv_lines))
    print(f"     Using Email Template: {template} " + str(TEMPLATES[template]))
    print(f"            Email Subject: {subject}")
    print(f"              SMTP Server: {SMTP_SETTINGS['smtp_host']}")
    print(f"                SMTP User: {SMTP_SETTINGS['smtp_user']}")
    print(f"          Send Email From: {mail_from}")
    if kwargs["dry_run"]:
        print(f"        DRY RUN Emails To: {kwargs['dry_run']}")
    else:
        print(f"     Send Email To Column: {email_column}")
    print("--------------------------------------------------")

    proceed = inquirer.confirm(message="Proceed?", default=False).execute()
    if not proceed:
        exit(0)

    for row in rows:
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

        msg = create_message(
            mail_to=mail_to,
            mail_from=mail_from,
            subject=subject,
            template_key=template,
            data=data
        )
        send_email(mail_to, mail_from_addr, msg)
        
        # a little delay to avoid spamming
        time.sleep(0.5)


if __name__ == "__main__":
    main()
