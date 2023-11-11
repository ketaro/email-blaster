This is a quick script I wrote to take a CSV file and do an email-merge using a Jinja-based template.  Over time I may clean it up a bit more and add more customizations as needed and maybe some testing if I really want to fancy it up.

Requires the following environment varilables set with the SMTP settings:

* SMTP_HOST
* SMTP_USER
* SMTP_PASSWD
* SMTP_MAIL_FROM (optional)

## Installation

```
cd email-blaster
python3 -mvenv env
source env/bin/activate
pip install -r requirements.txt
mkdir templates
```

Create the `templates` directory and place your email templates there.  Create a `.html` and `.txt` file with the same filename to send a MIME encoded multi-part email with both options (generally good email practice).


## Usage

```
Usage: email_blast.py [OPTIONS]

Options:
  --csv TEXT             The CSV file containing data for mail merge.
  --template TEXT        Email template to user for the mail merge.
  --mail-from-name TEXT  Name of the email sender.
  --subject TEXT         Email Subject
  --dry-run TEXT         Send to this email address for testing
  --help                 Show this message and exit.
 ```

