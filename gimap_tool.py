#! /usr/bin/env python

import argparse
import json
from pathlib import Path

import google.auth.transport.requests
from google.oauth2 import service_account
from imap_tools import MailBox
from imap_tools.utils import quote as quote_imap_string
from rich.console import Console
from rich.table import Table


def main(args):
    """
    Main program entry point.
    """
    subject = args.subject
    if subject is None:
        subject = args.email
    credentials = get_credentials(args.credentials, subject)
    subject = args.email
    if args.subject != subject:
        delegated_credentials = credentials.with_subject(subject)
        credentials = delegated_credentials
    request = google.auth.transport.requests.Request()
    credentials.refresh(request)
    access_token = delegated_credentials.token
    headers_only = True
    messages = []
    attachment_folder = args.attachment_folder
    if attachment_folder is not None:
        headers_only = False
        attachment_folder_path = Path(attachment_folder)
    with MailBox("imap.gmail.com").xoauth2(subject, access_token) as mailbox:
        criteria = "ALL"
        if args.criteria is not None:
            criteria = f"X-GM-RAW {quote_imap_string(args.criteria)}"
        for msg in mailbox.fetch(
            criteria=criteria, headers_only=headers_only, bulk=100, mark_seen=False
        ):
            messages.append(msg)
            if attachment_folder is not None:
                for attachment in msg.attachments:
                    fname = attachment_folder_path / Path(attachment.filename).name
                    with open(fname, "wb") as f:
                        f.write(attachment.payload)

    display_message_summaries(messages)


def display_message_summaries(messages):
    """
    Display message summaries.
    """
    table = Table(title="Email Messages")
    table.add_column("uid", justify="right", style="green")
    table.add_column("Date", justify="right", style="cyan", no_wrap=True)
    table.add_column("Subject", style="magenta")
    for msg in messages:
        table.add_row(msg.uid, msg.date.isoformat(), msg.subject)
    console = Console()
    console.print(table)


def get_credentials(fcreds, subject):
    """
    Get Google credentials
    """
    service_account_info = json.load(fcreds)
    scopes = [
        "https://mail.google.com/",
    ]
    credentials = service_account.Credentials.from_service_account_info(
        service_account_info,
        scopes=scopes,
        subject=subject,
    )
    return credentials


if __name__ == "__main__":
    parser = argparse.ArgumentParser("GMail IMAP tool")
    parser.add_argument(
        "email", action="store", help="Email address of mailbox to scan."
    )
    parser.add_argument(
        "-c",
        "--credentials",
        type=argparse.FileType("r"),
        action="store",
        help="Credentials file in JSON format.",
    )
    parser.add_argument(
        "--subject", action="store", help="Subject of credentials different from EMAIL."
    )
    parser.add_argument(
        "-a",
        "--attachment-folder",
        action="store",
        metavar="FOLDER",
        help="Download attachments to FOLDER.",
    )
    parser.add_argument("--criteria", action="store", help="GMail search criteria.")
    args = parser.parse_args()
    main(args)
