import json
import sys
from datetime import datetime, timedelta

import click
from O365 import MSGraphProtocol, Connection, Account
from O365.utils import ApiComponent
from dateutil.tz import tz

from credentials import credentials


class Calendar(ApiComponent):
    _endpoints = {"my_url_key": "/calendarview"}

    def __init__(self, *, parent=None, con=None, **kwargs):
        self.con = parent.con if parent else con
        protocol = parent.protocol if parent else kwargs.get("protocol")
        main_resource = parent.main_resource if parent else parent

        super().__init__(protocol=protocol, main_resource=main_resource)

    def read_calendar(self, date: datetime):
        # self.build_url just merges the protocol service_url with the endpoint passed as a parameter
        # to change the service_url implement your own protocol inheriting from Protocol Class
        url = self.build_url(self._endpoints.get("my_url_key"))

        query_parameters = {
            "startdatetime": date.strftime("%Y-%m-%d"),
            "enddatetime": (date + timedelta(days=1)).strftime("%Y-%m-%d"),
            "$select": "subject,organizer,attendees,start,end,location,isAllDay,webLink",
        }

        response = self.con.get(url, params=query_parameters)
        schedule = json.loads(response.text)

        my_day = []
        for event in schedule["value"]:
            # Filter some events that don't need minutes
            if event["subject"] in ["Fokuszeit", "Block", "Mittagessen", "Date Night"]:
                continue

            # I'm not interested in all day events
            if event["isAllDay"]:
                continue

            # Convert Start Time into local timezone
            start_time = datetime.fromisoformat(event["start"]["dateTime"][:-8])
            start_time = start_time.replace(tzinfo=tz.gettz(event["start"]["timeZone"]))
            start_time = start_time.astimezone(tz.tzlocal())

            # Participants
            participants = []

            # Organizer
            # I don't need myself in the participants list
            if event["organizer"]["emailAddress"]["address"] != "sascha.kiefer@sap.com":
                participants.append(
                    "#[[" + event["organizer"]["emailAddress"]["name"] + "]]"
                )

            # Attendees
            if len(event["attendees"]) < 20:  # More than 20 is a broadcast
                for attendee in event["attendees"]:
                    email = attendee["emailAddress"]["address"]
                    name = attendee["emailAddress"]["name"]

                    if (
                        email != "sascha.kiefer@sap.com"  # don't need me in the list
                        and ("#[[" + name + "]]")
                        not in participants  # filter the organizer from the attendees list
                        and name.startswith("DL ") is False  # Filter Delivery Lists
                    ):
                        participants.append("#[[" + name + "]]")

            my_day.append(
                {
                    "start_time": start_time,
                    "subject": event["subject"],
                    "participants": participants,
                    "link": event["webLink"],
                }
            )

        my_day.sort(key=lambda x: x["start_time"])

        # Create the output for Roam
        for event in my_day:
            participants_string = (
                ", ".join(event["participants"])
                if len(event["participants"]) > 0
                else ""
            )
            md = f"* __{event['start_time'].strftime('%H:%M')}__ - [{event['subject']}]({event['link']}) {participants_string} #[[Meeting Minutes]]"
            md = md.replace("  ", " ")

            click.echo(md)


def get_calendar():
    account = Account(credentials)

    if account.is_authenticated is False:
        click.echo(
            "You are not authenticated. Call 'my_schedule.py logon' first", err=True
        )
        sys.exit(-1)

    protocol = MSGraphProtocol()  # or maybe a user defined protocol
    con = Connection(
        credentials,
        scopes=[
            "https://graph.microsoft.com/User.Read",
            "https://graph.microsoft.com/Calendars.Read",
            "https://graph.microsoft.com/offline_access",
        ],
    )
    calendar = Calendar(con=con, protocol=protocol)
    return calendar


@click.group()
def main():
    """
    Query your Office365 calendar and return the results as a list of
    ROAM compatible entries
    """
    pass


@main.command()
def logon():
    """Follow the instructions on the screen to logon to Office365"""
    account = Account(credentials)
    if account.authenticate(scopes=["basic", "calendar"]):
        click.echo("You are authenticated")


@main.command()
def today():
    """Get the list for today's calendar entries"""
    get_calendar().read_calendar(date=datetime.now())


@main.command()
def tomorrow():
    """Get the list for tomorrow's calendar entries"""
    get_calendar().read_calendar(date=datetime.now() + timedelta(days=1))


if __name__ == "__main__":
    main()
