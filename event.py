from icalendar.cal import Event as ICalEvent
from exchangelib import (
    Account,
    Credentials,
    Configuration,
    DELEGATE,
    CalendarItem,
    EWSDateTime,
)
from exchangelib.items import SEND_TO_NONE

from datetime import datetime
import keyring
import os
import logging
from config import get_travel_time, at_different_location, is_async_online_lecture
from exchangelib.winzone import MS_TIMEZONE_TO_IANA_MAP

MS_TIMEZONE_TO_IANA_MAP["(UTC+01:00) Amsterdam, Berlin, Bern, Rom, Stockholm, Wien"] = (
    "Europe/Vienna"
)


class ExchangeAccountManager:
    def __init__(self):
        # These should be set as environment variables
        self.email = os.getenv("EXCHANGE_EMAIL")
        self.username = os.getenv(
            "EXCHANGE_USERNAME", self.email
        )  # fallback to email if not set
        self.password = keyring.get_password(
            "lecture_calendar_fixer_exchange", self.username
        )
        self.server = os.getenv(
            "EXCHANGE_SERVER"
        )  # optional, auto-discovery will be used if not set

        if not all([self.email, self.password]):
            logging.error(
                f"No password found in system keyring for user: {self.username}. Please set it beforehand"
            )
            logging.error(
                "To do so, run: `python -m keyring set lecture_calendar_fixer_exchange <username>`"
            )
            raise ValueError(
                "Missing Exchange credentials. Please set EXCHANGE_EMAIL and EXCHANGE_PASSWORD environment variables."
            )

        self._account = None

    def get_account(self) -> Account:
        """Get Exchange account instance"""
        if self._account:
            return self._account

        credentials = Credentials(username=self.username, password=self.password)

        if self.server:
            # Use specific server if provided
            config = Configuration(server=self.server, credentials=credentials)
            self._account = Account(
                primary_smtp_address=self.email,
                config=config,
                autodiscover=False,
                access_type=DELEGATE,
            )
        else:
            # Use autodiscovery
            self._account = Account(
                primary_smtp_address=self.email,
                credentials=credentials,
                autodiscover=True,
                access_type=DELEGATE,
            )

        return self._account

    def get_calendar_items(self, organizer_filter=None):
        """Get calendar items, optionally filtered by organizer"""
        account = self.get_account()
        # Get all calendar items
        calendar_folder = account.calendar
        items = list(calendar_folder.all())

        if organizer_filter:
            filtered_items = []
            for item in items:
                # Filter by the marker category that we add to all our events
                if hasattr(item, 'categories') and item.categories:
                    if any(organizer_filter in str(cat) for cat in item.categories):
                        filtered_items.append(item)
            return filtered_items

        return items


class EventWrapper:
    def __init__(
        self,
        subject: str,
        start: str,
        duration: int,
        location: str,
        organizer: str | None = None,
        start_dt: datetime | None = None,
        is_online: bool = False,
        kind: str = "Lehrveranstaltung",
    ) -> None:
        self.subject = subject
        self.start = start
        self.duration = duration
        self.location = location
        self.start_dt = start_dt
        self.organizer = (
            organizer if organizer else EventWrapper.get_default_organizer()
        )
        self.is_online = is_online
        self.kind = kind

    @classmethod
    def from_ical_event(cls, event: ICalEvent):
        subject = event["summary"]
        start_dt = event["dtstart"].dt
        end_dt = event["dtend"].dt
        dur = int((end_dt - start_dt).total_seconds() / 60)
        start = start_dt.strftime("%Y-%m-%d %H:%M")
        location = event.get("location")
        location = location if location else "-"

        organizer = event.get("UID")
        organizer = str(organizer) if organizer else cls.get_default_organizer()

        return cls(
            subject, start, dur, location, organizer=organizer, start_dt=start_dt
        )

    @classmethod
    def from_api_dict(cls, api_dict: dict):
        subject = api_dict["title"]

        start_dt = datetime.fromisoformat(api_dict["start"])
        end_dt = datetime.fromisoformat(api_dict["end"])
        dur = int((end_dt - start_dt).total_seconds() / 60)
        start = start_dt.strftime("%Y-%m-%d %H:%M")

        kind = api_dict["art"]
        is_online = api_dict["online"]

        raeume = api_dict.get("raeume", [])
        raum = raeume[0] if raeume else None  # returns list of dicts
        location = (
            f"{raum['raum']} / {raum['standort']}" if raum else "-"
        )  # print as "Room / Location"

        id = api_dict["id"]
        if id[0].isnumeric():
            organizer = f"{cls.get_default_organizer()}-{id}"
        else:
            organizer = cls.get_default_organizer()

        return cls(
            subject,
            start,
            dur,
            location,
            organizer=organizer,
            start_dt=start_dt,
            is_online=is_online,
            kind=kind,
        )

    @classmethod
    def from_outlook_event(cls, event):
        # event is now an exchangelib CalendarItem object
        start_dt = event.start.astimezone()
        end_dt = event.end.astimezone()
        duration = int((end_dt - start_dt).total_seconds() / 60)
        start_str = start_dt.strftime("%Y-%m-%d %H:%M")
        location = event.location if event.location else "-"
        organizer = (
            event.organizer.email_address
            if event.organizer
            else cls.get_default_organizer()
        )
        return cls(event.subject, start_str, duration, location, organizer, start_dt)

    def to_outlook_event(self, account: Account):
        from datetime import timedelta
        import pytz

        if not self.start_dt:
            raise ValueError("start_dt must be set for exchangelib integration")

        # Calculate end time from duration
        end_dt = self.start_dt + timedelta(minutes=self.duration)

        # Ensure timezone awareness - use local timezone if none is set
        if self.start_dt.tzinfo is None:
            # Assume local timezone if not set
            local_tz = pytz.timezone(
                "Europe/Vienna"
            )  # Adjust as needed for MCI location
            start_dt_tz = local_tz.localize(self.start_dt)
            end_dt_tz = local_tz.localize(end_dt)
        else:
            start_dt_tz = self.start_dt
            end_dt_tz = end_dt

        # Convert to EWSDateTime for exchangelib
        start_ews = EWSDateTime.from_datetime(start_dt_tz)
        end_ews = EWSDateTime.from_datetime(end_dt_tz)

        # by default handle events as if they were past events
        is_past = True
        if self.start_dt:
            is_past = self.start_dt < datetime.now(
                self.start_dt.tzinfo if self.start_dt.tzinfo else None
            )

        # default values for category and reminder time --> ToDo: move to config.py .env
        category = "Vorlesung"
        reminder_on = True
        reminder_time = 15
        legacy_free_busy_status = "Busy"
        
        # Always add a marker category to identify events created by this script
        marker_category = "MCI-DESIGNER-TERMIN"
        categories = [category, marker_category]

        if self.location != "-":
            room, mci_location, *_ = self.location.split(" / ")

            if is_async_online_lecture(self.subject, room):
                legacy_free_busy_status = "Free"
                reminder_on = False
            elif at_different_location(mci_location):
                category = "Vorlesung-Anderer-Standort"
                categories = [category, marker_category]
                reminder_time += get_travel_time(mci_location)
            elif self.kind not in ["Lehrveranstaltung", "Prüfung", "Sonstiges"]:
                legacy_free_busy_status = "Free"

        # Create calendar item with reminder settings
        if (not is_past) and reminder_on:
            item = CalendarItem(
                account=account,
                folder=account.calendar,
                subject=self.subject,
                start=start_ews,
                end=end_ews,
                location=self.location,
                legacy_free_busy_status=legacy_free_busy_status,
                categories=categories,
                reminder_is_set=True,
                reminder_minutes_before_start=reminder_time,
            )
        else:
            item = CalendarItem(
                account=account,
                folder=account.calendar,
                subject=self.subject,
                start=start_ews,
                end=end_ews,
                location=self.location,
                legacy_free_busy_status=legacy_free_busy_status,
                categories=categories,
                reminder_is_set=False,
            )

        # Save the item
        item.save(send_meeting_invitations=SEND_TO_NONE)

        return item

    @staticmethod
    def get_default_organizer() -> str:
        return "MCI-DESIGNER-TERMIN"

    def __eq__(self, __value: object) -> bool:
        if not isinstance(__value, EventWrapper):
            return False

        same_subject = self.subject == __value.subject

        # If the EventWrapper is coming from an icalendar event then the start is a string, but it also always has the start_dt attribute
        # If the EventWrapper is coming from an exchangelib event then the start is a datetime object
        # ToDo: Remove this distinction between the two ways a EventWrapper can be created
        self_start = self.start
        if isinstance(self_start, str) and self.start_dt:
            self_start = self.start_dt

        other_start = __value.start
        if isinstance(other_start, str) and __value.start_dt:
            other_start = __value.start_dt

        same_start = False
        if self_start and other_start:
            if hasattr(self_start, "ctime") and hasattr(other_start, "ctime"):
                same_start = self_start.ctime() == other_start.ctime()
            else:
                # fallback to string comparison if not datetime objects
                same_start = str(self_start) == str(other_start)

        same_duration = self.duration == __value.duration
        same_location = self.location == __value.location
        same_organizer = self.organizer == __value.organizer
        return (
            same_subject
            and same_start
            and same_duration
            and same_location
            and same_organizer
        )

    def __str__(self) -> str:
        return f"Subject: {self.subject}\n\tStart: {self.start}\n\tDuration: {self.duration}\n\tLocation: {self.location}\n\tOrganizer: {self.organizer}\n\tStart_dt: {self.start_dt}"

    # only define left and right addition for string representation of EventWrapper
    def __add__(self, other: str) -> str:
        return str(self) + other

    def __radd__(self, other: str) -> str:
        return other + str(self)
