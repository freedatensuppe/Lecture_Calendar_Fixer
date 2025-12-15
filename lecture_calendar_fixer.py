import logging
import datetime
import os
from pathlib import Path

from dotenv import load_dotenv

import icalendar
import requests

from event import EventWrapper, ExchangeAccountManager
from api_call import load_from_mymci_api
import config


def delete_all_existing_lecture_events(exchange_manager):
    calendar_items = exchange_manager.get_calendar_items(
        EventWrapper.get_default_organizer()
    )

    logging.info(f"Found {len(calendar_items)} calendar items")

    should_retry = 1
    tries = 0
    while should_retry > 0 and tries < 5:
        tries += 1
        should_retry -= 1

        for item in calendar_items:
            logging.info(
                f"\nTrying to delete calendar item:\n\t{EventWrapper.from_outlook_event(item)}"
            )
            try:
                item.delete()
            except Exception as e:
                logging.warning(
                    f"Could not delete calendar item (Exception {e}), adding retry"
                )
                should_retry += 1


def add_lecture_events_to_outlook(webcalendar, exchange_manager):
    all_events = [
        subcomp for subcomp in webcalendar.subcomponents if subcomp.name == "VEVENT"
    ]
    lecture_events = [
        event for event in all_events if "Abgabetermin" not in event["summary"]
    ]

    logging.info(f"Found {len(lecture_events)} lecture events")

    account = exchange_manager.get_account()
    for event in lecture_events:
        wrapper = EventWrapper.from_ical_event(event)
        logging.info(f"\nAdding event:\n\t{wrapper}")
        wrapper.to_outlook_event(account)


def try_deleting_calendar_item(item) -> bool:
    attempts = 0
    while attempts < 5:
        try:
            item.delete()
            return True
        except Exception as e:
            logging.warning(
                f"Could not delete calendar item (Exception {e}), retrying..."
            )
        attempts += 1
    return False


def webcal_dict_to_wrapper(webcalendar_dict: list[dict]) -> list[EventWrapper]:
    lecture_events = []
    for event in webcalendar_dict:
        # skip submission dates
        if event["art"] not in ["Lehrveranstaltung", "Prüfung", "Sonstiges"]:
            continue

        lecture_events.append(EventWrapper.from_api_dict(event))

    return lecture_events


def webcal_to_wrapper(webcalendar) -> list[EventWrapper]:
    all_events = [
        subcomp for subcomp in webcalendar.subcomponents if subcomp.name == "VEVENT"
    ]
    lecture_events = []
    for event in all_events:
        # skip submission dates
        if "Abgabetermin" in event["summary"]:
            continue

        # skip all events that are stemming from your own SAKAI calendar
        # UID of event is either "MCI-DESIGNER-TERMIN-xxxx" or "MCI-SAKAI-TERMIN-xxxx"
        if "MCI-SAKAI-TERMIN" in event["uid"]:
            continue

        lecture_events.append(EventWrapper.from_ical_event(event))

    return lecture_events


def update_changed_events(wrapped_events: list[EventWrapper], exchange_manager):
    calendar_items = exchange_manager.get_calendar_items(
        EventWrapper.get_default_organizer()
    )

    logging.info(f"Found {len(calendar_items)} Exchange calendar items")
    calendar_item_dict = {}

    # Create dict with organizer email as key
    for item in calendar_items:
        organizer_email = item.organizer.email_address if item.organizer else ""
        calendar_item_dict[organizer_email] = item

    # create a dict of all found/imported events
    lecture_event_dict: dict[str, EventWrapper] = {}
    for event in wrapped_events:
        lecture_event_dict[event.organizer] = event

    logging.info(f"Found {len(wrapped_events)} lecture ical events")

    account = exchange_manager.get_account()

    for imported_event in lecture_event_dict.values():
        # get corresponding calendar item from dict
        corresponding_calendar_item = calendar_item_dict.get(imported_event.organizer)

        if corresponding_calendar_item:
            # if found wrap it for comparison
            calendar_item_wrapped = EventWrapper.from_outlook_event(
                corresponding_calendar_item
            )

            if imported_event != calendar_item_wrapped:
                # event has changed --> delete and add again
                logging.info(
                    f"\nTrying to delete calendar item:\n\t{calendar_item_wrapped}"
                )
                if try_deleting_calendar_item(corresponding_calendar_item):
                    # also remove from dict
                    del calendar_item_dict[imported_event.organizer]

                logging.info(f"\nAdding event:\n\t{imported_event}")
                imported_event.to_outlook_event(account)
            else:
                logging.info(f"\nCalendar item is up to date:\n\t{imported_event}")
        else:
            # if it is not available then add it
            logging.info(f"\nAdding event:\n\t{imported_event}")
            imported_event.to_outlook_event(account)

    # More calendar items than ical events --> something has been deleted in the ical events
    if len(calendar_item_dict) > len(lecture_event_dict):
        for calendar_item in calendar_item_dict.values():
            organizer_email = (
                calendar_item.organizer.email_address if calendar_item.organizer else ""
            )
            if not lecture_event_dict.get(organizer_email):
                calendar_item_wrapped = EventWrapper.from_outlook_event(calendar_item)

                # if the calendar item is not in the ical events then delete it only if it is in the future
                if (
                    calendar_item_wrapped.start_dt
                    and calendar_item_wrapped.start_dt
                    > datetime.datetime.now(
                        calendar_item_wrapped.start_dt.tzinfo
                        if calendar_item_wrapped.start_dt.tzinfo
                        else None
                    )
                ):
                    logging.info(
                        f"\nTrying to delete calendar item:\n\t{calendar_item_wrapped}"
                    )
                    try_deleting_calendar_item(calendar_item)


if __name__ == "__main__":
    load_dotenv()

    logfile_path = (
        Path(__file__).parent.resolve() / "full.log"
    )  # always logs into the same folder as the script, even when run from task scheduler
    logging.basicConfig(filename=logfile_path, encoding="utf-8", level=logging.DEBUG)
    logging.info(f"Running at {datetime.datetime.now()}")

    if config.use_ical_link() and config.use_api_call():
        logging.error(
            "Both use_ical_link and use_api_call are set to True in config.py. Please choose only one method to fetch events."
        )
        exit(1)

    if not (config.use_ical_link() or config.use_api_call()):
        logging.error(
            "Both use_ical_link and use_api_call are set to False in config.py. Please choose one method to fetch events."
        )
        exit(1)

    wrapped_events = []

    if config.use_ical_link():
        url = os.getenv("WEBCAL_URL")
        if url is None:
            logging.error("No webcal url found in .env file")
            exit(1)

        try:
            response = requests.get(url)
        except requests.exceptions.RequestException as e:
            logging.error(f"Could not fetch calendar: {e}")
            exit(1)

        webcalendar = icalendar.Calendar.from_ical(response.text)
        wrapped_events = webcal_to_wrapper(webcalendar)

    elif config.use_api_call():
        user = os.getenv("USER")
        if user is None:
            logging.error("No user found in .env file")
            exit(1)

        webcalendar = load_from_mymci_api(user)
        wrapped_events = webcal_dict_to_wrapper(webcalendar)

    try:
        exchange_manager = ExchangeAccountManager()
        update_changed_events(wrapped_events, exchange_manager)
    except ValueError as e:
        logging.error(f"Exchange configuration error: {e}")
        exit(1)
    except Exception as e:
        logging.error(f"Error connecting to Exchange server: {e}")
        exit(1)

