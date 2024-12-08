# -*- coding: utf-8 -*-

# This sample demonstrates handling intents from an Alexa skill using the Alexa Skills Kit SDK for Python.
# Please visit https://alexa.design/cookbook for additional examples on implementing slots, dialog management,
# session persistence, api calls, and more.
# This sample is built using the handler classes approach in skill builder.
import logging
import ask_sdk_core.utils as ask_utils

from ask_sdk_core.skill_builder import SkillBuilder
from ask_sdk_core.dispatch_components import AbstractRequestHandler
from ask_sdk_core.dispatch_components import AbstractExceptionHandler
from ask_sdk_core.handler_input import HandlerInput

from ask_sdk_model import Response

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# 
from google.oauth2 import service_account
from googleapiclient.discovery import build

from datetime import datetime, timedelta

scope = ["https://www.googleapis.com/auth/calendar"]

creds = service_account.Credentials.from_service_account_file("creds.json", scopes=scope)

API_NAME = 'calendar'
API_VERSION = 'v3'

calendar_id = "60b64ad4c3f932a896b4f25927c7dc8905684dfc0961b62b51f698b6e58c6625@group.calendar.google.com"


class LaunchRequestHandler(AbstractRequestHandler):
    """Handler for Skill Launch."""
    def can_handle(self, handler_input):
        # type: (HandlerInput) -> bool

        return ask_utils.is_request_type("LaunchRequest")(handler_input)

    def handle(self, handler_input):
        # type: (HandlerInput) -> Response
        speak_output = "Welcome to your calendar assistant. Say meeting to begin"

        return (
            handler_input.response_builder
                .speak(speak_output)
                .ask(speak_output)
                .response
        )


class CreateEventIntentHandler(AbstractRequestHandler):
    def can_handle(self, handler_input):
        # Check if the request is for CreateEventIntent
        return ask_utils.is_intent_name("CreateEventIntent")(handler_input)

    def handle(self, handler_input):
        slots = handler_input.request_envelope.request.intent.slots
        date = str(slots["date"].value)
        time = str(slots["time"].value)
        eventName = str(slots["eventName"].value)
        
        dateSlot = datetime.strptime(date, "%Y-%m-%d")
        hour = int(time.split(":")[0])
        mins = int(time.split(":")[1])
        time_min = datetime(dateSlot.year, dateSlot.month, dateSlot.day, hour, mins)
        time_max = time_min + timedelta(hours = 1)
        reserve_event(time_min, time_max,eventName)
        
        speak_output = (
                f"{eventName} has been successfully booked for {date} from "
                f"{time_min.strftime('%I:%M %p')} to {time_max.strftime('%I:%M %p')}."
        )
        
        return (
            handler_input.response_builder
            .speak(speak_output)
            .set_should_end_session(True)
            .response
        )

class EntireScheduleIntentHandler(AbstractRequestHandler):
    def can_handle(self, handler_input):
        return ask_utils.is_intent_name("EntireScheduleIntent")(handler_input)
    
    def handle(self, handler_input):
        # Get the date slot from the Alexa request
        slots = handler_input.request_envelope.request.intent.slots
        date = str(slots["date"].value)

        # Parse the date and set the time range
        date_obj = datetime.strptime(date, "%Y-%m-%d")
        time_min = datetime(date_obj.year, date_obj.month, date_obj.day, 0, 0).isoformat() #+ "Z"
        time_max = datetime(date_obj.year, date_obj.month, date_obj.day, 23, 59).isoformat() #+ "Z"
        
        # Fetch events from Google Calendar
        service = build(API_NAME, API_VERSION, credentials=creds)
        
        events = []
        page_token = None
        while True:
            # Fetch events for the given date range
            events_result = service.events().list(
                calendarId=calendar_id,
                timeMin=time_min,
                timeMax=time_max,
                singleEvents=True,  # Ensures recurring events are split
                orderBy="startTime", # Orders by event start time
                pageToken=page_token,
            ).execute()

            events.extend(events_result.get("items", []))
            page_token = events_result.get("nextPageToken")
            if not page_token:
                break

        # Prepare Alexa's response
        if not events:
            speak_output = f"You have no events scheduled for {date}."
        else:
            event_list = []
            for event in events:
                start_time = event["start"].get("dateTime", event["start"].get("date"))
                event_time = datetime.fromisoformat(start_time.replace("Z", "+00:00")).strftime("%H:%M")
                event_list.append(f"{event['summary']} at {event_time}")
            event_text = ", ".join(event_list)
            speak_output = f"Your schedule for {date} includes: {event_text}."

        # Respond to the user
        return (
            handler_input.response_builder
                .speak(speak_output)
                .set_should_end_session(True)
                .response
        )

class HelpIntentHandler(AbstractRequestHandler):
    """Handler for Help Intent."""
    def can_handle(self, handler_input):
        # type: (HandlerInput) -> bool
        return ask_utils.is_intent_name("AMAZON.HelpIntent")(handler_input)

    def handle(self, handler_input):
        # type: (HandlerInput) -> Response
        speak_output = "You can say hello to me! How can I help?"

        return (
            handler_input.response_builder
                .speak(speak_output)
                .ask(speak_output)
                .response
        )


class CancelOrStopIntentHandler(AbstractRequestHandler):
    """Single handler for Cancel and Stop Intent."""
    def can_handle(self, handler_input):
        # type: (HandlerInput) -> bool
        return (ask_utils.is_intent_name("AMAZON.CancelIntent")(handler_input) or
                ask_utils.is_intent_name("AMAZON.StopIntent")(handler_input))

    def handle(self, handler_input):
        # type: (HandlerInput) -> Response
        speak_output = "Goodbye!"

        return (
            handler_input.response_builder
                .speak(speak_output)
                .response
        )

class FallbackIntentHandler(AbstractRequestHandler):
    """Single handler for Fallback Intent."""
    def can_handle(self, handler_input):
        # type: (HandlerInput) -> bool
        return ask_utils.is_intent_name("AMAZON.FallbackIntent")(handler_input)

    def handle(self, handler_input):
        # type: (HandlerInput) -> Response
        logger.info("In FallbackIntentHandler")
        speech = "Hmm, I'm not sure. You can say Hello or Help. What would you like to do?"
        reprompt = "I didn't catch that. What can I help you with?"

        return handler_input.response_builder.speak(speech).ask(reprompt).response

class SessionEndedRequestHandler(AbstractRequestHandler):
    """Handler for Session End."""
    def can_handle(self, handler_input):
        # type: (HandlerInput) -> bool
        return ask_utils.is_request_type("SessionEndedRequest")(handler_input)

    def handle(self, handler_input):
        # type: (HandlerInput) -> Response

        # Any cleanup logic goes here.

        return handler_input.response_builder.response


class IntentReflectorHandler(AbstractRequestHandler):
    """The intent reflector is used for interaction model testing and debugging.
    It will simply repeat the intent the user said. You can create custom handlers
    for your intents by defining them above, then also adding them to the request
    handler chain below.
    """
    def can_handle(self, handler_input):
        # type: (HandlerInput) -> bool
        return ask_utils.is_request_type("IntentRequest")(handler_input)

    def handle(self, handler_input):
        # type: (HandlerInput) -> Response
        intent_name = ask_utils.get_intent_name(handler_input)
        speak_output = "You just triggered " + intent_name + "."

        return (
            handler_input.response_builder
                .speak(speak_output)
                # .ask("add a reprompt if you want to keep the session open for the user to respond")
                .response
        )


class CatchAllExceptionHandler(AbstractExceptionHandler):
    """Generic error handling to capture any syntax or routing errors. If you receive an error
    stating the request handler chain is not found, you have not implemented a handler for
    the intent being invoked or included it in the skill builder below.
    """
    def can_handle(self, handler_input, exception):
        # type: (HandlerInput, Exception) -> bool
        return True

    def handle(self, handler_input, exception):
        # type: (HandlerInput, Exception) -> Response
        logger.error(exception, exc_info=True)

        speak_output = "Sorry, I had trouble doing what you asked. Please try again."

        return (
            handler_input.response_builder
                .speak(speak_output)
                .ask(speak_output)
                .response
        )

sb = SkillBuilder()

sb.add_request_handler(LaunchRequestHandler())
sb.add_request_handler(CreateEventIntentHandler())
sb.add_request_handler(EntireScheduleIntentHandler())
sb.add_request_handler(HelpIntentHandler())
sb.add_request_handler(CancelOrStopIntentHandler())
sb.add_request_handler(FallbackIntentHandler())
sb.add_request_handler(SessionEndedRequestHandler())
sb.add_request_handler(IntentReflectorHandler()) # make sure IntentReflectorHandler is last so it doesn't override your custom intent handlers


sb.add_exception_handler(CatchAllExceptionHandler())

lambda_handler = sb.lambda_handler()
def reserve_event(time_min, time_max, event_name):
    # Initialize Google Calendar API service
    service = build(API_NAME, API_VERSION, credentials=creds)

    # Define event payload
    event = {
        'summary': event_name,
        'description': "Automated Event Description",
        'start': {
            'dateTime': time_min.strftime("%Y-%m-%dT%H:%M:%S"),
            'timeZone': 'America/Los_Angeles'
        },
        'end': {
            'dateTime': time_max.strftime("%Y-%m-%dT%H:%M:%S"),
            'timeZone': 'America/Los_Angeles'
        },
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'popup', 'minutes': 30}
            ],
        },
    }

    # Insert event into Google Calendar
    try:
        created_event = service.events().insert(calendarId=calendar_id, body=event).execute()
        print(f"Event created: {created_event.get('htmlLink')}")
    except Exception as e:
        print(f"An error occurred: {e}")

def check_availability(time_min, time_max):
    service = build(API_NAME, API_VERSION, credentials=creds)

    # Query for Free/Busy information
    payload = {
        "timeMin": time_min.strftime("%Y-%m-%dT%H:%M:%S"),
        "timeMax": time_max.strftime("%Y-%m-%dT%H:%M:%S"),
        "items": [{"id": calendar_id}]
    }

    try:
        freebusy = service.freebusy().query(body=payload).execute()
        print(f"Free/Busy Response: {freebusy}")
        return freebusy
    except Exception as e:
        print(f"An error occurred while checking availability: {e}")
