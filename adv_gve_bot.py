# -*- coding: utf-8 -*-
"""
Sample code for using webexteamsbot
"""

import os
import requests
from webexteamsbot import TeamsBot
from webexteamsbot.models import Response
import sys
import json
from excel_db import *
from card import *


list_of_projects = []
next_project = None

# Retrieve required details from environment variables
bot_email = os.getenv("TEAMS_BOT_EMAIL")
teams_token = os.getenv("TEAMS_BOT_TOKEN")
bot_url = os.getenv("TEAMS_BOT_URL")
bot_app_name = os.getenv("TEAMS_BOT_APP_NAME")

# Example: How to limit the approved Webex Teams accounts for interaction
#          Also uncomment the parameter in the instantiation of the new bot
# List of email accounts of approved users to talk with the bot
# approved_users = [
#     "josmith@demo.local",
# ]

# If any of the bot environment variables are missing, terminate the app
if not bot_email or not teams_token or not bot_url or not bot_app_name:
    print(
        "sample.py - Missing Environment Variable. Please see the 'Usage'"
        " section in the README."
    )
    if not bot_email:
        print("TEAMS_BOT_EMAIL")
    if not teams_token:
        print("TEAMS_BOT_TOKEN")
    if not bot_url:
        print("TEAMS_BOT_URL")
    if not bot_app_name:
        print("TEAMS_BOT_APP_NAME")
    sys.exit()

# Create a Bot Object
#   Note: debug mode prints out more details about processing to terminal
#   Note: the `approved_users=approved_users` line commented out and shown as reference
bot = TeamsBot(
    bot_app_name,
    teams_bot_token=teams_token,
    teams_bot_url=bot_url,
    teams_bot_email=bot_email,
    debug=True,
    # approved_users=approved_users,
    webhook_resource_event=[
        {"resource": "messages", "event": "created"},
        {"resource": "attachmentActions", "event": "created"},
    ],
)


# Create a custom bot greeting function returned when no command is given.
# The default behavior of the bot is to return the '/help' command response
def greeting(incoming_msg):
    # Loopkup details about sender
    sender = bot.teams.people.get(incoming_msg.personId)

    # Create a Response object and craft a reply in Markdown.
    response = Response()
    response.markdown = f"Hello {sender.firstName}, I'm a chat bot. {sender.userName}"
    response.markdown += "See what I can do by asking for **/help**."
    return response


# Create functions that will be linked to bot commands to add capabilities
# ------------------------------------------------------------------------

# A simple command that returns a basic string that will be sent as a reply
def do_something(incoming_msg):
    """
    Sample function to do some action.
    :param incoming_msg: The incoming message object from Teams
    :return: A text or markdown based reply
    """
    return "i did what you said - {}".format(incoming_msg.text)


# This function generates a basic adaptive card and sends it to the user
# You can use Microsofts Adaptive Card designer here:
# https://adaptivecards.io/designer/. The formatting that Webex Teams
# uses isn't the same, but this still helps with the overall layout
# make sure to take the data that comes out of the MS card designer and
# put it inside of the "content" below, otherwise Webex won't understand
# what you send it.
def show_card(incoming_msg):
    attachment =  {
    
        "contentType": "application/vnd.microsoft.card.adaptive",
        "content": {
            "type": "AdaptiveCard",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "GVE SP Tracker (beta)",
                    "size": "Medium",
                    "color": "Dark"
                },
                {
                    "type": "Input.Text",
                    "placeholder": "Customer name",
                    "id": "customer"
                },
                {
                    "type": "Input.Text",
                    "placeholder": "Name of the project",
                    "id": "project"
                },
                {
                    "type": "Input.ChoiceSet",
                    "choices": [
                        {
                            "title": "RFP",
                            "value": "RFP"
                        },
                        {
                            "title": "RFI",
                            "value": "RFI"
                        },
                        {
                            "title": "GVE Support",
                            "value": "GVE Support"
                        },
                        {
                            "title": "Complex Design",
                            "value": "Complex Design"
                        }
                    ],
                    "placeholder": "Category of this engagement",
                    "id": "category"
                },
                {
                    "type": "Input.ChoiceSet",
                    "choices": [
                        {
                            "title": "High",
                            "value": "High"
                        },
                        {
                            "title": "Medium",
                            "value": "Medium"
                        },
                        {
                            "title": "Low",
                            "value": "Low"
                        }
                    ],
                    "placeholder": "Estimated complexity",
                    "id": "complexity"
                },
                {
                    "type": "Input.ChoiceSet",
                    "choices": [
                        {
                            "title": "WiP",
                            "value": "WiP"
                        },
                        {
                            "title": "CE Pending",
                            "value": "CE Pending"
                        },
                        {
                            "title": "Close",
                            "value": "Close"
                        }
                    ],
                    "placeholder": "Status",
                    "id": "status",
                },
                {
                    "type": "Input.ChoiceSet",
                    "choices": [
                        {
                            "title": "AMER",
                            "value": "AMER"
                        },
                        {
                            "title": "EMEA",
                            "value": "EMEA"
                        },
                        {
                            "title": "APJC",
                            "value": "APJC"
                        }
                    ],
                    "placeholder": "Cisco region",
                    "id": "region",
                },
                {
                    "type": "Input.ChoiceSet",
                    "choices": [
                        {
                            "title": "Alejandro Hernandez",
                            "value": "AH"
                        },
                        {
                            "title": "Francisco Quiroz",
                            "value": "FQ"
                        }
                    ],
                    "placeholder": "TSA Assigned",
                    "id": "tsa",
                },
                {
                    "type": "Input.ChoiceSet",
                    "choices": [
                        {
                            "title": "Queue",
                            "value": "Queue"
                        },
                        {
                            "title": "GVE",
                            "value": "GVE"
                        },
                        {
                            "title": "Field",
                            "value": "Field"
                        }
                    ],
                    "placeholder": "Requester",
                    "id": "user",
                }     
            ],
            "actions": [{
                    "type": "Action.Submit",
                    "title": "Create",
                    "data": "add",
                    "style": "positive",
                    "id": "button1"
                }
            ],
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.0"
        }
    }
    
    backupmessage = "This is an example using Adaptive Cards."

    c = create_message_with_attachment(
        incoming_msg.roomId, msgtxt=backupmessage, attachment=attachment
    )
    #print(f"Esto es el mensaje: {c}")
    return ""

# make sure to take the data that comes out of the MS card designer and
# put it inside of the "content" below, otherwise Webex won't understand
# what you send it.
def card_update(incoming_msg):
    global list_of_projects, next_project
    project_info = []
    list_of_projects = find_matching_rows_by_email(incoming_msg.personEmail)
    
    print(f" --------1----> lista de proyectos {list_of_projects}")
    if len(list_of_projects) > 0:
        next_project = list_of_projects.pop(0)
        
        project_info.append(next_project['Customer'])
        project_info.append(next_project['Project'])
        project_info.append(next_project['Category'])
        project_info.append(next_project['Complexity'])
        project_info.append(next_project['Status'])
        print(f"------ Primer renglon ------- {project_info} ")
    
        
    attachment =  update_attachment(project_info)
    

    backupmessage = "This is an example using Adaptive Cards."

    c = create_message_with_attachment(
            incoming_msg.roomId, msgtxt=backupmessage, attachment=attachment )
        #print(f"Esto es el mensaje: {c}")
    return ""



# An example of how to process card actions
def handle_cards(api, incoming_msg):
    global list_of_projects, next_project
    project_info = []
    """
    Sample function to handle card actions.
    :param api: webexteamssdk object
    :param incoming_msg: The incoming message object from Teams
    :return: A text or markdown based reply
    """
    m = get_attachment_actions(incoming_msg["data"]["id"])
    print(f"Esto es el incoming_msg mensaje de respuesta: {incoming_msg}")
    print(f"Esto es el mensaje de respuesta: {m}")
    if len(list_of_projects)>0:
        if m['inputs']['data'] == 'add':
            r = add_gve_record(m["inputs"])
        else:
            print(f"------ Se trata de ------- {m['inputs']['data']}")
            r= m['inputs']['data']
            next_project = list_of_projects.pop(0)
            print(f"el primer proyecto es {next_project}")
            
            project_info.append(next_project['Customer'])
            project_info.append(next_project['Project'])
            project_info.append(next_project['Category'])
            project_info.append(next_project['Complexity'])
            project_info.append(next_project['Status'])
       

        attachment =  update_attachment(project_info)

        backupmessage = "This is an example using Adaptive Cards."

        c = create_message_with_attachment(m['roomId'], msgtxt=backupmessage, attachment=attachment )
        #print(f"Esto es el mensaje: {c}")
        return ""
    else:
        ### Corregir esta salida y poner el codigo para actualizar segur next o delete
        sender = bot.teams.people.get(m['personId'])
        response = Response()
        response.markdown = f"Hello {sender.firstName}, there a no more projects"
        
   
    return f"card action was - {r}"


# Temporary function to send a message with a card attachment (not yet
# supported by webexteamssdk, but there are open PRs to add this
# functionality)
def create_message_with_attachment(rid, msgtxt, attachment):
    headers = {
        "content-type": "application/json; charset=utf-8",
        "authorization": "Bearer " + teams_token,
    }

    url = "https://api.ciscospark.com/v1/messages"
    data = {"roomId": rid, "attachments": [attachment], "markdown": msgtxt}
    response = requests.post(url, json=data, headers=headers)
    return response.json()


# Temporary function to get card attachment actions (not yet supported
# by webexteamssdk, but there are open PRs to add this functionality)
def get_attachment_actions(attachmentid):
    headers = {
        "content-type": "application/json; charset=utf-8",
        "authorization": "Bearer " + teams_token,
    }

    url = "https://api.ciscospark.com/v1/attachment/actions/" + attachmentid
    response = requests.get(url, headers=headers)
    return response.json()


# An example using a Response object.  Response objects allow more complex
# replies including sending files, html, markdown, or text. Rsponse objects
# can also set a roomId to send response to a different room from where
# incoming message was recieved.
def ret_message(incoming_msg):
    """
    Sample function that uses a Response object for more options.
    :param incoming_msg: The incoming message object from Teams
    :return: A Response object based reply
    """
    # Create a object to create a reply.
    response = Response()

    # Set the text of the reply.
    response.text = "Here's a fun little meme."

    # Craft a URL for a file to attach to message
    u = "https://sayingimages.com/wp-content/uploads/"
    u = u + "aaaaaalll-righty-then-alrighty-meme.jpg"
    response.files = u
    return response


# An example command the illustrates using details from incoming message within
# the command processing.
def current_time(incoming_msg):
    """
    Sample function that returns the current time for a provided timezone
    :param incoming_msg: The incoming message object from Teams
    :return: A Response object based reply
    """
    # Extract the message content, without the command "/time"
    timezone = bot.extract_message("/time", incoming_msg.text).strip()

    # Craft REST API URL to retrieve current time
    #   Using API from http://worldclockapi.com
    u = "http://worldclockapi.com/api/json/{timezone}/now".format(timezone=timezone)
    r = requests.get(u).json()

    # If an invalid timezone is provided, the serviceResponse will include
    # error message
    if r["serviceResponse"]:
        return "Error: " + r["serviceResponse"]

    # Format of returned data is "YYYY-MM-DDTHH:MM<OFFSET>"
    #   Example "2018-11-11T22:09-05:00"
    returned_data = r["currentDateTime"].split("T")
    cur_date = returned_data[0]
    cur_time = returned_data[1][:5]
    timezone_name = r["timeZoneName"]

    # Craft a reply string.
    reply = "In {TZ} it is currently {TIME} on {DATE}.".format(
        TZ=timezone_name, TIME=cur_time, DATE=cur_date
    )
    return reply


# Create help message for current_time command
current_time_help = "Look up the current time for a given timezone. "
current_time_help += "_Example: **/time EST**_"

# Set the bot greeting.
bot.set_greeting(greeting)

# Add new commands to the bot.
bot.add_command("attachmentActions", "*", handle_cards)
bot.add_command("/n", "New GVS SP Tracker record", show_card)
bot.add_command("/u", "Update GVS SP Tracker record", card_update)
bot.add_command("/dosomething", "help for do something", do_something)
bot.add_command("/demo", "Sample that creates a Teams message to be returned.", ret_message)
bot.add_command("/time", current_time_help, current_time)

# Every bot includes a default "/echo" command.  You can remove it, or any
# other command with the remove_command(command) method.
bot.remove_command("/echo")

if __name__ == "__main__":
    # Run Bot
    bot.run(host="0.0.0.0", port=7001)