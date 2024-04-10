import os
from webexteamsbot import TeamsBot

#export TEAMS_BOT_URL=https://f3c9-201-110-158-77.ngrok-free.app
#export TEAMS_BOT_TOKEN=MTE1N2YwMmEtZGUzZS00MjNmLThlNTYtODczY2Q2NGNkYWIzOGM5ZDQ2YjQtZjI2_PF84_1eb65fdf-9643-417f-9974-ad72cae0e10f
#export TEAMS_BOT_EMAIL=gve_sp_tracker@webex.bot
#export TEAMS_BOT_APP_NAME=GVE_SP_Tracker

# Retrieve required details from environment variables
bot_email = os.getenv("TEAMS_BOT_EMAIL")
teams_token = os.getenv("TEAMS_BOT_TOKEN")
bot_url = os.getenv("TEAMS_BOT_URL")
bot_app_name = os.getenv("TEAMS_BOT_APP_NAME")

# Create a Bot Object
bot = TeamsBot(
    bot_app_name,
    teams_bot_token=teams_token,
    teams_bot_url=bot_url,
    teams_bot_email=bot_email,
)


# A simple command that returns a basic string that will be sent as a reply
def do_something(incoming_msg):
    """
    Sample function to do some action.
    :param incoming_msg: The incoming message object from Teams
    :return: A text or markdown based reply
    """
    return "i did what you said - {}".format(incoming_msg.text)


# Add new commands to the box.
bot.add_command("/dosomething", "help for do something", do_something)


if __name__ == "__main__":
    # Run Bot
    bot.run(host="0.0.0.0", port=7001)
