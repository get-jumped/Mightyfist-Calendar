import discord
from discord.ext import commands
import logging
from dotenv import load_dotenv
import os
import aiohttp
from google.cloud import secretmanager

def get_discord_secret():
    client = secretmanager.SecretManagerServiceClient()
    name = "projects/mightyfist-calendar/secrets/DISCORD_TOKEN/versions/latest"
    response = client.access_secret_version(request={"name": name})
    return response.payload.data.decode("UTF-8")


def get_apps_secret():
    client = secretmanager.SecretManagerServiceClient()
    name = "projects/mightyfist-calendar/secrets/APPS_SCRIPT_ENDPOINT/versions/latest"
    response = client.access_secret_version(request={"name": name})
    return response.payload.data.decode("UTF-8")

# USE FOR TESTING NOT ON THE CLOUD
# load_dotenv()

# token = os.getenv('DISCORD_TOKEN')

# url = os.getenv("APPS_SCRIPT_ENDPOINT")

token = get_discord_secret()

url = get_apps_secret()

# handler = logging.FileHandler(filename='discord.log', encoding='utf-8', mode='w')
intents = discord.Intents.default()
intents.message_content = True
intents.members = True

bot = commands.Bot(command_prefix='!', intents=intents)

@bot.event
async def on_ready():
    print(f"We are ready to go in, {bot.user.name}")


@bot.event
async def on_member_join(member):
    await member.send(f"Welcome to the server {member.name}")


@bot.command()
async def unavailable(ctx, *, msg: str = None):
    if msg is None:
        await ctx.send("Did not provide a list of unavailable days e.g !unavailable 10 14 15")

    member = ctx.author
    days = parseMessage(msg)

    params = {
        "command": "unavailable",
        "username": str(member.display_name),
        "days": days
    }

    try:
        async with aiohttp.ClientSession() as session:
            async with session.post(
                url,
                json=params,
                timeout=aiohttp.ClientTimeout(total=15),
            ) as response:
                if response.status != 200:
                    await ctx.send(f"Failed to connect to Apps Script. HTTP Status: {response.status}")
                    return

                result = await response.json(content_type=None)

                if result.get("status") == "success":
                    message = ''
                    for m in result.get('message'):
                        message += m
                        message += '\n'
                    
                    await ctx.send(message)
                else:
                    await ctx.send(f"Apps Script error: {result.get('message')}")

    except aiohttp.ContentTypeError:
        await ctx.send("Apps Script returned something that wasn't JSON.")
    except (aiohttp.ClientError, TimeoutError) as e:
        await ctx.send(f"An error occurred while sending data: {e}")    


@bot.command()
async def confirm(ctx):
    member = ctx.author

    params = {
        "command": "confirm",
        "username": str(member.display_name),
        "channel": str(ctx.channel)
    }

    try:
        async with aiohttp.ClientSession() as session:
            async with session.post(
                url,
                json=params,
                timeout=aiohttp.ClientTimeout(total=15),
            ) as response:
                if response.status != 200:
                    await ctx.send(f"Failed to connect to Apps Script. HTTP Status: {response.status}")
                    return

                result = await response.json(content_type=None)

                if result.get("status") == "success":
                    await ctx.send(result.get('message'))
                else:
                    await ctx.send(f"Apps Script error: {result.get('message')}")

    except aiohttp.ContentTypeError:
        await ctx.send("Apps Script returned something that wasn't JSON.")
    except (aiohttp.ClientError, TimeoutError) as e:
        await ctx.send(f"An error occurred while sending data: {e}")    



# ------------------------------------- HELPER FUNCITONS -------------------------------------
def parseMessage(msg):
    days = []
    spaceIndex = msg.find(' ')

    while spaceIndex != -1:
        days.append(msg[:spaceIndex])
        msg = msg[spaceIndex + 1:]
        spaceIndex = msg.find(' ')

    days.append(msg)

    return days



# bot.run(token, log_handler=handler, log_level=logging.DEBUG)
bot.run(token, log_handler=None)