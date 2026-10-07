import discord
from discord.ext import commands
import logging
from dotenv import load_dotenv
import os
import aiohttp

load_dotenv()

token = os.getenv('DISCORD_TOKEN')

handler = logging.FileHandler(filename='discord.log', encoding='utf-8', mode='w')
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
async def unavailable(ctx, *, msg):
    params = {
        "command": "unavailable",
        "username": str(ctx.author),
        "message": msg
    }

    url = os.getenv("APPS_SCRIPT_ENDPOINT")
    if not url:
        await ctx.send("APPS_SCRIPT_URL is not configured.")
        return

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
                    await ctx.send(f"Data successfully synchronized with Google Apps Script! {result.get('username')}, {result.get('message')}")
                else:
                    await ctx.send(f"Apps Script error: {result.get('message')}")

    except aiohttp.ContentTypeError:
        await ctx.send("Apps Script returned something that wasn't JSON.")
    except (aiohttp.ClientError, TimeoutError) as e:
        await ctx.send(f"An error occurred while sending data: {e}")    




bot.run(token, log_handler=handler, log_level=logging.DEBUG)