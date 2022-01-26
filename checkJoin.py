import discord
#This imports Discord. Named thing.py because my old bots had their main files in thing.js, and I'm sentimental. 

from datetime import date
from datetime import datetime
from discord.utils import get
from openpyxl import load_workbook
from openpyxl import Workbook


wb = Workbook()


intents = discord.Intents.all()
client = discord.Client(intents=intents)

@client.event
async def on_ready():
    CADiscord = client.get_guild(523962430179770369)
    memberList = ["Ethanarcade44", "Astroturtle", "TheStevenofSuburbia", "Monitor Lizard", "The Invisible Man", "OctetOcckson", "minerharry", "the danger of no confusion", "MasterOfShadow", "Starfall", "Patchkat", "blockhead"]
    for member in CADiscord.members:
        if member.name in memberList:
            print(member.name + ": " + str(member.joined_at))

    print("Completed!")



client.run('NjYyNzg4MTk1NzM3NDAzNDQy.Xg_Dmw.dyg4Kch4KxX6C6bDZAcx-Le2TVs')