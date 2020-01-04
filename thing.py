import discord
#This imports Discord. Named thing.py because my old bots had their main files in thing.js, and I'm sentimental. 
import xlrd
import xlwt
from xlwt import Workbook
wb = Workbook()
wr = xlrd.open_workbook("C:\\Users\\Sebastian_Polge\\OneDrive-CaryAcademy\\Documents\\meNewBot\\Verity\\StatsBot\\Beep.xls") 

client = discord.Client()

@client.event
async def on_ready():
    print('We have logged in as {0.user}'.format(client))

@client.event
async def on_message(message):
    if message.author == client.user:
        return
    if message.content.startswith('$setUpServer'):
        sheet1 = wb.add_sheet('Sheet 1')
        i = 1
        memberList = message.guild.members
        for member in memberList:
            sheet1.write(i, 0, member.name)
            i+=1
            print("Added " + member.name)
        channelList = message.guild.text_channels
        x = 1
        for channel in channelList:
            sheet1.write(0, x, channel.name)
            x+=1
            print("Added #" + channel.name)
        wb.save("Beep.xls")
                
    if message.content.startswith('$serverStatCount'):
        server = message.guild.text_channels
        for channel in server:
            async for message in channel.history(limit=200):
                x = 0
                y = 0
                sheet = wr.sheet_by_index(0) 
                for i in range(sheet.nrows):
                    if sheet.cell_value(i, 0) == message.channel:
                        x = i
                for i in range(sheet.ncols):
                    if sheet.cell_value(0,i) == message.author:
                        y = i
                val = int(sheet.cell_value(x,y))
                var = val + 1
                sheet.write(x,y,var)
client.run('NjYyNzg4MTk1NzM3NDAzNDQy.Xg_Dqg.RvP5k1D4dWeg0tqomlXiaHz7QQg')