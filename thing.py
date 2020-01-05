import discord
#This imports Discord. Named thing.py because my old bots had their main files in thing.js, and I'm sentimental. 
import xlrd
import xlwt
from xlwt import Workbook
from xlutils.copy import copy 
from xlrd import open_workbook



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
        
                
    if message.content.startswith('$serverStatCount'):
        wb = Workbook()
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
        server = message.guild.text_channels
        z = 0
        for channel in server:
            if (channel.name != "starfall-private-space" and channel.name != "robot-game" and channel.name != "ca-nerd-squad"):
                for member in message.guild.members:
                    messageCount = 0
                    x = 1
                    y = 1
                    z+=1
                    print("Author: " + member.name)
                    print ("Channel: " + message.channel.name)
                    sheet = wr.sheet_by_index(0) 
                    for i in range(sheet.ncols):
                        #print (sheet.cell_value(0, i) + "," + message.channel.name)
                        if sheet.cell_value(0,i) == channel.name:
                            print("Channel is found!") 
                            y = i
                    for i in range(sheet.nrows):
                        #print(sheet.cell_value(i, 0) + "," + member.name) 
                        if sheet.cell_value(i,0) == member.name:
                            print("Author is found!") 
                            x = i
                    print(str(x) + ";" + str(y))
                
                    async for message in channel.history(limit=None):
                        if (message.author.name == member.name):
                            messageCount+=1
                    print("x: " + str(x) + ", y: " + str(y))
                    sheet1.write(x, y, int(messageCount))
                    wb.save("Beep.xls")
                print("Completed first channel")



                

client.run('NjYyNzg4MTk1NzM3NDAzNDQy.Xg_Dqg.RvP5k1D4dWeg0tqomlXiaHz7QQg')