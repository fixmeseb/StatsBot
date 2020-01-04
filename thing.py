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
    if message.content.startswith('$setUpServer'):
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
        for i in range(len(memberList)):
            for j in range(len(channelList)):
                sheet1.write(i + 1,j + 1,0)
        wb.save("Beep.xls")
                
    if message.content.startswith('$serverStatCount'):
        server = message.guild.text_channels
        z = 0
        for channel in server:
            async for message in channel.history(limit=200):
                x = 0
                y = 0
                z+=1
                rb = open_workbook("C:\\Users\\Sebastian_Polge\\OneDrive-CaryAcademy\\Documents\\meNewBot\\Verity\\StatsBot\\Beep.xls",formatting_info=True)
                r_sheet = rb.sheet_by_index(0) # read only copy to introspect the file
                wb = copy(rb) # a writable copy (I can't read values out of this, only write to it)
                w_sheet = wb.get_sheet(0) # the sheet to write to within the writable copy     
                
                print("Author: " + message.author.name)
                print ("Channel: " + message.channel.name)
                sheet = wr.sheet_by_index(0) 
                for i in range(sheet.ncols):
                    print (sheet.cell_value(0, i) + "," + message.channel.name)
                    if sheet.cell_value(0,i) == message.channel.name:
                        print("Channel is found!") 
                        y = i
                for i in range(sheet.nrows):
                    print(sheet.cell_value(i, 0) + "," + message.author.name) 
                    if sheet.cell_value(i,0) == message.author.name:
                        print("Author is found!") 
                        x = i
                print(str(x) + ";" + str(y))
                beep = sheet.cell_value(x, y) + 1
                print("beep value: " + str(sheet.cell_value(x,y)))
                sheetR.write(x, y, beep)
                wb.save("Beep.xls")

                

client.run('NjYyNzg4MTk1NzM3NDAzNDQy.Xg_Dqg.RvP5k1D4dWeg0tqomlXiaHz7QQg')