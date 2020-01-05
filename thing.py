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
        zed = 0
        memberList = message.guild.members
        memberQuant = len(memberList)
        for member in memberList:
            sheet1.write(i, 0, member.name)
            i+=1
            print("Added " + member.name)
        channelList = message.guild.text_channels
        channelQuant = len(channelList)
        x = 1
        for channel in channelList:
            sheet1.write(0, x, channel.name)
            x+=1
            print("Added #" + channel.name)
        wb.save("Beep.xls")
        server = message.guild.text_channels       
        for channel in server:
            y = 1
            authorMessageQuant = {
            }
            for member in memberList:
                authorMessageQuant[member.name] = 0
            async for message in channel.history(limit=None):
                for member in memberList:
                    if member.name == message.author.name:
                        #print("Found author!")
                        authorMessageQuant[member.name] = str(int(authorMessageQuant[member.name])+ 1)
            sheet = wr.sheet_by_index(0) 
            for i in range(sheet.ncols):
                if sheet.cell_value(0,i) == channel.name:
                    print("Channel is found!") 
                    y = i
            for member in authorMessageQuant:
                x = 1
                for i in range(sheet.nrows):
                    if sheet.cell_value(i,0) == member:
                        print("Author is found!") 
                        x = i
                print(str(x) + "/" + str(memberQuant) + ";" + str(y) + "/" + str(channelQuant))
                messageCount = authorMessageQuant[member]
                sheet1.write(x, y, int(messageCount))
                zed = y
                wb.save("Beep.xls")
        x = 1
        zed +=1
        sheet1.write(0, zed, "Message Total:")
        for member in memberList:
            timesThroughAlphabet = len(channelList) / 26
            alphabet = {
                
            }
            sheet1.write(x, zed, xlwt.Formula("=SUM"))
            x+=1
            print("Zed: " + str(zed) + ", x: " + str(x))
        wb.save("Beep.xls")
        x = 1
        zed +=1
        sheet1.write(0, zed, "Date Joined:")
        for member in memberList:
            dateTimeFull = member.joined_at
            dateTimeSplit = str(dateTimeFull).split()
            sheet1.write(x, zed, dateTimeSplit[0])
            x+=1
            print("Zed: " + str(zed) + ", x: " + str(x))
        wb.save("Beep.xls")
        print("Complete!")


                

                

            
                


                    




                

client.run('NjYyNzg4MTk1NzM3NDAzNDQy.Xg_Dqg.RvP5k1D4dWeg0tqomlXiaHz7QQg')