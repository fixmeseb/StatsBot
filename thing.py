import discord
#This imports Discord. Named thing.py because my old bots had their main files in thing.js, and I'm sentimental. 
import xlrd
import xlwt
from xlwt import Workbook
from xlutils.copy import copy 
from xlrd import open_workbook
from datetime import date
from datetime import datetime
from discord.utils import get



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
        
    if message.content.startswith('&userInfo'):
        messageContentList = message.content.split(" ")
        if len(messageContentList) > 1:
            userPing = messageContentList[1]
            print(userPing)
            userPingList = userPing.split('!')
            userPingThing = userPingList[1].split('>')
            userID = userPingThing[0]
            print(userID)
            print("366709133195476992")
            user = client.get_user(int(userID))
            print(user.name)
            userURL = str(user.avatar_url)
            wz = xlrd.open_workbook("C:\\Users\\Sebastian_Polge\\OneDrive-CaryAcademy\\Documents\\meNewBot\\Verity\\StatsBot\\Complete.xls")
            sheet = wz.sheet_by_index(0)
            sheetList = wz.sheets()
            string = message.guild.name
            numberOfChannels = 0
            positionOfUser = 0                
            guildSheet = wz.sheet_by_name(string)
            numberOfChannels = guildSheet.ncols - 1
            for i in range(guildSheet.nrows): 
                print(guildSheet.cell_value(i, 0))
                if guildSheet.cell_value(i, 0) == user.name:
                    positionOfUser = i
            print (numberOfChannels - 3)
            print (positionOfUser + 1)
            totalMessagesSentServer = guildSheet.cell_value(positionOfUser, numberOfChannels - 3)
            for server in sheetList:
                serverMembers = []
                for i in range(server.nrows): 
                    print(server.cell_value(i, 0))
                    serverMembers.add(server.cell_value(i, 0))
                if serverMembers.__contains__
                     
            embed = discord.Embed(title=user.name, color=0xFF9900)
            embed.add_field(name="Total Messages Sent on this Server", value = totalMessagesSentServer, inline=False)
            embed.add_field(name="Most Active Server", value="Beep", inline=False)
            embed.add_field(name="Total Messages Sent", value="Bloop", inline=False)
            embed.set_thumbnail(url=userURL)
            await message.channel.send(embed=embed)                
        
    if message.content.startswith('&serverStatCount') and message.author.id == 366709133195476992:
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
            if channel.name != "robot-game" and channel.name != "starfall-private-space" and channel.name != "ca-nerd-squad":
                sheet1.write(0, x, channel.name)
                x+=1
                print("Added #" + channel.name)
        wb.save("Beep.xls")
        server = message.guild.text_channels       
        for channel in server:
            if channel.name != "robot-game" and channel.name != "starfall-private-space" and channel.name != "ca-nerd-squad":
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
                    print(str(x) + "/" + str(memberQuant) + ";" + str(y) + "/" + str(channelQuant) + " #" + str(channel.name))
                    messageCount = authorMessageQuant[member]
                    sheet1.write(x, y, int(messageCount))
                    zed = y + 1
                    wb.save("Beep.xls")
        x = 1
        sheet1.write(0, zed, "Message Total:")
        sheet1.write(0, zed + 1, "Date Joined:")
        sheet1.write(0, zed + 2, "Days on Server:")
        sheet1.write(0, zed + 3, "Message Average:")
        for member in memberList:
            string = ""
            n = y
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                string = chr(65 + remainder) + string
            print(string)
            value1 = "B" + str(x + 1)
            value2 = string + str(x + 1)
            sheet1.write(x, zed, xlwt.Formula("SUM(" + value1 + ":" + value2 + ")"))
            print("Zed: " + str(zed) + ", x: " + str(x))
            dateTimeFull = member.joined_at
            dateTimeSplit = str(dateTimeFull).split()
            sheet1.write(x, zed + 1, dateTimeSplit[0])
            print("Zed: " + str(zed) + ", x: " + str(x))
            dateTimeFull = member.joined_at
            dateJoined = datetime.date(dateTimeFull)
            dateTimeSplit = str(dateTimeFull).split()
            currentDate = datetime.date(datetime.now())
            delta = currentDate - dateJoined
            sheet1.write(x, zed + 2, delta.days)
            sheet1.write(x, zed + 3, xlwt.Formula("SUM(" + value1 + ":" + value2 + ")/" + str(delta.days))) 
            print("Zed: " + str(zed) + ", x: " + str(x))
            x+=1
        wb.save("Beep.xls")
        print("Complete!")


                

                

            
                


                    




                

client.run('NjYyNzg4MTk1NzM3NDAzNDQy.Xg_Dqg.RvP5k1D4dWeg0tqomlXiaHz7QQg')