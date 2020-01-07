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
    
    if message.content.startswith('&userInfo') or message.content.startswith('&userinfo'):
        messageContentList = message.content.split(" ")
        if len(messageContentList) > 1:
            userPing = messageContentList[1]
            print(userPing)
            userPingList = userPing.split('!')
            userPingThing = userPingList[1].split('>')
            userID = userPingThing[0]
            print(userID)
            user = client.get_user(int(userID))
            print("user name: " + user.name)                    
            messageContentList = message.content.split(" ")
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
                    print (positionOfUser)
                    totalMessagesSentServer = guildSheet.cell_value(positionOfUser, numberOfChannels - 3)
                    messageAverageOnServer = guildSheet.cell_value(positionOfUser, numberOfChannels)
                    totalMessagesSent = 0
                    messageAverageServer = ""
                    messageAverage = 0
                    for server in sheetList:
                        serverMembers = []
                        for i in range(server.nrows): 
                            print(server.cell_value(i, 0))
                            serverMembers.append(server.cell_value(i, 0))
                        if user.name in serverMembers:
                            locationUser = serverMembers.index(user.name)
                            totalMessageLoction = server.ncols - 4
                            averageMessageLocation = server.ncols - 1
                            print("Row: " + str(locationUser) + ", Col: " + str(totalMessageLoction))
                            totalMessagesSent = totalMessagesSent + server.cell_value(locationUser, totalMessageLoction)
                            print("Server Messages: " + str(server.cell_value(locationUser, totalMessageLoction)))
                            if messageAverage < server.cell_value(locationUser, averageMessageLocation) and server.cell_value(locationUser, averageMessageLocation) > 0:
                                messageAverage = server.cell_value(locationUser, averageMessageLocation)
                                messageAverageServer = server.name
                            print (str(messageAverage) + " on " + messageAverageServer)
                    embed = discord.Embed(title=user.name, color=0xFF9900)
                    if message.guild.name == "The CA Discord" and int(guildSheet.cell_value(positionOfUser, 12)) >= 1000:
                        embed.add_field(name="Messages Sent in Counting and Recursion: ", value=int(guildSheet.cell_value(positionOfUser, 12)), inline=False)
                        embed.add_field(name="Messages Sent Not in Counting and Recursion: ", value=int(totalMessagesSentServer) - int(guildSheet.cell_value(positionOfUser, 12)), inline=False)
                    else:
                        embed.add_field(name="Total Messages Sent on this Server", value=int(totalMessagesSentServer), inline=False)
                    embed.add_field(name="Average Messages on this Server Per Day", value=messageAverageOnServer, inline=False)
                    embed.add_field(name="Most Active Server", value=messageAverageServer, inline=False)
                    embed.add_field(name="Total Messages Sent", value=int(totalMessagesSent), inline=False)
                    embed.set_thumbnail(url=userURL)
                    embed.set_footer(text="Created by The Invisible Man", icon_url="https://cdn.discordapp.com/avatars/366709133195476992/01cb7c2c7f2007d8b060e084ea4eb6fd.png?size=512")
                    await message.channel.send(embed=embed)
    if message.content.startswith('&serverInfo'):
        wz = xlrd.open_workbook("C:\\Users\\Sebastian_Polge\\OneDrive-CaryAcademy\\Documents\\meNewBot\\Verity\\StatsBot\\Complete.xls")
        guildSheet = wz.sheet_by_name(message.guild.name)
        positionOfTotalColumn = guildSheet.ncols - 4
        positionOfAverageMessage = guildSheet.ncols -1
        activeMember = ""
        activeMemberActivity = 0
        activeMemberNoCount = ""
        activeMemberNoCountNum = 0
        totalMessages = 0
        for i in range(guildSheet.nrows):
            if (i != 0):
                totalMessages += guildSheet.cell_value(i, positionOfTotalColumn)
                if (guildSheet.cell_value(i, positionOfAverageMessage) > activeMemberActivity):
                    activeMemberActivity = guildSheet.cell_value(i, positionOfAverageMessage)
                    activeMember = guildSheet.cell_value(i, 0)
        for i in range(guildSheet.nrows):
            if (i != 0) and message.guild.name == "The CA Discord":
                beep = guildSheet.cell_value(i, guildSheet.ncols - 4) - guildSheet.cell_value(i, 12)/guildSheet.cell_value(i, guildSheet.ncols - 2)
                print(guildSheet.cell_value(i, 0), ": " + str(beep))
                if (beep > activeMemberNoCountNum):
                    activeMemberNoCountNum = beep
                    activeMemberNoCount = guildSheet.cell_value(i, 0)
        totalPerChannel = 0
        sumTotal = 0
        channelName = ""
        countingAndRecursionInt = -1
        endCol = guildSheet.ncols - 4
        for i in range(guildSheet.ncols):
            sumTotal = 0
            if (i != 0 and i < endCol):
                for v in range(guildSheet.nrows):
                  if (v != 0):
                      sumTotal+=int(guildSheet.cell_value(v, i))
                if (sumTotal > totalPerChannel):
                    if guildSheet.cell_value(0, i) != "counting-and-recursion":
                        channelName = guildSheet.cell_value(0, i)
                        totalPerChannel = sumTotal
                    else:
                        countingAndRecursionInt = sumTotal
        print (channelName +": " + str(totalPerChannel))
        memberCount = guildSheet.nrows - 1
        channelCount = guildSheet.ncols - 5
        owner = message.guild.owner.name
        embed = discord.Embed(title=message.guild.name, color=0xFF9900)
        embed.add_field(name="Total Messages Sent on this Server", value = int(totalMessages), inline=False)
        embed.add_field(name="Most Active Member", value=activeMember, inline=False)
        if message.guild.name == "The CA Discord":
            embed.add_field(name="Most Active Member, Not Including #counting-and-recursion", value=activeMemberNoCount, inline=False)
        if (countingAndRecursionInt > totalPerChannel):
            embed.add_field(name="Highest Message Count per Channel, Not Including #counting-and-recursion", value="#" + channelName, inline=False)
            embed.add_field(name="Highest Message Count per Channel", value="#counting-and-recursion", inline=False)
        else:
            embed.add_field(name="Highest Message Count per Channel", value=channelName, inline=False)
        embed.add_field(name="Number of Members", value=memberCount, inline=False)
        embed.add_field(name="Number of Channels", value=channelCount, inline=False)
        embed.add_field(name="Owner", value=owner, inline=False)
        embed.set_thumbnail(url=message.guild.icon_url)
        embed.set_footer(text="Created by The Invisible Man", icon_url="https://cdn.discordapp.com/avatars/366709133195476992/01cb7c2c7f2007d8b060e084ea4eb6fd.png?size=512")
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