import requests
import json
from datetime import datetime
import xlwt 
from xlwt import Workbook 
import sys
# python3 <filename> <token>
n = len(sys.argv)
if n != 2:
    print('Please pass valid token')
    exit(0)
token = sys.argv[1]
print("Start of pulling slack data")
headers = {'Authorization': 'Bearer ' + token}
print("Start of users data  ")
baseUrl = 'https://slack.com/'

url = baseUrl + 'api/users.list'
wb = Workbook() 
style_string = "font: bold on; borders: bottom dashed"
style = xlwt.easyxf(style_string)
users = wb.add_sheet('USERS') 
users.write(0, 0, "ID",style=style)
users.write(0, 1, "NAME",style=style)
users.write(0, 2, "TEAM_ID",style=style)
users.write(0, 3, "FULL_NAME",style=style)
users.write(0, 4, "TITLE",style=style)
users.write(0, 5, "PHONE",style=style)
users.write(0, 6, "SKYPE",style=style)
usersRow = 1  
usersRemaining = True
cursor = str() 
while  usersRemaining:
    response = requests.get(url,headers=headers,timeout=300)
    result =  response.json()
    
    if result["ok"]:
        if "members" in result:
            for user in result["members"]:
                users.write(usersRow, 0, user["id"])
                users.write(usersRow, 1, user["name"])
                users.write(usersRow, 2, user["team_id"])
                users.write(usersRow, 3, user["profile"]["real_name"])
                users.write(usersRow, 4, user["profile"]["title"])
                users.write(usersRow, 5, user["profile"]["phone"])
                users.write(usersRow, 6, user["profile"]["skype"])    
                usersRow += 1  
        else:
            break
        cursor = result["response_metadata"]["next_cursor"] 
        if len(cursor) == 0:
            break
        else:
            url = baseUrl + 'api/users.list?limit=200' + '&cursor=' + cursor
	
    else:
        print("Request failed for url:",url)
        print("Error :",result["error"])
        break

print("End of users data")

print("Start of channels data  ")
channelUrl= baseUrl + 'api/conversations.list?limit=100&types=public_channel,private_channel'
channels = wb.add_sheet('CHANNELS') 
channels.write(0, 0, "CHANNEL_ID",style=style)
channels.write(0, 1, "CHANNEL_NAME",style=style)
channels.write(0, 2, "CREATOR",style=style)
channels.write(0, 3, "PURPOSE",style=style)
channels.write(0, 4, "TOPIC",style=style)
channelsRow = 1

messages = wb.add_sheet('MESSAGES') 
messages.write(0, 0, "CHANNEL_ID",style=style)
messages.write(0, 1, "USER",style=style)
messages.write(0, 2, "MESSAGE",style=style)
messages.write(0, 3, "TIMESTAMP",style=style)
messagesRow = 1  


replies = wb.add_sheet('REPLIES') 
replies.write(0, 0, "CHANNEL_ID",style=style)
replies.write(0, 1, "USER",style=style)
replies.write(0, 2, "MESSAGE",style=style)
replies.write(0, 3, "TIMESTAMP",style=style)
repliesRow = 1  

channelsRemaining = True
cursor = str() 
while channelsRemaining:
    chResponse = requests.get(channelUrl,headers=headers,timeout=300)
    chResult =  chResponse.json()
    
    if chResult["ok"]:
        if "channels" in chResult:
            for channel in chResult["channels"]:
                channels.write(channelsRow, 0, channel["id"])
                channels.write(channelsRow, 1, channel["name"])
                channels.write(channelsRow, 2, channel["created"])
                channels.write(channelsRow, 3, channel["purpose"]["value"])
                channels.write(channelsRow, 4, channel["topic"]["value"])
                channelsRow += 1
                # start of messages
                messagecursor = str() 
                messageUrl = baseUrl + 'api/conversations.history?channel=' +channel["id"]
                messagesRemain = True
                while  messagesRemain:
                    memResponse = requests.get(messageUrl,headers=headers,timeout=300)
                    memResult = memResponse.json()
                    
                    if memResult["ok"]:
                        if "messages" in memResult:
                            for message in memResult["messages"]:
                                messages.write(messagesRow, 0, channel["id"] + '&ts=' +message["ts"])
                                if "user" in message:
                                    messages.write(messagesRow, 1, message["user"])
                                if "username" in message:
                                    messages.write(messagesRow, 1, message["username"])
                                messages.write(messagesRow, 2, message["text"])
                                messages.write(messagesRow, 3, message["ts"])
                                messagesRow +=1
                                # start of   replies
                                repliecursor = str() 
                                replieUrl = baseUrl +  'api/conversations.replies?channel=' +channel["id"] + '&ts=' +message["ts"]
                                repliesRemain = True
                                while  repliesRemain:
                                    repliesResponse = requests.get(replieUrl,headers=headers,timeout=300)
                                    repliesResult = repliesResponse.json()
                                    
                                    if repliesResult["ok"]:
                                        if "messages" in repliesResult:
                                            for replie in repliesResult["messages"]:
                                                replies.write(repliesRow, 0, channel["id"])
                                                if "user" in replie:
                                                    replies.write(repliesRow, 1, replie["user"])
                                                if "username" in replie:
                                                    replies.write(repliesRow, 1, replie["username"])
                                                replies.write(repliesRow, 2, replie["text"])
                                                replies.write(repliesRow, 3, replie["ts"])
                                                repliesRow +=1

                                        
                                        else:
                                            break
                                        if "response_metadata" in repliesResult:
                                            repliecursor = repliesResult["response_metadata"]["next_cursor"]
                                        else:
                                            repliecursor=""
                                        if len(repliecursor) == 0:
                                            break
                                        else:
                                            replieUrl = baseUrl + 'api/conversations.replies?limit=200' + '&cursor=' + repliecursor     + '&channel=' +channel["id"]   + '&ts=' +message["ts"]  
                                        print(repliecursor)
                                        

                                    else:
                                        print("Request failed for url:",replieUrl)
                                        print("Error:",repliesResult["error"])
                                        break

                                #end of replies
                        
                        else:
                            break
                        if "response_metadata" in memResult:
                            messagecursor = memResult["response_metadata"]["next_cursor"]
                        else:
                            messagecursor=""
                        if len(messagecursor) == 0:
                            break
                        else:
                            messageUrl = baseUrl + 'api/conversations.history?limit=200' + '&cursor=' + messagecursor     + '&channel=' +channel["id"]
                        print(messagecursor)
                    else:
                        print("Request failed for url:",messageUrl)
                        print("Error:",memResult["error"])
                        break


                # end here for messages  
        
        else:
            break
        channelCursor = chResult["response_metadata"]["next_cursor"] 
        if len(channelCursor) == 0:
            break
        else:
            channelUrl = baseUrl + 'api/conversations.list?types=public_channel,private_channel&limit=200' + '&cursor=' + channelCursor
    
    else:
        print("Request failed for url:",channelUrl)
        print("Error :",chResult["error"])
        break
        
	
print("End of channels data")


file =datetime.now().strftime("%d-%m-%Y %H:%M:%S") + "-SlackData.xls"
wb.save(file)
print("End of pulling slack data")
