import requests
import json
from datetime import datetime
import xlwt 
from xlwt import Workbook 
import sys
# python3 <filename>  <owner>  <token>  <repo>
n = len(sys.argv)
if n != 4:
    print('Please pass valid 3 arguments  owner, token repo')
    exit(0)

owner = sys.argv[1]
token = sys.argv[2]
repo = sys.argv[3]
print("Start of pulling data from Github")


print("Start of contributors data  ")
baseUrl = 'https://api.github.com/'
url = baseUrl + 'repos/' + owner + '/' + repo + '/contributors?page{0}&per_page=100'
headers = {'Authorization': 'Bearer ' + token}
wb = Workbook() 
style_string = "font: bold on; borders: bottom dashed"
style = xlwt.easyxf(style_string)
collaborators = wb.add_sheet('Contributors') 
collaborators.write(0, 0, "ID",style=style)
collaborators.write(0, 1, "NAME",style=style)
collaborators.write(0, 2, "REPOSITORY",style=style)
collaboratorsrow = 1  

colabRemaining = True

while  colabRemaining:
    response = requests.get(url,headers=headers,timeout=300)
    result =  response.json()
    
    if response.status_code == 200:
        if "link" in response.headers:
            links = response.headers['link'].split(',')
            url = None
            for link in links:
                if 'rel="next"' in link:
                    url = link[link.find("<")+1:link.find(">")]
            for item in result:
                collaborators.write(collaboratorsrow, 0, item["id"])
                collaborators.write(collaboratorsrow, 1, item["login"])
                collaborators.write(collaboratorsrow, 2, repo)
                collaboratorsrow += 1

            if url is None:
                break	
        else:
            for item in result:
                collaborators.write(collaboratorsrow, 0, item["id"])
                collaborators.write(collaboratorsrow, 1, item["login"])
                collaborators.write(collaboratorsrow, 2, repo)
                collaboratorsrow += 1
            break
    else:
        print("Request got failed for url :",url)	
        print(result["message"])
        break
	
print("End of contributors data")

print("Start Issues Comments")
issueApiUrl = baseUrl + 'repos/' + owner + '/' + repo + '/issues/comments?page{0}&per_page=100'	
issuesComments = wb.add_sheet('Issues-Comments') 
issuesComments.write(0, 0, "ISSUEID",style=style)
issuesComments.write(0, 1, "USERID",style=style)
issuesComments.write(0, 2, "USERNAME",style=style)
issuesComments.write(0, 3, "URL",style=style)
issuesComments.write(0, 4, "DESCRIPTION",style=style)
issuesComments.write(0, 5, "CREATED_AT",style=style)
issuesComments.write(0, 6, "UPDATED_AT",style=style)
issuesCommentsRow = 1

issueRemaining = True

while  issueRemaining:
    issueResponse = requests.get(issueApiUrl,headers=headers,timeout=300)  
    issueResult = issueResponse.json()
   
    if issueResponse.status_code == 200:
        if "link" in issueResponse.headers:
            links = issueResponse.headers['link'].split(',')
            issueApiUrl = None
            for link in links:
                if 'rel="next"' in link:
                    issueApiUrl = link[link.find("<")+1:link.find(">")]
            for issue in issueResult :      
                issuesComments.write(issuesCommentsRow, 0, issue["id"])
                issuesComments.write(issuesCommentsRow, 1, issue["user"]["id"])  
                issuesComments.write(issuesCommentsRow, 2, issue["user"]["login"])
                if issue["html_url"] is not None:        
                    issuesComments.write(issuesCommentsRow, 3, issue["html_url"][:32766])
                if issue["body"] is not None:
                    issuesComments.write(issuesCommentsRow, 4, issue["body"][:32766])  
                issuesComments.write(issuesCommentsRow, 5, issue["created_at"])     
                issuesComments.write(issuesCommentsRow, 6, issue["updated_at"])      
                issuesCommentsRow += 1
            

            if issueApiUrl is None:
                break	
        else:
            for issue in issueResult :      
                issuesComments.write(issuesCommentsRow, 0, issue["id"])
                issuesComments.write(issuesCommentsRow, 1, issue["user"]["id"])  
                issuesComments.write(issuesCommentsRow, 2, issue["user"]["login"])
                if issue["html_url"] is not None:        
                    issuesComments.write(issuesCommentsRow, 3, issue["html_url"][:32766])
                if issue["body"] is not None:
                    issuesComments.write(issuesCommentsRow, 4, issue["body"][:32766])  
                issuesComments.write(issuesCommentsRow, 5, issue["created_at"])     
                issuesComments.write(issuesCommentsRow, 6, issue["updated_at"])      
                issuesCommentsRow += 1
            break
    else:
        print("Request got failed for url :",issueApiUrl)
        print(issueResult["message"])
        break



print("End Issues Comments")


print("Start Pull Requests")
pullUrl = baseUrl + 'repos/' + owner + '/' + repo + '/pulls?page{0}&per_page=100'
pullReq = wb.add_sheet('PullRequests') 
pullReq.write(0, 0, "PRNUMBER",style=style)
pullReq.write(0, 1, "USERID",style=style)
pullReq.write(0, 2, "USERNAME",style=style)
pullReq.write(0, 3, "URL",style=style)
pullReq.write(0, 4, "TITLE",style=style)
pullReq.write(0, 5, "STATUS",style=style)
pullReq.write(0, 6, "DESCRIPTION",style=style)
pullReq.write(0, 7, "REVIEWERS",style=style)
pullReq.write(0, 8, "CREATED_AT",style=style)
pullReq.write(0, 9, "UPDATED_AT",style=style)
pullReqRow = 1
pullRemaining = True
while  pullRemaining:
    pullResponse = requests.get(pullUrl,headers=headers,timeout=300)  
    pullResult = pullResponse.json()
   
    if pullResponse.status_code == 200:
        if "link" in pullResponse.headers:
            links = pullResponse.headers['link'].split(',')
            pullUrl = None
            for link in links:
                if 'rel="next"' in link:
                    pullUrl = link[link.find("<")+1:link.find(">")]
            for pull in pullResult :    
                pullReq.write(pullReqRow, 0, pull["number"])
                pullReq.write(pullReqRow, 1, pull["user"]["id"])  
                pullReq.write(pullReqRow, 2, pull["user"]["login"])
                if pull["html_url"] is not None:        
                    pullReq.write(pullReqRow, 3, pull["html_url"][:32766])
                if pull["title"] is not None: 
                    pullReq.write(pullReqRow, 4, pull["title"][:32766])  
                pullReq.write(pullReqRow, 5, pull["state"])  
                pullReq.write(pullReqRow, 6, pull["body"])  
                reviewersList= []
                for reviewer in pull["requested_reviewers"] :
                    reviewersList.append(reviewer["login"])    
                pullReq.write(pullReqRow,7, reviewersList)     

                pullReq.write(pullReqRow,8, pull["created_at"])     
                pullReq.write(pullReqRow, 9,pull["updated_at"])      
                pullReqRow += 1

            if pullUrl is None:
                break	
        else:
            for pull in pullResult :    
                pullReq.write(pullReqRow, 0, pull["number"])
                pullReq.write(pullReqRow, 1, pull["user"]["id"])  
                pullReq.write(pullReqRow, 2, pull["user"]["login"]) 
                if pull["html_url"] is not None:       
                    pullReq.write(pullReqRow, 3, pull["html_url"][:32766])
                if pull["title"] is not None:
                    pullReq.write(pullReqRow, 4, pull["title"][:32766])  
                pullReq.write(pullReqRow, 5, pull["state"])  
                pullReq.write(pullReqRow, 6, pull["body"])  
                reviewersList= []
                for reviewer in pull["requested_reviewers"] :
                    reviewersList.append(reviewer["login"])    
                pullReq.write(pullReqRow,7, reviewersList)     

                pullReq.write(pullReqRow,8, pull["created_at"])     
                pullReq.write(pullReqRow, 9,pull["updated_at"])      
                pullReqRow += 1
            break
    else:
        print("Request got failed for url :",pullUrl)
        print(pullResult["message"])
        break
  
print("End Pull Requests")

print("Start Repos Stars")
starsApiUrl = baseUrl + 'repos/' + owner + '/' + repo + '/stargazers?page{0}&per_page=100'	
stars = wb.add_sheet('Stargazers') 
stars.write(0, 0, "USERID",style=style)
stars.write(0, 1, "USERNAME",style=style)
stars.write(0, 2, "URL",style=style)
starsRow = 1
starsRemaining = True
while  starsRemaining:
    starsResponse = requests.get(starsApiUrl,headers=headers,timeout=300)  
    starsResult = starsResponse.json()
   
    if starsResponse.status_code == 200:
        if "link" in starsResponse.headers:
            links = starsResponse.headers['link'].split(',')
            starsApiUrl = None
            for link in links:
                if 'rel="next"' in link:
                    starsApiUrl = link[link.find("<")+1:link.find(">")]
            for star in starsResult :      
                stars.write(starsRow, 0, star["id"])
                stars.write(starsRow, 1, star["login"])  
                stars.write(starsRow, 2, star["html_url"])          
                starsRow += 1

            if starsApiUrl is None:
                break	
        else:
            for star in starsResult :      
                stars.write(starsRow, 0, star["id"])
                stars.write(starsRow, 1, star["login"])  
                stars.write(starsRow, 2, star["html_url"])          
                starsRow += 1
            break
    else:
        print("Request got failed for url :",starsApiUrl)
        print(starsResult["message"])
        break

  
print("End Repo Stars")


print("Start Repos Forks")
forksApiUrl = baseUrl + 'repos/' + owner + '/' + repo + '/forks?page{0}&per_page=100'	
forks = wb.add_sheet('Forks') 
forks.write(0, 0, "OWNER_ID",style=style)
forks.write(0, 1, "OWNER_NAME",style=style)
forks.write(0, 2, "REPO_NAME",style=style)
forks.write(0, 3, "REPO_URL",style=style)
forks.write(0, 4, "DESCRIPTION",style=style)
forks.write(0, 5, "CREATED_AT",style=style)
forks.write(0, 6, "UPDATED_AT",style=style)
forksRow = 1
forksRemaining = True
while  forksRemaining:
    forksResponse = requests.get(forksApiUrl,headers=headers,timeout=300)  
    forksResult = forksResponse.json()
   
    if forksResponse.status_code == 200:
        if "link" in forksResponse.headers:
            links = forksResponse.headers['link'].split(',')
            forksApiUrl = None
            for link in links:
                if 'rel="next"' in link:
                    forksApiUrl = link[link.find("<")+1:link.find(">")]
            for fork in forksResult :      
                forks.write(forksRow, 0, fork["owner"]["id"])
                forks.write(forksRow, 1, fork["owner"]["login"])  
                forks.write(forksRow, 2, fork["name"])  
                if fork["html_url"] is not None: 
                    forks.write(forksRow, 3, fork["html_url"][:32766]) 
                if fork["description"] is not None:     
                    forks.write(forksRow, 4, fork["description"][:32766])          
                forks.write(forksRow,5, fork["created_at"])     
                forks.write(forksRow, 6,fork["updated_at"])    
                forksRow += 1

            if forksApiUrl is None:
                break	
        else:
            for fork in forksResult :      
                forks.write(forksRow, 0, fork["owner"]["id"])
                forks.write(forksRow, 1, fork["owner"]["login"])  
                forks.write(forksRow, 2, fork["name"])  
                if fork["html_url"] is not None: 
                    forks.write(forksRow, 3, fork["html_url"][:32766])  
                if fork["description"] is not None:    
                    forks.write(forksRow, 4, fork["description"][:32766])          
                forks.write(forksRow,5, fork["created_at"])     
                forks.write(forksRow, 6,fork["updated_at"])    
                forksRow += 1 
            break
    else:
        print("Request got failed for url :",forksApiUrl)
        print(forksResult["message"])
        break 
print("End Repo Forks")
print("Start Repos Issues")
issuesApiUrl = baseUrl + 'repos/' + owner + '/' + repo + '/issues?page{0}&per_page=100'	
issues = wb.add_sheet('Issues') 
issues.write(0, 0, "ISSUE_ID",style=style)
issues.write(0, 1, "PULL_NUMBER",style=style)
issues.write(0, 2, "TITLE",style=style)
issues.write(0, 3, "URL",style=style)
issues.write(0, 4, "USER_ID",style=style)
issues.write(0, 5, "USER_NAME",style=style)
issues.write(0, 6, "STATUS",style=style)
issues.write(0, 7, "CREATED_AT",style=style)
issues.write(0, 8, "UPDATED_AT",style=style)

issuesRow = 1
issuesRemaining = True
while  issuesRemaining:
    
    issuesResponse = requests.get(issuesApiUrl,headers=headers,timeout=300)  
    issuesResult = issuesResponse.json()
    if issuesResponse.status_code == 200:
        if "link" in issuesResponse.headers:
            links = issuesResponse.headers['link'].split(',')
            issuesApiUrl = None
            for link in links:
                if 'rel="next"' in link:
                    issuesApiUrl = link[link.find("<")+1:link.find(">")]
            for issue in issuesResult :      
                issues.write(issuesRow, 0, issue["id"])
                issues.write(issuesRow, 1, issue["number"])  
                if issue["title"] is not None:
                    issues.write(issuesRow, 2, issue["title"][:32766])  
                if issue["html_url"] is not None:
                    issues.write(issuesRow, 3, issue["html_url"][:32766])     
                issues.write(issuesRow, 4, issue["user"]["id"])          
                issues.write(issuesRow,5, issue["user"]["login"])     
                issues.write(issuesRow, 6,issue["state"])  
                issues.write(issuesRow, 7,issue["created_at"])  
                issues.write(issuesRow, 8,issue["updated_at"])  
                issuesRow += 1

            if issuesApiUrl is None:
                break	
        else:
            for issue in issuesResult :      
                issues.write(issuesRow, 0, issue["id"])
                issues.write(issuesRow, 1, issue["number"])
                if issue["title"] is not None:  
                    issues.write(issuesRow, 2, issue["title"][:32766])
                if issue["html_url"] is not None:  
                    issues.write(issuesRow, 3, issue["html_url"][:32766])     
                issues.write(issuesRow, 4, issue["user"]["id"])          
                issues.write(issuesRow,5, issue["user"]["login"])     
                issues.write(issuesRow, 6,issue["state"])  
                issues.write(issuesRow, 7,issue["created_at"])  
                issues.write(issuesRow, 8,issue["updated_at"])  
                issuesRow += 1  
            break
    else:
        print("Request got failed for url :",issuesApiUrl)
        print(issuesResult["message"])
        break 
print("End Repo Issues")
print("Start Discussions")
teamsApiUrl = baseUrl + 'orgs/' + owner + '/teams?page{0}&per_page=100' 	
discussion = wb.add_sheet('Discussion') 
discussion.write(0, 0, "TEAM_ID",style=style)
discussion.write(0, 1, "TEAM_NAME",style=style)   
discussion.write(0, 2, "TEAM_SLUG",style=style)                  
discussion.write(0, 3, "TEAM_DESCRIPTION",style=style)                  
discussion.write(0, 4, "DISCUSSION_ID",style=style)    
discussion.write(0, 5, "TITLE",style=style)                  
discussion.write(0, 6, "AUTHOR_ID",style=style)                  
discussion.write(0, 7, "AUTHOR_NAME",style=style)                  
discussion.write(0, 8, "BODY",style=style)                  
discussionRow = 1 
disucssRemaining = True
while  disucssRemaining:
    teamsResponse = requests.get(teamsApiUrl,headers=headers,timeout=300)  
    teamsResult = teamsResponse.json()
    
    if teamsResponse.status_code == 200:
        if "link" in teamsResponse.headers:
            links = teamsResponse.headers['link'].split(',')
            teamsApiUrl = None
            for link in links:
                if 'rel="next"' in link:
                    teamsApiUrl = link[link.find("<")+1:link.find(">")]
            for team in teamsResult :      
                discussionsApiUrl = baseUrl + 'orgs/' + owner + '/teams/' + team["slug"] + "/discussions?page{0}&per_page=100"
                discussionResponse = requests.get(discussionsApiUrl,headers=headers,timeout=300)  
                discussionResult = discussionResponse.json()                
                remaining = True
                if discussionResponse.status_code == 200:
                    while  remaining:
                        discussionResponse = requests.get(discussionsApiUrl,headers=headers,timeout=300)  
                        discussionResult = discussionResponse.json()
                        if "link" in discussionResponse.headers:
                            links = discussionResponse.headers['link'].split(',')
                            discussionsApiUrl =None
                            for link in links:
                                if 'rel="next"' in link:
                                    discussionsApiUrl = link[link.find("<")+1:link.find(">")]	
                            for record in discussionResult:
                                discussion.write(discussionRow, 0, team["id"])
                                discussion.write(discussionRow, 1, team["name"])  
                                discussion.write(discussionRow, 2, team["slug"]) 
                                if team["description"] is not None:  
                                    discussion.write(discussionRow, 3,team["description"][:32766])  
                                discussion.write(discussionRow, 4,record["number"])  
                                if record["title"] is not None: 
                                    discussion.write(discussionRow, 5,record["title"][:32766])  
                                discussion.write(discussionRow, 6,record["author"]["id"])  
                                discussion.write(discussionRow, 7,record["author"]["login"]) 
                                if record["body"] is not None:  
                                    discussion.write(discussionRow, 8,record["body"][:32766])  
                                discussionRow += 1				
                            if discussionsApiUrl  is None:
                                break  

                        else:
                            for record in discussionResult:
                                discussion.write(discussionRow, 0, team["id"])
                                discussion.write(discussionRow, 1, team["name"])  
                                discussion.write(discussionRow, 2, team["slug"])  
                                if team["description"] is not None: 
                                    discussion.write(discussionRow, 3,team["description"][:32766])  
                                discussion.write(discussionRow, 4,record["number"])
                                if record["title"] is not None:   
                                    discussion.write(discussionRow, 5,record["title"][:32766])  
                                discussion.write(discussionRow, 6,record["author"]["id"])  
                                discussion.write(discussionRow, 7,record["author"]["login"])  
                                if record["body"] is not None: 
                                    discussion.write(discussionRow, 8,record["body"][:32766])  
                                discussionRow += 1	
                            break
                else:
                    print("Request got failed for url :",discussionsApiUrl)
                    print(discussionResult["message"])
                    break

            if teamsApiUrl is None:
                break	
        else:
            for team in teamsResult :      
                discussionsApiUrl = baseUrl + 'orgs/' + owner + '/teams/' + team["slug"] + "/discussions?page{0}&per_page=100"
                remaining = True
                while  remaining:
                   
                    discussionResponse = requests.get(discussionsApiUrl,headers=headers,timeout=300)  
                    discussionResult = discussionResponse.json()
                    if discussionResponse.status_code == 200:
                        if "link" in discussionResponse.headers:
                            links = discussionResponse.headers['link'].split(',')
                            discussionsApiUrl =None
                            for link in links:
                                if 'rel="next"' in link:
                                    discussionsApiUrl = link[link.find("<")+1:link.find(">")]	
                            for record in discussionResult:
                                discussion.write(discussionRow, 0, team["id"])
                                discussion.write(discussionRow, 1, team["name"])  
                                discussion.write(discussionRow, 2, team["slug"]) 
                                if team["description"] is not None: 
                                    discussion.write(discussionRow, 3,team["description"][:32766])  
                                discussion.write(discussionRow, 4,record["number"])  
                                if record["title"] is not None:
                                    discussion.write(discussionRow, 5,record["title"][:32766])  
                                discussion.write(discussionRow, 6,record["author"]["id"])  
                                discussion.write(discussionRow, 7,record["author"]["login"])  
                                if record["body"] is not None:
                                    discussion.write(discussionRow, 8,record["body"][:32766])  
                                discussionRow += 1				
                            if discussionsApiUrl  is None:
                                break  

                        else:
                            for record in discussionResult:
                                discussion.write(discussionRow, 0, team["id"])
                                discussion.write(discussionRow, 1, team["name"])  
                                discussion.write(discussionRow, 2, team["slug"])
                                if team["description"] is not None:  
                                    discussion.write(discussionRow, 3,team["description"][:32766])  
                                discussion.write(discussionRow, 4,record["number"]) 
                                if record["title"] is not None: 
                                    discussion.write(discussionRow, 5,record["title"][:32766])  
                                discussion.write(discussionRow, 6,record["author"]["id"])  
                                discussion.write(discussionRow, 7,record["author"]["login"]) 
                                if record["body"] is not None: 
                                    discussion.write(discussionRow, 8,record["body"][:32766])  
                                discussionRow += 1	
                            break
                    else:
                        print("Request got failed for url :",discussionsApiUrl)
                        print(discussionResult["message"])
                        break

                    
            break
    else:
        print("Request got failed for url :",teamsApiUrl)
        print(teamsResult["message"])
        break

    
     
            
print("End Discussions")

file =datetime.now().strftime("%d-%m-%Y %H:%M:%S") + "-GitHubData.xls"
wb.save(file)

print("End of pulling data from Github")


