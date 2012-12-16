try:
    from config_secrets import *
    myplan_secret_existed=True
except:
    myplan_secret_existed=False

__author__ = 'Max'
  #!/usr/bin/env python
#import atom
import gdata.docs.data
import gdata.docs.client
import gdata.docs.service
import gdata.spreadsheet.service


default_source   = 'domain-spreadhseetdownloader-0.01'
default_spreadsheetid='spreadsheet:0Ao993IR9m-38dDdjN3pGNFViQlotQXY3TkRVV2c1MUE'

def extract_token():



    gs_client     = gdata.spreadsheet.service.SpreadsheetsService()
    gs_client.ssl = True
    #gs_client.debug = True

    gs_client.ClientLogin(
                      email,
                      password,
                      source,
                     )

    token=gs_client.GetClientLoginToken()
    print token
    return token

def SpreadsheetsEnum(token):
    gs_client  = gdata.spreadsheet.service.SpreadsheetsService()
    gs_client.ssl = True
    gs_client.SetClientLoginToken(token)
    q = gdata.spreadsheet.service.DocumentQuery()
    q['title'] = "CRM"
    #q['title-exact'] = 'true'
    feed = gs_client.GetSpreadsheetsFeed(query=q)


    #gd_client.GetResource()
    #allres=gd_client.GetAllResources()
    for res in feed.entry:
        spreadsheet_id = res.id.text.rsplit('/',1)[1]
        title=res.title.text
        print title,spreadsheet_id
        break
        sfeed = gs_client.GetWorksheetsFeed(spreadsheet_id)
        for sres in sfeed.entry:
            worksheet_id = sres.id.text.rsplit('/',1)[1]
            title=sres.title.text
            print "   ",title,worksheet_id
        #gs_client.Export(spreadsheet_id, 'spreadsheeet.csv')
        rows = gs_client.GetListFeed(spreadsheet_id, worksheet_id).entry
        #for row in rows:
        #    for key in row.custom:
        #        print " %s: %s" % (key, row.custom[key].text)
        #    print

def loginnangetentry(resid,login):
    client     = gdata.docs.client.DocsClient()
    client.ssl = True

    uname=""
    pwd=""
    if login:
        uname=login[0]
        pwd=login[1]
    else:
        if myplan_secret_existed:
            uname=googledocs_login
            pwd=googledocs_password
    client.ClientLogin(
        uname,
        pwd,
        default_source,
        )

    if len(resid)<1:
            resid=default_spreadsheetid

    entry=client.GetResourceById(resid)

    return client, entry
def download_spreadsheet(resid,filename, login=None):
    download_item('spreadsheet:'+resid,filename,login)
def download_item(resid,filename,login):



    client, entry =loginnangetentry(resid,login)

    client.DownloadResource(entry,filename,extra_params={'gid': 0, 'exportFormat': 'xls'})

def upload_item(resid,filename, login=None):

    client, entry =loginnangetentry(resid,login)

    media= gdata.data.MediaSource(file_path=filename, content_type="application/vnd.ms-excel")


    try:
        client.UpdateResource(entry, media)
    except gdata.client.RequestError as re:
        if re.message.find("<internalReason>Sorry, there was an error saving the file. Please try again.</internalReason>")>0:
            pass
        else:
            raise
def findgdoc():
    client     = gdata.docs.client.DocsClient()
    email = raw_input("login: ")
    password = raw_input("password: ")
    str = raw_input("search pattern: ")
    source   = 'domain-spreadhseetdownloader-0.01'

    print "login..."
    if len(email)<1:
        email=googledocs_login

    if len(password)<1:
        password=googledocs_password


    token=client.ClientLogin(
                      email,
                      password,
                      source,
                     )


    print "enum documents..."
    #if len(str)<1:
    #    str="books"
    str=str.lower()

    allres=client.GetAllResources()
    print len(allres), "documents total"
    for doc in allres:
        t=doc.title.text.lower()
        id=doc.id.text
        link=doc.GetSelfLink()
        res_id=doc.resource_id.text

        if len(str)>0:
            if t.find(str)>=0:
                print "'"+t+"'", res_id
        else:
            print "'"+t+"'", res_id

    print "done."

if __name__ == "__main__":
    findgdoc()



#upload_item('spreadsheet:0Ao993IR9m-38dDdjN3pGNFViQlotQXY3TkRVV2c1MUE',"c:/5b.xls")