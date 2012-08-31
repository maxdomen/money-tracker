__author__ = 'Max'
  #!/usr/bin/env python
import atom
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

def loginnangetentry(resid):
    client     = gdata.docs.client.DocsClient()
    client.ssl = True

    client.ClientLogin(
                          default_email,
                          default_password,
                          default_source,
                         )

    if len(resid)<1:
            resid=default_spreadsheetid

    entry=client.GetResourceById(resid)

    return client, entry

def download_item(resid,filename):



    client, entry =loginnangetentry(resid)

    client.DownloadResource(entry,filename,extra_params={'gid': 0, 'exportFormat': 'xls'})

def upload_item(resid,filename):

    client, entry =loginnangetentry(resid)

    media= gdata.data.MediaSource(file_path=filename, content_type="application/vnd.ms-excel")


    try:
        client.UpdateResource(entry, media)
    except gdata.client.RequestError as re:

        print "ignore ", re
def findgdoc():
    client     = gdata.docs.client.DocsClient()
    email = raw_input("login: ")
    password = raw_input("password: ")
    str = raw_input("search pattern (books): ")
    source   = 'domain-spreadhseetdownloader-0.01'

    print "login..."
    if len(email)<1:
        email=default_email

    if len(password)<1:
        password=default_password


    token=client.ClientLogin(
                      email,
                      password,
                      source,
                     )


    print "enum documents..."
    if len(str)<1:
        str="books"
    str=str.lower()

    allres=client.GetAllResources()
    print len(allres), "documents total"
    for doc in allres:
        t=doc.title.text.lower()
        id=doc.id.text
        link=doc.GetSelfLink()
        res_id=doc.resource_id.text

        if t.find(str)>=0:
            print "'"+t+"'", res_id


    print "done."

#if __name__ == "__main__":
#    findgdoc()


download_item('spreadsheet:0Ao993IR9m-38dE9Yc1ZKSVR0R084enl0elpKRHJlSmc','data/home/2012/2012 sveta.xls')
#upload_item('spreadsheet:0Ao993IR9m-38dDdjN3pGNFViQlotQXY3TkRVV2c1MUE',"c:/5b.xls")