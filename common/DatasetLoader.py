import os
import xlrd
import pickle
from gdatafetch import download_spreadsheet
from gdatafetch import download_spreadsheet, upload_item
__author__ = 'Max'
# -*- coding: utf8 -*-

CONFIG_SPREADSHEET = "0Ao993IR9m-38dFZQcVZMRHFQeFhSY1lTNEVkd3RkUmc"

class User:
    def __init__(self):
        self.Name=""
        self.Name_lowered=""
        self.Journal_Spradsheet_Id=""
        self.Report_Spreadsheet_Id=""
        self.Resource_Cost_h_RUB=""
        self.Report_Quant_Size_minutes=""
        self.Sharing=True
        self.Archived=""
        self.TablesToPublish=[]


class DatasetLoader:
    def __init__(self,data_dir=None, localOnly=False):

        if not data_dir:
            data_dir="data/"

        self.localOnly=localOnly
        data_dir=os.path.realpath(data_dir)+"\\"
        print "Data dir",data_dir
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)

        #self.remote_local_map={}
        self.data_dir=data_dir
        self.users_list=None
        self.config_book=None
        self.google_user=None
        self.google_password=None
    def init_login_and_password(self):
        pwdcachefilename=self.data_dir+'pwdcache.txt'

        if os.path.exists(pwdcachefilename):
                print "Load login from {0}...".format(pwdcachefilename)
                f1 = file(pwdcachefilename, 'rb')
                p=pickle.load(f1)
                gl,gp=p
                print "Use {0}".format(gl)
        else:
                gl=raw_input("Enter google login: ")
                gp=raw_input("Enter google password: ")
                if len(gl)<1 or len(gp)<1:
                    print "Please inter correct parameters"
                    return

                f1 = file(pwdcachefilename, 'wb')
                pickle.dump((gl,gp),f1)
                print "The login has been saved insecurely to {0}".format(pwdcachefilename)


        self.google_user=gl
        self.google_password=gp

    def config(self):
        if self.config_book:
            return self.config_book
        config_id="0Ao993IR9m-38dFZQcVZMRHFQeFhSY1lTNEVkd3RkUmc"
        localpath=self.get(config_id)
        self.config_book=xlrd.open_workbook(localpath)
        return self.config_book

    def get_users_list(self):
        if self.users_list:
            return self.users_list
        self.users_list={}
        sheet = self.config().sheet_by_index(0)
        for rowi in range(1, sheet.nrows):
            r = sheet.row(rowi)
            u=User()
            u.Journal_Spradsheet_Id = r[1].value
            u.Report_Spreadsheet_Id = r[2].value

            if len(u.Journal_Spradsheet_Id) > 0:
                u.Name= r[0].value
                u.Name_lowered=u.Name.lower()
                u.Resource_Cost_h_RUB = r[3].value
                u.Report_Quant_Size_minutes = int(r[4].value)
                str_sharing = r[5].value
                u.Sharing=True
                if str_sharing=="no":
                    u.Sharing=False

                self.users_list[u.Name_lowered]=u

                #index.append((uname, path, resid2, resid, hourly_rate, quantsize, sharing))
        return self.users_list


    def get_local_or_download(self, google_id,forceDownload=False):
        localpath=self.id_to_local(google_id)
        needDownload=not os.path.exists(localpath)
        if forceDownload:
            needDownload=forceDownload
        if needDownload:
            localpath=self.download(google_id)

        return localpath

    def id_to_local(self, google_id):
        confpath=self.data_dir+google_id+".xlsx"
        return confpath

    def download(self, google_id):
        confpath=self.id_to_local(google_id)
        print "download", google_id, confpath
        login=None
        if self.google_user:
            login=(self.google_user,self.google_password)
        download_spreadsheet(google_id, confpath,login)
        return confpath

    def get(self,google_id):
        localpath=self.id_to_local(google_id)
        #if not os.path.exists(localpath):
        #    localpath=self.download(google_id)
        #else:
        #    print "local file",google_id,"exists"

        needDownload=True

        if self.localOnly:
            needDownload=not os.path.exists(localpath)
        if needDownload:
            localpath=self.download(google_id)

        return localpath
    def upload(self,local_xls,output_googleid):
        login=None
        if self.google_user:
            login=(self.google_user,self.google_password)
        upload_item('spreadsheet:' + output_googleid, local_xls,login)


###########################################################
class DatasetLoader2:
    def __init__(self,data_dir=None, localOnly=False):

        if not data_dir:
            data_dir="data/"

        self.localOnly=localOnly
        data_dir=os.path.realpath(data_dir)+"\\"
        print "Data dir",data_dir
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)

        #self.remote_local_map={}
        self.data_dir=data_dir
        #self.users_list=None
        #self.config_book=None
        self.google_user=None
        self.google_password=None
    def init_login_and_password(self):

        userprofile=os.environ.get('USERPROFILE')
        if len(userprofile)>0:
            pwdcachefilename=userprofile+'\\pwdcache.txt'
        else:
            pwdcachefilename=self.data_dir+'pwdcache.txt'



        if os.path.exists(pwdcachefilename):
            print "Load login from {0}...".format(pwdcachefilename)
            f1 = file(pwdcachefilename, 'rb')
            p=pickle.load(f1)
            gl,gp=p
            print "Use {0}".format(gl)
        else:
            gl=raw_input("Enter google login: ")
            gp=raw_input("Enter google password: ")
            if len(gl)<1 or len(gp)<1:
                print "Please inter correct parameters"
                return

            f1 = file(pwdcachefilename, 'wb')
            pickle.dump((gl,gp),f1)
            print "The login has been saved insecurely to {0}".format(pwdcachefilename)


        self.google_user=gl
        self.google_password=gp


    def download(self, localpath,google_id):
        #confpath=self.id_to_local(google_id)
        confpath=self.data_dir+localpath

        #isfolder exist
        dir=os.path.dirname(confpath)
        if not os.path.exists(dir):
            print "create dir", dir
            os.makedirs(dir)

        print "download", confpath,google_id
        login=self._check_login()

        download_spreadsheet(google_id, confpath,login)
        return confpath
    def _check_login(self):
        login=None
        if self.google_user:
            login=(self.google_user,self.google_password)
        else:
            self.init_login_and_password()
            login=(self.google_user,self.google_password)
        return login

    def upload(self,local_xls,output_googleid):
        confpath=self.data_dir+local_xls
        fullpath=os.path.realpath(confpath)
        login=self._check_login()
        #login=None
        #if self.google_user:
        #    login=(self.google_user,self.google_password)

        print "upload", fullpath,output_googleid
        upload_item('spreadsheet:' + output_googleid, fullpath,login)

def makesure_directory(confpath):
    fullpath=os.path.realpath(confpath)
    #dir=os.path.dirname(fullpath)
    #print "check",fullpath
    if not os.path.exists(fullpath):
        print "create dir", fullpath
        os.makedirs(fullpath)