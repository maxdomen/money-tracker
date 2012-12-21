# -*- coding: utf8 -*-

import codecs
from datetime import datetime, timedelta
#from decimal import *
#import accounts
import xlrd
from model import accounts
import model.accounts
#import pyparsing 

import os
import sys
#from pyparsing import Word, alphas
#from accounts import *
from model.currency import Money


__author__ = 'Max'


#autocategirization
#merchant

class TxSource:
    def __init__(self,fn,wn,line,col):
        self.filename=fn
        self.worksheet_name=wn
        self.line=line
        self.column=col

    def verbose(self):
        str="{0}:{1} {2}:{3}".format(self.filename, self.worksheet_name, self.line, self.column)
        return str


class ObjectFromRow:
        def __init__(self, fnames):
            for n in fnames:
                self.__dict__[n]=None

class SimpleSheetReader:

    def rows(self):
        return self.sheet.nrows

    def __init__(self, book, worksheetindex=0):

        #self.sheet=book.sheet_by_name(sheetName)
        self.sheet=book.sheet_by_index(worksheetindex)

        self.names={}
        self.allnames=""
        ##find column indexes
        header=self.sheet.row(0)
        cindex=0
        for c in range(len(header)):
            colname="c"+str(cindex)
            cindex=cindex+1

            self.names[colname]=c
            if self.allnames>"":self.allnames+=","
            self.allnames+=colname

    def row_dict(self, index):
        result={}
        for c in self.names:

            cindex=self.names[c]
            cval=self.sheet.row(index)[cindex].value

            result[c]=cval


        return result
    def row(self, index):
        result=ObjectFromRow(self.names)
        for c in self.names:

            cindex=self.names[c]
            cell=self.sheet.row(index)[cindex]
            cval=cell.value

            setattr(result, c, cval)

        return result

class SheetReader:

    def rows(self):
        return self.sheet.nrows

    def __init__(self, book, sheetName):

        #self.sheet=book.sheet_by_name(sheetName)
        self.sheet=book.sheet_by_index(0)

        self.names={}
        header=self.sheet.row(0)
        ##find column indexes
        self.allnames=""
        for c in range(len(header)):
            colname=header[c].value
            if colname=="from":
                colname="from_"

            self.names[colname]=c
            if self.allnames>"":self.allnames+=","
            self.allnames+=colname

            #setattr(self.__class__, name, property(func))
        #load cels

    #def setalias(self, curname,alias):
    #    ci=self.names[curname]
    #    self.names.

    def row_dict(self, index):
        result={}
        for c in self.names:

            cindex=self.names[c]
            cval=self.sheet.row(index)[cindex].value

            result[c]=cval


        return result
    def row(self, index):
        result=ObjectFromRow(self.names)
        for c in self.names:

            cindex=self.names[c]
            cval=self.sheet.row(index)[cindex].value


            setattr(result, c, cval)

        return result



class XlsLeftoversJournalReader:
    def __init__(self,filename,sheetname, config):
        self.filename=filename
        self.sheetname=sheetname
        self.config=config
        return
    def parse_to(self, accstoread):

        print "  parse",self.filename
        book = xlrd.open_workbook(self.filename)
        sheet=book.sheet_by_name(self.sheetname)

        frow=self.config['first_row']

        #sheet.row(frow)[cindex]
        #cashconfig={'first_row':4,'col_date':0, 'col_in':3,'col_out':4,'col_balance':5,'col_op':2}

        for rowi in range(frow,sheet.nrows):

            r=sheet.row(rowi)
            xlsdate=r[self.config['col_date']].value
            tdate=xlrd.xldate_as_tuple(xlsdate,0)
            h=tdate[3]
            m=tdate[4]
            s=tdate[5]
            if h==0 and m==0 and s==0:
                h=23
                m=59
                s=59
            date=datetime(tdate[0],tdate[1],tdate[2],h,m,s)


            for acc in accstoread.keys():
                ind=accstoread[acc]
                sout=r[ind].value
                if isinstance(sout,float):
                    aout=Money(sout)
                    l= accounts.LeftOver(date,aout)
                    src=TxSource(self.filename,self.sheetname,rowi,ind )
                    l.src=src
                    acc.leftover(l)

class XlsReader:
    def __init__(self,filename,sheetname, config):
        self.filename=filename
        self.sheetname=sheetname
        self.config=config
        return
    def addrec(self,acc, time, op, amount_in,amount_out, balance,srcdef):


        if balance>0:
            timediff=timedelta(24.0/24/60)
            lo=accounts.LeftOver(time+timediff, balance)
            lo.src=srcdef
            acc.leftover(lo)


        if amount_in>0:
            tx=accounts.Tx(amount_in,time)
            tx.src=srcdef
            acc.income(tx)

            if amount_out>0:
                print  "in and out in the same record", time,amount_in,amount_out
                tx.comment="[income from undefined source]"
                tx=accounts.Tx(amount_out,time)
                tx.src=srcdef
                acc.out(tx)
        else:
            if amount_out<=0:
                return
            tx=accounts.Tx(amount_out,time)
            tx.src=srcdef
            acc.out(tx)

        tx.comment=op
        #tx.src=srcdef

        return tx

    def parse_to(self, accs):

        print "  parse",self.filename
        book = xlrd.open_workbook(self.filename)
        sheet=book.sheet_by_name(self.sheetname)

        accsobj={}
        for a in accs:
            accsobj[a.name]=a

        frow=self.config['first_row']

        #sheet.row(frow)[cindex]
        #cashconfig={'first_row':4,'col_date':0, 'col_in':3,'col_out':4,'col_balance':5,'col_op':2}
        accnamecolind=self.config['col_acc']
        for rowi in range(frow,sheet.nrows):
            r=sheet.row(rowi)


            xlsdate=r[self.config['col_date']].value
            if len(str(xlsdate))<1:
                xlsdate=prevxlsdate
            else:
                tdate=xlrd.xldate_as_tuple(xlsdate,0)
                date=datetime(tdate[0],tdate[1],tdate[2],tdate[3],tdate[4],tdate[5])
                prevxlsdate=xlsdate



            op=r[self.config['col_op']].value
            if len(op)<1:
                continue

            if len(accs)==1:
                acc=accs[0]
            else:
                accname=r[accnamecolind].value
                if len(accname)<1:
                    raise Exception("Cell {0}:{1} refers to not existent account (the cell is empty)".format(rowi,accnamecolind))
                if not accsobj.has_key(accname):
                    raise Exception("Cell {0}:{1} refers to unknown account '{2}'".format(rowi,accnamecolind,accname))
                acc=accsobj[accname]



            ain=self.getamount(r,'col_in')
            aout=self.getamount(r,'col_out')
            abal=self.getamount(r,'col_balance')

            src=TxSource(self.filename,self.sheetname,rowi,0)




            tx=self.addrec(acc,date,op,ain,aout,abal,src)

            tag1=r[self.config['col_tag1']].value
            tag2=r[self.config['col_tag2']].value
            if tx!=None:
                if len(tag1)>0: tx.add_tag(tag1)
                if len(tag2)>0: tx.add_tag(tag2)

        return
    def getamount(self,r,pname):
        sout=r[self.config[pname]].value
        aout=0
        if isinstance(sout,float):
            aout=Money(sout)

        return aout

class AutoReader:
    def __init__(self,filename):
        self.filename=filename
        return
    def parse_to(self, acc):
        txs=[]
        return