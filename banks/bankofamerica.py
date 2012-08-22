# -*- coding: utf8 -*-
from decimal import Decimal
from datetime import datetime, timedelta
from decimal import *
from StatementReader import TxSource
import accounts

__author__ = 'Max'
import csv

class BankOfAmericaReader:
    def __init__(self,filename):
        self.filename=filename
        self.leftovers=[]
        return
    def addrec(self,acc, samount, stime, op, endbalance,srcdef,seci):
        #print u"Add Tx: {0} {1} {2}".format(stime, amount,op)

        ptime=stime.split("/")
        #print ptime
        time=datetime(int(ptime[2]),int(ptime[0]),int(ptime[1]),0,0,seci)
        time2=datetime(int(ptime[2]),int(ptime[0]),int(ptime[1]), 0,0,seci+1)

        amount=float(samount)
        income=True
        if amount<0:
            income=False
            amount=amount*(-1)

        if op.find("Beginning balance")>=0:
            lo=accounts.LeftOver(time, amount)
            lo.src=srcdef
            acc.leftover(lo)
            return

        tx=accounts.Tx(amount,time)
        
        tx.comment=op
        tx.src=srcdef
        if income:
            acc.income(tx)
        else:
            acc.out(tx)


        lo=accounts.LeftOver(time2, endbalance)
        lo.src=srcdef
        self.leftovers.append(lo)
        #acc.leftover(lo)

        #DEPOSIT
        #Check

    def parse_to(self, acc):


        print "  parse",self.filename
        f=open(self.filename)
        csv_reader=csv.reader(f)

        process=False
        seci=0
        for row in csv_reader:
            if len(row)<1:
                continue
            if row[0]=="Date" and row[1]=="Description":
                process=True
                continue
            if process:
                time=row[0]
                amnt=row[2]
                if not amnt:
                    amnt='0'

                src=TxSource(self.filename,"[0]",row,0)
                self.addrec(acc,amnt,time,row[1],row[3],src,seci)
                seci+=2

        #сохряняем только последний остаток в сутках
        self.leftovers.sort(key=lambda lo: lo.time, reverse=True)
        day=None#self.leftovers[0].time
        for lo in self.leftovers:
            if not day:
                acc.leftover(lo)
                day=lo.time
                continue
            diff=day-lo.time
            if diff<timedelta(days=1):
                continue
            acc.leftover(lo)
            day=lo.time


        return