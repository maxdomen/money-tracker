# -*- coding: utf8 -*-
import codecs
from currency import usd, Currency, Money

__author__ = 'Max'
from datetime import datetime, timedelta
#from decimal import Decimal
import unicodedata
import xlrd
from StatementReader import TxSource
import accounts


class TCSBankReader:
    def __init__(self,filename):
        self.filename=filename
        return

    def str_to_validate(self, str):
        parts=str.split('.')
        if len(parts)<3:
             return (False,None)

        syear=parts[2]

        spaceind=syear.find(" ")
        if spaceind>0:
            syear=syear[0:spaceind]
            if len(syear)<3:
                short=int(syear)
                syear=2000+short

        sday=parts[0]
        txdate=datetime(int(syear),int(parts[1]),int(sday))

        return True, txdate
    def find_currency(self,amount):

        currency, cur_index=self.extract_currency(amount)
        part_decim=amount[0:cur_index]
        part_decim=part_decim.replace(' ','')


        num=part_decim.split(',')
        orignum=Money(num[0]+"."+num[1])
        return currency,orignum
    def extract_currency(self,amount):
        #currency= accounts.rub
        #cur_index=amount.find('RUB')
        #if cur_index<0:
        ##    cur_index=amount.find('USD')
        #   currency= accounts.usd
        #    if cur_index<0:
        #        cur_index=amount.find('EUR')
        #        currency= accounts.eur
        #        if cur_index<0:
        #            cur_index=amount.find('GBR')
        #            currency= accounts.gbr
        #            if cur_index<0:
        #                cur_index=amount.find('GBP')
        #                currency= accounts.gbr
        #                if cur_index<0:
        #                    raise Exception("Unknown currency in "+amount)
        currency, cur_index=Currency.str_to_currency_code(amount)
        return currency, cur_index
    def addrec_2011(self,acc, amount, date, desc, srcdef):
        validdate, txdate=self.str_to_validate(date)
        if not validdate:
           return



        currency,orignum=self.find_currency(amount)
        amount_rub=accounts.Currency.convert(currency, acc.currency,txdate,orignum)

        self.addrec2(acc,txdate,amount_rub,orignum,currency,desc,srcdef )

    def addrec2(self,acc, txdate, amount_rub,amount_orig, currency_orig , desc, srcdef):

       

        tx=accounts.Tx(amount_rub,txdate)

        tx.original_amount=Money(amount_orig)
        tx.original_currency=currency_orig

        tx.src=srcdef
        tx.comment=desc



        #if dnum==5000:
        #    pass

        income=False
        if desc.find(u"Пополнение")>=0 or desc.find(u"Операция возврата")>=0:
            income=True
            tx.dest="[owner]"

        if desc.find(u"Отказано")>=0 or desc.find(u"Отмена")>=0:
            return

        if desc.find(u"Проценты по кредиту")>=0 or desc.find(u"Комиссия за выдачу наличных")>=0 or desc.find(u"Плата за предоставление услуги")>=0:
            tx.dest="[bank]"
            tx.comment=desc

        t=u"Оплата в"
        i=desc.find(t)
        if i<0:
            t=u"Покупка"
            i=desc.find(t)
            if i<0:
                t=u"Операция без присутствия карты"
                i=desc.find(t)
        if i>=0:
            tx.dest=desc[i+len(t):len(desc)]
            tx.comment=""




        if income:
            acc.income(tx)
        else:
            acc.out(tx)
        #print tx.verbose(acc.currency)
        return
    def addrec_2012(self,acc, txdate, amount_rub,amount_orig, currency_orig , desc, srcdef):
        self.addrec2(acc, txdate, amount_rub,amount_orig, currency_orig , desc, srcdef)
        return
    def parse2012_to(self,acc):

        #Дата и время операции	Дата списания	Сумма в валюте операции	 Сумма в р.		Описание
        print "  parse",self.filename

        #f=open(self.filename,'rb')
        #f=codecs.open(self.filename,'r', encoding='cp1251')
        f=codecs.open(self.filename,'r', encoding='utf-8')


        content=f.read()
        lines=content.split('\n')
        parsed=[]
        linecount=0
       # print len(lines)
        for l in lines:
            #if len(l)<1:
            #    continue
                
            linecount=linecount+1
            fields=l.split(';')
            fc=0
            f=fields[0]
            if f[3]=='"': #специальный случай первой строки
                f=f[3:len(f)-1]
                fields[0]=f
            if f[1]=='"': #специальный случай первой строки
                f=f[1:len(f)-1]
                fields[0]=f

            for f in fields:
                first=f[0:1]
                if first=='"':
                    f=f[1:len(f)-1]

                fields[fc]=f
                fc+=1
            validdate, date_of_operation=self.str_to_validate(fields[0])
            validdate, date_of_draw=self.str_to_validate(fields[1])
            if not validdate:
                date_of_draw=date_of_operation

            amount_rub=fields[6]
            amount_orig=fields[4]
            currency_orig,t=self.extract_currency(fields[5])
            desc=unicode(fields[8])
            desc+=" "
            desc+=unicode(fields[9])
            srcdef=TxSource(self.filename, "[0]",linecount,0)

            if len(amount_rub.strip())<1:
                amount_rub=amount_orig
            self.addrec_2012(acc,date_of_draw, amount_rub,amount_orig, currency_orig , desc, srcdef)


    def parse2011_to(self, acc):
      #  print os.curdir


        print "  parse",self.filename
        f=codecs.open(self.filename,'r', encoding='cp1251')

        content=f.read()
       # print len(content)



        lines=content.split('\n')
        parsed=[]
        linecount=0
       # print len(lines)
        for l in lines:
            linecount=linecount+1
           # print "process>> "+l
            l=l[1:len(l)-2]
            fields=l.split(';')

            if len(fields)<3:
                continue

            index=0
            srcdef=TxSource(self.filename, "[0]",linecount,0)
            parsedline=['', '', '',srcdef]
            parsed.append(parsedline)


            for f in fields:
                #print f
                if f[0:2]=='""':
                    f=f[2:len(f)]

                if f[len(f)-2:len(f)]=='""':
                    f=f[0:len(f)-2]

                if f[len(f)-1:len(f)]=='"':
                    f=f[0:len(f)-1]

                parsedline[index]=f
                index=index+1
        #print parsed

        start=False
        for rec in parsed:
            if rec[1]==u"Описание":
                start=True
                continue
            if start:
                self.addrec_2011(acc,rec[0],rec[2],rec[1],rec[3])

       # print lines

        #1251
        txs=[]
        return
    def parse2011v2_to(self, acc):
      #  print os.curdir


        print "  parse",self.filename
        f=codecs.open(self.filename,'r', encoding='cp1251')

        content=f.read()
       # print len(content)



        lines=content.split('\n')
        parsed=[]
        linecount=0
       # print len(lines)
        for l in lines:
            linecount=linecount+1
           # print "process>> "+l
            #l=l[1:len(l)-2]
            fields=l.split(';')

            if len(fields)<3:
                continue

            index=0
            srcdef=TxSource(self.filename, "[0]",linecount,0)
            parsedline=['', '', '',srcdef]
            parsed.append(parsedline)


            for f in fields:
                #print f
                if f[0:1]=='"':
                    f=f[1:len(f)]
                if f[len(f)-1:len(f)]=='"':
                    f=f[0:len(f)-1]

                parsedline[index]=f
                index=index+1
        #print parsed

        start=False
        for rec in parsed:
            if rec[1]==u"Описание":
                start=True
                continue
            if start:
                self.addrec_2011(acc,rec[0],rec[2],rec[1],rec[3])

       # print lines

        #1251
        txs=[]
        return