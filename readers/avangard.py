# -*- coding: utf8 -*-
from datetime import datetime, timedelta
#from decimal import Decimal
import xlrd
from StatementReader import TxSource, SimpleSheetReader
from model import accounts
import model.accounts
from model.currency import Money

__author__ = 'Max'
#

class AvangardReader:
    def __init__(self,filename):
        self.filename=filename
        return



    def addrec(self,acc, amount,income, time, op, merchant,orig_amount, orig_cur,srcdef, scard_id):

        if op==u"Предоставление овердрафта": return



        tx=accounts.Tx(amount,time)

        tx.recordSource=srcdef
        tx.comment=op
        if len(merchant)>0:
            tx.comment=merchant+" "+op
        #tx.dest=merchant
        tx.src=srcdef

        #if op.find(u"Погашение процентов")>=0:
        #    tx.dest="[bank]"
        #if op.find(u"Комиссия за")>=0:
        #    tx.dest="[bank]"

        if len(orig_cur)>0:
            curmap={'RUR':accounts.rub,'USD':accounts.usd,'EUR':accounts.eur,'GBR':accounts.gbr,'GBP':accounts.gbr,'CAD':accounts.cad}
            tx.original_currency=curmap[orig_cur]
            tx.original_amount=Money(orig_amount)

        if op.find(u"Наличные")>=0:
            tx.add_tag("To Cash")


        if len(scard_id)>0:
            tx.props['card_id']=scard_id

        if income:
            acc.income(tx)
        else:
            acc.out(tx)


    def date_from_xls(self,xlsdate,book):
        tdate=xlrd.xldate_as_tuple(xlsdate,book.datemode)
        date=datetime(tdate[0],tdate[1],tdate[2],tdate[3],tdate[4],tdate[5])
        return date
    def parse_corporate_to(self, acc):
        print "  parse",self.filename
        book = xlrd.open_workbook(self.filename)
        sheet=book.sheet_by_index(0)


        for rowi in range(10, sheet.nrows):

            r=sheet.row(rowi)
            xlsdate=r[2].value

            if not isinstance(xlsdate, float):
                continue

            #print xlsdate,"<-"
            date=self.date_from_xls(xlsdate, book)
            #tdate=xlrd.xldate_as_tuple(xlsdate,book.datemode)
            #date=datetime(tdate[0],tdate[1],tdate[2],tdate[3],tdate[4],tdate[5])
            docnum=r[6].value
            accnum=r[8].value
            destination=r[21].value.strip()
            opcode=r[24].value
            income=False
            samnt=r[26].value
            if not isinstance(samnt, float):
                samnt=r[29].value
                income=True
            amount=Money(samnt)
            opdescr=r[31].value.strip()


            #print date, docnum,accnum, samnt, opdescr, destination
            src=TxSource(self.filename,"[0]",rowi,0)
            self.addrec(acc,amount,income,date,opdescr+" "+destination,destination,"","",src, "")

    def parse_private_withreserved_to(self, acc):
        
        print "  parse",self.filename
        book = xlrd.open_workbook(self.filename)
        sheet=book.sheet_by_index(0)
        self.book=book

        keyword1=u"Проведенные по картсчету операции"
        keyword2=u"Авторизованные(зарезервированные), но еще не поступившие в банк операции"
        keyword3=u"Входящий остаток средств на картсчете / остаток задолженности"

        cur_date=None
        #определяем начало и конец подтаблиц
        table1_s,table1_e=0,0
        table2_s,table2_e=0,0

        for rowi in range(10, sheet.nrows):
            r=sheet.row(rowi)
            title=r[1].value
            if title==keyword1:
                mode=1
                table1_s=rowi+3
            if title==keyword2:
                mode=2
                table2_s=rowi+2

            if title==u"Итого":
                 table1_e=rowi-1
            if title==u"Общая сумма покупок":
                 table2_e=rowi-1


        crow=sheet.row(5)
        s_xlsdate=crow[3].value
        report_date_start=self.str_to_date(s_xlsdate,2000)+timedelta(seconds=-1)

        s_xlsdate=crow[7].value
        report_date_finish=self.str_to_date(s_xlsdate,2000)+timedelta(seconds=-1)+timedelta(hours=23, minutes=59,seconds=59)

        #сканирование авторизованных приходов
        for rowi in range(table1_s,table1_e+1):
            #break
            #print table1_e,rowi
            r=sheet.row(rowi)
            prihod=r[2].value
            if isinstance(prihod, float):
                #приход
                current_caption=self.scan_all_title(r)
                xlsdate=r[1].value
                if not isinstance(xlsdate, float):
                    continue
                cur_date=self.date_from_xls(xlsdate, book)
                src=TxSource(self.filename,"[0]",rowi,0)
                self.addrec2_1(acc,True,cur_date,prihod,current_caption,src)

                #возможно есть погашение процентов
                next_r=sheet.row(rowi+1)
                current_caption=self.scan_all_title(next_r)
                if current_caption==u"Погашение процентов по предоставленному овердрафту":
                    interest=0
                    interest=next_r[7].value
                    self.addrec2_1(acc,False,cur_date,interest,current_caption,src)
                continue
        #сканирование авторизованных расходов
        for rowi in range(table1_s,table1_e+1):

            #break
            r=sheet.row(rowi)
            rashod=r[7].value
            if isinstance(rashod, float):
                current_caption=self.scan_all_title(r)
                if current_caption.find(u"Погашение")==0:
                    continue

                current_caption=self.exapnd_caption(current_caption,sheet, rowi)


                xlsdate=r[1].value
                if isinstance(xlsdate, float):
                    cur_date=self.date_from_xls(xlsdate, book)
                src=TxSource(self.filename,"[0]",rowi,0)
                self.addrec2_1(acc,False,cur_date,rashod,current_caption,src)
                continue

        #сканирование зарезервированных расходов
        if table2_s>0:
            for rowi in range(table2_s,table2_e+1):
                #break
                #print "R",table2_e,rowi
                r=sheet.row(rowi)
                rashod=r[7].value
                if isinstance(rashod, float):
                    current_caption=self.scan_all_title(r)
                    current_caption=self.exapnd_caption(current_caption,sheet, rowi)
                    r_date=sheet.row(rowi+1)
                    xlsdate=r_date[9].value
                    cur_date=self.date_from_xls(xlsdate, book)

                    if cur_date<report_date_start:
                        scurdate=self.date_to_str(cur_date)
                        current_caption="Not processed from {0} ".format(scurdate)+current_caption
                        cur_date= report_date_finish+timedelta(seconds=-1)
                    src=TxSource(self.filename,"[0]",rowi,0)
                    self.addrec2_1(acc,False,cur_date,rashod,current_caption,src)
                    continue

         #извлечение остатков



        #баланс в конце
        #это кредитная карта или нет?
        ostatok_out_r=sheet.row(1)
        columdind=30
        if len(ostatok_out_r)<30:
            columdind=24


        siscredit=sheet.row(1)[columdind].value
        is_credit=False
        #if len(siscredit)>0:
        if isinstance(siscredit,float):
            is_credit=(siscredit!=0.0)


        headrow=sheet.row(4)
        headrow2=sheet.row(8)
        headrow3=sheet.row(9)
        if is_credit:
            number=headrow[columdind].value+headrow2[columdind].value
        else:
            number=headrow3[columdind].value

        if isinstance(number,float):
            self.make_leftover(acc,book,sheet,number,report_date_finish)

        #баланс в начале
        number=sheet.row(table1_s-2)[columdind+1].value
        if isinstance(number,float):
            self.make_leftover(acc,book,sheet,number,report_date_start)


    def str_to_date(self, str, add_y):
        
        ptime=str.split(' ')[0].split('.')
        mo=int(ptime[1])
        time=datetime(int(ptime[2])+add_y,mo,int(ptime[0]))
        return time
    def make_leftover(self,acc,book,sheet,number, date):


        lo=accounts.LeftOver(date, number)
        src=TxSource(self.filename,"[0]",0,0)
        lo.src=src
        acc.leftover(lo)

    def exapnd_caption(self,current_caption, sheet, rowi):
         if current_caption==u"Оплата товаров и услуг":
                    more1=self.scan_all_title(sheet.row(rowi+1))
                    more2=self.scan_all_title(sheet.row(rowi+2))
                    current_caption=more2+more1+current_caption
                    #current_caption=more2+more1
                    #if current_caption.find("SPORT")>=0:
                    #    print "SPORT3"

         if current_caption==u"Получение наличных в банкомате":
                    more1=self.scan_all_title(sheet.row(rowi+1))
                    more2=self.scan_all_title(sheet.row(rowi+2))
                    current_caption=more2+more1+current_caption
         return current_caption
    def scan_all_title(self,r):
        current_caption=""
        for ci in range(9,30):
            cell=r[ci]
            s=cell.value
            if not isinstance(s, float):
                if s!="":
                    if s!=u"Место":
                        if len(current_caption)>0:
                            current_caption+=" "
                        current_caption+=s
            else:
                if cell.ctype==3:
                    opdate=self.date_from_xls(s,self.book)
                    opdate=self.date_to_str(opdate)
                    current_caption+=u" от {0} ".format(opdate)
                else:
                    current_caption+=" {0:.2f} ".format(s)

        return current_caption
    




    def extract_human_date(self,op):
        found=False
        human_date=None
        pos=op.find(u" от ")
        if pos>0:
            fragment=op[pos+4:len(op)]
            nums=fragment.split(".")
            if len(nums)>=3:
                sday=nums[0]
                smonth=nums[1]
                syear=nums[2]
                syear=syear[0:4]
                if syear[2]==" ":
                    syear=syear[0:2]

                day=int(sday)
                month=int(smonth)
                year=int(syear)
                if year<2000:
                    year=year+2000
                #print u"-{0}.{1}.{2}-".format(sday, smonth, syear)

                human_date=datetime(year,month, day)
                found=True
        return found,human_date

    def addrec2_1(self,acc,income, time, amount,op,src):
        tx=accounts.Tx(amount,time)
        tx.comment=op

        #test=u"TOCHKA ZRENIYA  OPTIKA SANKT-PETERSB RU от 15.12.2012 10:38  Карта *1723. Сумма 1297.00  RUR.Оплата товаров и услуг"
        #found_human, human_date=self.extract_human_date(test)

        #if op.find(u" от ")>0:
        found_human, human_date=self.extract_human_date(op)

        tx.src=src
        if income:
            acc.income(tx)
        else:
            acc.out(tx)

        if found_human:
            #print "found_human",human_date,op
            tx.human_date=human_date

    def parse_private_to(self, acc):


        print "  parse",self.filename
        book = xlrd.open_workbook(self.filename)

        actreader=SimpleSheetReader(book,worksheetindex=0)
        for a in range(1,actreader.rows()):
            row=actreader.row(a)

            if row.c3==u"Предоставление овердрафта":
                #бессмфсденная строчка
                continue
            if row.c3==u"Погашение овердрафта":
                #бессмфсденная строчка
                continue



            xlsdate=row.c0

            tdate=xlrd.xldate_as_tuple(xlsdate,book.datemode)
            #print tdate
            date=datetime(tdate[0],tdate[1],tdate[2],tdate[3],tdate[4],tdate[5])

            amount=row.c1 #это приход
            income=True
            if not isinstance(amount, float):
                amount=row.c2
                income=False


            caption=row.c3

            opdate=row.c4
            if isinstance(opdate, float):
                opdate=self.date_from_xls(opdate,book)
                opdate=self.date_to_str(opdate)
                caption=caption+u" от "+opdate
                
            src=TxSource(self.filename,"[0]",a,0)
            self.addrec(acc,amount,income,date,caption,row.c9,row.c6, row.c7,src, row.c5)



        return
    def date_to_str(self, opdate):
        sopdate=u"{0}.{1}.{2} {3}:{4}".format(opdate.day, opdate.month, opdate.year,opdate.hour,opdate.minute)
        return sopdate