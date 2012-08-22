from datetime import datetime, timedelta
#from decimal import Decimal
import unicodedata
import xlrd
from StatementReader import TxSource
import model.accounts

__author__ = 'Max'

class ChaseBankReader:
    def __init__(self,filename,sheetname):
        self.filename=filename
        self.sheetname=sheetname
        #self.config=config
        return
    def addrec(self,acc, time, op, amount_in,amount_out, balance,srcdef):

        #print time, op, amount_in,amount_out, balance

        if amount_in>0:
            tx=accounts.Tx(amount_in,time)
            acc.income(tx)
        else:
            tx=accounts.Tx(amount_out,time)
            acc.out(tx)

        tx.comment=op
        tx.src=srcdef

        return tx

    def parse_to(self, acc):

        print "  parse",self.filename
        book = xlrd.open_workbook(self.filename)
        sheet=book.sheet_by_name(self.sheetname)


        prevdate=None
        incrementor=0
        for rowi in range(sheet.nrows-1, 0, -1):
           # print rowi
            r=sheet.row(rowi)
            xlsdate=r[0].value
            if isinstance(xlsdate, unicode):
                xlsdate=unicodedata.normalize('NFKD', xlsdate).encode('ascii','ignore')

            if isinstance(xlsdate, str):
                xlsdate=xlsdate.strip()
                if len(xlsdate)<5:
                    continue
                    
                date = datetime.strptime(xlsdate, '%m/%d/%Y')
            else:
                tdate=xlrd.xldate_as_tuple(xlsdate,0)
                date=datetime(tdate[0],tdate[1],tdate[2],tdate[3],tdate[4],tdate[5])



            date+=  timedelta(1.0/24)

            op=r[2].value


            ain=self.getamount(r,4)
            aout=self.getamount(r,3)
            abal=self.getamount(r,5)



    
            src=TxSource(self.filename,self.sheetname,rowi,0 )

            #incrementor+=1
            if date!=prevdate:
                incrementor=0

            wdate=date
            partial=1.0/24/60


            wdate+=timedelta(partial*incrementor)
            tx=self.addrec(acc,wdate,op,ain,aout,abal,src)


            if abal>0:

                wdate+=timedelta(partial*1)
                incrementor+=1
                #print wdate, abal
                lo=accounts.LeftOver(wdate, abal)
                lo.src=src
                acc.leftover(lo)

            incrementor+=2;
            
            prevdate=date

        return
    def getamount(self,r,pname):
        aout=0
        val=r[pname].value
        if isinstance(val, float):
            aout=Decimal(val)
        else:
            sout=val.strip()

            if len(sout)>0:
                d=sout[0]
                if d=="$":
                    sout=sout[1:len(sout)]

                parts=sout.split(',')
                sout=""
                for p in parts:
                    sout+=p

                aout=Decimal(sout)

        return aout