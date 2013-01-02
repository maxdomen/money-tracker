# -*- coding: utf8 -*-
import copy
from accounts import Statement, StatementRow, RowType, Account, Tx, Pool

from common.Classification import Period, TagTools
from common.Table import Style
from currency import usd, rub, Currency, Money
from readers.StatementReader import TxSource

__author__ = 'Max'


class DebtRow:
    def __init__(self,title,max_period_index):
        self.title=title
        self.value=0.0
        self.total=0.0
class Debts:
    def __init__(self, statement,start=None, end=None):
        self.rows=[]

        if start==None:
             start=statement.get_time_start()
        if end==None:
             end=statement.get_time_finish()

        self._start=start
        self._end=end

        self.periods=Period.CreateSet(Period.Month,self._start,self._end,10)

        self._process_statement(statement)

    def _process_statement(self,statement):
        self.accs={}
        debtops=[]


        self.accs["tcs"]=DebtRow("tcs",100)
        self.accs["avu"]=DebtRow("avu",100)
        self.accs["CM"]=DebtRow("CM",100)

        for r in statement.Rows:

            if r.type!=RowType.Tx:
                continue
            tags=list(r.tags)
            amount=r.amount.as_float()
            date=r.date


            if tags.count("debt")>0:
                tags.remove("debt")

                if r.tx.direction==1:
                    pass

                else:
                    amount=-1*amount

                stags=TagTools.TagsToStr(tags)
                #print "debt found",amount,stags,date
                dr=self.accs.get(stags)
                if not dr:
                    dr=DebtRow(stags,100)
                    self.accs[stags]=dr

                debtops.append( (0,date,amount,tags,stags) )

        #все возможные каналы долгов известны
        emptyaccs=copy.deepcopy(self.accs)
        for p in self.periods:
            p.accs=copy.deepcopy(emptyaccs)

        #credit cards
        for p in self.periods:
            lastknownr=None
            for r in statement.Rows:
                if r.date>p._end:
                    break
                lastknownr=r



            for accname, amount in lastknownr.cumulatives.items():
                if accname=="tcs" or accname=="avu":
                    if amount>0:
                        amount=0
                    amount=-1*amount
                    debtops.append( (1,lastknownr.date,amount,[],accname) )

        self.debtops=debtops

    def define_debt(self,accname,date,amount):

        self.debtops.append( (1,date,amount,[],accname) )


    def xsl_to(self,table,period, debts_due_to_date):

        for p in self.periods:

                    for acc in p.accs.values():
                        acc.total=self.accs[acc.title].value

                    for optype,date,amount,tags,key in self.debtops:
                        if date>=p._start and date<=p._end:

                            if optype==0:
                                p.accs[key].value+=amount
                                self.accs[key].value+=amount
                                p.accs[key].total=self.accs[key].value

                            if optype==1:
                                p.accs[key].total=amount

        #названия долгов
        baserowi=10
        table[baserowi-2,0]=u"Сумма долгов"
        rowi=baserowi
        for acc in self.accs.values():
            table[rowi,0]=acc.title
            rowi+=1


        coli=1


        for p in self.periods:
            if  not (p._end<debts_due_to_date):
                break

            if p._start<period.start or p._end>period.end:
                continue

            alldebts=0
            rowi=baserowi
            for acc in p.accs.values():

                table[rowi,coli]=acc.total, Style.Money
                alldebts+=acc.total
                rowi+=1

            table[baserowi-2,coli]=alldebts

            coli+=1

