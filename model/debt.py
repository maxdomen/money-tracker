# -*- coding: utf8 -*-
from datetime import datetime, timedelta
#from decimal import Decimal
import xlrd
#from StatementReader import TxSource
from accounts import Statement, StatementRow, RowType, Account, Tx, Pool
#from aggregatereport import Period
from common.Classification import Period
from currency import usd, rub, Currency, Money
from readers.StatementReader import TxSource

__author__ = 'Max'


class DebtRow:
    def __init__(self,title,max_period_index):
        self.title=title
        #self.current=0
        #self.latest_known_date=None
        #self.history=[]
        self._cells=[Money(0)]*max_period_index
class Debts:
    def __init__(self, start=None, end=None,statement=None):
        self.rows=[]

        if start==None:
             start=statement.Rows[0].date
        if end==None:
             end=statement.Rows[len(statement.Rows)-1].date+timedelta(seconds=1)

        self._start=start
        self._end=end

        self.periods=Period.CreateSet(Period.Month,self._start,self._end,10)

    def add_credit_card_as_account(self, statement, account, mode,qualificator=1):
        debt=DebtRow(account.name, len(self.periods))
        #debti=len(self.rows)
        self.rows.append(debt)



        #обработка строк
        for row in statement.Rows:
            #row_date=self._get_logical_date(row)
            if mode==1:
                if row.account!=account:
                    continue

            row_date=row.date

            v=row.left_acc
            if mode==2:
                v=row.left_pool
            v=v*qualificator
                #debt.current=v
                #debt.latest_known_date=row_date
            pi=0
            for p in self.periods:
                if row_date>=p._start and row_date<=p._end:
                    debt._cells[pi]=v
                    break
                        #p._cells[debti]=v
                pi+=1


    def calc_total(self):
        sumdebt=DebtRow("Total",len(self.periods))
        #debti=len(self.rows)
        self.rows.append(sumdebt)

        pi=0
        for p in self.periods:
            #di=0
            sum=Money(0)
            for debt in self.rows:
                v=debt._cells[pi]
                #v=p._cells[di]
                sum+=v
                #di+=1
            sumdebt._cells[pi]=sum
            pi+=1
        #if row_date>=p._start and row_date<=p._end:
        #p._cells[debti]=v



class BudgetFreq:
    Monthly=0
    Weekly=1
    Annually=2
    OneTime=3
    Daily=4

class BudgetBehaviour:
    Std=0
    Expectation=1
    

class BudgetRow:
    def __init__(self, period, day,debit, credit, currency,tags, description, exactdate=None, behaviour=None):
        self.period=period
        self.day=day
        self.debit=debit
        self.credit=credit
        self.currency=currency
        self.tags=tags
        self.description=description
        self.exactdate=exactdate
        self.start=None
        self.end=None
        self.behaviour=BudgetBehaviour.Std
        if behaviour:
            self.behaviour=behaviour
class Budget:
    def __init__(self):
        self.rows=[]
        self.account=Account('budgetacc',rub)
    def Add(self, row):
        self.rows.append(row)
    def make_statement(self, currency=usd, forNyears=1):
        start=datetime.now()
        y=start.year
        self._start=datetime(y,1,1)
        #end=datetime(y+forNyears,12,31, hour=23, minute=59)

        #res=Statement()
        #res.Accounts=[]
        #res.currency=currency

        for budget in self.rows:
            if budget.period== BudgetFreq.Annually:
                for year_repeater in range(0,forNyears):
                    dt=datetime(self._start.year+year_repeater,budget.exactdate.month,budget.exactdate.day )
                    self.createline(budget, dt)



            if budget.period== BudgetFreq.OneTime:
                self.createline(budget, budget.exactdate)

            if budget.period== BudgetFreq.Monthly:
                for year_repeater in range(0,forNyears):
                    day=1
                    if budget.day!=0:
                        day=budget.day
                    for mo in range(1,13):
                        date=datetime(self._start.year+year_repeater, mo,day)
                        self.createline(budget, date)

            if budget.period==BudgetFreq.Daily:
                cur=self._start
                for day in range(1,365*forNyears):
                    self.createline(budget, cur)
                    cur+=timedelta(days=1)

            if budget.period== BudgetFreq.Weekly:
                cur=self._start
                di=cur.weekday()
                ti=1 #monday

                if budget.day!=0:
                    ti=budget.day

                cur+=timedelta(days=(di-ti))

                for we in range(1,(365*forNyears)/7):
                    #date=datetime(self._start.year, mo,day)
                    self.createline(budget, cur)
                    cur+=timedelta(days=7)



        p = Pool()
        p.link_account(self.account)
        res=p.make_statement(currency)
        return res
    def createline(self, budget, date):

        if budget.start:
            if date<budget.start:
                return

        if budget.end:
            if date>budget.end:
                return
        amnt=budget.debit
        if budget.credit>0:
            amnt=budget.credit
        tx=Tx(amnt,date)
        tx._tags=budget.tags
        tx.comment=budget.description

        tx.src=TxSource("fn","sn",0,0)
        tx.source_budget=budget

        if budget.credit>0:
            self.account.income(tx)
        else:
            self.account.out(tx)

            #(self, period, day,debit, credit, currency,tags, description, exactdate=None):
    def xsldate(self, xslsvalue):
        res=None

        if isinstance(xslsvalue, float):

             tdate=xlrd.xldate_as_tuple(xslsvalue,0)
             res=datetime(tdate[0],tdate[1],tdate[2])
        return res
    def read(self, filename, sheetname):
        print "  budget",filename
        book = xlrd.open_workbook(filename)
        sheet=book.sheet_by_name(sheetname)

        behavemap={"expectation":BudgetBehaviour.Expectation}

        periodsmap={"monthly":BudgetFreq.Monthly,"weekly":BudgetFreq.Weekly,"annually": BudgetFreq.Annually,"onetime": BudgetFreq.OneTime, "daily": BudgetFreq.Daily}
        for rowi in range(1,sheet.nrows):
            r=sheet.row(rowi)
            speriod=r[1].value.lower()
            if len(speriod)<1:
                continue
                
            if not periodsmap.has_key(speriod):
                 raise Exception("Unknown Period '{0}' at row {1}".format(speriod, rowi))

            period=periodsmap[speriod]

            scur=r[9].value.upper().strip()
            currency=rub
            if len(scur)>0:
                currency, cur_index=Currency.str_to_currency_code(scur)

            sbehave=r[10].value.lower().strip()
            behave=BudgetBehaviour.Std
            if len(sbehave)>0:
                behave=behavemap[sbehave]


            tags=r[6].value.split(',')
            etags=[]
            for t in tags:
                st=t.strip()
                if len(st)>0:
                    etags.append(st)
            tags=etags

            sday=r[2].value
            day=1
            exactdate=None

            if isinstance(sday, float):
                if sday>31:
                    xlsdate=sday
                    tdate=xlrd.xldate_as_tuple(xlsdate,0)
                    exactdate=datetime(tdate[0],tdate[1],tdate[2])
                else:
                    day=int(sday)
            debit=r[3].value
            if not isinstance(debit, float): debit=0
            credit=r[4].value
            if not isinstance(credit, float): credit=0
            br=BudgetRow(period,day, debit,credit,currency,tags,r[5].value,exactdate=exactdate,behaviour=behave)
            self.Add(br)



            br.start=self.xsldate(r[7].value)
            br.end=self.xsldate(r[8].value)
            #tdate=xlrd.xldate_as_tuple(xlsdate,0)
            #date=datetime(tdate[0],tdate[1],tdate[2],tdate[3],tdate[4],tdate[5])
    def diff(self,plan, actual):

        pass