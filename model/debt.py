# -*- coding: utf8 -*-
import copy
from datetime import datetime, timedelta
#from decimal import Decimal
import xlrd
#from StatementReader import TxSource
import common.CalendarHelper
from accounts import Statement, StatementRow, RowType, Account, Tx, Pool
#from aggregatereport import Period
from common.Classification import Period, TagTools
from common.Table import Style
from currency import usd, rub, Currency, Money
from readers.StatementReader import TxSource

__author__ = 'Max'


class DebtRow:
    def __init__(self,title,max_period_index):
        self.title=title
        #self.current=0
        #self.latest_known_date=None
        #self.history=[]
        #self._cells=[float(0)]*max_period_index
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
    #def debt_tx(self, time, acccount,value):
    #    pass
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
                    #tags.remove("__in")
                    #print "INCOME"
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


    def xsl_to(self,table):

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


        baserowi=10
        rowi=baserowi
        for acc in self.accs.values():
            table[rowi,0]=acc.title
            rowi+=1
        dtnow=datetime.now()+timedelta(days=31)


        coli=1


        for p in self.periods:
            if  not (p._end<dtnow):
                break

            alldebts=0
            rowi=baserowi
            for acc in p.accs.values():
                #table[rowi,coli]=acc.value, Style.Money
                table[rowi,coli]=acc.total, Style.Money
                alldebts+=acc.total
                rowi+=1

            table[baserowi-2,coli]=alldebts
            #table[baserowi]
            coli+=1


class BudgetFreq:
    Monthly=0
    Weekly=1
    Annually=2
    OneTime=3
    Daily=4

class BudgetBehaviour:
    Std=0
    Expectation=1
    Done=2
    

class BudgetRow:
    def __init__(self, period, day,debit, credit, currency,tags, description, id,saccomplished="",exactdate=None, behaviour=None):
        self.period=period
        self.day=day
        self.debit=debit
        self.credit=credit
        self.currency=currency
        self.tags=tags
        self.description=description
        self.id=id
        if len(self.id)<1:
            self.id=description.lower()
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
        self.bying_targets=None
        self.executions=[]
    def Add(self, row):
        self.rows.append(row)
    def get_buying_targets(self):
        return self.bying_targets

    def check_item_execution(self,budget_item, foradate):
        is_todo=False
        is_overdue=False
        is_executed=self._check_is_executed(budget_item,foradate)
        if not is_executed:
            is_overdue=self._is_item_overdue(budget_item, foradate)

        if budget_item.exactdate and budget_item.debit>0:
            if (not is_overdue) and (not is_executed):
                if budget_item.exactdate>=foradate:
                    diff=budget_item.exactdate- foradate
                    if diff.days<31:
                        is_todo=True

                    if budget_item.period!=BudgetFreq.OneTime:
                        if budget_item.debit<2000:
                            is_todo=False

        return is_overdue, is_executed, is_todo

    def _check_is_executed(self,budget_item,date):
        res=False

        if budget_item.behaviour==BudgetBehaviour.Done:
            return True

        #if len(budget_item.id)<1:
        #    return False

        p=None
        if budget_item.period== BudgetFreq.Annually:
           # p=common.CalendarHelper.Period(datetime(date.year,1,1),datetime(date.year,12,31,23,59,59))
            p=common.CalendarHelper.Period(datetime(budget_item.exactdate.year,1,1),datetime(budget_item.exactdate.year,12,31,23,59,59))
        if budget_item.period== BudgetFreq.OneTime:
            #p=common.CalendarHelper.Period(datetime(date.year,date.month, date.day),datetime(date.year,date.month, date.day, 23,59,59))
            p=common.CalendarHelper.Period(datetime(2000,1,1),datetime(3000,1,1))
        if budget_item.period== BudgetFreq.Monthly:
            if budget_item.exactdate:
                s=datetime(budget_item.exactdate.year,budget_item.exactdate.month,1)

                if s.month+1>12:
                    e=datetime(s.year+1,1,1)-timedelta(seconds=1)
                else:
                    e=datetime(s.year,s.month+1,1)-timedelta(seconds=1)
                p=common.CalendarHelper.Period(s,e)

        if not p:
            return False

        for eid, edate in self.executions:
            if eid!=budget_item.id and (eid!=budget_item.description):
                continue
            if edate>=p.start and edate<=p.end:
                res=True
                break

        return res
    def _is_item_overdue(self, budget_item, foradate):

        if budget_item.debit<1:
            return False

        if budget_item.behaviour==BudgetBehaviour.Expectation:
            return False

        item_date=budget_item.exactdate
        if not item_date:

            if budget_item.period==BudgetFreq.Monthly:
                day=1
                if budget_item.day!=0:
                    day=budget_item.day
                item_date=datetime(foradate.year, foradate.month,day)

            if not item_date:
                return False

        if item_date<foradate:

            #фильтруем шум периодических расходов
            if budget_item.period!=BudgetFreq.OneTime:
                diff=foradate-item_date
                if diff.days>30:
                    return False
                if budget_item.debit<2000:
                    return False
            #фильтруем end
            return True
        return False
    def create_buying_target(self, srcitem, dt):
        if not self.in_time_limit(srcitem, dt):
            return srcitem
        budget2=srcitem
        if srcitem.behaviour!=BudgetBehaviour.Expectation:
            budget2=copy.copy(srcitem)
            #budget2._description=budget2.description
            budget2._description=budget2.description+"({0})".format(dt.year)
            if budget2.period== BudgetFreq.Monthly:
                budget2._description=budget2.description+"({0}-{1})".format(dt.year,dt.month)
            budget2.id=budget2.description.lower()
            budget2.exactdate=dt
            self.bying_targets.append(budget2)
        return budget2
    def make_statement(self, currency=usd, forNyears=1):
        self.bying_targets=[]
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

                    bcopy=self.create_buying_target(budget, dt)
                    self.createline(bcopy, dt)


            if budget.period== BudgetFreq.OneTime:
                if budget.behaviour!=BudgetBehaviour.Done:
                    if budget.debit>0:
                        self.bying_targets.append(budget)
                        #budget.isoverdue=True
                self.createline(budget, budget.exactdate)

            if budget.period== BudgetFreq.Monthly:
                #self.bying_targets.append(budget)
                for year_repeater in range(0,forNyears):
                    day=1
                    if budget.day!=0:
                        day=budget.day
                    for mo in range(1,13):
                        date=datetime(self._start.year+year_repeater, mo,day)


                        #if budget.description.find(u"ртпла")>0:
                        #    print "ртпла"
                        bcopy=self.create_buying_target(budget, date)
                        self.createline(bcopy, date)


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
    def in_time_limit(self,budget,date):
        if budget.start:
            if date<budget.start:
                return False

        if budget.end:
            if date>budget.end:
                return False
        return True
    def createline(self, budget, date):
        if not self.in_time_limit(budget, date):
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
    def read_executions(self,filename, sheetname):
        book = xlrd.open_workbook(filename)
        sheet=book.sheet_by_name(sheetname)

        for rowi in range(1,sheet.nrows):
            r=sheet.row(rowi)
            sid=r[1].value.lower()
            if len(sid)>0:
                xlsdate=r[2].value
                tdate=xlrd.xldate_as_tuple(xlsdate,0)
                exactdate=datetime(tdate[0],tdate[1],tdate[2])
                self.executions.append( (sid,exactdate) )
    def read(self, filename, sheetname):
        print "  budget",filename
        book = xlrd.open_workbook(filename)
        sheet=book.sheet_by_name(sheetname)
                    #expectation

        behavemap={"expectation":BudgetBehaviour.Expectation, "done":BudgetBehaviour.Done}

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
            saccomplished=""
            if len(r)>11:
                saccomplished=r[11].value

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
            descr=r[5].value
            if period==BudgetFreq.OneTime:
                if not exactdate:
                    mes=u"Exact date for budget line '{0}' missed".format(descr)
                    print mes
                    raise Exception(mes)
            id=r[0].value
            br=BudgetRow(period,day, debit,credit,currency,tags,descr,id,saccomplished,exactdate=exactdate,behaviour=behave)
            self.Add(br)



            br.start=self.xsldate(r[7].value)
            br.end=self.xsldate(r[8].value)
            #tdate=xlrd.xldate_as_tuple(xlsdate,0)
            #date=datetime(tdate[0],tdate[1],tdate[2],tdate[3],tdate[4],tdate[5])
    def diff(self,plan, actual):

        pass