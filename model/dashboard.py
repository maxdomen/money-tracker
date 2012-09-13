# -*- coding: utf8 -*-
from datetime import datetime, timedelta
#import decimal
import xlwt
from accounts import RowType
from common.Classification import Period
from currency import Currency, Money
from debt import BudgetFreq, BudgetBehaviour

__author__ = 'Max'

class BigPicturePeriod:
    def __init__(self):
        self.reminder=0
        self.bills_topay=0
        self.avaialble_funds=0

        self.reminders={"Total":0.0}
        self.revenue_streams={"Total":0.0}

        self.costs={"Total":0.0}

        self.ebitda=0


class Widget:
    def __init__(self, title, cx=0,cy=0):
        self.title=title
        self.rows=[]
        #self.
        self.cells = [['' for col in range(cx)] for row in range(cy)]
        self.rows=self.cells
class DashboardDataset:
    def __init__(self,statement,plan=None):
        self.widgets=[]
        self.statement=statement
        self.plan=plan
        self.moment=datetime.now()

        w=self.widget_balance()
        self.totalAvailable=w.totalAvailable
        self.widget_lasttransactions()
        self.widget_next7days(self.totalAvailable)
        self.widget_next30days(self.totalAvailable)

    class AccInfo:
        def __init__(self,acc=None):
            self.ref=acc
            self.balance=0
            self.available=0

    def widget_balance(self):

        acclen=len(self.statement.Accounts)
        w=Widget(u"Баланс",cx=acclen+3,cy=3)
        self.widgets.append(w)
        rowscount=len(self.statement.Rows)
        accts={}


        w.cells[0][0]="Total"
        w.cells[1][0]=self.statement.Rows[rowscount-1].left_pool

        #hitc=0

        lastrow=self.statement.Rows[rowscount-1]
        
        for ref in self.statement.Accounts:
            v=lastrow.cumulatives[ref.name]
            accts[ref.name]=self.AccInfo(ref)
            accts[ref.name].balance=v



        coli=2
        w.cells[1][1]="balance"
        w.cells[2][1]="available"
        w.cells[2][0]=0.0
        w.totalAvailable=0.0
        for acc2 in self.statement.Accounts:
            w.cells[0][coli]=acc2.name
            #if accts.has_key(acc2.name):
            acc=accts[acc2.name]
            #if acc.ref:
            bal=acc.balance
            limit=acc2.limit

         

            avail=bal
            if limit!=0:
                limit=Currency.convert(acc2.currency,self.statement.currency,self.moment,limit)
                avail=bal-limit

            print acc2.name, limit, bal, avail

            w.cells[1][coli]=bal
            w.cells[2][coli]=avail
            w.totalAvailable+=bal-limit
            coli+=1

        w.cells[2][0]=w.totalAvailable
        return w

    def widget_lasttransactions(self):
        w=Widget(u"Последние транзакции")
        self.widgets.append(w)
        rowscount=len(self.statement.Rows)
        rowc=0
        for ri in range(rowscount-1, 1,-1):
            row=self.statement.Rows[ri]



            if not (row.type==RowType.Tx or row.type==RowType.Lost):
                continue
          
            if not row.account:
               continue

            debit=row.amount
            credit=0

            if row.tx and row.tx.direction==1:
                credit=row.amount
                debit=0

            txrow=[]
            txrow.append(self.strdate(row.date))
            txrow.append(debit)
            txrow.append(credit)
            txrow.append(row.description)
            w.rows.append(txrow)
            rowc+=1
            if rowc>15:
                break

    def widget_next7days(self,startfund):
        self.widget_next_N_days(7,u"Следующие 7 дней", [BudgetFreq.Weekly,BudgetFreq.Daily],startfund)
    def widget_next30days(self,startfund):
        self.widget_next_N_days(30,u"Следующие 30 дней", [BudgetFreq.Weekly,BudgetFreq.Daily,BudgetFreq.Monthly],startfund)
    def strdate(self,dt):
        res="{:%m-%d}".format(dt)
        #res="{0:00}/{1:00}".format(dt.month,dt.day)
        return res
    def widget_next_N_days(self, days,title, collapseperiods,startfund):

        w=Widget(title)
        self.widgets.append(w)
        start=self.moment
        end=start+timedelta(days=days)
        #firstrow=['',0,0,0,u"Расходы до 500р"]

        datarows=[]

        #datarows.append(firstrow)

        totalcredit=0
        totaldebit=0
        commoncosts={}
        commondates={}

        funds=startfund
        prevdescr=""
        for row in self.plan.Rows:
            if row.date>=start and row.date<=end:
                budget=row.tx.source_budget

                if budget.behaviour==BudgetBehaviour.Expectation:
                    continue

                if row.tx.direction==1:
                    funds+=row.amount
                else:
                    funds-=row.amount

                #period=budget.period
                #if row.amount<500:
                #    firstrow[1]+=row.amount
                #    continue
                #iscollapsed=collapseperiods.count(period)>0
                #iscollapsed=False
                #if iscollapsed:
                #    if commoncosts.has_key(budget):
                #        commoncosts[budget]+=row.amount
                #    else:
                #        commoncosts[budget]=row.amount
                #        commondates[budget]=row.date
                #    continue
   
                txrow=[]


                debit=row.amount
                credit=0
                if row.tx.direction==1:
                    credit=row.amount
                    debit=0

                fundsstatus=""
                if funds.float<0:
                    fundsstatus="[OUT OF MONEY]"

                if prevdescr==row.description and len(datarows)>0:
                    i=len(datarows)-1
                    cr=datarows[i]
                    cr[1]+=debit
                    cr[2]+=credit
                    cr[3]=funds
                    cr[4]=fundsstatus
                else:
                    txrow.append(self.strdate(row.date))
                    txrow.append(debit)
                    txrow.append(credit)
                    txrow.append(funds)
                    txrow.append(fundsstatus)
                    txrow.append(row.description)
                    datarows.append(txrow)
                prevdescr=row.description


        #for budget, amount in commoncosts.items():
        #    txrow=[]
        #    date=commondates[budget]
        #    txrow.append(self.strdate(date))

        #    debit=amount
        #    credit=0
        #    if budget.credit>0:
        #        debit=0
        #        credit=amount

        #    txrow.append(debit)
        #    txrow.append(credit)
        #    txrow.append(0)
        #    txrow.append('')
        #    txrow.append(budget.description)
        #    datarows.append(txrow)

        totalcredit=0
        totaldebit=0
        for row in datarows:
            totalcredit+=row[2]
            totaldebit+=row[1]

        datarows=sorted(datarows, key=lambda r: r[0])
        w.rows.extend(datarows)

        txrow=[]
        txrow.append('')
        txrow.append(totaldebit)
        txrow.append(totalcredit)
        txrow.append(0)
        txrow.append("Total")
        w.rows.insert(0,txrow)
        w.rows.insert(1,[])
        


class DashboardPublisher:
    def __init__(self, dataset,filename, sheetname,existing_workbook=None):
        self.style_money=xlwt.easyxf(num_format_str='#,##0')
        if existing_workbook:
           self.wb = existing_workbook
        else:
           self.wb = xlwt.Workbook()

        self.ws = self.wb.add_sheet(sheetname)

        rowi=0
        for w in dataset.widgets:
            rowi=self.print_widget(w,rowi)
            rowi+=1

    def print_widget(self, w,rowi):
        coli=0
        self.ws.write(rowi, coli, w.title)
        rowi+=1
        for row in w.rows:
            coli=1
            for cell in row:


                if isinstance(cell, Money):
                    if cell!=0:
                        self.ws.write(rowi, coli, cell.as_float(),self.style_money)
                else:
                    self.ws.write(rowi, coli, cell)
                coli+=1
            rowi+=1


        return rowi
