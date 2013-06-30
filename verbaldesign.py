from datetime import datetime
from itertools import *

__author__ = 'Max'


class Line:
    def __init__(self, date, amount,comment):
        self.date=date
        self.comment=comment
        self.amount=amount
        self.balance=None
        self.currency=None
        self.acc="<unk>"
        self.tags=[]



def set_currency(pool, target_currency):
    pass

def categorize(pool):
    for l in pool:
        l.parent_categories=[1,0]


    #self.parent_categories=[]

def calculate_balances(pool):
    pass
    #self.balances={}
    #self.balance_total=0

def by_month(line):
    d=line.date
    d2=datetime(d.year,d.month,1)
    return d2

pool=[]
pool.append(Line(datetime(2013,1,1),100,"test"))
pool.append(Line(datetime(2013,1,10),101,"test"))
pool.append(Line(datetime(2014,1,3),101,"2test"))

pool2=[]
pool2.append(Line(datetime(2013,1,1),100,"test"))
pool2.append(Line(datetime(2013,1,10),101,"test"))
pool2.append(Line(datetime(2014,1,3),101,"2test"))

#pool3=xlstoarray()


#pool=chain(pool,pool2)
for l in chain(pool,pool2):
    print l

#set_currency(pool, rub)
categorize(pool)
calculate_balances(pool)


#company relationships
#fromcompany=select(pool, "parent_categories","tx_in" )
fromcompany=(line for line in pool if line.parent_categories.count(0)>0 and line.comment[0]!="2")
#print sum(fromcompany)

for s in fromcompany:
    print s

for k,g in groupby(pool,by_month):
    print k,g
    for line in g:
        print " ",line

#tocompany=select(pool, "parent_categories","tx_out" )

#companyrelationsheeps=merge(fromcompany,tocompany)
#calculate_balances(companyrelationsheeps)
#sortbydate(companyrelationsheeps)
#prints(companyrelationsheeps)


#bycategurt
#currentmonth=select(pool, "date","2011-may" )
#bycategory=groupby(currentmonth,"parent_categories")
#sums=sum(bycategory)