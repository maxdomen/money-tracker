#from decimal import Decimal
import accounts

__author__ = 'Max'
#from decimal import *
from datetime import datetime

class CurrencyExhangeRate:
    def __init__(self):
        self.date = ''
        self.currency1 = 0
        self.currency2 = 0
        self.rate = 0

#class C

rub = 1
usd = 2
eur=3
gbr=4
cad=5

class Money:
    __slots__ = ['float']
    def __init__(self, value=0.0):

        if isinstance(value, float):
            self.float=value
            return
        if isinstance(value, str):
            self.float=float(value)
            return
        if isinstance(value, unicode):
            self.float=float(value)
            return
        if isinstance(value, Money):
            self.float=value.float
            return
        if isinstance(value, int):
            self.float=float(value)
            return

        if not isinstance(value, float):
            raise Exception("bad Money constructor parameter {1}({0})".format(value, value.__class__.__name__))
        self.float=value
    def __eq__(self, other):
        v=Money(other)
        return v.float == self.float
    def __ne__(self, other):
        if other==None:
            return False

        v=Money(other)
        return v.float != self.float

    def as_float(self):
        return self.float
    def __repr__(self):
        return "Money('%s')" % str(self)

    def __str__(self, eng=False, context=None):
        return str(self.float)

    def __iadd__(self, other, context=None):
        #Decimal
        v=Money(other)
        return Money(self.float+v.float)

    def __add__(self, other, context=None):

        v=Money(other)
        return Money(self.float+v.float)
    __radd__ = __add__
    def __sub__(self, other, context=None):
        v=Money(other)
        return Money(self.float-v.float)
    def __rsub__(self, other, context=None):
        v=Money(other)
        return Money(v.float-self.float)
    #__rsub__ = __sub__
    def __div__(self, other):
        v=Money(other)
        return Money(self.float/v.float)
    def __rdiv__(self, other):
        v=Money(other)
        return Money(v.float/self.float)
    #__rdiv__ = __div__
    #def __truediv__(self, other):
    #    v=Money(other)
    #    return self.float.__truediv__(v.float)

    def __mul__(self, other):
        v=Money(other)
        return Money(self.float*v.float)

    def __abs__(self):
        #v=Money(other)
        return Money(abs(self.float))

    def __format__(self, specifier, context=None, _localeconv=None):
        return self.float.__format__(specifier)


class ExhangeRate:
    rates = []

    def __init__(self, date,cfrom, to, rate):
        self.date=date
        self.cfrom=cfrom
        self.to=to
        self.rate=Money(rate)



class Currency:

    rates=[]
    #def __init__(self):
    #    rates=[]
    @staticmethod
    def addrate(date,cfrom, to, rate):
        rec=ExhangeRate(date,cfrom, to, rate)
        Currency.rates.append(rec)
    @staticmethod
    def getrate(date,cfrom, to):

        if cfrom==to:
            return Money(1)
        #find nearest date
        #nearestdate=datetime.now()
        #nearestdiff=nearestdate-date
        bestrec=None
        bestdiff=datetime.now()-datetime(1,1,1)
        for rec in Currency.rates:
            if (rec.cfrom==cfrom and rec.to==to) or (rec.cfrom==to and rec.to==cfrom):
                if date<rec.date:
                    diff=rec.date-date
                else:diff=date-rec.date

                if diff<bestdiff:
                    bestdiff=diff
                    bestrec=rec

        if bestrec==None:
            sfrom=Currency.verbose(cfrom, 0)
            sto=Currency.verbose(to, 0)
            raise Exception("Cannot find conversion rate {0}->{1} date: {2}".format(sfrom, sto, date))

        if bestrec.cfrom==cfrom and bestrec.to==to:
            return bestrec.rate
        return Money(1)/bestrec.rate

    @staticmethod
    def str_to_currency_code(amount):
        currency= accounts.rub
        cur_index=amount.find('RUB')
        if cur_index<0:
            cur_index=amount.find('USD')
            currency= accounts.usd
            if cur_index<0:
                cur_index=amount.find('EUR')
                currency= accounts.eur
                if cur_index<0:
                    cur_index=amount.find('GBR')
                    currency= accounts.gbr
                    if cur_index<0:
                        cur_index=amount.find('GBP')
                        currency= accounts.gbr
                        if cur_index<0:
                            raise Exception("Unknown currency in "+amount)
        return currency, cur_index

    @staticmethod
    def convert(cfrom, cto, time, value):



        if cfrom == cto:
            return value

        rate=Currency.getrate(time,cfrom, cto)

        return value/rate
        #raise Exception("Cannot convert currency")

    @staticmethod
    def verbose(cfrom, value):
        if cfrom == rub:
            return "{0:.2f}p".format(value)

        if cfrom == usd:
            return "${0:.2f}".format(value)

        if cfrom == eur:
            return "{0:.2f}EUR".format(value)
        if cfrom == gbr:
            return "{0:.2f}GBP".format(value)

        if cfrom == cad:
            return "{0:.2f}CAD".format(value)
        raise Exception("Unknown currency {0}".format(cfrom))
  