# -*- coding: utf8 -*-
import copy
from accounts import Statement, StatementRow, RowType, Account, Tx, Pool

from common.Classification import Period, TagTools
from common.Table import Style
from currency import usd, rub, Currency, Money
from readers.StatementReader import TxSource

__author__ = 'Max'


class DebtAccountPeriod:
    def __init__(self):
        self.period_balance=None
class DebtTx:
    def __init__(self,accountname,date, balance):
        self.accountname=accountname
        self.date=date
        self.balance=balance
class Debts:
    def __init__(self, statement,start=None, end=None):
        self.rows=[]

        if start==None:
             start=statement.get_time_start()
        if end==None:
             end=statement.get_time_finish()

        self._start=start
        self._end=end

        self.periods=Period.CreateSet(Period.Month,self._start,self._end)

        self._process_statement(statement)

    def _enum_debt_tags(self,statement,debtxs):
        accum={}
        #найдем все ко-теги с тегом debt
        for r in statement.Rows:

            if r.type!=RowType.Tx:
                continue
            tags=list(r.tags)
            if tags.count("debt")>0:
                tags.remove("debt")
                amount=r.amount.as_float()

                if r.tx.direction==1:
                    pass
                else:
                    amount=-1*amount
                stags=TagTools.TagsToStr(tags)
                if not self.debt_accounts.has_key(stags):
                    #найден новый адресат долга
                    accum[stags]=0
                    self.debt_accounts[stags]=DebtAccountPeriod()
                accum[stags]+=amount
                debtxs.append(DebtTx(stags,r.date,balance=accum[stags]))

    def _process_statement(self,statement):
        self.debt_accounts={}

        debtxs=[]

        self.debt_accounts["tcs"]=DebtAccountPeriod()
        self.debt_accounts["avu"]=DebtAccountPeriod()
        self.debt_accounts["CM"]=DebtAccountPeriod()
        self._enum_debt_tags(statement,debtxs)

        #для каждого периода находим остатки по кредитынм картам
        for p in self.periods:
            lastknownr=None
            #находим последнюю известную строку в периоде
            for r in statement.Rows:
                if r.date>p._end:
                    break
                lastknownr=r

            #в строке бкдкт перечилены текущие остатки на счетах, в том числе и на кредитных картых
            if lastknownr:
                for accname, balance in lastknownr.cumulatives.items():
                    if accname=="tcs" or accname=="avu":
                        if balance>0:
                            balance=0
                        balance=-1*balance
                        debtxs.append(DebtTx(accname,lastknownr.date,balance=balance))


        self.debtxs=debtxs

    def define_debt_balance(self,accname,date,balance):

        #print "debt",date,amount,accname
        #self.debtops.append( (1,date,amount,[],accname) )
        self.debtxs.append(DebtTx(accname,date,balance=balance))

    def get_period(self, date):
        for p in self.periods:
            if date>=p._start and date<=p._end:
                return p
        return None
    def _finalize(self):
        #все возможные каналы долгов известны
        emptyaccs=copy.deepcopy(self.debt_accounts)
        for p in self.periods:
            p.accs=copy.deepcopy(emptyaccs)
        #надо рассчитать табличку периоды-долги по аккаунтам в них
        #для кредитных карт и
        #посичтать для авосгенернных аккаунтов и CM
        for dtx in self.debtxs:
            p=self.get_period(dtx.date)
            accdata=p.accs[dtx.accountname]
            #if dtx.balance:
            accdata.period_balance=dtx.balance

        #значение долга надо продолжить в будущее до последнего периода
        #чтобы не было стуации что долг внезапно обнулился
        prevp=None
        for p in self.periods:
            if not prevp:
                prevp=p
                continue
            for accname,accdata in p.accs.items():
                if not accdata.period_balance:
                    #баланс в периое не задан, заполеним его значением из преддущего периода
                    prev_p_balance=prevp.accs[accname].period_balance
                    if not prev_p_balance:
                        prev_p_balance=0
                    accdata.period_balance=prev_p_balance
            prevp=p

    def xsl_to(self,table,period, debts_due_to_date):
        self._finalize()


        baserowi=10
        table[baserowi-2,0]=u"Сумма долгов"


        #навзания долговых аккаунтов
        rowi=baserowi
        for name in self.debt_accounts.keys():
            table[rowi,0]=name
            rowi+=1


        coli=1


        for p in self.periods:
            #if  not (p._end<debts_due_to_date):
            #    break

            if p._start<period.start or p._end>period.end:
                continue

            alldebts=0
            rowi=baserowi
            for acc in p.accs.values():
                table[rowi,coli]=acc.period_balance, Style.Money
                alldebts+=acc.period_balance
                rowi+=1

            table[baserowi-2,coli]=alldebts
            coli+=1

