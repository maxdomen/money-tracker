# -*- coding: utf8 -*-
#from StatementReader import TCSBankReader, AvangardReader, BankOfAmericaReader
#from  StatementReader import *
#import bankofamerica
#from printstatement import PrintStatementToExcel2
import xlrd
from currency import *

#
__author__ = 'Max'
from datetime import datetime, timedelta
#from decimal import *

#class Journal:
#    def Available(self, str):
#        #22.03.2012 13:28 59 274.72RUR
#        return

#    def Available(self, date, amnt):
#        return


class Amount:
    def __init__(self, amnt):
        self.currency = ''
        self.amount = amnt


class Tx:
    def __init__(self, amnt,time):

        #if currency==None:
        #    currency=self.account.currency
        self._amount = Money(amnt)
        self.time = time
        self._txid=None
        self.similar=0

        self.src = ''
        #self.dest = ''

        self.comment = ''

        self.direction = 1
        #Комиссия за получение наличных в банкомате
        #Комиссия за конверсию по трансграничной операции
        self.subcharges = []
        self.recordSource=None
        #self.remainder = Amount(0)#nullable
        self.original_amount=None
        self.original_currency=None
        self._tags=[]
        self.dublicates=[]
        self.props={}
        self._cashedamount={}
        self.slices=[]
        self.logical_date=None
        self.human_date=None
    def __str__(self):
        return self.get_id()
    def slice(self, title, amount, tags, tags_to_remove):
        self.slices.append( (title,amount, tags,tags_to_remove))
    def set_logical_date(self,newdate):
        self.logical_date= newdate
        
    def add_tag(self,tag):

        if len(tag)<1:
            return
        if not isinstance(tag, unicode):
            tag=unicode(tag)
        if self._tags.count(tag)<1:
            self._tags.append(tag.lower())
            self._tags.sort()

    def remove_tag(self,tag):
        if self._tags.count(tag)>0:
            self._tags.remove(tag)

    def add_dublicate(self,tx):
        self.dublicates.append(tx)
    def get_amount(self,currency):


        cashed=self._cashedamount.get(currency, None)
        if cashed:
            return cashed

        cashed=self._get_amount(currency)
        self._cashedamount[currency]=cashed
        return cashed

    def _get_amount(self,currency):
        known_currency=self.account.currency
        known_amount=self._amount
        if known_currency==currency:
            return known_amount


        if self.original_amount!=None:
            known_amount=self.original_amount
            known_currency=self.original_currency

        if known_currency==currency:
            return known_amount

        targetamount=Currency.convert(known_currency,currency,self.time,known_amount)
        return targetamount

    def applytobalance(self, bal,currency):
        bal = bal + self.get_amount(currency)*self.direction
        return bal

    def _generate_id(self,accountname,am, t):
        numpart="{0}{1}{2}{3}{4}".format(t.year-2000,t.month,t.day,t.hour,t.minute)
        if self.similar>0:
            sid="{0}{1}{2:.2f}[{3}]".format(numpart, accountname, am, self.similar)
        else:
            sid="{0}{1}{2:.2f}".format(numpart, accountname, am)
        return sid
    def get_id(self):

        if self._txid==None:
            self._txid=self._generate_id(self.account.name, self._amount, self.time)
        return self._txid


    def verbose(self, cur):

        amnt=self.get_amount(cur)

        origamount=""
        if self.original_amount!=None:
            if self.original_currency!=cur:
                origamount="({0})".format(Currency.verbose(self.original_currency, self.original_amount))
        str = u"{0} {1}".format(origamount,self.comment)
        return str

class LeftOver:
    def __init__(self, amnt):
        self.time = datetime.now()
        self._amount = Money(amnt)
        self.src = None

    def __init__(self, sdate, amnt):
        self.time = sdate
        self._amount = Money(amnt)
        self.src = None
    def verbose(self, cur):
        str = "Statement {0}".format(Currency.verbose(cur, self.get_amount(cur)))
        return str
    def get_id(self):
        #yymmdd.tiacc.sum
        sid="{0}{1}{2}{3}".format(self.time.year,self.time.month,self.time.day, self.account.name)
        return sid
    def applytobalance(self, bal, currency):
        bal = self.get_amount(currency)
        return bal

    def get_amount(self,currency):
        known_currency=self.account.currency
        known_amount=self._amount
        if known_currency==currency:
            return known_amount


        targetamount=Currency.convert(known_currency,currency,self.time,known_amount)
        return targetamount

class Account:
    def __init__(self, name, currency):
        self.name = name
        self.Txs = []
        self.txsdictionary={}
        self.Statements = []
        self.currency = currency
        self.limit=0
        #self.Calendar=Calendar()

    def __str__(self):
        return self.name

    #def AvailbleFromBalance(self, balance):
    #    return balance-self.limit

    def leftover(self, state):
        self.Statements.append(state)
        state.account = self


    def _check_exist(self,tx):
        txid=tx.get_id()

        e=self.txsdictionary.get(txid)
        

        return e

    def _addtx(self,tx,direction):
        tx.direction = direction;
        tx.account = self


        e=self._check_exist(tx)
        if e and e.src:
            samesrcfile=e.src.filename==tx.src.filename

            #добликаты появляются если одна и таже транзакция оказалось в двух разных банковских отчетах
            #но если две одинаковых транзакции из одного источника
            #то это не дубликат, а две похожих транзакции. Это постоянно происходит с платежами по подписке. анпимер за сотовый
            if samesrcfile:
                oid=e.get_id()
                nid=oid
                while self.txsdictionary.has_key(nid):
                    e.similar+=1
                    e._txid=None
                    nid=e.get_id()

                self.txsdictionary[nid]=e
                self.txsdictionary[oid]=None

                self.Txs.append(tx)
                self.txsdictionary[tx.get_id()]=tx
                #raise Exception("Similar transactions from '{0}' on {1} amount={2}".format(tx.src.filename, tx.time, tx._amount))
               # print "Similar transactions from '{0}' on {1} amount={2}".format(tx.src.filename, tx.time, tx._amount)
            else:
                e.add_dublicate(tx)
        else:
            self.Txs.append(tx)
            self.txsdictionary[tx.get_id()]=tx

    def income(self, tx):
        self._addtx(tx,1)

    def out(self, tx):
        self._addtx(tx,-1)

    #def check_journaling_quality(self):
    #    return

    def account_balance_to_pool_balance(self, accountbalance):
        return accountbalance

    def allrecs(self):

        res=[]
        for tx in self.Txs:
            res.append(tx)
        for st in self.Statements:
            res.append(st)

        return res
        #timerecs = [(tx.time, 'tx', tx) for tx in self.Txs]

        #for st in self.Statements:
        #    rec = (st.time, 'st', st)
        #    timerecs.append(rec)

        #return timerecs

#    def verbose_rec(self):




#class VirtualAccount:
#    def __init__(self):
#        self.name = ''

class RowType:
    Unknown=0
    Tx=1
    LeftOver=2
    Lost=3
    Transition=4

class StatementRow:

    def __init__(self):
        self.type=RowType.Unknown
        self.account=None
        self.tx=None
        self.date=None
        self.amount=Money()
        self.description="<empty>"
        self.left_acc=Money()
        self.left_pool=Money()
        self.tags=[]
        self.left_acc_to=Money() #только для transition
    def get_logical_date2(self):
        dt=self.date
        if self.tx.logical_date:
            dt=self.tx.logical_date
        return dt

    def get_human_or_logical_date(self):
        dt=self.date
        if self.tx.logical_date:
            dt=self.tx.logical_date

        if hasattr(self.tx,"human_date"):
            if self.tx.human_date:
                dt=self.tx.human_date
        return dt
class Statement:
    def __init__(self):
          self.Rows = []
          return
    def get_time_start(self):
        return self.Rows[0].date
    def get_time_finish(self):
        end=self.Rows[len(self.Rows)-1].date+timedelta(seconds=1)
        return end
    def get_generator(self):
        for r in self.Rows:


            if r.type!=RowType.Tx:
                continue

            tx=r.tx
            tags=r.tags

            ltags=[]
            for t in tags:
                ltags.append(t.lower())

            if tx.direction==1:
                ltags.append("__in")

            r.normilized_tags=ltags

            #dt=r.get_logical_date()

            dt=r.get_human_or_logical_date()

            res=(dt,r.amount.as_float(),ltags)
            yield   res


class TransisitionsLoader():
    def __init__(self, pool, filename,sheetname):


        book = xlrd.open_workbook(filename)
        sheet=book.sheet_by_name(sheetname)

        #periodsmap={"monthly":BudgetFreq.Monthly,"weekly":BudgetFreq.Weekly,"annually": BudgetFreq.Annually,"onetime": BudgetFreq.OneTime, "daily": BudgetFreq.Daily}
        for rowi in range(1,sheet.nrows):
            r=sheet.row(rowi)
            tfrom=r[1].value

            if len(tfrom)<1:
                continue
                
            tto=r[2].value
            commission=None
            sco=r[3].value
            if len(sco)>1:
                commission=sco
            comission2=None
            sco=r[4].value
            if len(sco)>1:
                commission2=sco

            pool.add_transition( tfrom, tto, commission,comission2)

class Pool:
    def __init__(self):
        self._Accounts = []
        self.Transitions=[]
        return

    def get_tx_byid(self, txid):
        for acc in self._Accounts:
            arecs = acc.Txs
            for tx in arecs:
                id=tx.get_id()
                if id==txid:
                    return tx

        return None

    def link_account(self, acc):
        self._Accounts.append(acc)
    def add_transition(self, tfrom, tto, commission=None,comission2=None):
        trans=Transition(tfrom,tto, commission,comission2)
        self.Transitions.append(trans)
    def _checkmatch(self, tx, tags):
        res=False
        for t in tags:
            tc=tx._tags.count(t)
            if tc>0:
                res=True

                return res

        #у транзакции могут быть слайсы


        return res
    def generate_slices(self,origin,currency):

        slices=[]

        tx_full_amount= origin.get_amount(currency)
        amount_reminder=tx_full_amount

        for slice_title, slice_amount, slice_tags,tags_to_remove in origin.slices:
            target_slice_amount=Currency.convert(origin.account.currency,currency,origin.time,slice_amount)
            amount_reminder-=target_slice_amount
            full_slice_tags=[]
            full_slice_tags.extend(slice_tags)
            full_slice_tags.extend(origin._tags)

            for rt in tags_to_remove:
                if full_slice_tags.count(rt)>0:
                    full_slice_tags.remove(rt)
            slice=Tx(slice_amount,origin.time)
            slice.account=origin.account
            slice.src=origin.src
            slice.direction=origin.direction
            slice.comment="Slice "+slice_title+" "+origin.comment
            slice.human_date=origin.human_date
            slice._tags=full_slice_tags
            slices.append(slice)

            #slice=self.create_row(RowType.Tx, tx.account,tx,tx.time,new_leftover,poolbalance,"Slice "+slice_title+" "+descr,target_slice_amount,full_slice_tags )
            #res.Rows.append(slice)

        if amount_reminder<0:
            raise Exception(" reminder for {0} less then zero".format(origin.get_id()))

        #reminder=self.create_row(RowType.Tx, tx.account,tx,tx.time,new_leftover,poolbalance,"Reminder of "+descr,amount_reminder,tx._tags )
        #res.Rows.append(reminder)
        slice=Tx(amount_reminder,origin.time)
        slice.account=origin.account
        slice.direction=origin.direction
        slice.human_date=origin.human_date
        slice.src=origin.src
        slice.comment="Reminder of "+origin.comment
        slice._tags=origin._tags
        slices.append(slice)

        return slices

    def make_statement(self, currency=usd, virtual_account=None, filter_debit=None,filter_credit=None, skip_transitions=False):
        res=Statement()
        res.Accounts=list(self._Accounts)

        if virtual_account:
            res.Accounts.append(virtual_account)

        res.currency=currency
        allrecs = []

        #nativebalances = {acc.name: Money(0) for acc in res.Accounts}
        #acc_last_left = {acc.name: Money(0) for acc in res.Accounts}

        nativebalances = {acc.name: 0.0 for acc in res.Accounts}
        acc_last_left = {acc.name: 0.0 for acc in res.Accounts}

        it_was_something_known= {acc.name: False for acc in res.Accounts}

        allrecs_by_id_index={}

        #add transition destinations
        semi_auto_transitions=[]
        for trans in self.Transitions:
            if trans.tx_to_id[0]==">":
                accname=trans.tx_to_id[1:len(trans.tx_to_id)]
                print "SemiAuto Destination to {0} from {1}".format(accname, trans.tx_from_id)
                for acc in res.Accounts:
                    if acc.name==accname:
                        semi_auto_transitions.append((trans,acc))



        for acc in res.Accounts:
            arecs = acc.allrecs()
            for rec in arecs:
                #обработка слайсов
                if isinstance(rec,Tx) and len(rec.slices)>0:
                    slices=self.generate_slices(rec,currency)
                    allrecs.extend(slices)
                else:
                    allrecs.append(rec)

                #if rec.id
        pseudo_guid=0
        for tx in allrecs:
            txid=tx.get_id()

            for trans, trans_acc in  semi_auto_transitions:
                if trans.tx_from_id==txid:
                    #print "found"
                    pseudo_guid+=1
                    intx=accounts.Tx(tx._amount,tx.time)
                    intx.similar=pseudo_guid
                    intx.description="[FROM]"+tx.comment
                    trans_acc.income(intx)
                    allrecs.append(intx)
                    trans.tx_to_id=intx.get_id()
                    print "SemiAuto Destination Tx {0} created".format(trans.tx_to_id)

                    #tx.account.currency

        allrecs.sort(key=lambda x: x.time)
        poolbalance = 0

        for tx in allrecs:
            isstatement=False
            if not isinstance(tx,Tx):
                isstatement=True

            acc = tx.account
            accname=acc.name

            if isstatement and virtual_account:
                if accname!= virtual_account.name:
                    continue
                else:
                    for a in self._Accounts:
                        nativebalances[a.name]=0
                        acc_last_left[a.name]=0




            if (not isstatement) and tx.direction==-1 and filter_debit:
                match=self._checkmatch(tx,filter_debit)
                if not match:
                    continue
            if (not isstatement) and tx.direction==1 and filter_credit:
                match=self._checkmatch(tx,filter_credit)
                if not match:
                    continue

            leftover_before_statement = acc_last_left[accname]
            native_before_statement = nativebalances[accname]

            descr = tx.verbose(acc.currency)



            new_leftover=tx.applytobalance(leftover_before_statement,currency).as_float()
            new_native=tx.applytobalance(native_before_statement,tx.account.currency).as_float()

            if isstatement and it_was_something_known[acc.name]:
                native_diff= native_before_statement-new_native
                diff=leftover_before_statement-new_leftover
                epsilon=0.009
                isnativelost=abs(native_diff)>epsilon

                if tx.account==virtual_account:
                    diff=poolbalance-new_leftover
                    native_diff= diff


                islost=abs(diff)>epsilon

                if islost:
                    #lost data
                    r=StatementRow()
                    r.type=RowType.Lost
                    r.amount=Money(diff)
                    r.date=tx.time
                    r.account=tx.account
                    sdiff=Currency.verbose(currency,diff)
                    snatdiff=Currency.verbose(tx.account.currency,native_diff)
                    if isnativelost:
                        if sdiff==snatdiff:
                            r.description="Lost data for amount {0}".format(sdiff)
                        else:
                            r.description="Lost data for amount {0} [native diff: {1}]".format(sdiff,snatdiff)
                    else:
                        r.description="Lost data for amount {0} (Due to exchange rate)[native diff: {1}]".format(sdiff,snatdiff)
                    res.Rows.append(r)


            nativebalances[accname] = new_native
            acc_last_left[accname] = new_leftover
            it_was_something_known[accname] = True

            #poolbalance = Money(0)
            poolbalance = 0.0
            for acc in res.Accounts:
               poolbalance+=acc_last_left[acc.name]



            tx_full_amount= tx.get_amount(currency)
            r=self.create_row(RowType.Tx, tx.account,tx,tx.time,new_leftover,poolbalance,descr,tx_full_amount,None )

            allrecs_by_id_index[r.tx.get_id()]=r

            if isstatement:
               r.type=RowType.LeftOver
            else:
                r.tags=tx._tags


            res.Rows.append(r)




        res.poolendofperios=poolbalance

        if not skip_transitions:
            self.apply_transitions(res,allrecs_by_id_index)

        self.cumulate_lefts(res)
        return res
    def trans_get_comission_obj(self,allrecs_by_id_index,trans, tx_commission_id):
        res=None
        if tx_commission_id:
            res=allrecs_by_id_index.get(tx_commission_id)
            if not res:
                mes="Transition failed, tx_comission '{0}' not found".format(tx_commission_id)
                print mes
                return None

        return res
    def apply_transitions(self,st,allrecs_by_id_index):

        todel=[]
        toappend=[]

        for trans in self.Transitions:
            rec_from=allrecs_by_id_index.get(trans.tx_from_id)
            if not rec_from:
                mes="Transition failed2, tx_from '{0}' not found".format(trans.tx_from_id)
                print mes
                continue
                #raise Exception(mes)

            rec_to=None



	    rec_to=allrecs_by_id_index.get(trans.tx_to_id)
            
	    if not rec_to:
                mes="Transition failed, tx_to '{0}' not found".format(trans.tx_to_id)
                print mes
                continue
            #trans.tx_from_commission_obj=None
            rec_commission=self.trans_get_comission_obj(allrecs_by_id_index,trans,trans.tx_commission_id)
            if rec_commission:
               trans.tx_from_commission_obj=rec_commission.tx

            rec_commission2=self.trans_get_comission_obj(allrecs_by_id_index,trans,trans.tx_commission2)
            if rec_commission2:
                trans.tx_from_commission_obj2=rec_commission2.tx

            trans.tx_from_obj=rec_from.tx
            trans.tx_to_obj=rec_to.tx

            rec_to._marktoremove=True


            poolbalance=rec_to.left_pool
            fromacc_balance=rec_from.left_acc


            scomm=""
            if rec_commission:
                commam=rec_commission.amount
                if rec_commission2:
                    commam+=rec_commission2.amount
                scomm="(+{0})".format(commam)
                poolbalance=rec_commission.left_pool
                fromacc_balance=rec_commission.left_acc

            descr=" Transition {0} {1}{4} ->{2} {3}".format(rec_from.account.name,rec_from.amount,rec_to.account.name,rec_to.amount,scomm)
            #print descr
            tx_full_amount=0
            r=self.create_row(RowType.Transition, None,trans,rec_from.tx.time,fromacc_balance,poolbalance,descr,tx_full_amount,None )
            r.left_acc_to=rec_to.left_acc


            rec_to._markappend=r
            rec_to._marktoremove=True
            rec_from._marktoremove=True
            if rec_commission:
                rec_commission._marktoremove=True
            if rec_commission2:
                rec_commission2._marktoremove=True




        cleared=[]
        for row in st.Rows:
            if hasattr(row, '_markappend'):
                cleared.append(row._markappend)
            if hasattr(row, '_marktoremove'):
                continue


            cleared.append(row)

        st.Rows=cleared
        return

    def create_row(self, type, account, tx, date, left_acc, left_pool, description, amount,tags):
        r=StatementRow()
        r.type=type
        r.amount=Money(amount)
        r.account=account
        r.tx=tx
        r.date=date

        if isinstance(left_acc, Money):
            la=left_acc.as_float()
        else:
            la=left_acc

        if isinstance(left_pool, Money):
            lp=left_pool.as_float()
        else:
            lp=left_pool


        r.left_acc=la
        r.left_pool=lp
        r.description=description
        r.tags=tags

        if hasattr(tx,"logical_date"):
            if tx.logical_date:
                r.description=u"[logical {0}/{1}]{2}".format(tx.logical_date.month, tx.logical_date.day, r.description)

        #if hasattr(tx,"human_date"):
        #    if tx.human_date:
        #        diff=(tx.human_date-date).days
        #        if diff!=0:
        #            r.description=u"[{0}/{1}]{2}".format(tx.human_date.month, tx.human_date.day, r.description)


        return r

    def cumulate_lefts(self,res):
        acc_last_left={}
        #acc_last_left_knownletover={}

        st=res
        for acc in st.Accounts:
            acc_last_left[acc.name]=0.


        for row in st.Rows:
            if row.account:
                accname=row.account.name
                acc_last_left[accname]=row.left_acc
            if row.type==RowType.Transition:
                trans=row.tx
                acc_last_left[trans.tx_from_obj.account.name]=row.left_acc
                acc_last_left[trans.tx_to_obj.account.name]=row.left_acc_to

            row.cumulatives=dict(acc_last_left)

   
class Transition:
      def __init__(self,  TxFrom, TxTo, commission,commission2):
        self.tx_from_id=TxFrom
        self.tx_to_id=TxTo
        self.tx_commission2=commission2
        self.tx_commission_id=commission
        return
