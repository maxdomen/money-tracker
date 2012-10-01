# -*- coding: utf8 -*-
#import decimal
from xlwt.Style import XFStyle
from accounts import *

#from accounts import RowType
__author__ = 'Max'
from django.template import Template, Context
from django.conf import settings
import django.template.loader as loader
import xlwt
from currency import Currency
from datetime import datetime,timedelta



class PrintStatementToExcel2:
    def __init__(self, filename, sheetname, existing_workbook=None):
        self.filename=filename

        if existing_workbook:
            self.wb = existing_workbook
        else:
            self.wb = xlwt.Workbook()

        self.ws = self.wb.add_sheet(sheetname)


        self._chunktype=1

        self.style_money=xlwt.easyxf(num_format_str='#,##0.00')
        self.style_money_gray=xlwt.easyxf('font: color-index gray25',num_format_str='#,##0.00')
        self.style_money_bold=xlwt.easyxf('font: bold on',num_format_str='#,##0.00')

    def set_period(self,start,end):
        self._start=start
        self._end=end
        h=end.hour
        m=end.minute
        s=end.second

        #if h==0 and m==0 and s==0:
        h=23
        m=59
        s=59
        newend=datetime(end.year, end.month, end.day,h,m,s,100)
        self._end=newend

    def set_chunk(self,type):
        self._chunktype=type


    #def save(self):

    #    self.wb.save(self.filename)
        
    def do_print(self,st):

        ws=self.ws
        style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
	    num_format_str='#,##0.00')
        #style_time1 = xlwt.easyxf(num_format_str='D-MMM-YY')
        style_time1 = xlwt.easyxf(num_format_str='D-MMM')

        style_money_red=xlwt.easyxf('font: bold on, color-index red',num_format_str='#,##0.00')
        style_text_red=xlwt.easyxf('font: color-index red',num_format_str='#,##0.00')
        style_text_green=xlwt.easyxf('font: color-index green',num_format_str='#,##0.00')

        #ws = self.wb.add_sheet('Txs')

        ws.col(5).width=256*40
        ws.col(0).width=256*8
        ws.col(1).width=256*6
        ws.col(2).width=256*6

        ws.col(3).width=256*10
        ws.col(4).width=256*10
        #ws.set_panes_frozen(True)
        ws.panes_frozen = True
        ws.horz_split_pos = 1
        ws.vert_split_pos = 1
        ws.normal_magn=70

        for ci in range(6,12):
            ws.col(ci).width=256*12
        ws.col(13).width=256*14

        ri=0


        #ri+=1
        bc=0

        acc_col_idx={}
        #acc_last_left={}
        acc_last_left_knownletover={}

        t=0
        for acc in st.Accounts:
            acc_col_idx[acc.name]= t
            #acc_last_left[acc.name]=Money(0)
            acc_last_left_knownletover[acc.name]=False
            ws.write(ri, bc+6+t, acc.name)
            t+=1
        ws.write(ri, bc+6+t, "Pool")

        #start of period

        ri+=1
        for row in st.Rows:

            

            #надо учитывать остатки даже за пределами диапазаон времени, иначе неконсистенность в цифрах
            if row.account:
                accname=row.account.name
                #acc_last_left[accname]=row.left_acc
                acc_last_left_knownletover[accname]= (row.type==RowType.LeftOver)
            

            if row.date>self._end or row.date<self._start:
                continue

            descr_col_ind=bc+5
            if row.type==RowType.Lost:
                 #style=style_money_red

                 ws.write(ri, bc+0, row.date,style_time1)
                 ws.write(ri, bc+1, row.account.name)
                 ws.write(ri,descr_col_ind, row.description, style_text_red)
                 ri+=1
                 continue
            if row.type==RowType.Transition:
                trans=row.tx
                ws.write(ri, bc+0, row.date,style_time1)
                ws.write(ri, descr_col_ind, row.description,style_text_green)

                accs=[trans.tx_from_obj.account.name,trans.tx_to_obj.account.name]
                #acc_last_left[accs[0]]=row.left_acc
                #acc_last_left[accs[1]]=row.left_acc_to
                self.print_accounts_lefts(st,ws,ri,descr_col_ind,accs,False,acc_col_idx,row.cumulatives)
                ws.write(ri, descr_col_ind+1+t, row.left_pool,self.style_money)
                ri+=1
                continue


            amount_col_index=bc+3


            if row.type==RowType.Tx:
                if row.tx.direction==1: #income
                    amount_col_index=bc+4




            self.print_accounts_lefts(st,ws,ri,descr_col_ind,[accname],row.type==RowType.LeftOver,acc_col_idx,row.cumulatives)



            #ws.write(ri, bc+5+this_acc_col_idx, accleft,style_money)
            ws.write(ri, bc+0, row.date,style_time1)
            ws.write(ri, bc+1, accname)
            if row.type!=RowType.LeftOver:
                 ws.write(ri, amount_col_index, row.amount.as_float(),self.style_money)



            ws.write(ri, descr_col_ind, row.description)

            rate=Currency.getrate(row.date,row.tx.account.currency, st.currency)


            if row.type==RowType.Tx:
                if row.tx.original_currency!=None:
                    if row.tx.original_currency!=st.currency:
                        rate=row.tx._amount/row.tx.original_amount

            


            ws.write(ri,bc+2,rate.as_float(),self.style_money)


            pool_known_leftover=True
            for acc in st.Accounts:
                latval=row.cumulatives[acc.name]
                vt=acc_last_left_knownletover[acc.name]

                if not latval==0:
                    pool_known_leftover=(pool_known_leftover and vt)


            if pool_known_leftover:
                ws.write(ri, descr_col_ind+1+t, row.left_pool,self.style_money_bold)
            else:
                ws.write(ri, descr_col_ind+1+t, row.left_pool,self.style_money)

            srcinfoi=descr_col_ind+1+t+1

            src=row.tx.src
            if src:
                ws.write(ri,srcinfoi, row.tx.src.verbose())
            ws.write(ri,srcinfoi+1, row.tx.get_id())

            if row.type==RowType.Tx:
                stags=""
                for t1 in row.tags:
                    if len(stags)>0:
                        stags+=", "
                    stags+=t1
                ws.write(ri,srcinfoi+2, stags)

                if hasattr(row, 'classification'):
                    ws.write(ri,srcinfoi+3, row.classification.title)



            ri+=1

        #end of period    
        #str = 'End of period: {0} '.format(Currency.verbose(st.currency, st.poolendofperios))

    def print_accounts_lefts(self, st, ws,ri,descr_col_ind, accs_to_highlight, makebold,acc_col_idx,acc_last_left):
        for acc in st.Accounts:
                style=self.style_money_gray

                if accs_to_highlight.count(acc.name)>0:
                    style=self.style_money
                    if makebold:
                        style=self.style_money_bold

                #left=acc_last_left[acc.name].as_float()
                left=acc_last_left[acc.name]

                ws.write(ri, descr_col_ind+1+acc_col_idx[acc.name],left,style)



class BalanseObservation():
    def __init__(self, filename, sheetname,existing_workbook=None):
        self.filename=filename
        self.wb = existing_workbook
        self.ws = self.wb.add_sheet(sheetname)
        
        self.style_money=xlwt.easyxf(num_format_str='#,##0.00')
        #self._chunktype=1
        self.tags_credit=[u"Под отчет",u"Деньги CM"]
        self.tags_debit=[u"Reimbursment",u"Деньги CM"]
    def do(self,st):
        ws=self.ws
        ri=1
        bc=0
        style_time1 = xlwt.easyxf(num_format_str='D-MMM')
        self.ws.col(2).width=256*20
        self.ws.col(3).width=256*60
        self.ws.panes_frozen = True
        self.ws.horz_split_pos = 1
        self.ws.vert_split_pos = 1
        self.ws.normal_magn=80


        descr_col_ind=3
        balance_col_index=6
        balance=Money(0)

        for row in st.Rows:

            if row.type!=RowType.Tx:
                continue

            t=[]
            tags=row.tags
            amnt=row.amount
            found=False
            if row.tx.direction==1:
                for t in self.tags_credit:
                    tc=tags.count(t)
                    if tc>0:
                        found=True
                        balance+=amnt
                        break
            if row.tx.direction==-1:
                for t in self.tags_debit:
                    tc=tags.count(t)
                    if tc>0:
                        found=True
                        balance-=amnt
                        break
            if not found:
                continue

            ws.write(ri, bc+0, row.date,style_time1)
            #ws.write(ri, bc+1, accname)

            amount_col_index=4
            if row.tx.direction==1:
                amount_col_index=5

            ws.write(ri, amount_col_index, row.amount,self.style_money)
            ws.write(ri, balance_col_index, balance,self.style_money)

            ws.write(ri, descr_col_ind, row.description)

            stags=""
            for t in tags:
                if self.tags_debit.count(t)>0:
                    continue
                if len(stags)>0:
                        stags+=", "


                stags+=t
            ri+=1

            ws.write(ri, descr_col_ind-1, stags)



        return
