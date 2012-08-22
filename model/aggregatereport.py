# -*- coding: utf8 -*-
#import decimal

from xlwt.Style import XFStyle
from accounts import *

__author__ = 'Max'
import xlwt
from datetime import datetime,timedelta


class Group:
    def __init__(self,title):
        self._collapsed=False
        self._hided=False
        self.parent=None
        self.title=title
        self.childs=[]
        self.tags=[]
    def __str__(self):
        return self.title
    def get_short_title(self):

        st=list(self.tags)
        node=self.parent
        while node:
            for pt in node.tags:
                if st.count(pt)>0:
                    st.remove(pt)
            node=node.parent

        res=Layout.tags_to_str(st,-1).strip()

        if len(res)<1:
            res=self.title
        return res
    def addgroup(self,title=None, has_tag=None, has_all_tags=None, any_tag=None, options=None):
        g=Group(title)
        g.parent=self
        g.title=title

        if has_all_tags:
            g.tags=has_all_tags
        if has_tag:
            g.tags.append(has_tag)
        if any_tag:
            for at in any_tag:
                g.addgroup(has_tag=at)
        self.childs.append(g)

        if not title:
            g.title="<column>"
            if has_tag:
                g.title=has_tag
            if has_all_tags:
                g.title=Layout.tags_to_str(g.tags,-1)
            if any_tag:
                g.title=""
                for at in any_tag:
                    if len(g.title)>0:
                        g.title+=" or "
                    g.title+=at

        return g
    def collapse(self):
        self._collapsed=True
    def expand(self):
        self._collapsed=False
class Layout:
    def __init__(self):
        self.root=Group("/")
        self.income=self.root.addgroup("income")
        self.income.untagged=self.income.addgroup("untagged")
        self.spending=self.root.addgroup("spending")
        self.spending.untagged=self.spending.addgroup("untagged")
        self.lost=self.spending.addgroup("lost")

    def clone(self):
        pass

    @staticmethod
    def tags_to_str(tags, direction):
        smulti=""
        for tag in tags:

            if len(smulti)>0: smulti+="+"

            if direction==1:
                smulti+="In "
            smulti+=tag
        return smulti
    def create_automatic_groups(self, statement):
        #enum
        tags_combinations1={}
        tags_combinations2={}
        for row in statement.Rows:
            if row.type!=RowType.Tx:
                continue
            tags=row.tags
            enum=tags_combinations1  if row.tx.direction==-1 else tags_combinations2

            key=Layout.tags_to_str(tags,row.tx.direction)
            if not enum.has_key(key):
                enum[key]=tags

        #create
        #sort
        tags_combinations1=sorted(tags_combinations1.itervalues(),key=lambda x: len(x))
        tags_combinations2=sorted(tags_combinations2.itervalues(),key=lambda x: len(x))
        self.create_auto_columns(tags_combinations1,self.spending,-1  )
        self.create_auto_columns(tags_combinations2,self.income,1 )

    def _counthits(self,node,tags):
        hit=0
        for t in node.tags:
            if tags.count(t)>0:
                hit+=1
        return hit

    def _best_match(self,node,tags):
        target_c=len(tags)
        hits=0
        #create=True
        res=None

        if len(node.tags)>0:
            hits=self._counthits(node,tags)
            if hits<1:
                return None
            if hits>0:
                for te in node.tags:
                    if tags.count(te)>0:
                        tags.remove(te)
                res=node

        for c in node.childs:
            better=self._best_match(c,tags)
            if better:
                res=better
                break

        return res
    def find_best_group_for_tags(self,root,tags2):

        tags=list(tags2)
        better=self._best_match(root,tags)
        if not better:
            better=root

        create=True
        hits=self._counthits(better,tags)
        if len(tags)==hits:
            create=False



        return better, create

    def create_auto_columns(self, enum, root,dir):
        for tags in enum:
            if len(tags)<1:
                continue

            best_parent, create=self.find_best_group_for_tags(root,tags)
            if not create:
                #row.hint_column=best_group
                continue

            title=Layout.tags_to_str(tags,dir)
            best_parent.addgroup(title, has_all_tags=tags)


    def _finalize(self,node=None, hide=None):

        if not node:
            node=self.root
            self.array=[]
            self.maxindex=0
            hide=False

        node._hided=hide
        if node._collapsed:
            hide=True
        node.index=self.maxindex
        self.maxindex+=1
        self.array.append(node)


        for c in node.childs:
            self._finalize(c,hide)



    def verbose(self,gc=None, deep=None):
        self._finalize()
        res=""
        g=gc
        if g==None:
            g=self.root
        d=deep
        if d==None:
            d=0
        res+=str(g.index)+" "+g.title
        if len(g.childs)>0:
            res+="/\r\n"

        for c in g.childs:
            for e in range(0,d*2):
                res+=" "
            res+=self.verbose(c,d+1)+"\r\n"
        return res

class Period:
    def __init__(self, start, end, maxindex):
        self._start=start
        self._end=end
        self._cells=[Money(0)]*maxindex

    @staticmethod
    def CreateSet(chunktype=1,statement=None, start=None, end=None,maxindex=0):
        periods=[]

        if start==None:
            start=statement.Rows[0].date
        if end==None:
            end=statement.Rows[len(statement.Rows)-1].date+timedelta(seconds=1)

        h=end.hour
        m=end.minute
        s=end.second

        h=23
        m=59
        s=59
        newend=datetime(end.year, end.month, end.day,h,m,s,100)
        end=newend

        period=datetime(start.year,start.month,start.day)

            #totalrows=len(st.Rows)
            #currentrow=0

        while period<=end:

            if chunktype==1:#days
                    period_step=timedelta(1)
            if chunktype==2: #weeks
                    wd=period.weekday()
                    period_step=timedelta(7-wd)
                    #nextp=period+period_step
            if chunktype==3: #months
                    nextyear=period.year
                    nextmonth=period.month+1
                    if nextmonth>12:
                        nextmonth=1
                        nextyear+=1
                    firstdateofnextmonth=datetime(nextyear,nextmonth,1)
                    period_step=firstdateofnextmonth-period

            p=Period(period,period+period_step+timedelta(seconds=-1),maxindex)
            periods.append(p)
            period+=period_step
        return periods
class Dataset:
    def __init__(self, layout, statement, start=None, end=None,chunktype=1):
        layout._finalize()
        self.periods=[]
        self.layout=layout
        self._chunktype=chunktype
        #if start==None:
        #    start=statement.Rows[0].date
        #if end==None:
        #    end=statement.Rows[len(statement.Rows)-1].date+timedelta(seconds=1)
        #self._set_period(start,end)
        self._go(statement, start, end)
    #def _set_period(self,start,end):
    #    self._start=start
    #    self._end=end
    #    h=end.hour
    #    m=end.minute
    #    s=end.second

    #    h=23
    #    m=59
    #    s=59
    #    newend=datetime(end.year, end.month, end.day,h,m,s,100)
    #    self._end=newend

    #def set_chunk(self,type):
    #    self._chunktype=type

    def _get_logical_date(self,row):
        res=row.date
        if row.tx and row.type==RowType.Tx and  row.tx.logical_date:
            res=row.tx.logical_date
            #print "use logical date", res, row.tx.get_id()
        return res
    def _go(self,st,start, end):

        self.periods=Period.CreateSet(self._chunktype,st,start,end, self.layout.maxindex)

        #обработка строк
        for row in st.Rows:
            row_date=self._get_logical_date(row)
            for p in self.periods:
                if row_date>=p._start and row_date<=p._end:
                    self._row(p,row)
                    break


    def _row(self,period, row):



        group=None
        if row.type==RowType.Lost:
            group=self.layout.lost
        else:
            if row.type!=RowType.Tx:
                return
            root=self.layout.income if row.tx.direction==1 else self.layout.spending
            tags=list(row.tags)
            group=self.layout._best_match(root,tags)
            if not group:
                group=root.untagged

        if group._hided:
            while group._hided:
                group=group.parent
                
        period._cells[group.index]+=row.amount

class AggregateRow:
    def __init__(self, title,max_period_index):
        self.title= title
        self._cells=[Money(0)]*max_period_index
        self.sumgroups=None
        self.decrows=None
        self.cumulrow=None


class Aggregate:
    def __init__(self, dataset):
        self.dataset= dataset

        self.max_period_index=len(dataset.periods)
        self.rows=[]
        pass

    def CreateRow(self, title, sumgroups):
        row=AggregateRow(title,self.max_period_index)
        row.sumgroups=sumgroups
        self.rows.append(row)
        return row

    def CreateRowCalc(self,title, r1,r2):
        row=AggregateRow(title,self.max_period_index)
        self.rows.append(row)
        row.decrows=[]
        row.decrows.append(r1)
        row.decrows.append(r2)
        return row
    def CreateCumulative(self,title, r1):
        row=AggregateRow(title,self.max_period_index)
        self.rows.append(row)
        row.cumulrow=r1

        return row



    def go(self):
        self.periods=self.dataset.periods
        self.dataset.layout._finalize()

        for p in self.periods:
            dif=self.dataset.layout.maxindex-len(p._cells)
            if dif>0:
                #print "EXPAND"
                p._cells.extend( [0]*dif)

        #создаем связь группа - аггрегирующая строчка
        for group in self.dataset.layout.array:
            group._analytics_rows=[]
            for row in self.rows:
                if row.sumgroups:
                    for sumgroup in row.sumgroups:
                        if sumgroup==group:
                            group._analytics_rows.append(row)


        iperiod=0
        for period in self.dataset.periods:
            #аккумулирование значений
            for group in self.dataset.layout.array:
                self._appendcell(iperiod,period, group)
                #value=period._cells[group.index]

            #вычислистельные строки
            for row in self.rows:
                if row.decrows:
                    drow1=row.decrows[0]
                    drow2=row.decrows[1]
                    val=drow1._cells[iperiod]-drow2._cells[iperiod]
                    row._cells[iperiod]=val
                if row.cumulrow:
                    drow1=row.cumulrow
                    preval=0;
                    if iperiod>0:
                        preval=row._cells[iperiod-1]
                    #if row._prevval:
                    #    preval=row._prevval
                    newval=preval+drow1._cells[iperiod]
                    row._cells[iperiod]=newval
            iperiod+=1




    def _calc_group_value(self, period,group):
        res=period._cells[group.index]
        for c in group.childs:
            c_res=self._calc_group_value(period,c)
            res+=c_res
        return res

    def _appendcell(self,iperiod,period, group):

        if group.title==u"Зарплата":
            print group.title

        #группа учавствует в аналитических строках?
        if len(group._analytics_rows)<1:
            return

        #рассчитаем значение сроки и всех ее детей
        value=self._calc_group_value(period,group)

        for row in group._analytics_rows:
            row._cells[iperiod]+=value

class Publisher2:
    def __init__(self, table, filename, sheetname, existing_workbook=None, existing_sheet=None, after_row=None, sub_report_title=None ):
        self.table=table
        self.filename=filename
        self.after_row=after_row
        self.sub_report_title=sub_report_title
        if not existing_workbook:
            self.wb = xlwt.Workbook()
        else:
            self.wb = existing_workbook

        #self.wb.
        self.ws=existing_sheet


        if not existing_sheet:
            self.ws = self.wb.add_sheet(sheetname)


        self._printer()

    def save(self):
        self.wb.save(self.filename)


    def _printer(self):
        self.ws.panes_frozen = True


        self.ws.normal_magn=80


        #self.max_rowi=0
        #self._print_titles(self.dataset.layout.root,0,1)


        date_style_w1=xlwt.easyxf('',num_format_str='D-MMM')

        style_money=xlwt.easyxf(num_format_str='#,##0')
        #rowi=self.max_rowi+1

        startrowi=1

        if  self.after_row:
            startrowi= self.after_row

        if self.sub_report_title:
            #self.sub_report_title=sub_report_title
            self.ws.write(startrowi, 0, self.sub_report_title)
            startrowi+=1

        rowi=startrowi+1
        for row in self.table.rows:
            coli=0
            self.ws.write(rowi, coli, row.title)
            rowi+=1
        coli=1
        iperiod=0
        for period in self.table.periods:
            rowi=startrowi
            self.ws.write(rowi, coli, period._start,date_style_w1)
            #coli=1
            rowi=startrowi+1
            for row in self.table.rows:
                val=row._cells[iperiod]

                if isinstance(val, Money):
                    val=val.as_float()
                self.ws.write(rowi, coli, val,style_money)

                rowi+=1
            coli+=1
            iperiod+=1



class Publisher:
    def __init__(self, dataset, filename, sheetname, existing_workbook=None, ):
        self.dataset=dataset
        self.filename=filename
        if not existing_workbook:
            self.wb = xlwt.Workbook()
        else:
            self.wb = existing_workbook

        self.ws = self.wb.add_sheet(sheetname)
    def _print_titles(self, node, rowi, coli):
        if rowi>self.max_rowi:
            self.max_rowi=rowi

        if node!=self.dataset.layout.root:
            node.print_index=coli
            short=node.get_short_title()
            self.wswrite(rowi, coli, short)
        else:
            rowi-=1

        ci=coli

        cind=0
        for c in node.childs:
            cind+=1
            if c._hided:
                continue
            if len(node.tags)<1 and cind== 1 and  self.mode==2:
                ci-=1
            ci+=1
            ci=self._print_titles(c, rowi+1, ci)


        return ci

    def save(self):
        self.wb.save(self.filename)

    def horizontal(self):
        self.mode=1
        self._printer(self)
        self.ws.vert_split_pos = self.max_rowi+1
    def vertical(self):
        self.mode=2
        self._printer(self)
        self.ws.horz_split_pos = self.max_rowi+1
    def wswrite(self, x,y, str,style=None):

        rowi=x
        coli=y

        if self.mode==1:
            rowi=y
            coli=x

        if style:
            self.ws.write(rowi, coli, str, style)
        else:
            self.ws.write(rowi, coli, str)
    def _printer(self, mode):
        self.ws.panes_frozen = True

        #ensure cell room
        #for p in self.dataset.self.dataset.periods

        self.ws.normal_magn=80


        self.max_rowi=0
        self._print_titles(self.dataset.layout.root,0,1)


        date_style_w1=xlwt.easyxf('',num_format_str='D-MMM')

        style_money=xlwt.easyxf(num_format_str='#,##0')
        rowi=self.max_rowi+1
        for period in self.dataset.periods:
            sizediff=self.dataset.layout.maxindex-len(period._cells)
            if sizediff>0:
                for r in range (0,sizediff+1):
                    period._cells.append(Money(0))

            coli=0
            self.wswrite(rowi, coli, period._start,date_style_w1)
            coli=1
            for group in self.dataset.layout.array:
                value=period._cells[group.index]
                if value!=0:

                    self.wswrite(rowi, group.print_index, value.as_float(),style_money)
                coli+=1
            rowi+=1


