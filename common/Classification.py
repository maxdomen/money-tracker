# -*- coding: utf8 -*-
from datetime import datetime, timedelta

import xlrd
import xlwt

__author__ = 'Max'

class TagTools:
    @staticmethod
    def ConvertStringOfTagsToList(li,str):
        if len(str)>0:
            tsets=str.split(",")
            for s in tsets:
                a=s.strip().lower()
                hashpos=a.find("#")
                if hashpos>=0:
                    a=a[hashpos+1:len(a)]

                if li.count(a)<1:
                    li.append(a)

    @staticmethod
    def TagsToStr(tags):
        smulti=""
        for tag in tags:
            if len(smulti)>0:
                smulti+="+"
            smulti+=tag
        return smulti


class Category:
    def __init__(self,title, tags_expression=None):
        self.title=title.strip()
        self.tags=[]

        if tags_expression:
            self.tags=self._str_to_tags(tags_expression)

        self.childs=[]
        self._collapsed=False
        self._hided=False
        self._sid=""
        self._index=-1
        self.parent=None
        self._level_index=-1
        self._highlighted=False
    def add(self,cat):
        cat.parent=self
        self.childs.append(cat)
    def _add_tags_row(self, tags):
        row=[]
        for t in tags:
            row.append(t.lower())

        self.tags.append(row)
    def _str_to_tags(self, sexpr):
        res=[]
        parts=sexpr.lower().split(",")
        for p in parts:
            row=[]
            tags=p.split("+")
            for t in tags:
                st=t
                if st[0]=="@":
                    st=st[1:len(st)-1]
                row.append(st.strip())
            res.append(row)

        #print res
        return res
class Classification:
    def __init__(self,from_xls=None):

        self._root=Category(u"_root")


        if from_xls:
            filename=from_xls[0]
            spreadsheet=from_xls[1]
            book = xlrd.open_workbook(filename,spreadsheet)

            self._load_from_xls(book,spreadsheet)


        self._uncategorized=Category(u"Без категории")
        self._root.add(self._uncategorized)
        self._untagged=Category(u"Без тегов")
        self._uncategorized.add(self._untagged)
        self._auto_categorized=None

    def get_category_by_id(self,sid):
        for c in self.cat_array:
            if c._sid==sid:
                return c
        return None
    def create_auto_classification(self,statement):
        self._auto_categorized=Category(u"Автосозданные категории")
        self._root.add(self._auto_categorized)

        self.finalize()

        for row_date,value,tags in statement.get_generator():
            res=self._match_tags_to_category(tags)
            if not res:
                res=self._uncategorized
                if len(tags)>0:
                    title=TagTools.TagsToStr(tags)
                    #print "auto-create classification", title
                    autocat=Category(title)
                    #autocat.tags.append(list(tags))
                    autocat._add_tags_row(tags)
                    self._auto_categorized.add(autocat)
                    #res=autocat
                    self.finalize()


    def _load_from_xls(self,workbook,spreadsheet):
        sheet=workbook.sheet_by_name(spreadsheet)

        prevs={}
        prevs[1]=self._root
        for rowi in range(1, sheet.nrows):
            r=sheet.row(rowi)
            sid=r[0].value
            for coli in range(1,6):
                title=r[coli].value
                if len(title)<1:
                    continue

                salltags=""
                for coli2 in range(7,10):
                    st=r[coli2].value
                    if len(st)>0:
                        if len(salltags)>0:
                            salltags+=","
                        salltags+=st


                cat=Category(title,salltags)
                parent=prevs[coli]
                #print parent.title,"/",title,"(",salltags,")"

                spi=title.find("[")
                if spi>0:
                    command=title[spi+1:spi+2]
                    #print "COMMAND",command
                    title= title[0:spi]
                    cat.title=title
                    if command=="-":
                        cat._collapsed=True

                    if command=="=":
                        cat._highlighted=True
                if len(sid)>0:
                    cat._sid=sid
                parent.add(cat)
                prevs[coli+1]=cat
                break
    def _match_tags_to_category(self,tags):


        if len(tags)<1:
            return  self._untagged

        matches=[]
        max_combinationprecise=0
        for c in self.cat_array:


            for tags_combination in c.tags:
                matched=False
                combinationprecise=len(tags_combination)
                for tag in tags_combination:
                    matched=tags.count(tag)>0
                    if not matched:
                        break

                if matched:
                    #все теги комбинации есть во входном списке
                    #print "   match", c.title
                    matches.append( (c,combinationprecise) )
                    if combinationprecise>max_combinationprecise:
                        max_combinationprecise=combinationprecise
                    matched=False

        if len(matches)>0:

            return self.select_best(matches,tags,max_combinationprecise)

        #print "   match None"
        return None
    def select_best(self,matches,tags,max_combinationprecise):
        res= matches[0][0]
        if len(matches)>1:
            #print "many matches", len(matches),TagTools.TagsToStr(tags), "max_combinationprecise=",max_combinationprecise
            for m,hc in matches:
                allcombs=""
                for tt in m.tags:
                    allcombs=allcombs+" "+str(tt)
                #print "    ",m.title, hc,allcombs
                if hc==max_combinationprecise:
                    res=m
                    break
            #print "    ->", res.title
        return res
    def match_tags_to_category(self,tags):
        res=self._match_tags_to_category(tags)
        if not res:
            res=self._uncategorized

        return res
    def finalize(self):
        self.cat_array=[]
        self.cat_maxindex=0
        self._finalize_layout(0,self._root,False)

    def _finalize_layout(self,deep,node=None, hide=None):


        node._hided=hide
        node._level_index=deep
        #self.cat_array.append(node)

        if hide:
            node._index=-1
        else:
            node._index=self.cat_maxindex
            self.cat_maxindex+=1

        self.cat_array.append(node)

        if node._collapsed:
            hide=True

        for c in node.childs:
            self._finalize_layout(deep+1,c,hide)

class Period:
    Day=1
    Week=2
    Month=3
    def __init__(self, start, end, maxindex):
        self._start=start
        self._end=end
        self._cells=[0.0]*maxindex

    @staticmethod
    def CreateSet(chunktype,start,end,maxindex=100):
        periods=[]

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
            else:
                if chunktype==2: #weeks
                    wd=period.weekday()
                    period_step=timedelta(7-wd)
                    #nextp=period+period_step
                else:
                    if chunktype==3: #months
                        nextyear=period.year
                        nextmonth=period.month+1
                        if nextmonth>12:
                            nextmonth=1
                            nextyear+=1
                        firstdateofnextmonth=datetime(nextyear,nextmonth,1)
                        period_step=firstdateofnextmonth-period
                    else:
                        raise Exception("unknown chunk type {0}".format(chunktype))

            p=Period(period,period+period_step+timedelta(seconds=-1),maxindex)
            periods.append(p)
            period+=period_step
        return periods

class ClassificationDataset:
    def __init__(self,classification,period_quant,sourcedata=None, create_empty=None, reversed=False):

        if not isinstance(period_quant,int):
            raise Exception("bad parameter type. 'int' expected")
        self.periods=None
        classification.finalize()
        self.classification=classification
        if create_empty:
            #period_quant=create_empty[0]
            time_start=create_empty[0]
            time_finish=create_empty[1]
            self.periods=Period.CreateSet(period_quant,time_start,time_finish,classification.cat_maxindex+10)

        if sourcedata:
            #self._finalize_layout(classification._root,False)
            time_start=sourcedata.get_time_start()
            time_finish=sourcedata.get_time_finish()
            self.periods=Period.CreateSet(period_quant,time_start,time_finish,classification.cat_maxindex+10)
            self.process_list(sourcedata)

        if reversed and self.periods:
            self.periods=sorted(self.periods, key=lambda p:p._start,reverse=True)

    def merge(self,src):
        if not self.periods:
            self.periods=src.periods
            return

        cmap={}
        ind=0
        for srcc in src.classification.cat_array:
            cmap[ind]=None
            targeti=0
            for c in self.classification.cat_array:
                if c.title==srcc.title:
                    cmap[ind]=targeti
                targeti+=1
            ind+=1

        for srcp in src.periods:
            timepoint=srcp._start+timedelta(seconds=(srcp._end-srcp._start).seconds/2)
            for p in self.periods:
                if timepoint>=p._start and timepoint<=p._end:
                    srci=0
                   # print p
                    for srcc in src.classification.cat_array:
                        targeti= cmap[srci]
                        if srci>=len(srcp._cells):
                            break
                        srcv=srcp._cells[srci]
                        p._cells[targeti]+=srcv
                        srci+=1
    def process_list(self, sourcedata):

        for row_date,value,tags in sourcedata.get_generator():
            self.classify_value(row_date,value,tags)

    def classify_value(self,row_date,value,tags):
        for p in self.periods:
            if row_date>=p._start and row_date<=p._end:
                g=self._row(p,value,tags)
                return g

    def _row(self,period, value,tagsin):
        group=None
        #root=self.layout.income if row.tx.direction==1 else self.layout.spending
        if tagsin==None:
            tagsin=[]
        tags=list(tagsin)
        group=self.classification.match_tags_to_category(tags)
        if group._hided:
            while group._hided:
                group=group.parent

        #if isinstance(value, unicode) or isinstance(value, str):
        if isinstance(value, float):

            period._cells[group._index]+=value
        else:
            shouldbelist=period._cells[group._index]
            if not isinstance(shouldbelist, list):
                shouldbelist=list()
                period._cells[group._index]=shouldbelist

            shouldbelist.append(value)
        #else:

        return group
class ClassificationPrinter:
    @staticmethod
    def print_titles(dataset,ws,startrowi):
        rowi=startrowi+1
        hst = xlwt.easyxf('pattern:   pattern solid, fore_colour ice_blue ;')
        #hst = xlwt.easyxf('pattern: pattern solid;')


        #hst.pattern.pattern_fore_colour = 0x1F

        for category in dataset.classification.cat_array:
            #print category.title,category._index
            coli=0
            if category._index>=0:
                rowi=startrowi+category._index+1

                title=category.title
                #_level_index
                for i in range(0,category._level_index):
                    title="     "+ title


                if category._highlighted:
                    ws.write(rowi, coli, title,hst)
                else:
                    ws.write(rowi, coli, title)
            else:
                pass
                #print "hided category", category.title



    def __init__(self, dataset,existing_sheet,startrowi=1):
        self.ws=existing_sheet
        self.hst = xlwt.easyxf('pattern:   pattern solid, fore_colour ice_blue ;',num_format_str='#,##0')
        self.hsti = xlwt.easyxf('pattern:   pattern solid, fore_colour ice_blue ;font: italic on;',num_format_str='#,##0')


        #print startrowi
        ClassificationPrinter.print_titles(dataset,existing_sheet,startrowi)
        #rowi=startrowi
        coli=2
        self.style_time1 = xlwt.easyxf(num_format_str='D-MMM')
        self.style_money=xlwt.easyxf(num_format_str='#,##0')
        #startrowi+=1

        for p in dataset.periods:
            self.ws.write(startrowi, coli, p._start,self.style_time1)
            for category in dataset.classification.cat_array:
                if  category._hided:
                    continue
                #coli=0
                v=p._cells[category._index]
                rowi=startrowi+category._index+1
                subtotals=0
                if category._highlighted:
                    subtotals=self.calcsubtotals(category,p)

                    style=self.hst
                    if v==0:
                        style= self.hsti
                    v=v+subtotals
                    self.ws.write(rowi, coli, v,style)

                else:
                    if v>0:
                        self.ws.write(rowi, coli, v,self.style_money)

                #rowi+=1
            coli+=1
    def calcsubtotals(self,category,p):
        res=0
        for c in category.childs:
            v=p._cells[c._index]
            res=res+v
            res=res+self.calcsubtotals(c,p)
        return res