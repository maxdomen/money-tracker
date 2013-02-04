#! coding: utf-8
from collections import OrderedDict

__author__ = 'Max'
from  common.Classification import *
from common.Table import Table, Style, Color,DestinationXls
from model.budget import BudgetFreq

from model.accounts import *
import copy
class WeekDef:

    def __init__(self,t):
        self.startday=t
        self.lastday=None
        self.windex=0
        self.plan_total=0
        self.fact_total=0
        self.predict_total=0
        #self.use_fact=False

def budget_weekly_planner(caption,period, plan, clasfctn2, fact,budget):

    mtitle="{0} {1}".format(period.start.strftime("%B"),period.start.year)
    table=Table(mtitle)
    #self._normal_magn=0
    table[0,0]=mtitle #название месяца для которого делается отчет

    table._normal_magn=100
    table.define_style("totals", foreground_color=Color.Red)
    table.define_style("weekcaptions", bold=True, font_size=10)
    table.define_style("categoryline", bold=True, background_color=Color.LightGreen)
    table.define_style("categoryline_totals", italic=True,background_color=Color.LightGreen, formatting_style=Style.Money)
    table.define_style("item_plan",background_color=Color.LightGray, formatting_style=Style.Money)
    table.define_style("item_plan_overdue",background_color=Color.Red,foreground_color=Color.White, formatting_style=Style.Money)

    table.define_style("percent_green",foreground_color=Color.Green, formatting_style=Style.Percent)
    table.define_style("percent_red",foreground_color=Color.Red, formatting_style=Style.Percent)

    table.define_style("accum",bold=True, formatting_style=Style.Money)
    table.define_style("item_fact", formatting_style=Style.Money)

    clasfctn=copy.deepcopy(clasfctn2)

    t=period.start
    weeks=[]
    w=WeekDef(t)
    weeks.append(w)

    #находим первый и последний день недели
    while t<period.end:
        w.lastday=t
        if t.weekday()==6:
            #это воскресенье
            #ечли это последний день месяца, то не создаем  новую неделю
            firstDayOfNextWeek=t+timedelta(days=1)
            if firstDayOfNextWeek<period.end:
                i=w.windex
                w=WeekDef(firstDayOfNextWeek)
                w.windex=i+1
                weeks.append(w)
        t+=timedelta(days=1)

    #заголовки колонок недель
    rowi=1
    coli=3
    now=datetime.now()

    for w in weeks:
        w.coli=coli
        spast="[p]"
        if w.startday<now: spast=""
        if now>=w.startday and now<=w.lastday:spast="[c]"

        sdays="{0}-{1}".format(w.startday.day,w.lastday.day)
        table[rowi,coli]=u"{0}{1}".format(sdays,spast), "weekcaptions"
        coli+=2


    budget_weekly_planner_preprocessrows(plan,clasfctn,period,weeks,False)
    budget_weekly_planner_preprocessrows(fact,clasfctn,period,weeks, True)


    now=datetime.now()
    duedate=now
    if now>period.end:
        duedate=period.end
    if now<period.start:
        duedate=period.start

    #вывод массива данных
    overspendingReport={}
    lastrowi,outputedrecords=budget_weekly_planner_cat(table,clasfctn._root,7,  period,plan,weeks,budget,duedate,overspendingReport)

    plan_total=0
    fact_total=0
    prediction_total=0


    for w in weeks:
        table[3,w.coli]=u"План"
        table[3,w.coli+1]=w.plan_total, Style.Money

        table[4,w.coli]=u"Факт"
        table[4,w.coli+1]=w.fact_total, Style.Money

        plan_total+=w.plan_total
        fact_total+=w.fact_total
        prediction_total+=w.predict_total

        table.set_column_width(w.coli,   10)
        table.set_column_width(w.coli+1,  5)
        table.set_column_width(w.coli+2, 10)
        table.set_column_width(w.coli+3,  5)



    coli=0
    table[3,coli]=u"План"
    table[3,coli+1]=u"Факт"
    table[3,coli+2]=u"Предикт"
    table[4,coli]=plan_total, Style.Money
    table[4,coli+1]=fact_total, Style.Money
    table[4,coli+2]=prediction_total, Style.Money
    over=prediction_total/plan_total
    if over>1:
        table[5,coli+2]=over,"percent_red"
    else:
        table[5,coli+2]=over, "percent_green"

    table.set_column_width(0, 6)
    table.set_column_width(1, 6)
    table.set_column_width(2, 6)

    todorow=lastrowi+0-5
    table[todorow,0]=u"Просроченные бюджетные цели","weekcaptions"
    bt_row=1


    sum=0
    for budget_item in budget.get_buying_targets():

        is_overdue, is_executed, is_todo=budget.check_item_execution(budget_item,duedate)
        if is_overdue:
            table[todorow+bt_row,1]="", Style.Month
            descr=budget_item.description
            if hasattr(budget_item,"_description"):
                descr= budget_item._description
            table[todorow+bt_row,3]=u"{0}".format(descr)
            table[todorow+bt_row,6]=budget_item.debit, Style.Money
            sum+=budget_item.debit
            bt_row+=1
    table[todorow,6]=sum, Style.Money
    lastrowi=todorow

    bt_row=1
    sum=0
    table[lastrowi,7]=u"Бюджетные цели в ближайшие 30 дней","weekcaptions"
    for budget_item in budget.get_buying_targets():
        is_overdue, is_executed, is_todo=budget.check_item_execution(budget_item,duedate)

        if is_todo:
            if budget_item.exactdate<period.end:
                table[lastrowi+bt_row,8]=budget_item.exactdate, Style.Month
                table[lastrowi+bt_row,9]=u"{0}".format(budget_item.description)
                table[lastrowi+bt_row,12]=budget_item.debit, Style.Money
                sum+=budget_item.debit
                bt_row+=1
        #table[lastrowi,7]=u"Бюджетные цели ({0})".format(sum)
    table[lastrowi,12]=sum, Style.Money

    bt_row=lastrowi+3
    table[bt_row,0]=u"Отчет по перерасходу","weekcaptions"
    overspending_sum=0
    #overspendingReport=OrderedDict(overspendingReport)
    overspendingReport=OrderedDict(sorted(overspendingReport.items(), key=lambda t: t[1], reverse=True))
    for k, v in overspendingReport.items():
        if v>0:
            bt_row+=1
            table[bt_row,3]=k
            table[bt_row,6]=v, Style.Money
            overspending_sum+=v

    table[bt_row+2,3]=u"Итого:"
    table[bt_row+2,6]=overspending_sum, Style.Money

    return table


def budget_weekly_planner_preprocessrows(plan,clasfctn,period,weeks, isfact):
    for row in plan.Rows:
        if row.type!=RowType.Tx:
            continue
        dt=row.get_human_or_logical_date()

        if dt<period.start and dt>period.end:
            continue
        if row.tx.direction==1:
            continue
        for w in weeks:
            if dt>=w.startday and dt<=w.lastday:
                g=clasfctn.match_tags_to_category(row.normilized_tags)
                if not hasattr(g, 'txs'):
                    g.txs=[]
                g.txs.append(row)
                row.weekindex=w.windex
                row.palanner_isfact=isfact

def budget_weekly_planner_cat(table,category, rowi, period,plan,weeks,budget_def, duedate,overspendingReport):


    isshow=check_classification(category, "family_out")
    isfinhelp=check_classification(category, "fin_help")
    if isfinhelp:
        isshow=False

    isuncategorized=check_classification(category, "_uncategorized")
    if isuncategorized:
        isshow=True
    isauto=check_classification(category, "_auto")
    if isauto:
        isshow=True



    startrowi=rowi
    cattitle=category.title

    if isshow:
        table[startrowi-1, 0]= cattitle, "weekcaptions"
    rowi+=1


    outputedrecords=0

    trowis_plan=[]
    for w in weeks:
        trowis_plan.append(0)
        w.cat_plan_total=0
        w.cat_fact_total=0


    maxtrow=rowi
    cat_plan_total=0
    cat_fact_total=0
    cat_predict_total=0

    now=datetime.now()
    #все тразакции данной категории
    if hasattr(category, 'txs') and isshow:
        for row in category.txs:
            rowdt=row.get_human_or_logical_date()
            #мы заранее рассчитали, к какой неделе относится строка
            week=weeks[row.weekindex]
            row_isfact=row.palanner_isfact
            #находим индекс строки в таблицы, для транзакции
            trowi=  trowis_plan[row.weekindex]
            trowis_plan[row.weekindex]=trowi+1


            coli=week.coli
            amount=row.amount.as_float()

            descr=row.description
            style="item_plan"
            if row_isfact:
                week.fact_total+=amount
                week.cat_fact_total+=amount

                #if week.use_fact:
                if rowdt<=now:
                    cat_predict_total+=amount
                    week.predict_total+=amount
                cat_fact_total+=amount
                style="item_fact"
            else:
                week.plan_total+=amount
                week.cat_plan_total+=amount
                cat_plan_total+=amount
                budget=row.tx.source_budget

                is_overdue, is_executed, is_todo=budget_def.check_item_execution(budget,duedate)

                if is_executed:
                    descr="[+]"+descr

                if is_overdue:
                    style="item_plan_overdue"
                    descr="[!]"+descr

                #if (not week.use_fact) or (is_overdue):
                if rowdt>now or is_overdue:
                    if not is_executed:
                        cat_predict_total+=amount
                        week.predict_total+=amount

            table[rowi+trowi,coli]=descr, style
            table[rowi+trowi,coli+1]=amount,style

            if rowi+trowi>maxtrow:
                maxtrow=rowi+trowi
            outputedrecords+=1

    rowi=maxtrow+1

    if outputedrecords==0:
        table[rowi-3, 0]= ""
        rowi=startrowi
    else:
        table[startrowi, 0]= cat_plan_total,"categoryline_totals"
        table[startrowi, 1]= cat_fact_total,"categoryline_totals"
        table[startrowi, 2]= cat_predict_total,"categoryline_totals"
        overspendingReport[cattitle]=cat_fact_total-cat_plan_total
        for w in weeks:
            table[startrowi,w.coli]="","categoryline_totals"
            table[startrowi,w.coli+1]=w.cat_fact_total,"categoryline_totals"

    for c in category.childs:
        rowi_temp=rowi
        rowi, child_outputedrecords=budget_weekly_planner_cat(table,c, rowi,period,plan,weeks, budget_def,duedate,overspendingReport)
        outputedrecords+=child_outputedrecords
        if outputedrecords>0:
            rowi+=1
        else:
            rowi=rowi_temp
        #table[rowi, 0]= cattitle+" [end]"
    return rowi,outputedrecords

def check_classification(group, sid):
    if group._sid==sid:
        return True

    #for c in group.childs:
    if group.parent:
        if check_classification(group.parent, sid):
            return True
    return False