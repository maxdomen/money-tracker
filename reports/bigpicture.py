#! coding: utf-8
__author__ = 'Max'
from  common.Classification import *
from common.Table import Table, Style, Color,DestinationXls
from model.budget import BudgetFreq

def new_big_picture(clasfctn,statement,budgetstatement,budget, bigpict_cahsflow_checkpoints, period):
    monthlydataset=ClassificationDataset(clasfctn,Period.Month, statement)
    budgetmonthlydataset=ClassificationDataset(clasfctn,Period.Month, budgetstatement)

    table=Table(u"Годовой план")
    table.define_style("redmoney", foreground_color=Color.Red,formatting_style=  Style.Money)
    table.define_style("greenmoney", foreground_color=Color.Green,formatting_style=  Style.Money)
    table.define_style("blackmoney", foreground_color=Color.Black,formatting_style=  Style.Money)
    table.write_cells_vert(2,0,["family_in","family_out","EBIDTA","cashflow"])
    table[1,2]="test"
    coli=1
    cummulative=0
    dtnow=datetime.now()
    bt_row=0
    for p in monthlydataset.periods:


        if  not (p._end<dtnow):
            break



        if p._start<period.start or p._end>period.end:
            continue

        cummulative=big_pict_period(table,coli,p,clasfctn,monthlydataset,cummulative,bigpict_cahsflow_checkpoints)
        bt_row=show_buying_targets(bt_row,p,budget,coli,table, ispast=True)

        coli+=1

    for p in budgetmonthlydataset.periods:

        if  p._end<dtnow:
            continue
        else:
            table[0,coli]="Plan Y"+str(p._end.year-2000)



        if p._start<period.start or p._end>period.end:
            continue
        cummulative=big_pict_period(table,coli,p,clasfctn,budgetmonthlydataset,cummulative,bigpict_cahsflow_checkpoints)
        bt_row=show_buying_targets(bt_row,p,budget,coli,table, ispast=False)


        coli+=1

    return table


def big_pict_period(table,coli,p,clasfctn,monthlydataset,cummulative,bigpict_cahsflow_checkpoints):

    table[1,coli]=p._start,Style.Month

    category=clasfctn.get_category_by_id("family_in")
    income=monthlydataset.calcsubtotals(category,p)
    table[2,coli]=income

    category=clasfctn.get_category_by_id("family_out")
    losses=monthlydataset.calcsubtotals(category,p)

    category=clasfctn.get_category_by_id("fin_help")
    #fin_help=monthlydataset.calcsubtotals(category,p)
    fin_help=p._cells[category._index]
    losses=losses-fin_help

    #print "fin_help",fin_help,losses


    table[3,coli]=losses
    ebitda=income-losses


    if ebitda>0:
        table[4,coli]=ebitda, "greenmoney"
    else:
        table[4,coli]=ebitda, "redmoney"

    cummulative+=ebitda
    cummulative=cumulative_check_points(cummulative,bigpict_cahsflow_checkpoints,p)
    if cummulative>0:
        table[5,coli]=cummulative, "greenmoney"
    else:
        table[5,coli]=cummulative, "redmoney"

    return cummulative

def cumulative_check_points(cummulative,bigpict_cahsflow_checkpoints,p):

    for d,v in bigpict_cahsflow_checkpoints:
        if d>=p._start and d<=p._end:
            cummulative=v
            print "cummulative",cummulative,d,p._start
    return cummulative

def show_buying_targets(bt_row,p,budgetf,coli,table, ispast):
    sum=0.0
    base=22
    now=datetime.now()
    for budget_item in budgetf.get_buying_targets():
        if not budget_item.exactdate:
            continue
        if budget_item.exactdate>=p._start and budget_item.exactdate<=p._end:
            is_show=False
            is_overdue, is_executed, is_todo=budgetf.check_item_execution(budget_item,now)

            style="blackmoney"

            if is_overdue:
                is_show=True
            if budget_item.period== BudgetFreq.OneTime or budget_item.period== BudgetFreq.Annually:
                if (not is_executed):
                    is_show=True

            if is_show:
                descr=budget_item.description

                table[base+1+bt_row,coli]=u"{0}({1})".format(descr,budget_item.debit),style
                sum+=budget_item.debit
                bt_row+=1
    if bt_row>15:
        bt_row=0
    if sum>0:
        table[base,coli]=sum,Style.Money
    return bt_row
