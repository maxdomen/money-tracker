#! coding: utf-8
import cProfile
from common import CalendarHelper
from common.Classification import Classification, ClassificationDataset, ClassificationPrinter, Period
from common.Table import Table, Style, Color,DestinationXls
import common.CalendarHelper
from model import debt
from model import budget
from reports import bigpicture
from reports import weeklyplanner



import readers.StatementReader
from model.dashboard import DashboardDataset, DashboardPublisher
import model.debt
#from model.aggregatereport import Layout, Dataset, Publisher, Aggregate, Publisher2
#from debt import Debts
from reports.statement_monthly import classify_statement_monthly
from reports.statement_with_details import classify_statement_with_details

__author__ = 'Max'

from readers.avangard import AvangardReader
from readers.bankofamerica import BankOfAmericaReader
from readers.chase import ChaseBankReader
from readers.tcs import TCSBankReader

from model.accounts import *
from model.tags import AutoTagger
from model.printstatement import PrintStatementToExcel2
from readers.StatementReader import *
import xlwt
import common.CalendarHelper
from  common.Classification import *

#UT
def loadrates():
    Currency.addrate(datetime(2012, 03, 30),rub, usd,29.08)
    Currency.addrate(datetime(2012, 03, 30),rub, eur,38.76)
    Currency.addrate(datetime(2012, 03, 30),rub, gbr,46.31)
    Currency.addrate(datetime(2012, 03, 30),usd, eur,1.3)
    Currency.addrate(datetime(2012, 01, 16),usd, gbr,1.5335829333)

    Currency.addrate(datetime(2012, 01, 1),rub, usd,32.0000000000)
    Currency.addrate(datetime(2012, 01, 10),rub, usd,31.5626172211)
    Currency.addrate(datetime(2012, 01, 20),rub, usd,31.3238116083)
    Currency.addrate(datetime(2012, 01, 31),rub, usd,30.3403977095)


    Currency.addrate(datetime(2012, 02, 1),rub, usd,30.1341967113)
    Currency.addrate(datetime(2012, 02, 10),rub, usd,30.0485249718)
    Currency.addrate(datetime(2012, 02, 20),rub, usd,29.7705422178)
    Currency.addrate(datetime(2012, 02, 29),rub, usd,28.9700090541)


    Currency.addrate(datetime(2011, 12, 1),rub, usd,30.7272903392)
    Currency.addrate(datetime(2011, 12, 10),rub, usd,31.4625000000)
    Currency.addrate(datetime(2011, 12, 20),rub, usd,31.9865384130)
    Currency.addrate(datetime(2011, 12, 29),rub, usd,32.1417452109)


    Currency.addrate(datetime(2012, 2, 1),rub, cad,30.2268555650)

    Currency.addrate(datetime(2012, 4, 3),rub, usd,29.3479)
    Currency.addrate(datetime(2012, 5, 3),rub, usd,29.3708)
    Currency.addrate(datetime(2012, 5, 12),rub, usd,30.2306)
    Currency.addrate(datetime(2012, 5, 18),rub, usd,30.9417)
    Currency.addrate(datetime(2012, 5, 25),rub, usd,31.6247)


    Currency.addrate(datetime(2012, 5, 26),rub, usd,31.7572)
    Currency.addrate(datetime(2012, 6, 1),rub, usd,32.9173)
    Currency.addrate(datetime(2012, 6, 25),rub, usd,33.250)

    Currency.addrate(datetime(2012, 10, 1),rub, usd,31.0)

    Currency.addrate(datetime(2012, 10, 10),rub, cny,4.9374)

    Currency.addrate(datetime(2012, 10, 10),usd, cny,0.1596)

    Currency.addrate(datetime(2013, 5, 2),rub, usd,31.04)

def svetaaccounting(basedir,acc):
    #tcs = Account('scash',rub)
    #avr = Account('balt',rub)
    cashconfig={'first_row':1,'col_acc':1,'col_date':0, 'col_op':4,'col_in':2,'col_out':3,'col_balance':-1, 'col_tag1':5,'col_tag2':6}
    XlsReader(basedir+'home/2013/2013 sveta.xlsx','Records',cashconfig).parse_to([acc])


def parsing(basedir,avr,avu,avs,tcs,boa,wallet,safe, sveta, budget):
    print "Load sources"
    #BankOfAmericaReader(basedir+"home/2012/boa 2012.csv").parse_to(boa)

    #AvangardReader(basedir+"home/2012/avr 1.1.2012 - 1.4.2012.xls").parse_private_to(avr)
    #AvangardReader(basedir+"home/2012/avr apr 2012.xls").parse_private_to(avr)
    #AvangardReader(basedir+"home/2012/avr may 2012.xls").parse_private_to(avr)
    #AvangardReader(basedir+"home/2012/avr june 2012.xls").parse_private_withreserved_to(avr)
    #AvangardReader(basedir+"home/2012/avr july 2012.xls").parse_private_withreserved_to(avr)
    #AvangardReader(basedir+"home/2012/avr aug 2012.xls").parse_private_withreserved_to(avr)
    #AvangardReader(basedir+"home/2012/avr sep 2012 only.xls").parse_private_withreserved_to(avr)


    AvangardReader(basedir+"home/2013/avr jan-mar 2013.xls").parse_private_withreserved_to(avr)
    AvangardReader(basedir+"home/2013/avr apr 2013.xls").parse_private_withreserved_to(avr)


    #AvangardReader(basedir+"home/2012/avu 1.1.2012 - 1.4.2012.xls").parse_private_to(avu)
    #AvangardReader(basedir+"home/2012/avu apr 2012.xls").parse_private_to(avu)
    #AvangardReader(basedir+"home/2012/avu may 2012.xls").parse_private_to(avu)
    #AvangardReader(basedir+"home/2012/avu june 2012.xls").parse_private_withreserved_to(avu)
    #AvangardReader(basedir+"home/2012/avu july 2012.xls").parse_private_withreserved_to(avu)
    #AvangardReader(basedir+"home/2012/avu aug 2012.xls").parse_private_withreserved_to(avu)
    #AvangardReader(basedir+"home/2012/avu sep 2012 only.xls").parse_private_withreserved_to(avu)
    AvangardReader(basedir+"home/2013/avu jan-mar 2013.xls").parse_private_withreserved_to(avu)
    AvangardReader(basedir+"home/2013/avu apr 2013.xls").parse_private_withreserved_to(avu)



    #AvangardReader(basedir+"home/2012/avs sep 2012 only.xls").parse_private_withreserved_to(avs)
    AvangardReader(basedir+"home/2013/avs jan-mar 2013.xls").parse_private_withreserved_to(avs)
    AvangardReader(basedir+"home/2013/avs apr 2013.xls").parse_private_withreserved_to(avs)

    #TCSBankReader(basedir+"home/2012/tcs jan 2012new.csv").parse2012_to(tcs)
    #TCSBankReader(basedir+"home/2012/tcs feb-apr 2012new.csv").parse2012_to(tcs)
    #TCSBankReader(basedir+"home/2012/tcs may 2012.csv").parse2012_to(tcs)
    #TCSBankReader(basedir+"home/2012/tcs june 2012.csv").parse2012_to(tcs)
    #TCSBankReader(basedir+"home/2012/tcs july 2012.csv").parse2012_to(tcs)
    #TCSBankReader(basedir+"home/2012/tcs aug 2012.csv").parse2012_to(tcs)
    #TCSBankReader(basedir+"home/2012/tcs sep 2012 only.csv").parse2012b_to(tcs)

    TCSBankReader(basedir+"home/2013/tcs jan-mar 2013.csv").parse2012b_to(tcs)
    TCSBankReader(basedir+"home/2013/tcs apr 2013.csv").parse2012b_to(tcs)


    cashconfig={'first_row':1,'col_acc':1,'col_date':0, 'col_op':2,'col_in':3,'col_out':4,'col_balance':5, 'col_tag1':6,'col_tag2':7}
    XlsReader(basedir+'home/2013/2013 logs and cash.xls','Cash ops',cashconfig).parse_to([wallet,safe, sveta])



    cashconfig={'first_row':1,'col_date':0}
    accstoread={avr:1, avu:2, tcs:3, wallet:4, safe:5, sveta:6}
    #accstoread={avu:2}

    XlsLeftoversJournalReader(basedir+'home/2013/2013 logs and cash.xls','Account Log',cashconfig).parse_to(accstoread)

    budget.read(basedir+'home/2013/2013 logs and cash.xls','Budget')
    budget.read(basedir+'home/2013/2013 sveta.xlsx','Plan')

    budget.read_executions(basedir+'home/2013/2013 logs and cash.xls','Budget-Execution')
    budget.read_executions(basedir+'home/2013/2013 sveta.xlsx','Budget-Execution')
#FP=None
def tagging(basedir,familypool=None):
    print "Tagging"
    tagger=AutoTagger()

    def max_card_trans(tx):

        res=tx.props.get('card_id')
        if res!=None:
            if res.find('6159')>=0:
                return "max"
        return None

    def sveta_card_trans(tx):
       # props={}
        res=tx.props.get('card_id')
        if res!=None:
            if res.find('7568')>=0:
                return "sveta"
        return None

    def tx_to_outer_bank(tx):
        if tx.direction==-1:
            if tx.comment.find(u"Пополнение вклада")>=0:
                return "2outer"
            if tx.comment.find(u"Для зачисления на лицевой счет")>=0:
                return "2outer"
        return None

    tagger.handler(sveta_card_trans)
    tagger.handler(tx_to_outer_bank)


    tagger.load_declares(basedir+"home/2013/2013 logs and cash.xls","Auto Tags")
    tagger.load_manual_tags(basedir+"home/2013/2013 logs and cash.xls","Manual Tags")



    tagger.dotagforpool(familypool)


    TransisitionsLoader(familypool, basedir+"home/2013/2013 logs and cash.xls","Transitions")
    load_slices_and_logicaldates(familypool,basedir+"home/2013/2013 logs and cash.xls","Slices")

def load_slices_and_logicaldates(pool,filename,sheetname):

    book = xlrd.open_workbook(filename)
    sheet=book.sheet_by_name(sheetname)
    slicescount=0
    logicaldatecount=0
    for rowi in range(1,sheet.nrows):
        r=sheet.row(rowi)
        txid=r[1].value
        logical_date=r[6].value
        if isinstance(logical_date, float):
            tdate=xlrd.xldate_as_tuple(logical_date,0)
            res=datetime(tdate[0],tdate[1],tdate[2])
            txobj=pool.get_tx_byid(txid)
            if not txobj:
                print "Tx '{0}' not found for set_logical_date '{1}'".format(txid,res)
            else:
                txobj.set_logical_date(res)
            logicaldatecount+=1

        amount=r[3].value
        if isinstance(amount, float):
            comment=r[2].value
            tags_add=[]
            tags_remove=[]
            TagTools.ConvertStringOfTagsToList(tags_add,r[4].value)
            TagTools.ConvertStringOfTagsToList(tags_remove,r[5].value)
            tx=pool.get_tx_byid(txid)
            if tx:
                tx.slice(comment,amount,tags_add,tags_remove)
            slicescount+=1

    print "{0} slices added, {1} logical dates".format(slicescount,logicaldatecount)
def homeaccounting(basedir):
    loadrates()



    #account wo any statement


    pool = Pool()


    #actual
    tcs = Account('tcs',rub)
    tcs.limit=-75000
    avr = Account('avr',rub)

    avs = Account('avs',rub)

    avu = Account('avu',usd)
    avu.limit=-2000
    boa = Account('boa',usd)
    wallet = Account('wallet',rub)
    safe = Account('safe',rub)
    sveta = Account('sveta',rub)
    svetaaccounting("../money-data/",sveta)

    budgetf= budget.Budget()

    parsing(basedir,avr,avu,avs,tcs,boa,wallet,safe, sveta,budgetf)

    familypool = Pool()
    familypool.link_account(tcs)
    familypool.link_account(avr)
    familypool.link_account(avu)
    familypool.link_account(avs)
    familypool.link_account(boa)
    familypool.link_account(wallet)
    familypool.link_account(safe)
    familypool.link_account(sveta)

    FP=familypool
    tagging(basedir,familypool)
    #cProfile.run('tagging()')
    #cProfile.runctx('tagging(familypool)', globals(),locals())


    print "Generate statement"

    statement=familypool.make_statement(rub)


    #cProfile.runctx('statement=familypool.make_statement(rub)', globals(),locals())





    budgetstatement=budgetf.make_statement(datetime(datetime.now().year-1,1,1),rub,forNyears=2)


    #cProfile.runctx('budgetstatement=budget.make_statement(rub)', globals(),locals())


    dashboarddataset=DashboardDataset(statement,budgetstatement)


    debts=debt.Debts(statement,start=datetime(2012,1,1))




    clasfctn=load_and_organize_classfication(basedir,statement, False)





    #записываем в statement присловоенную категорию, чтобы показать в отчете
    for r in statement.Rows:
        r.classification=None
        if r.type==RowType.Tx:
            ts=list(r.tags)
            if r.tx.direction==1:
                ts.append("__in")
            group=clasfctn.match_tags_to_category(ts)
            #if group==clasfctn._uncategorized:
            #    print "clasfctn"
            r.classification=group#.title

    print "Print statement Txs"
    dashboardpublisher=DashboardPublisher(dashboarddataset,"test.xls","Dashboard")
    excel=PrintStatementToExcel2("test.xls","Txs",existing_workbook=dashboardpublisher.wb)
    excel.set_period(datetime(2012,1,1),datetime.now())
    excel.do_print(statement)

    wb=excel.wb

    #return excel.wb


    bigpict_cahsflow_checkpoints=[]
    #bigpict_cahsflow_checkpoints.append( (datetime(2012,11,30), -(222425.35+39002)) )
    #bigpict_cahsflow_checkpoints.append( (datetime(2012,11,30), 0) )

    bigpict_cahsflow_checkpoints.append( (datetime(2013,2,1), 0) )


    bigpict_period=CalendarHelper.Period(datetime(2012,11,1),datetime(2014,1,1))
    bigpicttable=bigpicture.new_big_picture(clasfctn,statement,budgetstatement, budgetf,bigpict_cahsflow_checkpoints, bigpict_period)


    classify_statement_monthly(clasfctn,statement,wb, "Monthly")
    relationshipwithcompany(statement,wb,debts)

    debts_due_to_date=datetime.now()+timedelta(days=31)
    debts.xsl_to(bigpicttable,bigpict_period,debts_due_to_date)



    DestinationXls(bigpicttable,wb)

    #chelp=common.CalendarHelper.CalendarHelper()





    classify_statement_with_details(clasfctn,statement,wb, "Details_Prev",True, common.CalendarHelper.month_prev())
    classify_statement_with_details(clasfctn,statement,wb, "Details_Cur",True, common.CalendarHelper.month_current())




    clasfctn=load_and_organize_classfication(basedir,statement, True)
    classify_statement_monthly(clasfctn,budgetstatement,wb, "BudgetMonthly")

    wb2 = xlwt.Workbook()


    table=weeklyplanner.budget_weekly_planner("Weekly_Cur",common.CalendarHelper.month_current(),budgetstatement,clasfctn,statement,budgetf)
    DestinationXls(table,wb,def_font_height=6)
    DestinationXls(table,wb2,def_font_height=6)

    table=weeklyplanner.budget_weekly_planner("Weekly_Prev",common.CalendarHelper.month_prev(),budgetstatement,clasfctn,statement,budgetf)
    DestinationXls(table,wb,def_font_height=6)
    DestinationXls(table,wb2,def_font_height=6)

    table=weeklyplanner.budget_weekly_planner("Weekly_Next",common.CalendarHelper.month_next(),budgetstatement,clasfctn,statement,budgetf)
    DestinationXls(table,wb,def_font_height=6)
    DestinationXls(table,wb2,def_font_height=6)

    wb.save("test.xls")


    DestinationXls(bigpicttable,wb2)
    wb2.save("familyreport.xls")



def relationshipwithcompany(statement,wb,debts):



    checkpoints=[]


    checkpoints.append([datetime(2012,11,24,17),117097, False])

    checkpoints.append([datetime(2013,04,10,16),-35427, False])

    table=Table("CM and Max")

    table.define_style("redmoney", foreground_color=Color.Red, formatting_style=Style.Money)

    table[0,0]=u"Отношения с компанией"
    rowi=0
    rbase=3
    mydebt=0

    table[rbase-1,1]=u"Мой долг компании"
    table[rbase-1,8]=u"Дала мне компания"
    table[rbase-1,7]=u"Я потратил на нужды компании или отдал долг"
    for row in statement.Rows:
        if row.type!=RowType.Tx:
            continue

        relation=False
        #print row.classification.title
        relation=weeklyplanner.check_classification(row.classification,"company_txs_in")
        if not relation:
            relation=weeklyplanner.check_classification(row.classification,"company_txs")
        if not relation:
            continue



        for cp in checkpoints:
            if row.date>=cp[0] and cp[2]==False:

                rowi+=1
                mydebt=cp[1]
                print_checkpoint(table,rbase+rowi,cp)
                debts.define_debt_balance("CM",cp[0],cp[1])
                rowi+=1
                break

        table[rbase+rowi,0]=row.date, Style.Day
        table[rbase+rowi,2]=row.classification.title
        table[rbase+rowi,4]=row.description
        v=row.amount.as_float()
        if row.tx.direction==1:
            mydebt+=v
            coli=8
        else:
            mydebt-=v
            coli=7
        table[rbase+rowi,1]=mydebt, Style.Money
        table[rbase+rowi,coli]=v, Style.Money

        table[rbase+rowi,15]=TagTools.TagsToStr(row.tags)

        debts.define_debt_balance("CM",row.date,mydebt)

        rowi+=1

    for cp in checkpoints:
             if cp[2]==False:
                 rowi+=1
                 mydebt=cp[1]
                 print_checkpoint(table,rbase+rowi,cp)
                 debts.define_debt_balance("CM",cp[0],cp[1])
                 rowi+=1
                 break

    DestinationXls(table,wb)
def print_checkpoint(table,rowi,cp):
    cp[2]=True
    table[rowi,0]=cp[0], Style.Day
    table[rowi,2]="Checkpoint"
    table[rowi,1]=cp[1], "redmoney"









def load_and_organize_classfication(basedir,statement, collapse_company_txs):
    classification=Classification(from_xls=(basedir+"home/2013/2013 logs and cash.xls","Classification"))


    #создаем категории для тегов, которые не попали созданные вручную категории
    classification.create_auto_classification(statement)
    if collapse_company_txs:
            classification.get_category_by_id("company_txs")._collapsed=True
            classification.finalize()

    #я хочу видеть неклассифицированные расходы в расходах семьи.
    family=classification.get_category_by_id("family_out")
    classification._uncategorized.moveto(family)
    classification._auto_categorized.moveto(family)


    classification.finalize()
    return classification



if __name__ == '__main__':
    homeaccounting("../money-data/")


#p = pstats.Stats('homeaccounting()')
#p.sort_stats('time')
