#! coding: utf-8
import cProfile
from common.Classification import Classification, ClassificationDataset, ClassificationPrinter, Period
from common.Table import Table, Style, Color,DestinationXls
from model import debt
import copy

import readers.StatementReader
from model.dashboard import DashboardDataset, DashboardPublisher
import model.debt
#from model.aggregatereport import Layout, Dataset, Publisher, Aggregate, Publisher2
#from debt import Debts

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



def svetaaccounting(basedir,acc):
    #tcs = Account('scash',rub)
    #avr = Account('balt',rub)
    cashconfig={'first_row':1,'col_acc':1,'col_date':0, 'col_op':4,'col_in':2,'col_out':3,'col_balance':-1, 'col_tag1':5,'col_tag2':6}
    XlsReader(basedir+'home/2012/2012 sveta.xlsx','Records',cashconfig).parse_to([acc])


def parsing(basedir,avr,avu,avs,tcs,boa,wallet,safe, sveta, budget):
    print "Load sources"
    BankOfAmericaReader(basedir+"home/2012/boa 2012.csv").parse_to(boa)

    AvangardReader(basedir+"home/2012/avr 1.1.2012 - 1.4.2012.xls").parse_private_to(avr)
    AvangardReader(basedir+"home/2012/avr apr 2012.xls").parse_private_to(avr)
    AvangardReader(basedir+"home/2012/avr may 2012.xls").parse_private_to(avr)
    AvangardReader(basedir+"home/2012/avr june 2012.xls").parse_private_withreserved_to(avr)
    AvangardReader(basedir+"home/2012/avr july 2012.xls").parse_private_withreserved_to(avr)
    AvangardReader(basedir+"home/2012/avr aug 2012.xls").parse_private_withreserved_to(avr)
    AvangardReader(basedir+"home/2012/avr sep 2012.xls").parse_private_withreserved_to(avr)

    AvangardReader(basedir+"home/2012/avu 1.1.2012 - 1.4.2012.xls").parse_private_to(avu)
    AvangardReader(basedir+"home/2012/avu apr 2012.xls").parse_private_to(avu)
    AvangardReader(basedir+"home/2012/avu may 2012.xls").parse_private_to(avu)
    AvangardReader(basedir+"home/2012/avu june 2012.xls").parse_private_withreserved_to(avu)
    AvangardReader(basedir+"home/2012/avu july 2012.xls").parse_private_withreserved_to(avu)
    AvangardReader(basedir+"home/2012/avu aug 2012.xls").parse_private_withreserved_to(avu)
    AvangardReader(basedir+"home/2012/avu sep 2012.xls").parse_private_withreserved_to(avu)


    AvangardReader(basedir+"home/2012/avs sep 2012.xls").parse_private_withreserved_to(avs)


    TCSBankReader(basedir+"home/2012/tcs jan 2012new.csv").parse2012_to(tcs)
    TCSBankReader(basedir+"home/2012/tcs feb-apr 2012new.csv").parse2012_to(tcs)
    TCSBankReader(basedir+"home/2012/tcs may 2012.csv").parse2012_to(tcs)
    TCSBankReader(basedir+"home/2012/tcs june 2012.csv").parse2012_to(tcs)
    TCSBankReader(basedir+"home/2012/tcs july 2012.csv").parse2012_to(tcs)
    TCSBankReader(basedir+"home/2012/tcs aug 2012.csv").parse2012_to(tcs)
    TCSBankReader(basedir+"home/2012/tcs sep 2012.csv").parse2012_to(tcs)

    #TCSBankReader("Data/home/2012/tcs june 2012b.csv").parse2011v2_to(tcs)

    cashconfig={'first_row':1,'col_acc':1,'col_date':0, 'col_op':2,'col_in':3,'col_out':4,'col_balance':5, 'col_tag1':6,'col_tag2':7}
    XlsReader(basedir+'home/2012/2012 logs and cash.xls','Cash ops',cashconfig).parse_to([wallet,safe, sveta])



    cashconfig={'first_row':1,'col_date':0}
    accstoread={avr:1, avu:2, tcs:3, wallet:4, safe:5, sveta:6}
    #accstoread={avu:2}

    XlsLeftoversJournalReader(basedir+'home/2012/2012 logs and cash.xls','Account Log',cashconfig).parse_to(accstoread)

    budget.read(basedir+'home/2012/2012 logs and cash.xls','Budget')
    budget.read(basedir+'home/2012/2012 sveta.xlsx','Plan')

#FP=None
def tagging(basedir,familypool=None):
    print "Tagging"
    #if not familypool:
    #    familypool=FP

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


    tagger.load_declares(basedir+"home/2012/2012 logs and cash.xls","Auto Tags")
    tagger.load_manual_tags(basedir+"home/2012/2012 logs and cash.xls","Manual Tags")



    tagger.dotagforpool(familypool)


    #cProfile.runctx('tagger.dotagforpool(familypool)', globals(),locals())


    TransisitionsLoader(familypool, basedir+"home/2012/2012 logs and cash.xls","Transitions")



    #feb
    familypool.get_tx_byid("12351837avr13682.00").set_logical_date(datetime(2012,2,28))
    familypool.get_tx_byid("123600wallet16318.00").set_logical_date(datetime(2012,2,29))
    #mar


    familypool.get_tx_byid("123201413avr22000.00").slice(u"В транзакции по зарплате только 20К",2000,["Reimbursment"],[u"Зарплата"])
    familypool.get_tx_byid("124500wallet31000.00").set_logical_date(datetime(2012,3,30))
    familypool.get_tx_byid("124500wallet15018.00").set_logical_date(datetime(2012,3,30))


    familypool.get_tx_byid("123271359avr20000.00").slice(u"В транзакции по зарплате 0К, остальное на Москву",20000,["Reimbursment"],[u"Зарплата"])


    #apr
    familypool.get_tx_byid("1252113avr23682.00").set_logical_date(datetime(2012,4,30))
    #may
    familypool.get_tx_byid("126900wallet18318.00[1]").set_logical_date(datetime(2012,5,30))
    familypool.get_tx_byid("126900avr41864.20").set_logical_date(datetime(2012,5,30))

    tx=familypool.get_tx_byid("127600avr41682.00").set_logical_date(datetime(2012,6,30))

    tx=familypool.get_tx_byid("1271200wallet18318.00[1]").set_logical_date(datetime(2012,6,30))

    #nov
    familypool.get_tx_byid("1211700tcs5522.00").slice(u"Вино 6 бутылок",1834,[u"спиртное"],["food"])
    familypool.get_tx_byid("1211700tcs5522.00").slice(u"Гель для душа(x2) и шампунь(x2)",562,[u"хоз"],["food"])
    familypool.get_tx_byid("1211700tcs5522.00").slice(u"Бритвы",345,[u"хоз"],["food"])
    tx=familypool.get_tx_byid("12102500sveta3700.00").set_logical_date(datetime(2012,11,1))

    familypool.get_tx_byid("1211200avr25375.00").set_logical_date(datetime(2012,10,30))




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

    budget= debt.Budget()

    parsing(basedir,avr,avu,avs,tcs,boa,wallet,safe, sveta,budget)

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




    budgetstatement=budget.make_statement(rub,forNyears=2)


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



    bigpicttable=new_big_picture(wb,clasfctn,statement,budgetstatement)


    classify_statement(clasfctn,statement,wb, "Monthly")
    relationshipwithcompany(statement,wb,debts)

    debts.xsl_to(bigpicttable)
    DestinationXls(bigpicttable,wb)


    m_t=datetime.now()
    m_cur_d_start=datetime(m_t.year,m_t.month,1,0,0,0)
    m_i=m_cur_d_start+timedelta(days=32)
    m_cur_d_finish=datetime(m_i.year,m_i.month,1,0,0,0)-timedelta(seconds=1)
    m_t=m_cur_d_start-timedelta(seconds=1)
    m_prev_d_finish=m_t
    m_prev_d_start=datetime(m_t.year,m_t.month,1)




    classify_statement_with_details(clasfctn,statement,wb, "Details_Prev",True, m_prev_d_start,m_prev_d_finish)
    classify_statement_with_details(clasfctn,statement,wb, "Details_Cur",True, m_cur_d_start,m_cur_d_finish)




    clasfctn=load_and_organize_classfication(basedir,statement, True)
    classify_statement(clasfctn,budgetstatement,wb, "BudgetMonthly")

    #budget_weekly_planner(wb,m_cur_d_start,m_cur_d_finish,budgetstatement,clasfctn,statement)
    budget_weekly_planner(wb,"Weekly_Prev",m_prev_d_start,m_prev_d_finish,budgetstatement,clasfctn,statement)
    budget_weekly_planner(wb,"Weekly_Cur",m_cur_d_start,m_cur_d_finish,budgetstatement,clasfctn,statement)


    wb.save("test.xls")
def relationshipwithcompany(statement,wb,debts):



    checkpoints=[]

    checkpoints.append([datetime(2012,2,22),-30454.64,False])
    checkpoints.append([datetime(2012,3,21),5407.42,False])
    checkpoints.append([datetime(2012,5,5),228423.96,False])
    checkpoints.append([datetime(2012,5,5),228423.96,False])
    checkpoints.append([datetime(2012,6,8),85318.81,False])
    checkpoints.append([datetime(2012,7,9, 17,0,0),67205.08,False])
    checkpoints.append([datetime(2012,7,21, 16,0,0),233207.07,False])
    checkpoints.append([datetime(2012,10,1,17),222878, False])
    checkpoints.append([datetime(2012,11,6,16),114975, False])



    table=Table("CM and Max")
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
        relation=check_classification(row.classification,"company_txs_in")
        if not relation:
            relation=check_classification(row.classification,"company_txs")
        if not relation:
            continue



        for cp in checkpoints:
            if row.date>=cp[0] and cp[2]==False:

                rowi+=1
                mydebt=cp[1]
                print_checkpoint(table,rbase+rowi,cp)
                debts.define_debt("CM",cp[0],cp[1])
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

        debts.define_debt("CM",row.date,mydebt)

        rowi+=1

    for cp in checkpoints:
             if cp[2]==False:
                 rowi+=1
                 mydebt=cp[1]
                 print_checkpoint(table,rbase+rowi,cp)
                 debts.define_debt("CM",cp[0],cp[1])
                 rowi+=1
                 break

    DestinationXls(table,wb)
def print_checkpoint(table,rowi,cp):
    cp[2]=True
    table[rowi,0]=cp[0], Style.Day
    table[rowi,2]="Checkpoint"
    table[rowi,1]=cp[1], Style.Money+Style.Red
def check_classification(group, sid):
    if group._sid==sid:
        return True

    #for c in group.childs:
    if group.parent:
        if check_classification(group.parent, sid):
            return True
    return False

def big_pict_period(table,coli,p,clasfctn,monthlydataset,cummulative):

    table[1,coli]=p._start,Style.Month

    category=clasfctn.get_category_by_id("family_in")
    income=monthlydataset.calcsubtotals(category,p)
    table[2,coli]=income
    category=clasfctn.get_category_by_id("family_out")
    losses=monthlydataset.calcsubtotals(category,p)
    table[3,coli]=losses
    ebitda=income-losses


    if ebitda>0:
         table[4,coli]=ebitda, Style.Money+Style.Green
    else:
         table[4,coli]=ebitda, Style.Money+Style.Red

    cummulative+=ebitda
    if cummulative>0:
        table[5,coli]=cummulative, Style.Money+Style.Green
    else:
        table[5,coli]=cummulative, Style.Money+Style.Red

    return cummulative

def new_big_picture(wb,clasfctn,statement,budgetstatement):
    monthlydataset=ClassificationDataset(clasfctn,Period.Month, statement)
    budgetmonthlydataset=ClassificationDataset(clasfctn,Period.Month, budgetstatement)

    table=Table("Big Picture")
    table.write_cells_vert(2,0,["family_in","family_out","EBIDTA","cashflow"])
    table[1,2]="test"
    coli=1
    cummulative=0
    dtnow=datetime.now()
    for p in monthlydataset.periods:

        if  not (p._end<dtnow):
            break
        cummulative=big_pict_period(table,coli,p,clasfctn,monthlydataset,cummulative)


        coli+=1

    for p in budgetmonthlydataset.periods:
        if  p._end<dtnow:
            continue
        else:
            table[0,coli]="Plan Y"+str(p._end.year-2000)

        cummulative=big_pict_period(table,coli,p,clasfctn,budgetmonthlydataset,cummulative)

        coli+=1

    return table


def classify_statement_with_details(clasfctn,statement,wb, sheetname2, collapse_company_txs, date_start, date_finish):
    for r in statement.Rows:
        if r.type!=RowType.Tx:
            continue


        dt=r.get_logical_date()

        if not (dt>=date_start and dt<=date_finish):
            continue

        g=clasfctn.match_tags_to_category(r.normilized_tags)

        if not hasattr(g, 'txs'):
            g.txs=[]
        g.txs.append(r)


    ws = wb.add_sheet(sheetname2)
    ws.normal_magn=70
    ws.col(1).width=256*12
    ws.col(2).width=256*12
    ws.col(7).width=256*12

    #rowi=0


    details_for_cat(ws,clasfctn._root,0, date_start, date_finish)


def build_category_path(category):
    cattitle=category.title
    p=category.parent
    while p:
        if p.title=="_root":
            break
        cattitle=p.title+"/"+cattitle
        p=p.parent
    return cattitle
def details_for_cat(ws,category, rowi, date_start, date_finish):

    bc=0
    style_time1 = xlwt.easyxf(num_format_str='D-MMM')
    style_money=xlwt.easyxf(num_format_str='#,##0.00')
    cattitle=build_category_path(category)
    ws.write(rowi, 0, cattitle)
    rowi+=1

    subtotal=0
    if hasattr(category, 'txs'):
        for r in category.txs:
            dt=r.get_logical_date()
            if not (dt>=date_start and dt<=date_finish):
                continue

            #print "   ",r.date,r.tx.direction, r.amount, r.description,"->", category.title
            ws.write(rowi, bc+0, dt,style_time1)
            if r.tx.direction==1:
                acoli=bc+2
            else:
                acoli=bc+1
            v=r.amount.as_float()
            subtotal+=v*r.tx.direction
            ws.write(rowi, acoli,v ,style_money)
            ws.write(rowi, bc+3, r.description)
            satags=TagTools.TagsToStr(r.tx._tags)
            if len(satags)<1:
                satags="<notags>"
            ws.write(rowi, bc+7,satags )

            rowi+=1

    for c in category.childs:
        rowi, childtotal=details_for_cat(ws,c, rowi,date_start, date_finish)
        rowi+=1
        subtotal+=childtotal

    ws.write(rowi, 3, u"Total in {0}".format(category.title))
    ws.write(rowi, 7,subtotal ,style_money)

    rowi+=1
    return rowi,subtotal
def load_and_organize_classfication(basedir,statement, collapse_company_txs):
    classification=Classification(from_xls=(basedir+"home/2012/2012 logs and cash.xls","Classification"))


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
def classify_statement(classification,statement,wb, sheetname):



    monthlydataset=ClassificationDataset(classification,Period.Month, statement)

    ws = wb.add_sheet(sheetname)
    ws.col(0).width=256*40
    ws.panes_frozen = True
    ws.horz_split_pos = 2
    ws.vert_split_pos = 1
    ws.normal_magn=70


    printer=ClassificationPrinter(monthlydataset, existing_sheet=ws)
    return classification

class WeekDef:
    pass
def budget_weekly_planner(wb, caption,d_start, d_finish, plan, clasfctn2, fact):
    table=Table(caption)
    table.normal_magn=80
    table.define_style("totals", foreground_color=Color.Red)
    table.define_style("weekcaptions", bold=True, font_size=10)
    table.define_style("categoryline", bold=True, background_color=Color.LightGreen)
    table.define_style("categoryline_totals", italic=True,background_color=Color.LightGreen, formatting_style=Style.Money)
    table.define_style("item_plan",background_color=Color.LightGray, formatting_style=Style.Money)
    table.define_style("accum",bold=True, formatting_style=Style.Money)
    table.define_style("item_fact", formatting_style=Style.Money)

    clasfctn=copy.deepcopy(clasfctn2)

    weeks=[]
    w=WeekDef()
    weeks.append(w)
    t=d_start
    w.startday=t
    w.windex=0
    w.plan_total=0
    w.fact_total=0
    while t<d_finish:
        w.lastday=t
        if t.weekday()==6:
            i=w.windex
            w=WeekDef()
            w.startday=t+timedelta(days=1)
            w.windex=i+1
            w.plan_total=0
            w.fact_total=0
            weeks.append(w)
        t+=timedelta(days=1)

    rowi=1
    coli=4
    #windex=1
    for w in weeks:
        print w.startday,w.lastday
        w.coli=coli
        table[rowi,coli]="Week {0}".format(w.windex+1), "weekcaptions"
        #windex+=1
        table[rowi,coli+1]="{0}-{1}".format(w.startday.day,w.lastday.day), "weekcaptions"
        coli+=4



                #print w.windex, g.title,row.description, row.amount

    budget_weekly_planner_preprocessrows(plan,clasfctn,d_start,d_finish,weeks,False)
    budget_weekly_planner_preprocessrows(fact,clasfctn,d_start,d_finish,weeks, True)

    lastrowi,outputedrecords=budget_weekly_planner_cat(table,clasfctn._root,6,  d_start, d_finish,plan,weeks)
    plan_total=0
    fact_total=0
    prediction_total=0

    table[lastrowi+2,0]=u"План, накапливающийся итог"
    table[lastrowi+3,0]=u"Факт(предсказание), накапливающийся итог"

    for w in weeks:
        table[3,w.coli+1]=u"План"
        table[4,w.coli+1]=w.plan_total, Style.Money
        plan_total+=w.plan_total
        table[lastrowi+1,w.coli+1]="Week {0}".format(w.windex+1)
        table[lastrowi+2,w.coli+1]=plan_total, "accum"


        table.set_column_width(w.coli, 10)
        table.set_column_width(w.coli+1, 5)
        table[3,w.coli+3]=u"Факт"
        table[4,w.coli+3]=w.fact_total, Style.Money
        fact_total+=w.fact_total
        if w.fact_total==0:
            prediction_total+=w.plan_total
        else:
            prediction_total+=w.fact_total
        table[lastrowi+3,w.coli+1]=prediction_total, "accum"


        table.set_column_width(w.coli+2, 10)
        table.set_column_width(w.coli+3, 5)
        table.set_column_width(w.coli+4, 1)

    mtitle="{0} {1}".format(d_start.strftime("%B"),d_start.year)
    table[0,0]=mtitle
    table[0,weeks[3].coli]=mtitle
    table[3,1]=u"План"
    table[3,2]=u"Факт"
    table[4,1]=plan_total, Style.Money
    table[4,2]=fact_total, Style.Money


    table.set_column_width(0, 8)
    table.set_column_width(1, 5)
    table.set_column_width(2, 5)
    table.set_column_width(3, 1)

    DestinationXls(table,wb,def_font_height=6)
def budget_weekly_planner_preprocessrows(plan,clasfctn,d_start,d_finish,weeks, isfact):
    for row in plan.Rows:
        if row.type!=RowType.Tx:
            continue

        if row.date<d_start and row.date>d_finish:
            continue
        if row.tx.direction==1:
            continue
        for w in weeks:
            #row.tx.
            if row.date>=w.startday and row.date<=w.lastday:
                g=clasfctn.match_tags_to_category(row.normilized_tags)
                if not hasattr(g, 'txs'):
                    g.txs=[]
                g.txs.append(row)
                row.weekindex=w.windex
                row.palanner_isfact=isfact

def budget_weekly_planner_cat(table,category, rowi, date_start, date_finish,plan,weeks):


    isfamily=check_classification(category, "family_out")
    #if not isfamily:
    #    return rowi,0

    startrowi=rowi
    cattitle=category.title

    if isfamily:
        table[startrowi, 0]= cattitle, "categoryline"
    rowi+=1


    outputedrecords=0


    #trowi=rowi
    trowis_plan=[]
    trowis_fact=[]
    for w in weeks:
        trowis_plan.append(0)
        trowis_fact.append(0)
        w.cat_plan_total=0
        w.cat_fact_total=0
    maxtrow=rowi
    cat_plan_total=0
    cat_fact_total=0

    if hasattr(category, 'txs') and isfamily:
        for row in category.txs:
            week=weeks[row.weekindex]
            trowis_n=trowis_plan

            if row.palanner_isfact:
                trowis_n=trowis_fact

            trowi=  trowis_n[row.weekindex]
            trowis_n[row.weekindex]=trowi+1

            coli=week.coli
            #print category.title,row.weekindex,trowi, row.description, row.amount
            amount=row.amount.as_float()
            #style=Style.Gray
            style="item_plan"
            if row.palanner_isfact:
                coli+=2
                week.fact_total+=amount
                week.cat_fact_total+=amount
                cat_fact_total+=amount
                style="item_fact"
            else:
                week.plan_total+=amount
                week.cat_plan_total+=amount
                cat_plan_total+=amount
                budget=row.tx.source_budget
                #if budget.

            table[rowi+trowi,coli]=row.description, style
            table[rowi+trowi,coli+1]=amount,style

            if rowi+trowi>maxtrow:
                maxtrow=rowi+trowi
            outputedrecords+=1

    rowi=maxtrow+1

    if outputedrecords==0:
        table[rowi-2, 0]= cattitle+" [empty]"
        table[rowi-2, 0]= ""
        rowi=startrowi
    else:
        table[startrowi, 1]= cat_plan_total,"categoryline_totals"
        table[startrowi, 2]= cat_fact_total,"categoryline_totals"
        for w in weeks:
            table[startrowi,w.coli]="","categoryline_totals"
            table[startrowi,w.coli+2]="","categoryline_totals"
            table[startrowi,w.coli+1]=w.cat_plan_total,"categoryline_totals"
            table[startrowi,w.coli+3]=w.cat_fact_total,"categoryline_totals"

    for c in category.childs:
        rowi, child_outputedrecords=budget_weekly_planner_cat(table,c, rowi,date_start, date_finish,plan,weeks)
        outputedrecords+=child_outputedrecords
        if outputedrecords>0:
            rowi+=1
        #subtotal+=childtotal
    return rowi,outputedrecords
def budget_weekly_planner_cat_enumrecs(category,plan,table,rowi,date_start, date_finish,weeks):
    for row in plan.Rows:
         if row.date<date_start and row.date>date_finish:
             continue
         if row.tx.direction==1:
             continue
         for w in weeks:
             if row.date>=w.startday and row.date<=w.lastday:
                 print w.windex, row.description, row.amount
    return rowi

#corpaccounting()
homeaccounting("../money-data/")


#p = pstats.Stats('homeaccounting()')
#p.sort_stats('time')
