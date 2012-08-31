#! coding: utf-8
import cProfile
from common.Classification import Classification, ClassificationDataset, ClassificationPrinter, Period
from model import debt

import readers.StatementReader
from model.dashboard import DashboardDataset, DashboardPublisher, BigPicture, BigPicturePublisher
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
from model.printstatement import PrintStatementToExcel2, BalanseObservation
from readers.StatementReader import *
#from debt import *
import pstats


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


def svetaaccounting(basedir,acc):
    #tcs = Account('scash',rub)
    #avr = Account('balt',rub)
    cashconfig={'first_row':1,'col_acc':1,'col_date':0, 'col_op':4,'col_in':2,'col_out':3,'col_balance':-1, 'col_tag1':5,'col_tag2':6}
    XlsReader(basedir+'home/2012/2012 sveta.xls','Records',cashconfig).parse_to([acc])


def parsing(basedir,avr,avu,tcs,boa,wallet,safe, sveta, budget):
    print "Load sources"
    BankOfAmericaReader(basedir+"home/2012/boa 2012.csv").parse_to(boa)

    AvangardReader(basedir+"home/2012/avr 1.1.2012 - 1.4.2012.xls").parse_private_to(avr)
    AvangardReader(basedir+"home/2012/avr apr 2012.xls").parse_private_to(avr)
    AvangardReader(basedir+"home/2012/avr may 2012.xls").parse_private_to(avr)
    AvangardReader(basedir+"home/2012/avr june 2012.xls").parse_private_withreserved_to(avr)
    AvangardReader(basedir+"home/2012/avr july 2012.xls").parse_private_withreserved_to(avr)
    AvangardReader(basedir+"home/2012/avr aug 2012.xls").parse_private_withreserved_to(avr)


    AvangardReader(basedir+"home/2012/avu 1.1.2012 - 1.4.2012.xls").parse_private_to(avu)
    AvangardReader(basedir+"home/2012/avu apr 2012.xls").parse_private_to(avu)
    AvangardReader(basedir+"home/2012/avu may 2012.xls").parse_private_to(avu)
    AvangardReader(basedir+"home/2012/avu june 2012.xls").parse_private_withreserved_to(avu)
    AvangardReader(basedir+"home/2012/avu july 2012.xls").parse_private_withreserved_to(avu)
    AvangardReader(basedir+"home/2012/avu aug 2012.xls").parse_private_withreserved_to(avu)




    TCSBankReader(basedir+"home/2012/tcs jan 2012new.csv").parse2012_to(tcs)
    TCSBankReader(basedir+"home/2012/tcs feb-apr 2012new.csv").parse2012_to(tcs)
    TCSBankReader(basedir+"home/2012/tcs may 2012.csv").parse2012_to(tcs)
    TCSBankReader(basedir+"home/2012/tcs june 2012.csv").parse2012_to(tcs)
    TCSBankReader(basedir+"home/2012/tcs july 2012.csv").parse2012_to(tcs)
    TCSBankReader(basedir+"home/2012/tcs aug 2012.csv").parse2012_to(tcs)


    #TCSBankReader("Data/home/2012/tcs june 2012b.csv").parse2011v2_to(tcs)

    cashconfig={'first_row':1,'col_acc':1,'col_date':0, 'col_op':2,'col_in':3,'col_out':4,'col_balance':5, 'col_tag1':6,'col_tag2':7}
    XlsReader(basedir+'home/2012/2012 logs and cash.xls','Cash ops',cashconfig).parse_to([wallet,safe, sveta])



    cashconfig={'first_row':1,'col_date':0}
    accstoread={avr:1, avu:2, tcs:3, wallet:4, safe:5, sveta:6}
    #accstoread={avu:2}

    XlsLeftoversJournalReader(basedir+'home/2012/2012 logs and cash.xls','Account Log',cashconfig).parse_to(accstoread)

    budget.read(basedir+'home/2012/2012 logs and cash.xls','Budget')
    budget.read(basedir+'home/2012/2012 sveta.xls','Plan')

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

    #familypool.get_tx_byid("1261300avr15100.00").slice(u"Плата за внешний перевод",100,["2bank"],[])

    #familypool.get_tx_byid("1261300avr15100.00").slice(u"Fee за межбанковский перевод",100,["2bank"],[])
    #familypool.add_transition("1261300avr15000.00","1261300tcs15000.00",commission="1261300avr100.00", comission2="1261300avr10.00[1]")


    #apr
    familypool.get_tx_byid("1252113avr23682.00").set_logical_date(datetime(2012,4,30))
    #may
    familypool.get_tx_byid("126900wallet18318.00[1]").set_logical_date(datetime(2012,5,30))
    familypool.get_tx_byid("126900avr41864.20").set_logical_date(datetime(2012,5,30))

    tx=familypool.get_tx_byid("127600avr41682.00").set_logical_date(datetime(2012,6,30))

    tx=familypool.get_tx_byid("1271200wallet18318.00[1]").set_logical_date(datetime(2012,6,30))
    #tx=familypool.get_tx_byid("1272000wallet10000.00").set_logical_date(datetime(2012,6,30))

def printdata(basedir,statement,dashboarddataset,bigpicture,virt_max_cm_statement,virt_private_debts):

    print "Print statement Txs"




    dashboardpublisher=DashboardPublisher(dashboarddataset,"test.xls","Dashboard")
    BigPicturePublisher(bigpicture,"test.xls","BigPicture",existing_workbook=dashboardpublisher.wb)

    excel=PrintStatementToExcel2("test.xls","Txs",existing_workbook=dashboardpublisher.wb)
    excel.set_period(datetime(2012,1,1),datetime.now())
    excel.do_print(statement)

    #dataset.layout._finalize()
    #pub=Publisher(dataset, "test.xls", "SheetVert",existing_workbook=excel.wb)
    #pub.vertical()


    #pub2=Publisher(datasetmonthly, "test.xls", "SheetVert2", existing_workbook=pub.wb)
    #pub2.horizontal()



    #pubAgg1=Publisher2(agg, "test.xls", "Agg1",existing_workbook=excel.wb)
    
    #pub3=Publisher2(debts, "test.xls", "Agg1",existing_workbook=excel.wb,existing_sheet=pubAgg1.ws, after_row=8, sub_report_title="Debts")


    observer=PrintStatementToExcel2("test.xls","CM Balance", excel.wb)
    observer.set_period(datetime(2012,1,1),datetime.now())
    observer.do_print(virt_max_cm_statement)


    #pub2=Publisher2(agg2, "test.xls", "BudgetAgg1",existing_workbook=excel.wb)

    observer2=PrintStatementToExcel2("test.xls","Personal Debts", excel.wb)
    observer2.set_period(datetime(2012,1,1),datetime.now())
    observer2.do_print(virt_private_debts)



    #pub2=Publisher(budgetmonthly, "test.xls", "BudgetVert2", existing_workbook=pub.wb)
    #pub2.horizontal()




    return excel.wb




def homeaccounting(basedir):
    loadrates()



    #account wo any statement


    pool = Pool()


    #actual
    tcs = Account('tcs',rub)
    tcs.limit=-75000
    avr = Account('avr',rub)
    avu = Account('avu',usd)
    avu.limit=-2000
    boa = Account('boa',usd)
    wallet = Account('wallet',rub)
    safe = Account('safe',rub)
    sveta = Account('sveta',rub)
    svetaaccounting("../money-data/",sveta)

    budget= debt.Budget()

    parsing(basedir,avr,avu,tcs,boa,wallet,safe, sveta,budget)

    familypool = Pool()
    familypool.link_account(tcs)
    familypool.link_account(avr)
    familypool.link_account(avu)
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




    budgetstatement=budget.make_statement(rub)


    #cProfile.runctx('budgetstatement=budget.make_statement(rub)', globals(),locals())


    dashboarddataset=DashboardDataset(statement,budgetstatement)



    bigpicture=BigPicture(statement,budgetstatement)


    #layout=createlayout()
    #layout.create_automatic_groups(statement)






    #dataset=Dataset(layout, statement, start=datetime(2012,1,1))


    #datasetmonthly=Dataset(layout, statement, start=datetime(2012,1,1), chunktype=3)



    #budgetlayout=createlayout()
    #budgetlayout.create_automatic_groups(budgetstatement)
    #budgetmonthly=Dataset(budgetlayout, budgetstatement, start=datetime(2012,1,1), chunktype=3)

    #agg=Aggregate(datasetmonthly)
    #r1=agg.CreateRow(u"Приход", sumgroups=[layout.familyin])
    #r2=agg.CreateRow(u"Расход", sumgroups=[layout.family, layout.annually, layout.spending.untagged, layout.lost])
    #r3=agg.CreateRowCalc(u"Семейная EBITDA", r1,r2)
    #r4=agg.CreateCumulative(u"Cumulative", r3)
    #agg.go()



    #agg2=Aggregate(budgetmonthly)
    #r1=agg2.CreateRow(u"Приход", sumgroups=[budgetlayout.income])
    #r2=agg2.CreateRow(u"Расход", sumgroups=[budgetlayout.spending])
    #r3=agg2.CreateRowCalc(u"EBITDA семьи", r1,r2)
    #r4=agg2.CreateCumulative(u"Cumulative", r3)
    #agg2.go()


    virt_private_debts_acc = Account('virt_private_debts',rub)
    virt_private_debts =familypool.make_statement(rub,virt_private_debts_acc ,
                                                    filter_debit=[u"debt"],
                                                    filter_credit=[u"debt"],
                                                    skip_transitions=True)



    virt_max_cm = Account('virt_max_cm',rub)
    virt_max_cm.leftover(LeftOver(datetime(2012,1,1),-25441))

    tx=Tx(-1477,datetime(2012,2,13) )
    tx.comment=u"списание денег за подарки "
    tx.add_tag("Reimbursment")
    virt_max_cm.out(tx)

    tx=Tx(2000,datetime(2012,2,22) )
    tx.comment=u"макс заплатил за корпоративные блины "
    tx.add_tag("Reimbursment")
    virt_max_cm.out(tx)

    tx=Tx(5000,datetime(2012,3,19) )
    tx.comment=u"выдали максу кешем"
    tx.add_tag("Reimbursment")
    virt_max_cm.income(tx)


    tx=Tx(350,datetime(2012,3,27) )
    tx.comment=u"выдали максу кешем"
    tx.add_tag("Reimbursment")
    virt_max_cm.income(tx)

    tx=Tx(200,datetime(2012,4,13) )
    tx.comment=u"выдали максу кешем (подарок мише на свадьбу)"
    tx.add_tag("Reimbursment")
    virt_max_cm.income(tx)


    #123271359avr20000.00 splice

    virt_max_cm.leftover(LeftOver(datetime(2012,2,22),-30454.64))
    virt_max_cm.leftover(LeftOver(datetime(2012,3,21),5407.42))
    virt_max_cm.leftover(LeftOver(datetime(2012,5,5),228423.96))
    virt_max_cm.leftover(LeftOver(datetime(2012,6,8),85318.81))
    virt_max_cm.leftover(LeftOver(datetime(2012,7,9, 17,0,0),67205.08))

    virt_max_cm.leftover(LeftOver(datetime(2012,7,21, 16,0,0),233207.07))


    #virt_max_cm.leftover(LeftOver(datetime(2012,6,12),67000.81))


    virt_max_cm_statement=familypool.make_statement(rub,virt_max_cm,
                                                    filter_debit=[u"Reimbursment",u"Деньги CM"],
                                                    filter_credit=[u"Reimbursment",u"Под отчет",u"Деньги CM"],
                                                    skip_transitions=True)







    #долги
    #по картам
    #  источник счет, его отрицательное значение
    # TCS
    # AVU
    # компания
    # источник - виртуальный аккаунт
    debts=debt.Debts(start=datetime(2012,1,1),statement=statement)
    debts.add_credit_card_as_account(statement,tcs,mode=1)
    debts.add_credit_card_as_account(statement,avu,mode=1)
    debts.add_credit_card_as_account(virt_max_cm_statement,virt_max_cm, mode=2, qualificator=-1)
    debts.add_credit_card_as_account(virt_private_debts,virt_private_debts_acc, mode=2, qualificator=-1)
        

    debts.calc_total()


    wb=printdata(basedir,statement,dashboarddataset,bigpicture,virt_max_cm_statement,virt_private_debts)
    classify_statement(basedir,statement,wb, "Monthly")
    classify_statement(basedir,budgetstatement,wb, "BudgetMonthly")
    wb.save("test.xls")

def classify_statement(basedir,statement,wb, sheetname):

    classification=Classification(from_xls=(basedir+"home/2012/2012 logs and cash.xls","Classification"))

    #classification.finalize()

    #classification.finalize()

    classification.create_auto_classification(statement)
    classification.get_category_by_id("company_txs")._collapsed=True
    classification.finalize()
    monthlydataset=ClassificationDataset(classification,Period.Month, statement)





    ws = wb.add_sheet(sheetname)
    ws.col(0).width=256*40

    ws.panes_frozen = True
    ws.horz_split_pos = 2
    ws.vert_split_pos = 1
    ws.normal_magn=70


    printer=ClassificationPrinter(monthlydataset, existing_sheet=ws)
####
def corpaccounting():
    loadrates()

    avr = Account('avr',rub)
    avu = Account('avu',usd)
    chase = Account('chase',usd)
    max = Account('max',rub)
    egor = Account('egor',rub)
    anna = Account('anna',rub)
    ak = Account('ak',rub)

    cmpool = Pool()
    cmpool.link_account(avr)
    cmpool.link_account(avu)
    cmpool.link_account(chase)
    cmpool.link_account(max)
    cmpool.link_account(egor)
    cmpool.link_account(anna)
    cmpool.link_account(ak)

    cashconfig={'first_row':1,'col_date':0, 'col_acc':1, 'col_op':2,'col_in':3,'col_out':4,'col_balance':5, 'col_tag1':6,'col_tag2':7}
    StatementReader.XlsReader('data/corp/2012/corp 2012 logs and cash.xls','Cash ops',cashconfig).parse_to( [max,egor,anna,ak])

    cashconfig={'first_row':1,'col_date':0}
    accstoread={avr:1, avu:2, chase:3, max:4, egor:5, anna:6, ak:7}
    StatementReader.XlsLeftoversJournalReader('data/corp/2012/corp 2012 logs and cash.xls','Account Log',cashconfig).parse_to(accstoread)

    #"C:\src\money\data\corp\2012\corp 2012 logs and ##cash.xls"

    ChaseBankReader("data/corp/2012/corp chase 2012 jan-mar.xls","Sheet1").parse_to(chase)
    AvangardReader("data/corp/2012/corp avr 2012 jan-mar.xls").parse_corporate_to(avr)
    AvangardReader("data/corp/2012/corp avr 2012 apr.xls").parse_corporate_to(avr)

    tagger=AutoTagger()
    tagger.load_declares("data/corp/2012/corp 2012 logs and cash.xls","Auto Tags")
    tagger.load_declares("data/corp/2012/corp 2012 logs and cash.xls","Auto Tags Chase")
    tagger.load_manual_tags("data/corp/2012/corp 2012 logs and cash.xls","Manual Tags")



    tagger.dotag(avr)
    tagger.dotag(avu)
    tagger.dotag(chase)
    tagger.dotag(max)
    tagger.dotag(egor)
    tagger.dotag(anna)

    #122100avr345501.03
    cmpool.get_tx_byid("122100avr345501.03").set_logical_date(datetime(2012,1,31)) #переносим транзакцию зарплаты на соотвествующий месяц
    cmpool.get_tx_byid("1221000avr36774.19").set_logical_date(datetime(2012,1,31))
    cmpool.get_tx_byid("123500avr13682.00").set_logical_date(datetime(2012,2,28))


    cmpool.get_tx_byid("122100avr345501.03").slice("Max Gannutin",65000,["maxg"])
    cmpool.get_tx_byid("122100avr345501.03").slice("Egor",15000,["egor"])
  #  cmpool.get_tx_byid("123500avr13682.00").slice("Max Gannutin",65000,["maxg"])





    #maxcost=Cost("Макс Ганнутин ЗП")
    #msxcost.add_tx()


    statement=cmpool.make_statement(usd)


    excel=PrintStatementToExcel2("test2.xls")
    excel.set_period(datetime(2012,1,1),datetime.now())
    excel.do_print(statement)


    print "Print statement Aggregate"

    #groups=["sveta","food",u"Рекуррентные","Reimbursment","2bank", "misc"]
    groups=["us","Salary",u"Под отчет","office"]
    excel.report_aggregate(statement,groups, True)

    #,excel.weekly
    #,excel.daily
    excel.set_chunk(3)
    excel.report_aggregate_horizontal(statement, groups, False)



    print "Write to file"
    excel.save()

    return

#corpaccounting()
homeaccounting("../money-data/")


#p = pstats.Stats('homeaccounting()')
#p.sort_stats('time')
