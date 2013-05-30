# -*- coding: utf8 -*-
from _main import load_slices_and_logicaldates
from common.CalendarHelper import Period, month_next_from
from common.Classification import Classification
from readers.StatementReader import XlsReader
from readers.avangard import AvangardReader
from reports.statement_monthly import classify_statement_monthly

__author__ = 'Max'
from readers.bankofamerica import BankOfAmericaReader
from model.accounts import *
import xlwt
from model.printstatement import PrintStatementToExcel2
from model.tags import AutoTagger
from reports.statement_with_details import classify_statement_with_details


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

    statement=cmpool.make_statement(usd)


    excel=PrintStatementToExcel2("test2.xls")
    excel.set_period(datetime(2012,1,1),datetime.now())
    excel.do_print(statement)


    print "Print statement Aggregate"

    #groups=["sveta","food",u"Рекуррентные","Reimbursment","2bank", "misc"]
    groups=["us","Salary",u"Под отчет","office"]
    excel.report_aggregate(statement,groups, True)

    excel.set_chunk(3)
    excel.report_aggregate_horizontal(statement, groups, False)



    print "Write to file"
    excel.save()

    return

def corpaccounting2013():

    boa = Account('boa',usd)

    cmpool = Pool()
    cmpool.link_account(boa)

    basedir="../money-data/corp/2013/"
    BankOfAmericaReader(basedir+'boa jan-may 2013.csv').parse_to(boa)


    print "Tagging"
    tagger=AutoTagger()
    tagger.load_declares(basedir+"2013 corp logs and cash.xls","Auto Tags")
    tagger.load_manual_tags(basedir+"2013 corp logs and cash.xls","Manual Tags")
    tagger.dotagforpool(cmpool)


    #statement по всем счетам
    print "Generate statement"
    clasfctn=Classification(from_xls=(basedir+"2013 corp logs and cash.xls","Classification"))
    clasfctn.finalize()

    wb = xlwt.Workbook()
    statement=cmpool.make_statement(usd)

    d=datetime(2012,12,31)
    while d<datetime.now():
        p=month_next_from(d)
        #print p.start
        classify_statement_with_details(clasfctn,statement,wb, "Boa#"+str(p.start.month),True, p)
        d=p.end


    classify_statement_monthly(clasfctn,statement,wb, "Monthly(BoA)")




    excel=PrintStatementToExcel2("test_corp.xls","Txs",existing_workbook=wb)
    excel.set_period(datetime(2013,1,1),datetime.now())
    excel.do_print(statement)


    #Авангард
    clasfctn=Classification(from_xls=(basedir+"2013 corp logs and cash.xls","Classification"))
    clasfctn.finalize()

    avr = Account('avr',rub)
    cashr = Account('cashr',rub)
    AvangardReader(basedir+"avr corp jan-may 2013.xls").parse_corporate2013_to(avr)
    rupool = Pool()
    rupool.link_account(avr)
    #rupool.get_tx_byid("1312800avr1718208.00").

    load_slices_and_logicaldates(rupool,basedir+"2013 corp logs and cash.xls","Slices")

    #663000

    cashconfig={'first_row':1,'col_acc':1,'col_date':0, 'col_op':2,'col_in':3,'col_out':4,'col_balance':5, 'col_tag1':6,'col_tag2':7}
    XlsReader(basedir+'2013 corp logs and cash.xls','Cash ops',cashconfig).parse_to([cashr])

    rupool.link_account(cashr)

    tagger=AutoTagger()
    tagger.load_declares(basedir+"2013 corp logs and cash.xls","Auto Tags Ru")
    tagger.load_manual_tags(basedir+"2013 corp logs and cash.xls","Manual Tags Ru")
    tagger.dotagforpool(rupool)
    ru_statement=rupool.make_statement(rub)
    excel=PrintStatementToExcel2("test_corp.xls","TxsRu",existing_workbook=wb)
    excel.set_period(datetime(2013,1,1),datetime.now())
    excel.do_print(ru_statement)

    classify_statement_monthly(clasfctn,ru_statement,wb, "Monthly(AVR)")


    d=datetime(2012,12,31)
    while d<datetime.now():
        p=month_next_from(d)
        #print p.start
        classify_statement_with_details(clasfctn,ru_statement,wb, "Avr#"+str(p.start.month),True, p)
        d=p.end

    wb.save("test_corp.xls")
    return
if __name__ == '__main__':
    corpaccounting2013()

