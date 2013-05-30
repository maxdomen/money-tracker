from common.Classification import ClassificationDataset, Period, ClassificationPrinter

__author__ = 'Max'
def classify_statement_monthly(classification,statement,wb, sheetname):



    monthlydataset=ClassificationDataset(classification,Period.Month, statement)

    ws = wb.add_sheet(sheetname)
    ws.col(0).width=256*40
    ws.panes_frozen = True
    ws.horz_split_pos = 2
    ws.vert_split_pos = 1
    ws.normal_magn=70


    printer=ClassificationPrinter(monthlydataset, existing_sheet=ws)
    return classification