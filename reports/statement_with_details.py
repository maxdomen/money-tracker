from common.Classification import TagTools
from model.accounts import RowType
import xlwt
__author__ = 'Max'
def classify_statement_with_details(clasfctn,statement,wb, sheetname2, collapse_company_txs, p):
    for r in statement.Rows:
        if r.type!=RowType.Tx:
            continue


        #dt=r.get_logical_date()
        dt=r.get_human_or_logical_date()


        if not (dt>=p.start and dt<=p.end):
            continue

        g=clasfctn.match_tags_to_category(r.get_normilized_tags())

        if not hasattr(g, 'txs'):
            g.txs=[]
        g.txs.append(r)


    ws = wb.add_sheet(sheetname2)
    ws.normal_magn=70
    ws.col(1).width=256*12
    ws.col(2).width=256*12
    ws.col(7).width=256*12

    #rowi=0


    details_for_cat(ws,clasfctn._root,0, p.start, p.end)


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
            #dt=r.get_logical_date()

            dt=r.get_human_or_logical_date()
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
            ws.write(rowi, bc+9,r.tx._txid )
            rowi+=1

    for c in category.childs:
        rowi, childtotal=details_for_cat(ws,c, rowi,date_start, date_finish)
        rowi+=1
        subtotal+=childtotal

    ws.write(rowi, 3, u"Total in {0}".format(category.title))
    ws.write(rowi, 7,subtotal ,style_money)

    rowi+=1
    return rowi,subtotal
def build_category_path(category):
    cattitle=category.title
    p=category.parent
    while p:
        if p.title=="_root":
            break
        cattitle=p.title+"/"+cattitle
        p=p.parent
    return cattitle