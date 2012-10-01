
from common.Table import Style

__author__ = 'Max'

def details_for_cat(table,rowi,category,print_fun):


    #style_time1 = xlwt.easyxf(num_format_str='D-MMM')
    #style_money=xlwt.easyxf(num_format_str='#,##0.00')
    cattitle=category.title
    p=category.parent
    while p:
            if p.title=="_root":
                break
            cattitle=p.title+"/"+cattitle
            p=p.parent

    rowi+=1
    rowi_start_of_category=rowi




    subtotal=0
    if hasattr(category, 'stories'):
        for r in category.stories:
            rowi+=1

            v=print_fun(table,rowi,r)


            subtotal+=v

    for c in category.childs:
        #rowi+=1
        rowi, childtotal=details_for_cat(table,rowi,c,print_fun)
        #rowi+=1
        subtotal+=childtotal

    category.total_sum=0
    if subtotal>0:
        rowi+=1
        table[rowi_start_of_category,0]=cattitle
        table[rowi, 2]= u"Total in"
        table[rowi, 4]= u"{0}".format(cattitle)
        table[rowi, 5]=subtotal ,Style.DetailedMoney
        category.total_sum=subtotal
        rowi+=1
    else:
        rowi-=1


    return rowi,subtotal
