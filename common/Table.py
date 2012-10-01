from datetime import datetime

__author__ = 'Max'
import xlwt

class DestinationXls:
    def __init__(self,table,wb):
        self.wb=wb
        self.table=table

        self._style_defs={}
        self._style_defs[Style.Month]= xlwt.easyxf(num_format_str='D-MMM')
        self._style_defs[Style.Day]= xlwt.easyxf(num_format_str='D-MMM')
        self._style_defs[Style.Money]= xlwt.easyxf(num_format_str='#,##0')
        self._style_defs[Style.DetailedMoney]= xlwt.easyxf(num_format_str='#,##0.00')
        self._style_defs[Style.Money+Style.Red]= xlwt.easyxf('font: color-index red',num_format_str='#,##0')
        self._style_defs[Style.Money+Style.Green]= xlwt.easyxf('font: color-index green',num_format_str='#,##0')

        self._print(self.table)



    def _print(self,table):
        ws = self.wb.add_sheet(table.title)
        self.ws=ws
        ws.normal_magn=70

        for rowi in range(0, len(table._cells)):
            for coli in range(0, len(table._cells[rowi])):
                c=table._cells[rowi][coli]

                if isinstance(c, tuple):
                    v=c[0]
                    s_ind=c[1]
                    xstyle=self._style_defs[s_ind]
                    ws.write(rowi, coli, v,xstyle)
                else:
                    if isinstance(c, float):
                        ws.write(rowi, coli, c,self._style_defs[Style.Money])
                    else:
                        ws.write(rowi, coli, c)

class Style:
    Unknown=0
    Month=1
    Day=2
    Money=4
    DetailedMoney=8
    #Gray=16
    Green=32
    Red=64



class Table:
    def __init__(self, title):
        self.title=title
        self._cells=[]
        #self._cells[0]=[]
    def __setitem__(self, key, value):
        #print key
        rowi=key[0]
        ifrom=len(self._cells)
        if rowi>=ifrom:
            for i in range(ifrom,rowi+2):
                self._cells.append([])

        r=self._cells[rowi]
        coli=key[1]
        ifrom=len(r)
        if coli>=ifrom:
            for i in range(ifrom,coli+2):
                r.append("")

        if isinstance(value,datetime):
            self._cells[rowi][coli]=(value, Style.Month)
        else:
            self._cells[rowi][coli]=value

    def write_cell(self,rowi, coli,str):
        self[rowi,coli]=str
    def write_cells_vert(self,rowi,coli, cells):
        for s in cells:
            self.write_cell(rowi,coli,s)
            rowi+=1
    def set_column_width(self, coli,):
        pass


