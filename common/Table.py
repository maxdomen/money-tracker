from datetime import datetime
from xlwt.Formatting import Pattern

__author__ = 'Max'
import xlwt

class DestinationXls:
    def __init__(self,table,wb, def_font_height=None):

        self.wb=wb
        ownedfile=False
        if isinstance(wb,str):
            self.wb_filename=self.wb
            self.wb=xlwt.Workbook()
            ownedfile=True

        self.table=table

        self._style_defs={}
        #for style in

        ps=[ Style.Month, Style.Day,Style.Money,Style.DetailedMoney,Style.Percent]

        for p in ps:
            self._style_defs[p]= xlwt.easyxf(num_format_str=Style.to_xls_formattingstyle(p))


        if def_font_height:
            for sd in self._style_defs.values():
                sd.font.height=def_font_height*20


        for id,cs in table._custom_styles.items():
            xfobj = xlwt.XFStyle()
            font_size, background_color,foreground_color,bold,italic,formatting_style=cs
            if def_font_height:
                xfobj.font.height=def_font_height*20
            if font_size!=8:
                xfobj.font.height=font_size*20

            if bold:
                xfobj.font.bold=True
            if italic:
                xfobj.font.italic=True
            if background_color!=Color.NoColor:
                color= Color.to_xls_color_index(background_color)
                xfobj.pattern.pattern=Pattern.SOLID_PATTERN
                xfobj.pattern.pattern_fore_colour=color
            if foreground_color!=Color.NoColor:
                color= Color.to_xls_color_index(foreground_color)
                xfobj.font.colour_index=color
            if formatting_style:
                xlsfs=Style.to_xls_formattingstyle(formatting_style)
                xfobj.num_format_str=xlsfs
            self._style_defs[id]=xfobj
                #xfobj.pattern.pattern_fore_colour=color

        self._print(self.table)

        if ownedfile:
            self.wb.save(self.wb_filename)


    def _print(self,table):
        ws = self.wb.add_sheet(table.title)
        self.ws=ws
        ws.normal_magn=table._normal_magn

        for coli, width_in_chars in table._col_widths:
            ws.col(coli).width=256*width_in_chars

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
    Month=1
    Day=2
    Money=4
    DetailedMoney=8

    Percent=17
    #Green=32
    #Red=64

    @staticmethod
    def to_xls_formattingstyle(src):
        colors={}
        colors[Style.Month]='D-MMM'
        colors[Style.Day]='D-MMM'
        colors[Style.Money]='#,##0'
        colors[Style.DetailedMoney]='#,##0.00'
        colors[Style.Percent]='0.00%'
        return colors[src]
class Color:
    NoColor=0
    Black=1
    Green=32
    LightGreen=33
    Red=64
    Gray=128
    LightGray=129

    @staticmethod
    def to_xls_color_index(src):
        colors={}
        colors[Color.NoColor]=0
        colors[Color.Black]=0x08
        colors[Color.Green]=0x3A
        colors[Color.LightGreen]=0x2A
        colors[Color.Red]=0x10
        colors[Color.Gray]=0x17
        colors[Color.LightGray]=0x16
        return colors[src]

class Table:
    def __init__(self, title):
        self.title=title
        self._cells=[]
        self._col_widths=[]
        self._custom_styles={}
        self._normal_magn=70
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
    def set_column_width(self, coli,width_in_chars):
        self._col_widths.append( (coli,width_in_chars) )

    def define_style(self, style_id,font_size=8, background_color=Color.NoColor,foreground_color=Color.Black,bold=False,italic=False, formatting_style=None):
        self._custom_styles[style_id]=(font_size, background_color,foreground_color,bold,italic,formatting_style)


