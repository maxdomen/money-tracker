# -*- coding: utf8 -*-
import unicodedata

__author__ = 'Max'
import xlrd

class Tag:
      def __init__(self, name):
          self.name=name

class tagdef:
    def __init__(self, firstword,str, tags):
        self.firstword=firstword
        self.fullstring=str.strip()

        if len(self.firstword)<1:
            self.firstword=self.fullstring.split(' ')[0]
        self.tags=list(tags)


        if len(self.firstword)<1:
            raise Exception("auto tag first word not defined for '{0}'".format(str))

        if len(self.tags)<1:
            raise Exception("auto tag no tags defined for '{0}'".format(str))

        #self.tags.exntend(tags)

class AutoTagger:

    def __init__(self):
        self.decls=[]
        self.handlers=[]
        self.manual_tags_input=[]
        self.manual_tags_remove=[]
    def declare2(self, firststword,str, tags):
        self.decls.append( tagdef(firststword,str,tags) )
    def declares(self, arr):
        self.decls.extend(arr)
    def handler(self, func):
        self.handlers.append(func)
    def dotagforpool(self,pool):
        for acc in pool._Accounts:
            self.dotag(acc)
    def dotag(self,acc):


        dict2={}
        for td in self.decls:
            #if td.firstword=="lufthansa":
            #    print "lufthansa"
            #first=td.pattern.split(' ')[0]
            prev=dict2.get(td.firstword)
            if prev:
                if not hasattr(prev,'collisions'):
                    prev.collisions=[]
                #print "add collision", td.firstword, td.fullstring
                prev.collisions.append(td)
            else:
                dict2[td.firstword]=td
            #dict2.put(first,tag)
            #self.rel[first]=

        for tx in acc.Txs:
            str2=tx.comment
            str=str2.lower().strip()
            tx._comment_lowerstrip=str

            #if str.find("globus")>=0:
            #    print "globus"

            words=str.split(' ')
            for w in words:
                key=w
                key=w.split(',')[0]
                #subw=w.split(',')[0]
                #if len(subw)>0:
                #    key=subw[0]
                ist=dict2.get(key)
                if ist:
                    if hasattr(ist,'collisions'):
                        #print "collision here",str,"key:", ist.firstword
                        for collision in ist.collisions:

                            if len(collision.fullstring)>0:
                                #print "   test", collision.fullstring
                                if str.find(collision.fullstring)>=0:

                                    ist=collision
                                    break
                        #print "   best match", ist.fullstring
                    for t in ist.tags:
                        tx.add_tag(t)


        map(self._dotag,acc.Txs)


        for txid, tag in self.manual_tags_input:
            tx=None

            if txid[0:2]=="--":
                txid=txid[2:len(txid)]
                #print "  search by comment", txid
                tx=self._searchbycomment(txid,acc.Txs)
                #if tx:
                #    print "   found", tx.comment
            else:
                tx=acc.txsdictionary.get(txid)

            if tx:
                tx.add_tag(tag)

        for txid, tag in self.manual_tags_remove:
            tx=None
            tx=acc.txsdictionary.get(txid)
            if tx:
                tx.remove_tag(tag)
    def _searchbycomment(self,comment,txss):
        for tx in txss:
            #if tx._comment_lowerstrip.find("sport")>=0:
            #    print "SPORT1"
            #    if comment.find("sport")>=0:
            #        print "SPORT2", comment,tx._comment_lowerstrip
            #print "   check",tx._comment_lowerstrip
            if tx._comment_lowerstrip==comment:
                return tx
        return None
    def _dotag2(self, tx):
        pass
    def _dotag(self, tx):

        #str2=tx.dest+tx.comment
        #str=str2.lower()


        for func in self.handlers:
            res=func(tx)
            if res!=None:
                tx.add_tag(res)



        #for pattern, tag in self.decls:
        #    if str.find(pattern)>=0:
        #        tx.add_tag(tag)


     

    def load_manual_tags(self,filename, sheetname):
        book = xlrd.open_workbook(filename)
        sheet=book.sheet_by_name(sheetname)



        for rowi in range(1,sheet.nrows):
            r=sheet.row(rowi)
            unicode_txid=r[1].value
            if isinstance(unicode_txid, unicode):
                txid=unicodedata.normalize('NFKD', unicode_txid).encode('ascii','ignore')
            else:
                txid=unicode_txid

            if len(txid) and  not txid[1:4].isdigit():
                #вместо id транзакции дано описание транзакции
                txid="--"+unicode_txid.lower().strip()
                #print "",txid


            tag_add1=r[2].value
            tag_add2=""
            tag_remove=""
            if len(r)>3:
                tag_add2=r[3].value
            if len(r)>4:
                tag_remove=r[4].value

            if len(tag_add1)>0:
                self.manual_tags_input.append((txid,tag_add1))

            if len(tag_add2)>0:
                self.manual_tags_input.append((txid,tag_add2))

            if len(tag_remove)>0:
                self.manual_tags_remove.append((txid,tag_remove))

    def load_declares(self,filename, sheetname):
        book = xlrd.open_workbook(filename)
        sheet=book.sheet_by_name(sheetname)



        for rowi in range(1,sheet.nrows):
            r=sheet.row(rowi)
            firststword=r[0].value.strip().lower()
            #if firststword=="firststword":
            #    print "firststword"

            str=r[1].value.lower()
            t1=r[2].value
            t2=r[3].value
            t3=r[4].value
            tags=[]

            if t1!="": tags.append(t1)
            if t2!="": tags.append(t2)
            if t3!="": tags.append(t3)
            self.declare2(firststword,str,tags)