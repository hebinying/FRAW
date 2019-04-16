#coding=utf-8
import xlrd,xlwt

C=["case_id","checkname","method","url","params","checkpoint"]
# 写入表格格式
datas=["编号","说明","Data","Data.analysis_model","Data.app","Data.date_unit","Data.dim_fields.[]","Data.filter.[]"\
       ,"Data.model_type","Data.from_date","Data.num_fields.[]","Data.order_by.[]","Data.page	Data.page_size","Data.product","Data.tenant","Data.to_date"]
# product,app,tenant数据准备
# 从txt文件中读取,按Jason格式返回

class readExcel():
    def __int__(self,filepath):
        self.filepath=filepath



class writeExcel():
    def __init__(self,datas):
        self.datas=datas
        self.wb=xlwt.Workbook()
        self.sheet=self.wb.add_sheet((datas["wordname"]).decode("gbk"))
        #分解数据
        self.filename=datas["wordname"]
        self.filepath=datas["wordpath"]
        self.tables=datas["content"]
        self.tablenumber=datas["tablenumber"]
    #分解content数据并保存到excel表中
    def write_tabledata(self):
        num=0
        dictname={}
        for i in range(len(C)):
            self.sheet.write(num,i,C[i])

        for i in range(len(self.tables)):
            num+=1
            dictname = {}
            if len(self.tables[i])!=0:
                dictname["case_id"]="case_"+str(num)
                dictname["checkname"]=(self.tables[i]["testcasename"])
                dictname["method"]=self.tables[i]["method"]
                dictname["url"]=self.tables[i]["url"]
                dictname["params"]=(self.tables[i]['params'])
                dictname["checkpoint"]=self.tables[i]['checkpoint']
                #dictname["sample"]=(self.tables[i]['sample'])
            self.write_excel(num,dictname)
        self.wb.save(self.filename+'.xlsx')

    def write_excel(self,rownum,data):
        ###############增加自动换行格式
        alignment = xlwt.Alignment()
        # alignment.horz = xlwt.Alignment.HORZ_CENTER
        alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT
        alignment.vert = xlwt.Alignment.VERT_CENTER
        style = xlwt.XFStyle()
        style.alignment = alignment
        ###########

        for i in range(len(C)):
            # print data[C[i]]
            self.sheet.write(rownum,i,data[C[i]],style)



if __name__=="__main__":
    a = readWord.WordUtil("m2c.scm.api-1.3.0.docx")
    data=a.get_tablesdata()
    b=writeExcel(data)
    b.write_tabledata()