# coding=utf-8
#soffice --headless --convert-to docx:"Office Open XML Text" --outdir=/home/ *.doc 

import docx
import os
import sys
import traceback

writeFileFlag = False
def replaceText(text):
    return text.replace("括号内填写原信息，下同",
            "").replace(" ", "").replace("（", "").replace("）",
            "").replace("姓名","").replace("：","").replace(":",
            "").replace("否，还兼职其他工作", "").replace("",
            "").replace("£","").replace("√","").replace("_",
            "").replace("_","").replace("","").replace("其他（请补充填写）",""
            ).replace("其他请补充填写",""
            ).replace('调查员',
        "").replace("请调查员记录","").replace("使用其他方式查灾请补充填写",
    "").replace("R","").replace("请记录","").replace("座机","").replace("；","").replace("✅","")

#取基本信息，括号内的内容
def getTextByPosition(data, rowno, colno,pno):
    try:
        return getTextin(data[rowno][colno][pno].text)
    except Exception as e:
        return ""
    pass

def getTextin(text):
    if text.find("括号内填写原信息，下同") != -1:
        return replaceText(text)
    try:
        #取括号内的值
        #（可以考虑直接全部采用替换，讲关键字全部替换掉即可）
        return replaceText(text.split("：")[1])
    except Exception as e:
        open("error.log","a").write(str(e))
        return replaceText(text)
    
def isSelected(paragrahObj):
    if paragrahObj._element.xml.find("√") != -1 or paragrahObj._element.xml.find('''<w:t>R</w:t>''') !=-1:
        return True
    # 纯方框没选中
    if paragrahObj._element.xml.find('''F0FE''') == -1:
        #未选中
        return False
    
def multiSelect(pList):
    text = ""
    for p in pList:
        if isSelected(p):
#            print(p._element.xml)
            #sys.exit(-1)
            text += replaceText(p.text)+"，"
    return text
    pass

def simpleSelect(paragrahObj):
    if isSelected(paragrahObj):
        return "是"
    else:
        return "否"
    
def getUpdate(paragrahObj):
    text = paragrahObj.text
    yesIndex = text.find("是")
    noIndex = text.find("否")
    selectIndex = text.find("√")
    if selectIndex != -1:
        if selectIndex < noIndex and selectIndex < yesIndex:
            return "否"
    return "是"
def parseFile(filename):
    global writeFileFlag
    #doc = docx.Document("/Users/desktop/Downloads/59.docx") 
    doc = docx.Document(filename) 
    table = doc.tables[0]
    data = {}
    for rowno, row in enumerate(table.rows):
         rowCellsText = []
         for colno, cell in enumerate(row.cells):
              if not cell.text in rowCellsText:
                    rowCellsText.append(cell.text) 
                    # print("rowno %s  colno %s: %s"%(rowno, colno, cell.text))
                    for pno, p in enumerate(cell.paragraphs):
                        if writeFileFlag:
                            open("text.txt", "a").write(filename+"\n")
                            open("text.txt", "a").write("rowno %s, colno %s, pno:%s, %s, %s\n" % (rowno, colno, pno, p.text, ''))
                        if not rowno in data:
                              data[rowno] = {}
                        if not colno in data[rowno]:
                              data[rowno][colno] = {}
                        if not pno in data[rowno][colno]:
                              data[rowno][colno][pno] = {}
                        data[rowno][colno][pno] = p
            
    base8Pno = 0
    if data[8][0][21].text.find("如为兼职，每天用于灾害信息员的工作时长") == -1:
        #缺少两个问题：1.如为兼职， 2.兼职时长，
        base8Pno = -8
        pass
    
#     baseSatisfy = 0
#     if data[8][0][base8Pno+70].text.find("如不满意，您觉得都存在哪些问题") !=1:
#         baseSatisfy = -1
        
    open("result.csv","a").write(
     """%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s\n"""
     % (
          getUpdate(data[1][0][0]),
          getTextByPosition(data,2,0,0),  # 姓名
          getTextByPosition(data,2,1,0),  # 性别
          getTextByPosition(data,2,2,0),  # 年龄
          getTextByPosition(data,3,0,0),  # 学历
          getTextByPosition(data,3,1,0),  # 专业
          getTextByPosition(data,3,2,0),  # 职称
          getTextByPosition(data,4,0,0),  # 政治面貌
          getTextByPosition(data,4,1,0),  # 移动电话
          "",  # 座机
          getTextByPosition(data,6,0,0),  # 区县
          getTextByPosition(data,6,1,0),  # 乡镇（街道）
          getTextByPosition(data,6,2,0),  # 村（社区）
          getTextByPosition(data,7,0,0),  # 单位
          getTextByPosition(data,7,1,0),  # 工龄
          getTextByPosition(data,7,2,0),  # 职务
          simpleSelect(data[8][0][2]),#是否为原民政灾害信息员
         (
             "不清楚" if(not isSelected(data[8][0][7])) else "否"
          ) if(not isSelected(data[8][0][6])) else "是",#是否为地震速报员
         simpleSelect(data[8][0][11]),#安全员？
         (
             "专职安全员"  if(isSelected(data[8][0][14]))  else "安全巡查员" 
          ) if (isSelected(data[8][0][11])) else "",# 继续 安全员
         "否"  if(not isSelected(data[8][0][18])) else  "是" ,#专职灾害信息员
         (replaceText(data[8][0][19].text))  if(not isSelected(data[8][0][14])) else  "" ,#是否兼职，可选
         ( 
             (
                 "" if(not isSelected(data[8][0][base8Pno+24])) else "3小时以上"
             ) if(not isSelected(data[8][0][base8Pno+23])) else "2-3小时"
             ) if(not isSelected(data[8][0][base8Pno+22])) else  "1小时以下",#时长 ，兼职时间，可选
         replaceText(data[8][0][base8Pno+27].text),#本职工作，可选
         ((
             ( 
             (
                 "" if(not isSelected(data[8][0][base8Pno+33])) else "5-10年"
             ) if(not isSelected(data[8][0][base8Pno+32])) else "3-5年"
             ) if(not isSelected(data[8][0][base8Pno+31])) else  "1-3年"
         ) if(not isSelected(data[8][0][base8Pno+30])) else  "少于1年" 
         )if(not isSelected(data[8][0][base8Pno+34])) else  "10年以上" ,#工作年限
         (
             ( 
             (
                 "" if(not isSelected(data[8][0][base8Pno+41])) else "以上皆无"
             ) if(not isSelected(data[8][0][base8Pno+40])) else replaceText(data[8][0][40].text)
             ) if(not isSelected(data[8][0][base8Pno+39])) else  "安全员"
         ) if(not isSelected(data[8][0][base8Pno+38])) else  "灾害信息员培训证书" ,#
         "",#z证书多选2
         ( 
             (
                 "" if(not isSelected(data[8][0][base8Pno+46])) else "否"
             ) if(not isSelected(data[8][0][base8Pno+47])) else "单位安排"
             ) if(not isSelected(data[8][0][base8Pno+48])) else  "个人安排",#体检
                  (
             ( 
             (
                 "" if(not isSelected(data[8][0][base8Pno+51])) else "非常健康"
             ) if(not isSelected(data[8][0][base8Pno+52])) else "健康"
             ) if(not isSelected(data[8][0][base8Pno+53])) else  "亚健康"
         ) if(not isSelected(data[8][0][base8Pno+54])) else  "健康状况较差" ,#健康情况,
        multiSelect([data[8][0][base8Pno+57],
                     data[8][0][base8Pno+58],
                     data[8][0][base8Pno+59],
                     data[8][0][base8Pno+60],
                     data[8][0][base8Pno+61],
                     data[8][0][base8Pno+62],
                     data[8][0][base8Pno+63],
                     data[8][0][base8Pno+64],
                     data[8][0][base8Pno+65],
                     ]),#请问您的健康状况是否在以下方面存在问题
        (
                 "" if(not isSelected(data[8][0][base8Pno+69])) else "不满意"
             ) if(not isSelected(data[8][0][base8Pno+70])) else "满意",#您对目前从事的灾害信息员工作是否满意
        "",#存在问题
        "",#建议
       ((
             ( 
             (
                 "" if(not isSelected(data[10][0][1])) else "手机"
             ) if(not isSelected(data[10][0][2])) else "座机"
             ) if(not isSelected(data[10][0][3])) else  "传真"
         ) if(not isSelected(data[10][0][4])) else  "电子邮件" 
         )if(not isSelected(data[10][0][5])) else  replaceText(data[10][0][5].text) ,#报警方式
       "", #报灾方式（其他）
        ((
                 "" if(not isSelected(data[10][0][8])) else "在村长、村支书、村干部带领下查灾，需要参考领导的意见"
             ) if(not isSelected(data[10][0][9])) else "独立查灾"
             ) if(not isSelected(data[10][0][10])) else  replaceText(data[10][0][10].text),#查灾方式（多选1）
        "",#查灾方式多选2
        multiSelect([
                    data[10][0][13],
                    data[10][0][14],
                    data[10][0][15],
                    data[10][0][16],
                    data[10][0][17],
                    data[10][0][18],
                    data[10][0][19],
                    data[10][0][20]
                    ]
                    ),#影响您不能及时报灾的因素有
        simpleSelect(data[10][0][23]),#是否上报过自然灾害灾情信息
        "",#报灾
        "",#救灾
        #"",#救灾建议
        simpleSelect(data[12][0][1]),#是否接受过灾害信息员的相关培训
        multiSelect([
                    data[12][0][5],
                    data[12][0][6],
                    data[12][0][7],
                    data[12][0][8],
                    data[12][0][9]
                    ]
                    ),#最近一次接受灾害信息员相关培训是在
        multiSelect([
                    data[12][0][12],
                    data[12][0][13],
                    data[12][0][14],
                    data[12][0][15],
                    data[12][0][16]
                    ]
                    ), #一共接受过多少次灾害信息员相关培训
        multiSelect([
                    data[12][0][19],
                    data[12][0][20],
                    data[12][0][21],
                    ]
                    ),        #培训的课程内容是否实用
        multiSelect([
                    data[12][0][24],
                    data[12][0][25],
                    data[12][0][26],
                    data[12][0][27],
                    ]
                    ),               #如果不实用，您认为主要原因是
        "",#您最希望得到哪方面的培训
        "",#如对培训工作有其他意见或建议，请说明
        multiSelect([
                    data[14][0][1],
                    data[14][0][2],
                    data[14][0][3],
                    ]
                    )        ,#作为灾害信息员的收入或补贴（月薪）情况
         multiSelect([
                    data[14][0][6],
                    data[14][0][7],
                    data[14][0][8],
                    ]
                    )  ,#是否有灾害信息员工作意外伤害保险
         multiSelect([
                    data[14][0][11],
                    data[14][0][12],
                    data[14][0][13],
                    data[14][0][14],
                    ]
                    )  ,         #是否有相应的奖励制度及表彰制度
         multiSelect([
                    data[14][0][17],
                    data[14][0][18],
                    data[14][0][19],
                    data[14][0][20],
                    data[14][0][21],
                    data[14][0][22],
                    data[14][0][23],
                    data[14][0][24],
                    data[14][0][25],
                    data[14][0][26],
                    ]
                    )  ,         #是否发放过个人防护用品（多选1）
         "",#多选2
         "",#其他
         multiSelect([
                    data[14][0][29],
                    data[14][0][30],
                    data[14][0][31],
                    data[14][0][32],
                    data[14][0][33],
                    data[14][0][34],
                    data[14][0][35],
                    data[14][0][36],
                    data[14][0][37],
                    ]
                    )  ,         #作为灾害信息员，您最急需的个人防护用品是（多选1）
         "",#多选2
         "",#多选3
         "",#多选4
         "",#多选5
         "",#其他
          replaceText(data[14][0][43].text),#您最急切的需要解决的问题是什么？
          simpleSelect(data[16][0][1]),#请问您是否能来参加座谈
          simpleSelect(data[18][0][7]),#是否推荐参加座谈会
          replaceText(data[18][0][0].text),
          filename.split("/")[-1]
     ))   

    #open("docx.xml", "a").write(filename)
    #open("docx.xml", "a").write(doc._element.xml)

if __name__ == '__main__':
    directoryName = "/Users/desktop/Downloads/docx20191129"
    #directoryName = "/Users/desktop/Downloads/docx_test"
    writeFileFlag = True
    if writeFileFlag:
        open("text.txt", "w").write("清理\n")
        open("docx.xml", "w").write("")
        open("error_file.txt","w").write("")
        open("result.csv","w").write("")
    
    for root, dirs, files in os.walk(directoryName):
        for fileno, f in enumerate(files):
            filename = os.path.join(root, f)
#             if fileno == 120:
#                 break
            print(filename)
            try:
                parseFile(filename)
            except Exception as e:
                traceback.print_exc()
                #sys.exit(-1-1)
                open("error_file.txt","a").write("error filename : %s %s\n"%(filename, str(e)))
        pass
#    pass
