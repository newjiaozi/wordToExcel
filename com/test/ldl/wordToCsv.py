from docx import Document
import xlrd
from xlutils.copy import copy
import re
from os import linesep

csv_data = '20161107_流量_遍历_stb.xls'

def writeInExcel(*results):
    data = xlrd.open_workbook(csv_data)
    table = data.sheet_by_index(0)
    write_data = copy(data)
    for i in range(1,len(results)+1):
        # print(i,results)
        write_data.get_sheet(0).write(i,2,results[i-1].get('summary',''))
        write_data.get_sheet(0).write(i,22,results[i-1].get('action',''))
        write_data.get_sheet(0).write(i,23,results[i-1].get('expected_result',''))
    write_data.save(csv_data)



def getResults():
    results = []
    document = Document('7.0.基本功能-测试用例-20180228.docx')
    l = [paragraph.text for paragraph in document.paragraphs]
    stopFlag= 0
    result = {}
    # count = 0
    for i in l:
        # match_obj = re.match(r'0\d{1}\.\d{2}\..*',i)
        match_obj = re.match(r'0\d{1}\..*',i)
        if match_obj:
            print(match_obj.group())
            stopFlag = 0
        elif i.startswith('Summary:'):
            if result:
                # count += 1
                results.append(result)
                result = {}
            stopFlag = 1
            new_str = i.replace('Summary:','')
            result['summary'] = new_str+linesep
            print("******",repr(i))
        elif i.startswith('前提：') or i.startswith('前提:') or i.startswith('前提'):
            stopFlag = 2
            result['action'] = i+linesep
            # print(i,'--',stopFlag,result)
        elif i.startswith('Action：') or i.startswith('Action:') or i.startswith('Action'):
            if stopFlag == 2:
                result['action'] += i+linesep
            else:
                stopFlag = 2
                result['action'] = i+linesep
            # print(i,'--',stopFlag,result)
        elif i.startswith('Effect：') or i.startswith('Effect:') or i.startswith('Effect'):
            stopFlag = 3
            result['expected_result'] = ''
            # print(i,'--',stopFlag,result)

        elif stopFlag == 1:
            result['summary'] += i+linesep
            # print(i,'--',stopFlag,result)

        elif stopFlag == 2:
            result['action'] += i+linesep
            # print(i,'--',stopFlag,result)

        elif stopFlag == 3:
            if i:
                result['expected_result'] += i+linesep
            else:
                continue
        else:
            print('##:',i)
    results.append(result)
    return results





if __name__ == '__main__':
    results = getResults()
    # with open('res','w') as f:
    #     f.write(str(results))
    # print(len(results),results)
    writeInExcel(*results)