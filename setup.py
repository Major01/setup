# _*_ coding:utf-8 _*_  
import xlwt 
import chardet 
import urllib.request, re  
  
def getdata():
    url_list = []
    for i in range(3001, 6175):
        url = 'http://furhr.com/?page={}'.format(i)
        print('抓取银行数据，数据来源：', url)

        try:
            html = urllib.request.urlopen(url).read()
            encode_type = chardet.detect(html)
            html = html.decode(encode_type['encoding'])
        except Exception as e:  
            print(e) 
            continue
        
        page_list = re.findall(r"<tr><td>(.*?)</td><td>(.*?)</td><td>(.*?)</td><td>(.*?)</td><td>(.*?)</td></tr>", html)
        url_list.append(page_list)

    return url_list

def excel_write(items):
    newTable = '银行基础信息1.xls'     
    wb = xlwt.Workbook(encoding='utf-8')  
    ws = wb.add_sheet('sheet1')  
    headdata = ['序号', '行号', '网点名称', '电话', '地址']  
    for colnum in range(0, 5):  
        ws.write(0, colnum, headdata[colnum], xlwt.easyxf('font:bold on'))  

    index = 1  
    for item in items:
        for j in range(0, len(item)):
            for i in range(0, 5):
                ws.write(index, i, item[j][i])
            index += 1
  
    wb.save(newTable)

if __name__ == '__main__':
    items = getdata()
    excel_write(items)