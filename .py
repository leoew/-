# -*- coding: utf-8 -*-
import requests
import re
import xlwt


def login():
    login_url = 'http://202.202.1.176:8080/_data/index_login.aspx'

    form_data = {
        '__VIEWSTATEGENERATOR': 'CAA0A5A7',
        'Sel_Type': 'STU',
        'txt_dsdsdsdjkjkjc': 'user_id',
        'txt_dsdfdfgfouyy': 'password',
        'efdfdfuuyyuuckjg': '517663DE4AB8E30EB33FD301EAFD64'
    }

    headers_base = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.8',
        'Connection': 'keep-alive',
        'Host': '202.202.1.176:8080',
        'Referer': 'http://202.202.1.176:8080/_data/index_login.aspx',
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/49.0'
                      '.2623.108 Safari/537.36'
    }

    s = requests.session()
    r = s.get(login_url, headers=headers_base)

    login_html = r.text.decode('gb2312')
    pattern_viewstate = re.compile('<input type="hidden" name="__VIEWSTATE" value="(.*?)"', re.S)
    viewstate = re.findall(pattern_viewstate, login_html)[0]
    form_data['__VIEWSTATE'] = viewstate

    res = s.post(login_url, headers=headers_base, data=form_data)

    grade_url = 'http://202.202.1.176:8080/xscj/Stu_MyScore_rpt.aspx'
    grade_data = {
        'sel_xn': '2015',   # 选择学年
        'sel_xq': '0',  # 选择学期，0对应第一学期，1对应第二学期
        'SJ': '0',  # 0对应原始成绩，1对应有效成绩
        'SelXNXQ': '2',  # 2对应学期，1对应学年，0对应入学以来
        'zfx_flag': '0',  # 0对应主修，1对应辅修
    }
    res = s.post(grade_url, headers=headers_base, data=grade_data)
    return res.text.decode('gb2312')


content = login()
pattern_info = re.compile(u'学号：(.*?)&nbsp;&nbsp; 姓名：(.*?)</td>.*?学年学期：(.*?)</td></tr>', re.S)
info = re.findall(pattern_info, content)

id = info[0][0]
name = info[0][1]
xnxq = info[0][2]

pattern_grade = re.compile('<td width=4% .*?>(.*?)<.*?>'
                           '<td width=25% .*?>(.*?)<.*?'
                           '<td width=5% .*?>(.*?)<.*?'
                           '<td width=16% .*?>(.*?)<.*?'
                           '<td width=8% .*?>(.*?)<.*?'
                           '<td width=9% .*?>(.*?)<.*?'
                           '<td width=10% .*?>(.*?)<.*?'
                           '<td width=9% .*?>(.*?)<.*?'
                           '<td width=10% .*?>(.*?)<.*?'
                           '<td width=4% .*?>(.*?)<.*?', re.S)

grade = re.findall(pattern_grade, content)

f2 = xlwt.Workbook()
sheet1 = f2.add_sheet(u'sheet1', cell_overwrite_ok=True)
sheet1.write(0, 0, u'学号')
sheet1.write(0, 1, id)
sheet1.write(1, 0, u'姓名')
sheet1.write(1, 1, name)
sheet1.write(2, 0, u'学年学期')
sheet1.write(2, 1, xnxq)

row = 5
for gra in grade:
    for i in range(8):
        sheet1.write(row, i, gra[i])
    row += 1
f2.save('grade.xls')
