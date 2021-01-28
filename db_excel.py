# -*- coding: utf-8 -*-
'''
DB -> EXCEL
'''
import json
import pymysql
from openpyxl import load_workbook, Workbook

# db 연동
m_conn = pymysql.connect(host='112.217.167.123', port=61000, user='epai', password='epai', db='epai', charset='utf8')

# excel basic
write_wb = Workbook()
write_ws = write_wb.active

# select * from v_flu_wlv_hour where `시군구 명` ='보령시' and `관측소 명` like '%동대교%' order by `관측 시간`;

# 연도
year = 2019
# 위치
location = '대관령'
# 테이블 명
table = 'v_flu_gdwether_hour'

##### flag
if table == 'v_flu_aws_hour' or table == 'v_flu_brrer_hour' or table == 'v_flu_dam_hour' \
        or table == 'v_flu_flux_hour' or table == 'v_flu_gdwether_hour' or table == 'v_flu_rainfall_hour' \
        or table == 'v_flu_wlv_hour':
    info = {'time': '관측 시간', 'loc': '관측소 명'}
elif table == 'v_msr_dpmn':
    info = {'time': '연도', 'loc': '퇴적물측정망 명'}
elif table == 'v_msr_rsmn':
    info = {'time': '연도', 'loc': '방사성물질측정망 명'}
elif table == 'v_msr_wqmn_totqy':
    info = {'time': '년도', 'loc': '수질측정망 명'}

with m_conn.cursor() as curs:
    try:
        query1 = '''
            SELECT 
                COLUMN_NAME
            FROM
                INFORMATION_SCHEMA.columns
            WHERE
                    table_name = '{table}';
                '''.format(table=table)
        curs.execute("set names utf8")
        curs.execute(query1)
        rows0 = curs.fetchall()
        column_list = [row for row in rows0]
        m_conn.commit()
    finally:
        pass
aa = []
for i in column_list:
    aa.append(i[0])

write_ws.append(aa)

with m_conn.cursor() as curs:
    try:
        query = '''
            select * 
            from {table}
            where `{loc}` = '{location}' and `{time}` like '{year}%'
            order by `{time}` asc;
                '''.format(table=table, year=year, location=location, time=info['time'], loc=info['loc'])
        print(query)
        curs.execute("set names utf8")
        curs.execute(query)
        rows1 = curs.fetchall()
        data_list = [row for row in rows1]
        m_conn.commit()
    finally:
        pass

for i in data_list:
    data = []
    for k in range(len(data_list[0])):
        data.append(i[k])
    write_ws.append(data)

write_wb.save('/Users/jhy/Desktop/' + info['loc'] + '_' + location + '_' + str(year) + '_db.xlsx')
