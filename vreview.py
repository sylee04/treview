# -*- coding: utf-8 -*-
"""
Created on Tue Jul 14 21:35:49 2020

@author: hana1602a
"""

# -*- coding: utf-8 -*-
"""
Created on Mon Mar 23 15:22:48 2020

@author: hana1602a
"""

# -*- coding: utf-8 -*-
"""
Created on Tue Dec 17 18:17:40 2019

@author: hana1602a
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Jul 18 14:51:27 2019

@author: hana1602a
"""

# -*- coding: utf-8 -*-
"""
Created on Tue Jun 25 20:17:38 2019

@author: hana1602a
"""

# -*- coding: utf-8 -*-
"""
Created on Tue Jun 25 18:03:06 2019

@author: hana1602a
"""

# -*- coding: utf-8 -*-
"""
Created on Fri Jan 18 16:00:04 2019

@author: hana1602a
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Dec 27 13:24:36 2018

@author: hana1602a
"""

# -*- coding: utf-8 -*-
"""
Created on Tue Oct 16 10:50:54 2018

@author: hana1602a
"""




import sqlite3
import os
#from pandas import Series, DataFrame
import pandas as pd
#import numpy as np
import easygui
import time
import re
import ctypes
import sys

#import time
start = time.time()

#startTime = time.strftime("%H_%M_%S")

서버컴여부 = 1


천단위콤마함수 = lambda x: format(x,'3,d')
천단위콤마함수2 = lambda x: format(int(x),'3,d')
소숫점둘째반올림함수 = lambda x: "%.2f" % x

toStringList = lambda L : [str(x) for x in L] #20200720
#toStringList(['a', 1]) Out[5]: ['a', '1']
#[str(x) for x in ['a',1]] Out[3]: ['a', '1']

def getfilesRev(dirpath):
    a = [s for s in os.listdir(dirpath) if os.path.isfile(os.path.join(dirpath, s))]
    #a.sort(key=lambda s: os.path.getmtime(os.path.join(dirpath, s)), reverse=True)
    a.sort(reverse=True)
    return a



checkDic = {}

if 1 : #팝업창 띄워서 엑셀화일선텍
    # 팝업창(메시지박스) 띄우기
    def Mbox(title, text, style):
        return ctypes.windll.user32.MessageBoxW(0, text, title, style)
    
    Mbox('leecta', '세무사랑 [매입매출전표(전체)] 엑셀자료를 선택하세요', 1)    
        
    dirFilenameExtension = easygui.fileopenbox()
else : #테스트할때
    dirFilenameExtension = '\\\\As5104t-9c9e\\업무폴더-201506~\\매입매출검토\\library\\해오라기\\기장산꼼장어_매입매출전표(전체)_20180701_20180831-test.xls'

startTime = time.strftime("%H_%M_%S")

dir1 = os.path.split(dirFilenameExtension)[0]
base1=os.path.basename(dirFilenameExtension)
filename1 = os.path.splitext(base1)[0]
extension2 = os.path.splitext(base1)[1]

filenameSplit = re.split('_|{|}|-', filename1)

if '매입매출전표(전체)' not in filenameSplit:
    Mbox('leecta', filename1+' ==> [세무사랑pro-매입매출전표입력-데이터변환(엑셀-전체)]가 아닙니다', 1)
    sys.exit()

기산일인덱스 = 0    
for i in range(len(filenameSplit)):
    if filenameSplit[i].isdigit() and len(filenameSplit[i]) == 8:
        기산일인덱스 = i
        break
if 기산일인덱스 == 0 :
    Mbox('leecta', filename1+' ==> 화일명에서 과세기간을 추출할수가 없습니다. 화일명을 수정하지마세요.', 1)
    sys.exit()
        
        
기초스트링 = filenameSplit[기산일인덱스]    
기말스트링 = filenameSplit[기산일인덱스+1]

기초일 = pd.to_datetime(기초스트링, format='%Y%m%d')
기말일 = pd.to_datetime(기말스트링, format='%Y%m%d')

과세기간 = str(기초일.year)+'년'+str(기초일.month)+'월'+str(기초일.day)+'일 ~ '+str(기말일.year)+'년'+str(기말일.month)+'월'+str(기말일.day)+'일'
회사명 = filenameSplit[0]
    
#엑셀자료위치 = 'E:\\TEST2018\\수원댁밥상_매입매출전표(전체)_20180101_20180630.xls'
#xls_file = pd.ExcelFile(엑셀자료위치)
xls_file = pd.ExcelFile(dirFilenameExtension)
df = xls_file.parse('2.매입매출')

df2 = xls_file.parse('5.거래처')
df20 = df2[['거래처코드','검색번호']]

df9 = xls_file.parse('6.전자세금계산서')

df13 = xls_file.parse('3.분개')
df131 = df13[df13['고정자산코드'].isin([1])]
고정자산공급가액 = df131['금액'].sum()
고정자산갯수 = df131['금액'].count()

df132 = df131[['매입매출키번호','고정자산코드']].set_index('매입매출키번호')
    
df17 = xls_file.parse('1.정보')
#회사정보사전 = df17.ix[0].to_dict() #{'사업자등록번호': '124-52-44924', '회사명': '수원댁밥상', 'Unnamed: 3': 1.1000000000000001, '매입매출전표전송': 'A1020001'}
회사정보사전 = df17.iloc[0].to_dict() #{'사업자등록번호': '124-52-44924', '회사명': '수원댁밥상', 'Unnamed: 3': 1.1000000000000001, '매입매출전표전송': 'A1020001'}


#엑셀자료위치2 = 'E:\\TEST2018\\검토사전엑셀.xls'
#엑셀자료위치2 = 'E:\\TEST2018\\검토사전엑셀.xlsx'
if 서버컴여부 == 1:
    #엑셀자료위치2 = 'E:\\TEST2018\\검토사전엑셀.xlsx'
    엑셀자료위치2 = '\\\\As5104t-9c9e\\업무폴더-201506~\\매입매출검토\\library\\검토사전엑셀.xlsx'
    엑셀자료위치2 = '\\\\Desktop-ekk32ts\\업무폴더-201912\\매입매출검토\\library\\검토사전엑셀.xlsx' #20191217
else : 엑셀자료위치2 = 'D:\\TEST2018\\검토사전엑셀.xlsx'
    
xls_file2 = pd.ExcelFile(엑셀자료위치2)
#데이터프레임2 = xls_file2.parse('Sheet1')
데이터프레임2 = xls_file2.parse('검토단어')
검토사전2 = 데이터프레임2.set_index('대상단어')['검토멘트'].to_dict()

ipchul = { '부가세\n유형' : list(range(11,25))+list(range(51,63)),'유형명' : ['과세', '영세', '면세', '건별', '간이', '수출', '카과', '카면', '카영', '면건', '전자', '현과', '현면', '현영', '과세', '영세', '면세', '불공', '수입', '금전', '카과', '카면', '카영', '면건', '현과', '현면'], '입출' : ['1.매출'] * len(range(11,25)) + ['2.매입'] * len(range(51,63)), '입출2' : ['매출'] * len(range(11,25)) + ['매입'] * len(range(51,63)), '카운트' : [1] * 26}
입출프레임 = pd.DataFrame(ipchul)

df3 = pd.merge(df, df20, left_on='거래처\n코드', right_on='거래처코드', how='left')
df31 = pd.merge(df3, 입출프레임, on='부가세\n유형', how='left')
df4 = df31.fillna(0)

#ser1 = xls_file2.parse('불공추정단어')
df15 = xls_file2.parse('불공추정단어')
#불공추정단어사전 = df15.to_dict()
#불공추정단어리스트 = list(df15.values)
불공추정단어리스트 = df15.추정단어.values.tolist()

#181007
if '서버컴여부 == 1':
    df16 = xls_file2.parse('기본매입업체') #신규개업 사업자번호 등록안한 업체 추정
    기본매입업체리스트 = df16.상호5.values.tolist()
    기본매입업체사전 = df16.set_index('상호5')['분류'].to_dict()


dff3 = xls_file2.parse('공통매입세액(마트)')
공통매입대상사업자등록번호리스트 = dff3.사업자등록번호2.values.tolist()

df4002 = df4[df4['검색번호'].isin(공통매입대상사업자등록번호리스트)]
if not df4002.empty:
    공통매입대상공급가액합계 = df4002['공급가액'].sum()
    공통매입대상부가세합계 = df4002['부가세'].sum()
    대상업체리스트 = df4002['거래처명'].drop_duplicates().apply(str).tolist()
    대상업체 = ', '.join(대상업체리스트)

#181007
dff4 = xls_file2.parse('접수증')
dff5 =xls_file2.parse('부가율')
#dff5['부가율2'] = dff5['부가율'].str.replace('%','')
dff5['부가율2'] = pd.to_numeric(dff5['부가율'].str.replace('%',''), errors = 'coerce')
dff5['주업종코드'] = dff5['주업종'].str[:8]
#grouped = dff5.groupby(['주업종코드','주업종'])['부가율2'].agg(['count','max','min','mean','median'])
grouped = dff5.groupby(['주업종코드'])['부가율2'].agg(['count','max','min','mean','median'])
grouped.columns = ['표본수','최대값','최소값','평균값','중간값']

if 회사정보사전['사업자등록번호'] in dff5.values:
    이회사의주업종코드 = dff5[dff5['사업자등록번호'].isin([회사정보사전['사업자등록번호']])].주업종코드.iloc[0]
    이회사의주업종 = dff5[dff5['사업자등록번호'].isin([회사정보사전['사업자등록번호']])].주업종.iloc[0]
    이회사의평균부가율등 = grouped[grouped.index.isin([이회사의주업종코드])]
    checkDic['2.4.1.2.평균부가율등'] = '▲ 주업종 : '+이회사의주업종+' ▲ 표본수 : '+str(이회사의평균부가율등['표본수'][0])+'개 ▲ 최대값 : '+소숫점둘째반올림함수(이회사의평균부가율등['최대값'][0])+'% ▲ 최소값 : '+소숫점둘째반올림함수(이회사의평균부가율등['최소값'][0])+'% ▲ 평균값 : '+소숫점둘째반올림함수(이회사의평균부가율등['평균값'][0])+'% ▲ 중간값 : '+소숫점둘째반올림함수(이회사의평균부가율등['중간값'][0])+'%'
else : 
    checkDic['2.4.1.2.평균부가율등'] = '조회내역이 없습니다'
                         
dff6 = xls_file2.parse('예정고지세액')
                     
dff7 = xls_file2.parse('총괄1',skiprows=5) #20200714
dff71 = xls_file2.parse('총괄1')
하나세무과세기간 = dff71.iloc[3,0]
semuincome7 = dff7.set_index('사업자번호')[['매수','공급가액']].to_dict()

기장료크로스리스트 = []
#기장료크로스리스트.append(하나세무과세기간)
if 회사정보사전['사업자등록번호'] in semuincome7['공급가액'] : #기장료 없을경우
    기장료크로스리스트.append(semuincome7['매수'][회사정보사전['사업자등록번호']])
    기장료크로스리스트.append(semuincome7['공급가액'][회사정보사전['사업자등록번호']])
#semuincome7['매수']['101-32-51254'] Out[13]: 5 semuincome7['공급가액']['101-32-51254'] Out[14]: 250000

#checkDic = {}
#checkList = [] # 불공추정단어리스트

#매출리스트
매출리스트 = list(range(11,25))
매출과세리스트 = [11,12,14,15,16,17,19,22,24]
매출면세리스트 = [13,18,20,23]
매출신카등리스트 = [17,18,19,22,23,24]
매출현금리스트 = [14,15,20]
매출신카등과세리스트 = [17,19,22,24]
매출신카등면세리스트 = [18,23]
세액있는매출리스트 = [11,14,15,17,21,22] #for 매출세액
매출세금계산서등리스트 = [11,12,13,16]

#매입리스트
매입리스트 = list(range(51,63))
매입과세리스트 = [51,52,54,55,56,57,59,61]
매입과세불공제외리스트 = [51,52,55,56,57,59,61]
매입면세리스트 = [53,58,60,62]
매입불공제리스트 = [54]
매입신카등리스트 = list(range(56,63))
매입세금계산서등리스트 = list(range(51,56))

#20200323
매입신카리스트1 = [57, 58]
매입현영리스트1 = [61, 62]
매입신카공급가액1 = df4[df4['부가세\n유형'].isin(매입신카리스트1)].공급가액.sum()
매입현영공급가액1 = df4[df4['부가세\n유형'].isin(매입현영리스트1)].공급가액.sum()



#a = [x for x in a if x != 20] Is there a simple way to delete a list element by value?
발송용부가세유형리스트 = [x for x in list(range(11,25))+list(range(51,63)) if x not in 매출현금리스트+매입신카등리스트]
발송용부가세유형리스트2 = [11, 12, 13, 16, 17, 18, 19, 21, 22, 23, 24, 51, 52, 53, 54, 55]
발송용부가세유형리스트3 = [11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 51, 52, 53, 54, 55] # 면건등 현금매출 반영
               
매출과세공급가액 = df4[df4['부가세\n유형'].isin(매출과세리스트)].공급가액.sum()
매입과세공급가액 = df4[df4['부가세\n유형'].isin(매입과세리스트)].공급가액.sum()
세액있는매출공급가액 = df4[df4['부가세\n유형'].isin(세액있는매출리스트)].공급가액.sum()

매출공급가액 = df4[df4['부가세\n유형'].isin(매출리스트)].공급가액.sum()
매출면세공급가액 = df4[df4['부가세\n유형'].isin(매출면세리스트)].공급가액.sum()

매입신카등공급가액 = df4[df4['부가세\n유형'].isin(매입신카등리스트)].공급가액.sum()
매입공급가액 = df4[df4['부가세\n유형'].isin(매입리스트)].공급가액.sum()

#매출세액 = 매출과세공급가액*0.1
#매출세액 = df4[df4['부가세\n유형'].isin(매출리스트)].부가세.sum()
매출세액 = 세액있는매출공급가액*0.1

의제매입세액한도율506065기산일 = pd.to_datetime('20180701', format='%Y%m%d') #20190118 의제매입한도율 2018년 2기부터 2019년까지 50% 60% 65% 적용
의제매입세액한도율506065만료일 = pd.to_datetime('20191231', format='%Y%m%d')

'''if 매출과세공급가액 > 200000000 :
    의제매입한도 = 매출과세공급가액*0.45
    적용된의제매입한도율 = 0.45
elif 매출과세공급가액 > 100000000 :
    의제매입한도 = 매출과세공급가액*0.55
    적용된의제매입한도율 = 0.55
else :
    의제매입한도 = 매출과세공급가액*0.6
    적용된의제매입한도율 = 0.6
'''
if 기말일 >= 의제매입세액한도율506065기산일 and 기말일 <= 의제매입세액한도율506065만료일: #20190118
    if 매출과세공급가액 > 200000000 :
        의제매입한도 = 매출과세공급가액*0.5
        적용된의제매입한도율 = 0.5
    elif 매출과세공급가액 > 100000000 :
        의제매입한도 = 매출과세공급가액*0.6
        적용된의제매입한도율 = 0.6
    else :
        의제매입한도 = 매출과세공급가액*0.65
        적용된의제매입한도율 = 0.65
else :
    if 매출과세공급가액 > 200000000 :
        의제매입한도 = 매출과세공급가액*0.45
        적용된의제매입한도율 = 0.45
    elif 매출과세공급가액 > 100000000 :
        의제매입한도 = 매출과세공급가액*0.55
        적용된의제매입한도율 = 0.55
    else :
        의제매입한도 = 매출과세공급가액*0.6
        적용된의제매입한도율 = 0.6
        
        
의제류대상금액 = df4['의제류\n대상금액'].sum()

if 의제류대상금액 > 의제매입한도 :
    의제매입한도적용후가액 = 의제매입한도
    의제한도멘트 = '의제류대상금액이 '+str(천단위콤마함수(int(의제류대상금액-의제매입한도)))+'원'+' 만큼 초과했습니다'
else :
    의제매입한도적용후가액 = 의제류대상금액
    의제한도멘트 = '의제류대상금액이 '+str(천단위콤마함수(int(의제매입한도-의제류대상금액)))+'원'+' 만큼 여유있습니다'


공제율상향년도 = [2018, 2019, 2020, 2021] #20200323 수정함

if 매출과세공급가액 <= 200000000 and 기말일.year in 공제율상향년도 :
    의제매입세액공제율은 = '109분의 9'
    의제매입세액공제율 = 9/109
else :
    의제매입세액공제율은 = '108분의 8'
    의제매입세액공제율 = 8/108
    
    
    
한도적용후의제매입공제세액 = 의제매입한도적용후가액*의제매입세액공제율

#매입세액 = df4[df4['부가세\n유형'].isin(매입리스트)].부가세.sum() + df4['의제류\n공제세액'].sum()
매입세액 = df4[df4['부가세\n유형'].isin(매입리스트)].부가세.sum() + 한도적용후의제매입공제세액 - df4[df4['부가세\n유형'].isin(매입불공제리스트)].부가세.sum()

신카발행세액공제 = df4[df4['부가세\n유형'].isin(매출신카등과세리스트)].합계금액.sum()*0.013
               
기준일 = pd.to_datetime('20180701', format='%Y%m%d') # 20181227 201901신고분부터 신카발행세액공제한도가 1천만원으로 상향
if 기말일 >= 기준일 and 신카발행세액공제 > 10000000: # 201901포함 이후신고분이라면
    신카발행세액공제 = 10000000
elif 기말일 < 기준일 and 신카발행세액공제 > 5000000:
    신카발행세액공제 = 5000000

if 회사정보사전['사업자등록번호'][4] in ['8']: #우리 거래처가 법인 : 8
    신카발행세액공제 = 0                    
               
'''if 회사정보사전['사업자등록번호'][4] in ['8']: #우리 거래처가 법인 : 8
    신카발행세액공제 = 0                
elif 신카발행세액공제 > 5000000:
    신카발행세액공제 = 5000000
#else : pass
'''


#신카세액공제전납부세액 = 매출세액 - 매입세액

if 매출면세공급가액 > 0 and not df4002.empty:
#    if not df4002.empty: #20190720
#        checkDic['9.1.공통매입'] = 'ㅇ공통매입공급가액 : '+천단위콤마함수2(공통매입대상공급가액합계)+' ㅇ공통매입세액 : '+천단위콤마함수2(공통매입대상부가세합계)+' ㅇ대상업체: '+대상업체
#    else:
#        checkDic['9.1.공통매입'] = 'N/A' #20190720 한전, 케이티등 공통매입안분대상 없체가 없다면
    checkDic['9.1.공통매입'] = 'ㅇ공통매입공급가액 : '+천단위콤마함수2(공통매입대상공급가액합계)+' ㅇ공통매입세액 : '+천단위콤마함수2(공통매입대상부가세합계)+' ㅇ대상업체: '+대상업체
    #checkDic['9.1.공통매입'] = 'N/A' #20190720 한전, 케이티등 공통매입안분대상 없체가 없다면
    면세공급가액비율 = 매출면세공급가액/매출공급가액
    불공제매입세액 = 공통매입대상부가세합계*면세공급가액비율
    checkDic['9.2.불공제매입세액'] = 'ㅇ면세비율 : '+소숫점둘째반올림함수(면세공급가액비율*100)+'%'+' ㅇ불공제매입세액 : '+천단위콤마함수2(불공제매입세액)
    #신카세액공제전납부세액 = 신카세액공제전납부세액 + 불공제매입세액
    #불공제미반영매입세액 = 매입세액 #20200323
    매입세액 = 매입세액 - 불공제매입세액
else : #20190720 매출면세공급가액이 없거나; ★한전, 케이티등 공통매입안분대상 없체가 없다면
    checkDic['9.1.공통매입'] = 'N/A'
    checkDic['9.2.불공제매입세액'] = 'N/A'

신카세액공제전납부세액 = 매출세액 - 매입세액    
    
    
if 신카세액공제전납부세액 < 신카발행세액공제 and 신카세액공제전납부세액 > 0:
    사중손실세액 = 신카발행세액공제 - 신카세액공제전납부세액
elif 신카세액공제전납부세액 <= 0:
    사중손실세액 = 신카발행세액공제
else : 사중손실세액 = 0

if 신카발행세액공제 > 0:
    checkDic['3.1.신카 사중손실세액'] = 천단위콤마함수(int(사중손실세액))+'원'
    checkDic['3.2.실제 공제된세액'] = 천단위콤마함수(int(신카발행세액공제 - 사중손실세액))+'원'
else :
    checkDic['3.1.신카 사중손실세액'] = 'N/A'
    checkDic['3.2.실제 공제된세액'] =  'N/A'
    

checkDic['3.0.신카발행공제세액'] = 천단위콤마함수(int(신카발행세액공제))+'원'

if 기말일 < 기준일:
    if 신카발행세액공제 - 사중손실세액 > 2500000:
        checkDic['3.4.세액공제한도체크'] = '반영된 신카발행공제세액이 250만원을 초과하였습니다(1년 500만원)'
    else :
        checkDic['3.4.세액공제한도체크'] = 'N/A'
elif 기말일 >= 기준일: #201901포함 이후신고분
    if 신카발행세액공제 - 사중손실세액 > 5000000:
        checkDic['3.4.세액공제한도체크'] = '반영된 신카발행공제세액이 500만원을 초과하였습니다(1년 1,000만원)'
    else :
        checkDic['3.4.세액공제한도체크'] = 'N/A'    
    
'''if 신카발행세액공제 - 사중손실세액 > 2500000:
    checkDic['3.4.세액공제한도체크'] = '반영된 신카발행공제세액이 250만원을 초과하였습니다(1년500만원)'
else :
    checkDic['3.4.세액공제한도체크'] = 'N/A'
'''

if 매출과세공급가액 > 500000000: # 20181227
    checkDic['3.5.세액공제제외체크'] = '당기 부가세과세표준이 5억을 초과하였습니다. 직전연도 공급가액이 10억이상이면 신용카드세액공제 대상에서 제외됩니다.'


if 신카세액공제전납부세액 > 신카발행세액공제:
#    납부세액 = 매출세액 - 매입세액 - 신카발행세액공제
    납부세액 =  신카세액공제전납부세액 - 신카발행세액공제
elif 신카세액공제전납부세액 > 0 :
#    납부세액 = 매출세액 - 매입세액 - (신카발행세액공제 - 사중손실세액)
    납부세액 = 신카세액공제전납부세액 - (신카발행세액공제 - 사중손실세액)
#else : 납부세액 = 매출세액 - 매입세액
else : 납부세액 = 신카세액공제전납부세액

checkDic['2.3.납부세액㉰'] = 천단위콤마함수(int(신카세액공제전납부세액))+'원'
               
#납부세액 = 매출세액 - 매입세액 - 신카발행세액공제
if 매출과세공급가액 > 0:
    부가율 = (매출과세공급가액 - (매입과세공급가액 + 의제류대상금액 - 고정자산공급가액))/매출과세공급가액*100
else : 부가율 = 0

#print('===================================================================================')
#checkDic['# 회사명 '] = 회사명
checkDic['# 회사명 '] = 회사정보사전['회사명']
checkDic['# 사업자번호 '] = 회사정보사전['사업자등록번호']
checkDic['# 과세기간 '] = 과세기간

#print('▲ 납부세액㉰: ', 천단위콤마함수(int(납부세액)))
checkDic['1.0.납부세액'] = 천단위콤마함수(int(납부세액))+'원'

if 매출면세공급가액 > 0 and not df4002.empty:
    checkDic['1.2.납부세액(공통불공미반영)'] = 천단위콤마함수(int(납부세액-불공제매입세액))+'원' #20200323

#if isinstance(조회된예정고지세액.iloc[0],int):
#    checkDic['1.3.납부세액(예정고지차감)'] = 천단위콤마함수(int(납부세액-조회된예정고지세액.iloc[0]))+'원' #20200323
             

#print('▲ 현매출기준 2018년~2019년 개인음식점 의제매입세액공제율 : ', 의제매입세액공제율은)
#checkDic['적용된 의제매입세액공제율'] = 의제매입세액공제율은
#checkDic['5.적용 의제공제율'] = 의제매입세액공제율은

if 기말일.year in 공제율상향년도:
    checkDic['5.0.매출기준 공제율'] = 의제매입세액공제율은
else :
    checkDic['5.0.매출기준 공제율'] = ' 8/108 '

#print('▲', 의제한도멘트)
checkDic['4.0.의제한도멘트'] = 의제한도멘트

if 의제류대상금액 > 0 :
    checkDic['4.1.의제매입현황'] = '▲ 의제류대상금액 : '+str(천단위콤마함수(int(의제류대상금액)))+'원\n'+'▲ 의제매입한도 : '+str(천단위콤마함수(int(의제매입한도)))+'원\n'+'▲ 의제매입한도적용후가액 : '+str(천단위콤마함수(int(의제매입한도적용후가액)))+'원\n'+'▲ 한도적용후의제매입공제세액 : '+str(천단위콤마함수(int(한도적용후의제매입공제세액)))+'원'
else :
    checkDic['4.1.의제매입현황'] = 'N/A'

#print('▲ 부가율 : ', 소숫점둘째반올림함수(부가율),'%')
checkDic['2.0.부가율'] = 소숫점둘째반올림함수(부가율)+'%'

카과현과매출 = df4[df4['부가세\n유형'].isin([17,22])].공급가액.sum()
건별매출 = df4[df4['부가세\n유형'].isin([14,15])].공급가액.sum()
if 카과현과매출 > 0:
    현금과세비율 = 건별매출/카과현과매출*100
    checkDic['2.1.현금과세비율'] = 소숫점둘째반올림함수(현금과세비율)+'%'
else :
    checkDic['2.1.현금과세비율'] = 'N/A'

#전체매출대비면세매입비율 = (df4[df4['부가세\n유형'].isin([53])].공급가액.sum())/(df4[df4['부가세\n유형'].isin(매출리스트)].공급가액.sum())
if 의제류대상금액 == 0:
    if 매출공급가액 > 0:
        전체매출대비면세매입비율 = (df4[df4['부가세\n유형'].isin([53])].공급가액.sum())/(매출공급가액)
        if 전체매출대비면세매입비율 > 0.01:
            #if 매출면세공급가액 > 0 and 전체매출대비면세매입비율 > 0.01:
            if 매출면세공급가액 > 0 :
                #면세매출마진율 = (df4[df4['부가세\n유형'].isin(매출면세리스트)].공급가액.sum()-df4[df4['부가세\n유형'].isin([53])].공급가액.sum())/df4[df4['부가세\n유형'].isin(매출면세리스트)].공급가액.sum()
                매출면세마진율 = (매출면세공급가액 - df4[df4['부가세\n유형'].isin([53])].공급가액.sum())/(매출면세공급가액)
                checkDic['2.2.면세마진율'] = 소숫점둘째반올림함수(매출면세마진율*100)+'%'
            #elif 전체매출대비면세매입비율 > 0.01 and df4[df4['부가세\n유형'].isin([53])].공급가액.sum() > 500000:
            elif df4[df4['부가세\n유형'].isin([53])].공급가액.sum() > 500000:
                checkDic['2.2.면세마진율'] = '면세매출 체크하세요'
else :
    checkDic['2.2.면세마진율'] = 'N/A'
    checkDic['2.2.면세마진율'] = 'N/A'
           
           
checkDic['7.고정자산갯수'] = 천단위콤마함수(int(고정자산갯수))+'개'

거래처명코드유형 = df4[["거래처명","거래처\n코드","부가세\n유형"]].drop_duplicates()
거래처명코드유형2 = 거래처명코드유형.reset_index(drop=True)
df5 = df4.drop_duplicates(["거래처명","거래처\n코드","부가세\n유형"])

def f(x):
    j = ''
    for i in x:
        j = j + i
    return j

df6 = df4.copy()

def 전자대체(x):
    if x == 1:
        return 'e'
#df6['전자여부2'] = df6[df6['전자여부'].isin([1])]
df6['전자여부2'] = df6['전자\n여부'].apply(전자대체)
df6['전자여부2'] = df6['전자여부2'].fillna(' ')
df6['전표일자및공급가액'] =  ' 【' + df4['전표일자'].apply(str).str[4:6]+'/'+df4['전표일자'].apply(str).str[6:8]+'】 '+ df4['공급가액'].apply(천단위콤마함수2)

df6['Time'] = pd.to_datetime(df6['전표일자'], format='%Y%m%d')
#df6['Date'] = pd.to_datetime(df6['전표일자'])
df6['월'] = df6.Time.dt.month
의제류공제율분자위치 = df6.columns.tolist().index('의제류\n공제율')
의제류공제율분모 = df6.columns.tolist()[의제류공제율분자위치+1] #엑셀에서 분모는 분자 우측에 있음
#df6['의제공제율'] = df6['의제류\n공제율'].apply(int).apply(str) + '/' + df6['Unnamed: 45'].apply(int).apply(str)+' '
df6['의제공제율'] = ' '+df6['의제류\n공제율'].apply(int).apply(str) + '/' + df6[의제류공제율분모].apply(int).apply(str)+' '
df6['의제공제율'] = df6[~df6['의제공제율'].isin([' 0/0 '])].의제공제율
df6['의제공제율'] = df6['의제공제율'].fillna(' ')
df91 = df9.set_index(['매입매출키번호'])
df6 = df6.join(df91, on='키번호')
df6 = df6.join(df132, on='키번호')


df6 = df6.fillna(0)
#df6.dtypes

dff1 = df6.copy() # VAT검토표용(프린트시 프린트날짜 인쇄되도록 할것)
#dff1['전표일자2'] = dff1.Time.dt.strftime('%y년 %m월 %d일')
dff1['전표일자2'] = dff1.Time.dt.strftime('%y-%m-%d')
VAT검토표용리스트 = ['전표일자2','입출', '부가세\n유형','매출매입\n구분','검색번호','거래처명','공급가액','부가세','합계금액','의제공제율']
VAT검토표용부가세유형리스트 = 매출세금계산서등리스트 + 매입세금계산서등리스트
dff2 = dff1[(~dff1['전자\n여부'].isin([1]))&(dff1['부가세\n유형'].isin(VAT검토표용부가세유형리스트))][VAT검토표용리스트]


df6 = df6.sort_values(by='전표일자') #일자별상세내역이 날짜역순도 발견되길래

if df6[df6['부가세\n유형'].isin([17])].공급가액.sum() > 0 and df6[df6['부가세\n유형'].isin([17])].부가세.sum() == 0:
    간이여부 = 1
else : 간이여부 = 0

if 간이여부 == 1:
    checkDic['2.3.1.간이여부'] = '간이로 추정되며 표시된 납부세액은 일반과세자기준입니다.'
else :
    checkDic['2.3.1.간이여부'] = 'N/A'


#grouped12 = df6.groupby(["거래처명","거래처\n코드","부가세\n유형"])['전표일자및공급가액'].agg([('일자별상세내역',f)])
#grouped13 = df6.groupby(['입출','부가세\n유형','매출매입\n구분','거래처명','검색번호'])['전표일자및공급가액'].agg([('일자별상세내역',f)])
grouped13 = df6.groupby(['입출','부가세\n유형','매출매입\n구분','의제공제율','거래처명','검색번호'])['전표일자및공급가액'].agg([('일자별상세내역',f)])
#grouped14 = df6.groupby(['입출','부가세\n유형','매출매입\n구분','거래처명','검색번호'])['카운트','전자\n여부','공급가액','부가세'].sum()
grouped14 = df6.groupby(['입출','부가세\n유형','매출매입\n구분','의제공제율','거래처명','검색번호'])['카운트','전자\n여부','공급가액','부가세'].sum()
concated = pd.concat([grouped14, grouped13], axis=1)
concated1 = concated.copy()

send10 = df6[df6['부가세\n유형'].isin(발송용부가세유형리스트3)]
send11 = send10.groupby(['입출','매출매입\n구분','거래처명','검색번호'])['전표일자및공급가액'].agg([('일자별상세내역',f)])
send12 = send10.groupby(['입출','매출매입\n구분','거래처명','검색번호'])['카운트','공급가액'].sum()
send1 = pd.concat([send12, send11], axis=1)

groupTotal1 = concated1.sum(level=['입출','부가세\n유형','매출매입\n구분','의제공제율'])
if '1.매출' in groupTotal1.index:
    groupTotal1.loc[('매출',' ','Total',' '),:] = groupTotal1.loc['1.매출'].sum()
if '2.매입' in groupTotal1.index:
    groupTotal1.loc[('매입',' ','Total',' '),:] = groupTotal1.loc['2.매입'].sum()

#업체매입매출장기장료 = concated.reset_index().set_index(['검색번호']).loc['317-02-14202','카운트':'공급가액']    
concated2 = concated.reset_index().set_index(['검색번호'])
if '317-02-14202' in concated2.index :
    #업체매입매출장기장료 = concated2.loc['317-02-14202','카운트':'공급가액']    
    업체매입매출장기장료 = concated2.loc['317-02-14202']    
    기장료크로스리스트.append(업체매입매출장기장료['카운트'])
    기장료크로스리스트.append(업체매입매출장기장료['공급가액'])
기장료크로스리스트2 = [천단위콤마함수2(s) for s in 기장료크로스리스트]
#기장료크로스리스트2 = [str(s) for s in 기장료크로스리스트]
기장료크로스리스트2.append(하나세무과세기간)
checkDic['9.6.1. 기장료 크로스체크'] = ' // '.join(기장료크로스리스트2)
    

신카매출크로스리스트 = []
dff72 = xls_file2.parse('신카매출합계')
dff73 = xls_file2.parse('현영매출합계')
#if 회사정보사전['사업자등록번호'] in semuincome7['공급가액']:
#if 회사정보사전['사업자등록번호'] in dff72['사업자등록번호'].tolist() :
if 회사정보사전['사업자등록번호'] in dff72['사업자등록번호'].values :
    신카매출크로스리스트.append(dff72[dff72['사업자등록번호'].isin([회사정보사전['사업자등록번호']])]['count'])
    신카매출크로스리스트.append(dff72[dff72['사업자등록번호'].isin([회사정보사전['사업자등록번호']])]['매출액계'])
concated3  = concated.reset_index().groupby(['부가세\n유형'])[['카운트', '전자\n여부', '공급가액','부가세']].sum()
#신카매출크로스리스트.append(concated3.loc[17:20]['카운트'].sum()) #20200724 20.면건이 포함되에 저장된다 why? slice가 원래 이런가? iloc이 아니라 loc이라서 그런가?
신카매출크로스리스트.append(concated3.loc[17:19]['카운트'].sum()) #20200724
#신카매출크로스리스트.append(concated3.loc[17:20]['공급가액'].sum()*1.1)
#신카매출크로스리스트.append(concated3.loc[17:20]['공급가액'].sum()+concated3.loc[17:20]['부가세'].sum()) #20200724
신카매출크로스리스트.append(concated3.loc[17:19]['공급가액'].sum()+concated3.loc[17:19]['부가세'].sum()) #20200724
#if len(신카매출크로스리스트) > 0 : #20200720
신카매출크로스리스트2 = [천단위콤마함수2(s) for s in 신카매출크로스리스트]
checkDic['9.6.2. 신카매출 크로스체크'] = ' // '.join(신카매출크로스리스트2)

현영매출크로스리스트 = []    
if 회사정보사전['사업자등록번호'] in dff73['사업자등록번호'].values :
    현영매출크로스리스트.append(dff73[dff73['사업자등록번호'].isin([회사정보사전['사업자등록번호']])]['count'])
    현영매출크로스리스트.append(dff73[dff73['사업자등록번호'].isin([회사정보사전['사업자등록번호']])]['총금액'])
concated3  = concated.reset_index().groupby(['부가세\n유형'])[['카운트', '전자\n여부', '공급가액','부가세']].sum()
현영매출크로스리스트.append(concated3.loc[22:25]['카운트'].sum())
#현영매출크로스리스트.append(concated3.loc[22:25]['공급가액'].sum()*1.1)
현영매출크로스리스트.append(concated3.loc[22:25]['공급가액'].sum()+concated3.loc[22:25]['부가세'].sum())
#if len(현영매출크로스리스트) > 0 : #현영매출크로스리스트 Out[21]: [0, 0]
현영매출크로스리스트2 = [천단위콤마함수2(s) for s in 현영매출크로스리스트]
checkDic['9.6.3. 현영매출 크로스체크'] = ' // '.join(현영매출크로스리스트2)    

if '1.매출' in groupTotal1.index: #20200719 소규모개인사업자부가가치세감면신청
    if groupTotal1.loc['1.매출']['공급가액'].sum() <= 40000000 :
        #checkDic['9.6.4. 소규모..감면신청 체크'] = 천단위콤마함수2(groupTotal1.loc['1.매출']['공급가액'].sum())+' ← 6개월 공급가액합계가 4천만원이하인지 검토하세요('+천단위콤마함수2(groupTotal1.loc['1.매출']['공급가액'].sum()*1.1)+')'
        checkDic['9.6.4. 소규모..감면신청 체크'] = 천단위콤마함수2(groupTotal1.loc['1.매출']['공급가액'].sum())+' ← 6개월 공급가액합계가 4천만원이하인지 검토하세요('+천단위콤마함수2(groupTotal1.loc['1.매출']['공급가액'].sum()+groupTotal1.loc['1.매출']['부가세'].sum())+')' #20200725 간이과세자 오해소지
    
    
if 10 : #20200714 asdf
    
    if 서버컴여부 == 1:
        #writer = pd.ExcelWriter(os.path.join('E:\\TEST2018\\세무사랑검토',filename1+'{당기전기'+time.strftime("%Y%m%d_%H%M%S")+'}'+'.xlsx'))
        #writer = pd.ExcelWriter(os.path.join('\\\\As5104t-9c9e\\업무폴더-201506~\\매입매출검토',filename1+'{당기전기검토'+time.strftime("%Y%m%d_%H%M%S")+'}'+'.xlsx'))
        #writer = pd.ExcelWriter(os.path.join('\\\\As5104t-9c9e\\업무폴더-201506~\\매입매출검토',filename1+'{부가세검토'+time.strftime("%Y%m%d_%H%M%S")+'}'+'.xlsx'))
        convertingTime = time.strftime("%Y%m%d_%H%M%S")
    
        if 1 : # 화일을 읽은폴더에 저장
            writer = pd.ExcelWriter(os.path.join(dir1,filename1+'{부가세검토'+time.strftime("%Y%m%d_%H%M%S")+'}'+'.xlsx'))
        else : 
            writer = pd.ExcelWriter(os.path.join('\\\\As5104t-9c9e\\업무폴더-201506~\\매입매출검토',filename1+'{부가세검토'+time.strftime("%Y%m%d_%H%M%S")+'}'+'.xlsx'))
            '''△△△△매입매출전표입력은 리버스문서보관함에 최초 저장되므로 읽을때 고정으로 보관함을 정하고  저장하는 위치는 매입매출폴더로 저정하면 더 손이편할듯하나 결국 판단사항이다.'''
    else : writer = pd.ExcelWriter(os.path.join('D:\\TEST2018\\세무사랑검토',filename1+'{당기전기검토'+convertingTime+'}'+'.xlsx'))





검토사항3 = []

for i3 in 검토사전2.keys():
    if i3 in df4.values:
        검토사항3.append(검토사전2[i3])

검토사전키값이공급처명일부와매치 = lambda x : [검토사전2[i] for i in 검토사전2 if i in x]

                
for i in df4.columns.tolist():   
#    t2.append(i)             
    for j in df4[i]:
#        t1.append(j)
        if type(j) is str:
#            if 검토사전키값이공급처명일부와매치(j) == []:
#                pass
#            else :
            if 검토사전키값이공급처명일부와매치(j) != []:
                if 검토사전키값이공급처명일부와매치(j)[0] not in 검토사항3:
                    검토사항3.append(검토사전키값이공급처명일부와매치(j)[0])
검토사항4 = []
#칼럼별매치 = []
for i2 in df4.columns.tolist():
    for j2 in 검토사전2:
        if df4[i2].dtypes == object: #str 쓰려면 값이 숫자이면 안됨
            검토사항4 += df4[df4[i2].str.contains(j2, na=False)][i2].unique().tolist()

검토사항5 = {}            
for i3 in df4.columns.tolist():
    for j3, k3 in 검토사전2.items():
        if df4[i3].dtypes == object:
            매치값 = df4[df4[i3].str.contains(j3, na=False)][i3].unique().tolist()
            for m3 in 매치값:
                if m3 not in 검토사항5:
                    검토사항5[m3] = k3
                
검토사항6 = []
#불공추정대상프레임 = df4[df4['부가세\n유형'].isin(매입과세불공제외리스트)]
df4001 = df4[df4['부가세\n유형'].isin(매입과세불공제외리스트)]
for i4 in df4001.columns.tolist():
#for i4 in 불공추정대상프레임.columns.tolist():    
    for j4 in 불공추정단어리스트:
#        if df4[i4].dtypes == object: #str 쓰려면 값이 숫자이면 안됨
        if df4001[i4].dtypes == object: #str 쓰려면 값이 숫자이면 안됨
            검토사항6 += df4001[df4001[i4].str.contains(j4, na=False)][i4].unique().tolist()
            
#181007            
if '서버컴여부 == 1':
    그런단어없어 = []
    for val in 기본매입업체리스트:
        if df4001[df4001['거래처명'].str.contains(val, na=False)]['거래처명'].unique().tolist() == []:
            그런단어없어.append(val)

            
    검토사항70 = [] # 있는분류
    검토사항71 = [] # 전체분류(중복있음)
    검토사항72 = [] # 없는분류; 청호나이스오 코웨이 둘다 없으면 '정수기'분류 없는것임
    for k, v in 기본매입업체사전.items():
        if df4001[df4001['거래처명'].str.contains(k, na=False)]['거래처명'].unique().tolist() != []:
            if v not in 검토사항70:
                검토사항70.append(v)
                
    검토사항71 = [v for k,v in 기본매입업체사전.items()]
    검토사항72 = list(set(검토사항71).difference(set(검토사항70))) # 차집합



부가세유형검토사전 = {'영세':'영세 유형 조회됩니다', '수입':'수입 유형 조회됩니다','수출':'수출 유형 조회됩니다','불공':'불공 유형 조회됩니다'}
부가세유형검토사항 = []
for i in df4['매출매입\n구분']:
    if i in 부가세유형검토사전:
        if 부가세유형검토사전[i] not in 부가세유형검토사항:
            부가세유형검토사항.append(부가세유형검토사전[i])
                



def printDF(df):
    if df.empty:
        print('중복이 없습니다(데이터프레임이 비어있습니다)')
    else : print(df)

df7 = df6.copy()
df8 = df6.copy()
df12 = df6.copy()
    
df13 = df4
df14 = df4[:]
#print(df4 is df13, '→ df13 is view of df4')
#print(df4 is df14, '→ df14 is copy of df4')

#print('▲ 검토사항(단어): ')
#print(검토사항3)
#checkDic['8.검토사항'] = ', '.join(검토사항3) # 리스트의 항목이 2개이상이면 엑셀의 셀에 삽입니 안되는듯해서 str로 변경함
검토결과리스트 = []
for k, v in 검토사항5.items():
    검토결과리스트.append(k+' → '+v)
checkDic['8.검토사항'] = ' // '.join(검토결과리스트)
if 검토사항6 != []:
    checkDic['9.0.불공추정단어'] = ', '.join(검토사항6)
else :
    checkDic['9.0.불공추정단어'] = '없음'
    
#181007    
if '서버컴여부 == 1':
    checkDic['9.3.조회안되는상호'] = ', '.join(그런단어없어)
    checkDic['9.4.조회안되는분류'] = ', '.join(검토사항72)


#print('▲ 부가세유형검토사항: ')
#print(부가세유형검토사항)
if 부가세유형검토사항 != []:
    checkDic['6.부가세유형검토'] = ', '.join(부가세유형검토사항)
else :
    checkDic['6.부가세유형검토'] = 'N/A'

checkShowlist = ['전표일자','부가세\n유형','거래처명','검색번호','공급가액', '부가세', '전자여부2', '의제공제율']

전체칼럼리스트 = df6.columns.tolist()
키번호제외리스트 = [x for x in 전체칼럼리스트 if x !='키번호']

#print('▲ 전체중복')
키번호제외부울 = df6.duplicated(키번호제외리스트,keep=False)
#printDF(df6[키번호제외부울][checkShowlist])
df151 = df6[df6.duplicated(키번호제외리스트,keep=False)][checkShowlist]
df151['구분'] = 'A.전체중복'

#print("▲ 일부중복['전표일자','부가세\n유형','거래처명','검색번호','공급가액']:")
일부중복체크부울 = df6.duplicated(['전표일자','부가세\n유형','거래처명','검색번호','공급가액'],keep=False)&(~키번호제외부울)
#printDF(df6[일부중복체크부울][checkShowlist])
df152 = df6[일부중복체크부울][checkShowlist]
df152['구분'] = "B.일부중복"            

#20200323 공급대가 절대값 고려해서 중복체크
df6_1 = df6.copy()
df6_1['공급가액abs'] = df6_1['공급가액'].abs() #절대값
일부중복체크부울_1 = df6_1.duplicated(['전표일자','부가세\n유형','거래처명','검색번호','공급가액abs'],keep=False)&(~키번호제외부울)
df152_1 = df6_1[일부중복체크부울_1][checkShowlist]
df152_1['구분'] = "B.일부중복abs" 


#print("▲ 중복추정")
중복추정부울 = df6.duplicated(['월','부가세\n유형','거래처명','검색번호','공급가액'],keep=False)&(~일부중복체크부울)&(~키번호제외부울)
#printDF(df6[중복추정부울][checkShowlist])
df1531 = df6[중복추정부울][checkShowlist]
df153 = df1531[~df1531['부가세\n유형'].isin(매입신카등리스트)]
df153['구분'] = "C.중복추정"

#print("▲ 전자종이 병존")
df71 = df7[['검색번호','전자\n여부']].drop_duplicates()
df72 = df71[df71.duplicated(['검색번호'],keep=False)]
#printDF(df7[df7.index.isin(df72.index)][checkShowlist])
df154 = df7[df7.index.isin(df72.index)][checkShowlist]
df154['구분'] = "D.전자,종이 병존"

#print("▲ 한업체의 세금계산서중에 공제와불공제 병존체크")
df7120 = df7[~df7['검색번호'].isin([0])]
df712 = df7120[['검색번호','부가세\n유형']].drop_duplicates()
df7121 = df712[df712['부가세\n유형'].isin([51, 54])]
df722 = df7121[df7121.duplicated(['검색번호'],keep=False)]
#printDF(df7[df7.index.isin(df722.index)][checkShowlist])
df155 = df7[df7.index.isin(df722.index)][checkShowlist]
df155['구분'] = "E.공제,불공제 병존"

#print("▲ 한업체의 계산서중에 의제매입 유무 병존 및 공제율 다른것 체크")
df750 = df12[~df12['검색번호'].isin([0])]
df7500 = df750[df750['부가세\n유형'].isin(매입면세리스트)]
df751 = df7500[['검색번호','의제류\n공제율']].drop_duplicates()
df752 = df751[df751.duplicated(['검색번호'],keep=False)]
#printDF(df12[df12.index.isin(df752.index)][checkShowlist])
df156 = df12[df12.index.isin(df752.index)][checkShowlist]
df156['구분'] = "F1.의제매입체크"


#print("▲ 의제매입 공제율 체크")
df75000 = df7500[~df7500['의제류\n공제율'].isin([0])]
df755 = df75000['의제류\n공제율'].drop_duplicates()
#printDF(df12[df12.index.isin(df755.index)][checkShowlist])
df157 = df12[df12.index.isin(df755.index)][checkShowlist]
df157['구분'] = "F2.의제매입공제율"

매출기준공제율과타이핑공제율에차이가있음부울 = False  
if not df157.empty: #의제입력한게 있다면
    for i in range(len(df6[~df6['의제류\n공제율'].isin([0])])):
        분자1 = (df6[~df6['의제류\n공제율'].isin([0])].iloc[i,:])['의제류\n공제율']
        분모1 = (df6[~df6['의제류\n공제율'].isin([0])].iloc[i,:])[의제류공제율분모]
        타이핑된공제율 = 분자1/분모1
        if abs(의제매입세액공제율 - 타이핑된공제율) > 0.001:
            checkDic['5.1.매출기준vs입력'] = '의제매입세액공제율의 매출기준('+의제매입세액공제율은+')과 다른 입력값'+'('+str(int(분자1))+'/'+str(int(분모1))+')'+'이 있습니다'
            매출기준공제율과타이핑공제율에차이가있음부울 = True
            기준과다른공제율 = 타이핑된공제율
            기준과다른공제율은 = '1.1.납부세액('+str(int(분자1))+'/'+str(int(분모1))+')'
else :
    checkDic['5.1.매출기준vs입력'] = 'N/A'

if 매출기준공제율과타이핑공제율에차이가있음부울 :
    타이핑기준한도적용후의제매입공제세액 = 의제매입한도적용후가액*기준과다른공제율
    타이핑기준매입세액 = df4[df4['부가세\n유형'].isin(매입리스트)].부가세.sum() + 타이핑기준한도적용후의제매입공제세액 - df4[df4['부가세\n유형'].isin(매입불공제리스트)].부가세.sum()
    타이핑기준신카세액공제전납부세액 = 매출세액 - 타이핑기준매입세액
    
    if 타이핑기준신카세액공제전납부세액 < 신카발행세액공제 and 타이핑기준신카세액공제전납부세액 > 0:
        타이핑기준사중손실세액 = 신카발행세액공제 - 타이핑기준신카세액공제전납부세액
    elif 타이핑기준신카세액공제전납부세액 <= 0:
        타이핑기준사중손실세액 = 신카발행세액공제
    else : 타이핑기준사중손실세액 = 0
    
    if 타이핑기준신카세액공제전납부세액 > 신카발행세액공제:
        타이핑기준납부세액 = 매출세액 - 타이핑기준매입세액 - 신카발행세액공제
    elif 신카세액공제전납부세액 > 0 :
        타이핑기준납부세액 = 매출세액 - 타이핑기준매입세액 - (신카발행세액공제 - 타이핑기준사중손실세액)
    else : 타이핑기준납부세액 = 매출세액 - 타이핑기준매입세액
    checkDic[기준과다른공제율은] = 천단위콤마함수(int(타이핑기준납부세액))
else :
    #checkDic[기준과다른공제율은] = 'N/A'
    checkDic['1.1.납부세액(다른공제율)'] = 'N/A'
            
#print("▲ 세금계산서와 신카현영 중복체크")
df811 = df8[~df8['검색번호'].isin([0])]
df81 = df811[['부가세\n유형','검색번호']].drop_duplicates()
df82 = df81[df81.duplicated(['검색번호'],keep=False)]
df83 = df8[df8.index.isin(df82.index)]
df8301 = df83[df83['부가세\n유형'].isin(매입세금계산서등리스트)][checkShowlist]
df8302 = df83[df83['부가세\n유형'].isin(매입신카등리스트)][checkShowlist]
df83011 = df8301[df8301.검색번호.isin(df8302.검색번호)]
df83022 = df8302[df8302.검색번호.isin(df8301.검색번호)]
df8303 = pd.concat([df83011,df83022],axis=0)
df158 = df8303.sort_values(by='거래처명')
#printDF(df158)
df158['구분'] = "G.세금,신카중복"


#print('▲ 국세청승인번호 중복체크:')
#printDF(df6[df6.duplicated(['국세청승인번호'],keep=False)&(~df6['국세청승인번호'].isin([0]))][checkShowlist])
df159 = df6[df6.duplicated(['국세청승인번호'],keep=False)&(~df6['국세청승인번호'].isin([0]))][checkShowlist]
df159['구분'] = "H.전자승인번호중복"


#print('▲ 매입과 매출 모두 있는 업체')
입출검색번호고유프레임 = df6[~df6.duplicated(['입출','검색번호'])]
입출검색번호고유프레임 = 입출검색번호고유프레임[입출검색번호고유프레임['검색번호'] != 0]
#printDF(입출검색번호고유프레임[입출검색번호고유프레임.duplicated(['검색번호'],keep=False)][checkShowlist])
df160 = 입출검색번호고유프레임[입출검색번호고유프레임.duplicated(['검색번호'],keep=False)][checkShowlist]
df160['구분'] = "I.매입,매출 병존"

#print('▲ 예정신고누락분')
df161 = df6[df6['예정신고\n누락분여부'].isin([1])][checkShowlist]
#printDF(df161)
df161['구분'] = "J.예정신고누락분"

#print('▲ 봉사료')
df162 = df6[~df6['봉사료'].isin([0])][checkShowlist]
#printDF(df162)
df162['구분'] = "K.봉사료"

#print('▲ 가산세')
df163 = df6[~df6['가산세\n구분'].isin([0])][checkShowlist]
#printDF(df163)
df163['구분'] = "L.가산세"

#print('▲ 공급자와 공급받는자 바뀜')
df164 = df6[df6['검색번호'].isin([회사정보사전['사업자등록번호']])][checkShowlist]
#printDF(df164)
df164['구분'] = "M.공급자vs공급받는자"

#print('▲ 누락추정(갯수가 넷다섯/전자는 제외)')
임대료패턴리스트 = ['부가세\n유형','거래처명','검색번호','공급가액']
df1651 = df6[df6.duplicated(['부가세\n유형','거래처명','검색번호','공급가액'],keep=False)][checkShowlist]
df1651['거래처명'] = df1651['거래처명'].apply(str) # AttributeError: 'int' object has no attribute 'apply'
#df1652 = df1651.set_index(임대료패턴리스트).sort_index() #20190721 미선씨컴 에러일으켜서 주석처리함
df1653 = df6.groupby(임대료패턴리스트)['전표일자'].count()
df16530 = df6.groupby(임대료패턴리스트)['전표일자']
df165300 = df16530.agg([('카운트', 'count')])
df165301 = df6.groupby(임대료패턴리스트).size().to_frame('사이즈')
df1653011 = df6.groupby(임대료패턴리스트).size()
df16531 = (df1651.groupby(임대료패턴리스트)['전표일자'].count()>3)&(df1651.groupby(임대료패턴리스트)['전표일자'].count()<6)
df16532 = (df1651.groupby(임대료패턴리스트)[['전표일자','거래처명']].count()>3)&(df1651.groupby(임대료패턴리스트)[['전표일자','거래처명']].count()<6)
df16533 = (df6.groupby(임대료패턴리스트)[['전표일자','거래처명']].count()>3)&(df6.groupby(임대료패턴리스트)[['전표일자','거래처명']].count()<6)
df165310 = pd.DataFrame({'구분': df16531.index, '카운트':df16531.values})
df1653101 = df16531.to_frame('부울')
df16534 = df1651.groupby(임대료패턴리스트).size().isin([4,5])
df1653401 = df16534.reset_index(name='4~5')
df1653402 = df1653401[df1653401['4~5']==True][임대료패턴리스트]
df165341 = df16534.to_frame('넷다섯')
df1653410 = df165341[df165341['넷다섯']==True]
df1653411 = df1653410.reset_index()
df1653412 = df1653411[임대료패턴리스트]
df1655 = pd.merge(df1651,df1653412)
df1656 = df1655[~df1655['전자여부2'].isin(['e'])]
df16551 = pd.merge(df1651,df1653402)
df16561 = df16551[~df16551['전자여부2'].isin(['e'])]
                  
df165 = df1656[~df1656['부가세\n유형'].isin(매입신카등리스트)]
#printDF(df165)
df165['구분'] = "N.넷다섯"

#print('▲ 법인&종이세금발행')
if 회사정보사전['사업자등록번호'][4] in ['8']: #우리 거래처가 법인 : 8
    df1661 = df6[df6['부가세\n유형'].isin([11,12,13])][checkShowlist]
    df1662 = df1661[~df1661['전자여부2'].isin(['e'])]
    df166 = df1662.copy()
    #printDF(df166)
    df166['구분'] = "O.11유형&종이발행"
else :
    df166 = pd.DataFrame()

#print('▲ 51유형&면세업체')
df1671 = df6[df6['부가세\n유형'].isin([51])][checkShowlist]
df1672 = df1671[df1671['검색번호'].str[4:5].isin(['9'])][checkShowlist] # ***-9*-*****인데 매입과세인경우
df167 = df1672.copy()
#printDF(df167)
df167['구분'] = "P.51유형-9*-업체"

#print('▲ 공급가액10%와 부가세의 차이가 1천원이상')
부가세공란아닌과세리스트 = [11, 14, 17, 22, 51, 54, 56, 57, 61]
df1681 = df6[df6['부가세\n유형'].isin(부가세공란아닌과세리스트)][checkShowlist]
df1682 = df1681[abs(df1681['공급가액']*0.1 - df1681['부가세']) >= 10000]
df168 = df1682.copy()
#printDF(df168)
df168['구분'] = "Q.공급가액vs세액"

#print('▲ 고정자산')
df169 = df6[~df6['고정자산코드'].isin([0])][checkShowlist]
#printDF(df169)
df169['구분'] = "R.고정자산"

#print('▲ 상호가 공란인업체')
df1701 = df6[df6['거래처명'].isin([0])][checkShowlist]
df170 = df1701[df1701['부가세\n유형'].isin(매입세금계산서등리스트+매출세금계산서등리스트)]
#printDF(df170)
df170['구분'] = "S.임대추정 ← 상호가 공란인업체"

#print('▲ 부가세 > 공급가액10% : 100원 이상')
df1711 = df6[df6['부가세\n유형'].isin(부가세공란아닌과세리스트)][checkShowlist]
df171 = df1711[(df1711['부가세'] - df1711['공급가액']*0.1) >= 100]
#printDF(df171)
df171['구분'] = "T.세액>공급가액x0.1"


#중복체크합치기 = pd.concat([df151, df152, df153, df154, df155, df156, df157, df158, df159, df160, df161, df162, df163, df164, df165, df166, df167, df168, df169, df170, df171], axis=0)
중복체크합치기 = pd.concat([df151, df152, df152_1, df153, df154, df155, df156, df157, df158, df159, df160, df161, df162, df163, df164, df165, df166, df167, df168, df169, df170, df171], axis=0) #20200323
중복체크합치기['거래처명'] = 중복체크합치기['거래처명'].apply(str)
중복체크그룹 = 중복체크합치기.set_index(['구분','거래처명'])
중복체크그룹 = 중복체크그룹.sort_index()







###########################################################################
###########################################################################

#180928 여러 세무대리 거래처중 1개만 선택해서 가공


선택 = 1 # 1:db, 2:xls ; 직전기자료 어디서 가져오나

if 선택 == 1 :
    if 서버컴여부 == 1:
        #con = sqlite3.connect("E:\\TEST2018\\매입매출입력2018년1기.db")
        #con = sqlite3.connect("\\\\As5104t-9c9e\\업무폴더-201506~\\매입매출검토\\library\\매입매출입력2018년1기.db")
        #fileListinfolder = getfilesRev("\\\\As5104t-9c9e\\업무폴더-201506~\\매입매출검토\\library")
        fileListinfolder = getfilesRev("\\\\Desktop-ekk32ts\\업무폴더-201912\\매입매출검토\\library") #20191217
        for i in fileListinfolder:
            if i.split('_')[0] in ['매입매출입력']: #내림차순으로 정렬해서 처음 발견되는 DB가 최근자료이므로 선택함
                whatIwant = i # whatIwant Out[5]: '매입매출입력_2018년2기.db' #20191217
                break
        if 1 :
            #con = sqlite3.connect(os.path.join("\\\\As5104t-9c9e\\업무폴더-201506~\\매입매출검토\\library", whatIwant))
            con = sqlite3.connect(os.path.join("\\\\Desktop-ekk32ts\\업무폴더-201912\\매입매출검토\\library", whatIwant))
        else :
            con = sqlite3.connect("E:\\TEST2018\\appendToDB\\매입매출입력_2018년1기.db") # for test
        #'\\\\As5104t-9c9e\\업무폴더-201506~\\매입매출검토\\library\\매입매출입력_2018년1기.db'        
                
    else : con = sqlite3.connect("D:\\TEST2018\\매입매출입력2018년1기.db")

    선택된세무대리거래처 = 회사정보사전['사업자등록번호']
    #x = "SELECT * FROM '2018년1기' WHERE 사업자등록번호1 = "+"'"+선택된세무대리거래처+"'"
    term1 = os.path.splitext(whatIwant)[0].split('_')[1] #'2018년1기'는 DB화일명 뒷부분이면서 DB의 테이블 이름이다
    #term1 Out[51]: '2018년2기' 20190625
    x = "SELECT * FROM "+"'"+term1+"'"+" WHERE 사업자등록번호1 = "+"'"+선택된세무대리거래처+"'"
    #"SELECT * FROM '2018년1기' WHERE 사업자등록번호1 = '110-08-69321'"
    priorYear1 = pd.read_sql(x, con, index_col = 'index')

elif 선택 == 2:
    
    #dirFilenameExtension = 'D:\\TEST2018\\다수의엑셀처리\\출력\\매입매출전표입력appended{출력일20180921_092839}.xlsx'
    dirFilenameExtension = 'D:\\TEST2018\\다수의엑셀처리\\출력\\매입매출전표입력appended{출력일20180920_175649}.xlsx'
    #dirFilenameExtension = 'D:\\TEST2018\\다수의엑셀처리\\출력\\매입매출전표입력appended{출력일20180920_175649}-1.xlsx'
    
    dir1 = os.path.split(dirFilenameExtension)[0]
    base1=os.path.basename(dirFilenameExtension)
    filename1 = os.path.splitext(base1)[0]
    extension2 = os.path.splitext(base1)[1]
    
    if 'appended' not in filename1:
        Mbox('leecta', filename1+' ==> [매입매출전표입력appended]가 아닙니다', 1)
        sys.exit()
    
    workbook = pd.ExcelFile(dirFilenameExtension)
    priorYear = workbook.parse('dfs')
    priorYear1000 = priorYear.copy()
    #선택된세무대리거래처 = '124-52-44924' # str
    선택된세무대리거래처 = 회사정보사전['사업자등록번호']
    priorYear1 = priorYear1000[priorYear1000['사업자등록번호1'].isin([선택된세무대리거래처])].fillna(0)
    
    '''def 전자대체(x):
        if x == 1:
            return 'e'
    '''
    
    def 일자별상세내역누적(x):
        j = ''
        for i in x:
            j = j + i
        return j
    
    priorYear1['전자여부2'] = priorYear1['전자\n여부'].apply(전자대체)
    priorYear1['전자여부2'] = priorYear1['전자여부2'].fillna(' ')
    priorYear1['전표일자및공급가액'] =  ' 【' + priorYear1['전표일자'].apply(str).str[4:6]+'/'+priorYear1['전표일자'].apply(str).str[6:8]+'】 '+ priorYear1['공급가액'].apply(천단위콤마함수2)
    
    priorYear1['Date'] = pd.to_datetime(priorYear1['전표일자'], format='%Y%m%d')
    priorYear1['월'] = priorYear1.Date.dt.month
    priorYear1['년도'] = priorYear1.Date.dt.year
    의제류공제율분자위치 = priorYear1.columns.tolist().index('의제류\n공제율')
    의제류공제율분모 = priorYear1.columns.tolist()[의제류공제율분자위치+1] #엑셀에서 분모는 분자 우측에 있음
    priorYear1['의제공제율'] = ' '+priorYear1['의제류\n공제율'].apply(int).apply(str) + '/' + priorYear1[의제류공제율분모].apply(int).apply(str)+' '
    priorYear1['의제공제율'] = priorYear1[~priorYear1['의제공제율'].isin([' 0/0 '])].의제공제율
    priorYear1['의제공제율'] = priorYear1['의제공제율'].fillna(' ')
    priorYear1['의제류공제율'] = priorYear1['의제류\n공제율']/(priorYear1[의제류공제율분모].apply(lambda num : num if num != 0 else 1)) #분모가 0 안되도록


    
if not priorYear1.empty: # DB에 직전기 자료가 있다면
    #여기서부터 업체별 그루핑
    priorYear1 = priorYear1.fillna(0)
    gb = pd.DataFrame() # groupby
    gb['매출과세공급가액'] = priorYear1[priorYear1['부가세\n유형'].isin(매출과세리스트)].groupby(['회사명1','사업자등록번호1']).공급가액.sum()
    gb['매입과세공급가액'] = priorYear1[priorYear1['부가세\n유형'].isin(매입과세리스트)].groupby(['회사명1','사업자등록번호1']).공급가액.sum()
    gb['세액있는매출공급가액'] = priorYear1[priorYear1['부가세\n유형'].isin(세액있는매출리스트)].groupby(['회사명1','사업자등록번호1']).공급가액.sum()
    gb['매출공급가액'] = priorYear1[priorYear1['부가세\n유형'].isin(매출리스트)].groupby(['회사명1','사업자등록번호1']).공급가액.sum()
    gb['매출면세공급가액'] = priorYear1[priorYear1['부가세\n유형'].isin(매출면세리스트)].groupby(['회사명1','사업자등록번호1']).공급가액.sum()
    gb['매출세액'] = gb['세액있는매출공급가액']*0.1
    gb['과세년도'] = priorYear1.groupby(['회사명1','사업자등록번호1']).년도.max()
    gb['의제류대상금액'] = priorYear1.groupby(['회사명1','사업자등록번호1'])['의제류\n대상금액'].sum()
    gb['의제류공제율2'] = priorYear1.groupby(['회사명1','사업자등록번호1']).의제류공제율.max() # min하면 0이 저장됨; 의제류공제율 칼럼에 0이 많음; fillna값이 0이므로 max를 써야함

    #20200323    
    gb['매입신카공급가액0'] = priorYear1[priorYear1['부가세\n유형'].isin(매입신카리스트1)].groupby(['회사명1','사업자등록번호1']).공급가액.sum()
    gb['매입현영공급가액0'] = priorYear1[priorYear1['부가세\n유형'].isin(매입현영리스트1)].groupby(['회사명1','사업자등록번호1']).공급가액.sum()
    
    
    def 의제매입한도함수2(group): # group -> 매출과세공급가액; 리턴값은 의제매입한도
        ret1 = 0
        if group > 200000000:
    #        return group*0.45
            ret1 = group*0.45
        elif group > 100000000:
    #        return group*0.55
            ret1 = group*0.55
    #    else : return group*0.6
        else :
            ret1 = group*0.6
        return ret1
    
    def 의제매입한도율함수(group): # group -> 매출과세공급가액; 리턴값은 의제매입한도율
        ret2 = 0
        if group > 200000000:
            ret2 = 0.45
    #        return 0.45
        elif group > 100000000:
            ret2 = 0.55
    #        return 0.55
        else :
            ret2 = 0.6
        return ret2
    
    gb['의제매입한도'] = gb['매출과세공급가액'].apply(의제매입한도함수2)
    gb['적용된의제매입한도율'] = gb['매출과세공급가액'].apply(의제매입한도율함수)
    
    gb['한도초과액'] = gb['의제류대상금액']-gb['의제매입한도']
    의제매입한도적용후가액1 = gb[gb['의제류대상금액'] > gb['의제매입한도']].의제매입한도
    의제매입한도적용후가액2 = gb[gb['의제류대상금액'] <= gb['의제매입한도']].의제류대상금액
    gb['의제매입한도적용후가액'] = 의제매입한도적용후가액1.add(의제매입한도적용후가액2, fill_value=0) # empty series가 있으면 NaN값이 저장됨
    
    gb['의제매입한도적용후매입세액'] = gb['의제매입한도적용후가액']*gb['의제류공제율2'] #의제매입한도때문에 부가세칼럼에서 바로 추출못하고 공급가액에서 공제율곱해서 세액구함
    
    #공제율상향년도 = [2018, 2019]
    
    gb['매입리스트부가세'] = priorYear1[priorYear1['부가세\n유형'].isin(매입리스트)].groupby(['회사명1','사업자등록번호1']).부가세.sum()
    gb['매입불공제리스트부가세'] = priorYear1[priorYear1['부가세\n유형'].isin(매입불공제리스트)].groupby(['회사명1','사업자등록번호1']).부가세.sum()
    gb['매입세액'] = gb['매입리스트부가세'].add(gb['의제매입한도적용후매입세액'], fill_value=0)
    gb['매입세액'] = gb['매입세액'].add(-gb['매입불공제리스트부가세'], fill_value=0)

    if gb['매출면세공급가액'][0] > 0:
        gb['공통매입대상부가세합계'] = priorYear1[priorYear1['검색번호'].isin(공통매입대상사업자등록번호리스트)].groupby(['회사명1','사업자등록번호1']).부가세.sum()
        gb['불공제매입세액'] = (gb['공통매입대상부가세합계'])*(gb['매출면세공급가액'])/(gb['매출공급가액'])
        #gb['납부세액'] = gb['납부세액'] + gb['p불공제매입세액']
        gb['매입세액'] = gb['매입세액'] - gb['불공제매입세액']
        

    
    gb['신카발행세액공제'] = priorYear1[priorYear1['부가세\n유형'].isin(매출신카등과세리스트)].groupby(['회사명1','사업자등록번호1']).합계금액.sum()*0.013
    
    if 회사정보사전['사업자등록번호'][4] in ['8']: #우리 거래처가 법인 : 8
        gb['신카발행세액공제'] = 0                
    elif gb['신카발행세액공제'][0] > 5000000:
        gb['신카발행세액공제'] = 5000000   
       
    '''사업자번호리스트 = gb['신카발행세액공제'].index.get_level_values(level=1).tolist()
    법인사업자번호리스트 = []
    for i in 사업자번호리스트:
        if i[4] in ['8']:
            법인사업자번호리스트.append(i)
    '''

    def 신카발행세액공제한도체크함수2(group): # group -> 신카발행세액공제
        if group > 5000000:
            return 5000000
        else : return group
    
    gb['한도체크후신카발행세액공제'] = gb['신카발행세액공제'].fillna(0).apply(신카발행세액공제한도체크함수2)
    
    gb['신카세액공제전납부세액'] = gb['매출세액'] - gb['매입세액']
    
    def 사중손실세액함수2(x):
    #    if 신카세액공제전납부세액 - 한도체크후신카발행세액공제 < 0 and 신카세액공제전납부세액 > 0 :
        if (x[0] - x[1] < 0)&(x[0] > 0) :
            return x[1] - x[0]
        elif x[0] <= 0:
            return x[1]
        else :
            return 0
    
    gb['사중손실세액'] = gb[['신카세액공제전납부세액', '한도체크후신카발행세액공제']].fillna(0).apply(사중손실세액함수2, axis=1)
    
    def 납부세액함수(group):
        if gb['신카세액공제전납부세액'] > gb['한도체크후신카발행세액공제']:
            return gb['매출세액'] - gb['매입세액'] - gb['한도체크후신카발행세액공제']
        elif gb['신카세액공제전납부세액'] > 0 :
            return gb['매출세액'] - gb['매입세액'] - (gb['한도체크후신카발행세액공제'] - gb['사중손실세액'])
        else : gb['매출세액'] - gb['매입세액']
    
    def 납부세액함수2(gb): #gb[0]:신, gb[1]:한, gb[2]:사, gb[3]:매출세액, gb[4]:매입세액
        if gb[0] > gb[1]:
            return gb[0] - gb[1]
        elif gb[0] > 0 :
            return gb[0] - (gb[1] - gb[2])
        else :
            return gb[0] #환급일때
    
    gb = gb.fillna(0)
    
    gb['납부세액'] = gb[['신카세액공제전납부세액', '한도체크후신카발행세액공제', '사중손실세액']].fillna(0).apply(납부세액함수2, axis=1)

    '''if gb['매출면세공급가액'][0] > 0:
        gb['p공통매입대상부가세합계'] = priorYear1[priorYear1['검색번호'].isin(공통매입대상사업자등록번호리스트)].groupby(['회사명1','사업자등록번호1']).부가세.sum()
        gb['p불공제매입세액'] = (gb['p공통매입대상부가세합계'])*(gb['매출면세공급가액'])/(gb['매출공급가액'])
        gb['납부세액'] = gb['납부세액'] + gb['p불공제매입세액']
    '''

    if gb['매출과세공급가액'][0] > 0:
        gb['직전기부가율은'] = 소숫점둘째반올림함수((gb['매출과세공급가액'][0] - (gb['매입과세공급가액'][0] + gb['의제류대상금액'][0] - priorYear1[priorYear1['고정자산코드'].isin([1])].공급가액.sum()))/gb['매출과세공급가액'][0]*100)+'%'
    else : gb['직전기부가율은'] = 'N/A'    
    
    checkDic['2.4.0.직전기 납부세액'] = 천단위콤마함수(int(gb['납부세액'][0]))+'원(직전기 부가율 : '+gb['직전기부가율은'][0]+')←예정고지 미반영금액임'
    '''if 간이여부 == 1:
        checkDic['2.4.1.간이여부'] = '간이로 추정되며 표시된 납부세액은 일반과세자기준입니다.'
    else :
        checkDic['2.4.1.간이여부'] = 'N/A'
        '''
    if not 회사정보사전['사업자등록번호'][4] in ['8']: #우리 거래처가 법인 : 8
        if 간이여부 == 0:
            if gb['납부세액'][0]/2 > 200000:
                gb['당기추정예정고지세액'] = (gb['납부세액'][0]/2)//1000*1000
                checkDic['2.5.당기추정예정고지'] = 천단위콤마함수(int(gb['당기추정예정고지세액'][0]))+'원'
            elif gb['납부세액'][0]/2 > 0:
                checkDic['2.5.당기추정예정고지'] = '20만원이하 고지생략'
            else : 
                checkDic['2.5.당기추정예정고지'] = '없음'
        else : checkDic['2.5.당기추정예정고지'] = 'N/A'
                
    priorGroup1 = priorYear1.groupby(['입출','부가세\n유형','매출매입\n구분','의제공제율','거래처명','검색번호'])['전표일자및공급가액'].agg([('일자별상세내역',f)])
    priorGroup2 = priorYear1.groupby(['입출','부가세\n유형','매출매입\n구분','의제공제율','거래처명','검색번호'])['카운트','전자\n여부','공급가액','부가세'].sum()
    priorGroup = pd.concat([priorGroup2, priorGroup1], axis=1)
    
    #priorGrptotal = priorYear1.sum(level=['입출','부가세\n유형','매출매입\n구분','의제공제율'])
    priorGrptotal = priorYear1.groupby(['입출','부가세\n유형','매출매입\n구분','의제공제율'])['카운트','전자\n여부','공급가액','부가세'].sum()
    if '1.매출' in priorGrptotal.index:
        priorGrptotal.loc[('매출',' ','Total',' '),:] = priorGrptotal.loc['1.매출'].sum()
    if '2.매입' in priorGrptotal.index:
        priorGrptotal.loc[('매입',' ','Total',' '),:] = priorGrptotal.loc['2.매입'].sum()
        
        
    직전기현금매출 = priorYear1[priorYear1['부가세\n유형'].isin([14, 15])].공급가액.sum()      
    직전기신카등매출 = priorYear1[priorYear1['부가세\n유형'].isin(매출신카등과세리스트)].공급가액.sum()
    과세매출증감 = 매출과세공급가액 - gb['매출과세공급가액'][0]
    if gb['매출과세공급가액'][0] > 0:
        과세매출증감율은 = 소숫점둘째반올림함수(과세매출증감/gb['매출과세공급가액'][0]*100)+'%'
    else :
        과세매출증감율은 = '..' #20190724
    과세매입증감 = 매입과세공급가액 - gb['매입과세공급가액'][0]
    if gb['매입과세공급가액'][0] > 0:
        과세매입증감율은 = 소숫점둘째반올림함수(과세매입증감/gb['매입과세공급가액'][0]*100)+'%'
    else : 과세매입증감율은 = '..' #20190718
    의제매입증감 = 의제류대상금액 - gb['의제류대상금액'][0]
    if gb['의제류대상금액'][0] > 0:
        의제매입증감율은 = 소숫점둘째반올림함수(의제매입증감/gb['의제류대상금액'][0]*100)+'%'
    else : 의제매입증감율은 = '..'
    if 직전기신카등매출 > 0 :
        직전기현금매출비율은 = 소숫점둘째반올림함수(직전기현금매출/직전기신카등매출*100)+'%'
    else :
        직전기현금매출비율은 = '신카등매출이 없음'

    #20200323        
    매입신카증감 = 매입신카공급가액1 - gb['매입신카공급가액0'][0]
    if gb['매입신카공급가액0'][0] > 0:
        매입신카증감율은 = 소숫점둘째반올림함수(매입신카증감/gb['매입신카공급가액0'][0]*100)+'%'
    else :
        매입신카증감율은 = '..'            
    매입현영증감 = 매입현영공급가액1 - gb['매입현영공급가액0'][0]
    if gb['매입현영공급가액0'][0] > 0:
        매입현영증감율은 = 소숫점둘째반올림함수(매입현영증감/gb['매입현영공급가액0'][0]*100)+'%'
    else :
        매입현영증감율은 = '..'            
        
        
    checkDic['2.4.1.직전대비증감'] = '▲ 직전기현금매출 : '+str(천단위콤마함수(int(직전기현금매출)))+'원'+' // 현금비율 : '+직전기현금매출비율은+'\n'\
                                    +'▲ 과세매출증감 : '+str(천단위콤마함수(int(과세매출증감)))+'원'+' // 증감율 : '+과세매출증감율은+'\n'\
                                    +'▲ 과세매입증감 : '+str(천단위콤마함수(int(과세매입증감)))+'원'+' // 증감율 : '+과세매입증감율은+'\n'\
                                    +'▲ 매입신카증감 : '+str(천단위콤마함수(int(매입신카증감)))+'원'+' // 증감율 : '+매입신카증감율은+'\n'\
                                    +'▲ 매입현영증감 : '+str(천단위콤마함수(int(매입현영증감)))+'원'+' // 증감율 : '+매입현영증감율은+'\n'\
                                    +'▲ 의제매입증감 : '+str(천단위콤마함수(int(의제매입증감)))+'원'+' // 증감율 : '+의제매입증감율은
        

    
    
    
else :
    checkDic['2.4.0.직전기 납부세액'] = 'N/A'
    checkDic['2.4.1.직전대비증감'] = 'N/A'
    checkDic['2.5.당기추정예정고지'] = 'N/A'
    checkDic['2.4.1.직전대비증감'] = 'N/A'


if 회사정보사전['사업자등록번호'] in dff4.values:
    #181007
    #접수증납부세액 = dff4[dff4['사업자(주민)\n등록번호'].isin([회사정보사전['사업자등록번호']])]['실제납부할세액(본세)']
    접수증납부세액 = dff4[dff4['사업자(주민)\n등록번호'].isin([회사정보사전['사업자등록번호']])]['실제납부할세액\n(본세)']
    접수일시 = dff4[dff4['사업자(주민)\n등록번호'].isin([회사정보사전['사업자등록번호']])]['접수일시']
    #checkDic['2.4.1.1.직전기 접수증 납부세액'] = 천단위콤마함수2(접수증납부세액.iloc[0])+'원(접수일시 : '+접수일시.iloc[0]+')'
    checkDic['2.4.1.1.직전기 접수증 납부세액'] = 천단위콤마함수2(접수증납부세액.iloc[0])+'원(접수일시 : '+str(접수일시.iloc[0])+')' #20191218 str(접수일시.iloc[0]) Out[9]: '20190722115123'
    #checkDic['2.4.1.2.평균부가율등'] = '▲ 주업종 : '+이회사의주업종+' ▲ 표본수 : '+str(이회사의평균부가율등['표본수'][0])+'개 ▲ 최대값 : '+소숫점둘째반올림함수(이회사의평균부가율등['최대값'][0])+'% ▲ 최소값 : '+소숫점둘째반올림함수(이회사의평균부가율등['최소값'][0])+'% ▲ 평균값 : '+소숫점둘째반올림함수(이회사의평균부가율등['평균값'][0])+'% ▲ 중간값 : '+소숫점둘째반올림함수(이회사의평균부가율등['중간값'][0])+'%'
else : 
    checkDic['2.4.1.1.직전기 접수증 납부세액'] = '신고내역이 없습니다'
    #checkDic['2.4.1.2.평균부가율등'] = '신고내역이 없습니다'
    

if 회사정보사전['사업자등록번호'] in dff6.values:
    조회된예정고지세액 = dff6[dff6['납세자번호'].isin([회사정보사전['사업자등록번호']])]['예정고지세액']
    조회된납부기한 = dff6[dff6['납세자번호'].isin([회사정보사전['사업자등록번호']])]['납부기한'] 
    if type(조회된예정고지세액.iloc[0]) is int:
        checkDic['2.6.조회된 예정고지세액'] = 천단위콤마함수2(조회된예정고지세액.iloc[0])+'원(납부기한 : '+조회된납부기한.iloc[0]+')'
    elif isinstance(조회된예정고지세액.iloc[0],int):#20190625
        checkDic['2.6.조회된 예정고지세액'] = 천단위콤마함수2(조회된예정고지세액.iloc[0])+'원(납부기한 : '+조회된납부기한.iloc[0]+')'
    elif isinstance(조회된예정고지세액.iloc[0],float):
        checkDic['2.6.조회된 예정고지세액'] = 천단위콤마함수2(조회된예정고지세액.iloc[0])+'원(납부기한 : '+조회된납부기한.iloc[0]+')'
    else :
        checkDic['2.6.조회된 예정고지세액'] = 조회된예정고지세액.iloc[0]
else : 
    checkDic['2.6.조회된 예정고지세액'] = '스크래핑한 업체가 아닙니다'

if 회사정보사전['사업자등록번호'] in dff6.values:
    if isinstance(조회된예정고지세액.iloc[0],int):
        checkDic['1.3.납부세액(예정고지차감)'] = 천단위콤마함수(int(납부세액-조회된예정고지세액.iloc[0]))+'원' #20200323



df181002_1 = df6[df6['부가세\n유형'].isin(매입세금계산서등리스트)][['거래처명','검색번호','전자\n여부']].drop_duplicates()
df181002_2 = priorYear1[priorYear1['부가세\n유형'].isin(매입세금계산서등리스트)][['거래처명','검색번호','전자\n여부']].drop_duplicates()
df181002_5 = priorYear1[priorYear1['부가세\n유형'].isin([54])][['거래처명','검색번호','전자\n여부']].drop_duplicates() #20190720 직전기 불공업체리스트
df181002_3 = pd.concat([df181002_2, df181002_1, df181002_1]).drop_duplicates(keep=False) # keep=False : 중복발생부분은 남겨두지않고 drop
'''
pd.concat([df181002_2, df181002_1]).drop_duplicates(keep=False)
Out[10]: 
                 거래처명          검색번호  전자\n여부
index                                      
8572     롯데제과(주)화성영업소  186-81-00924     1.0
8733             농가축산  123-91-43940     0.0
25               농가축산  123-91-43940     1.0
58                맬랜드  543-15-01188     1.0
172    용일 농업회사법인 유한회사  727-86-01239     0.0
→ 전기에만 있는것과 당기에만 있는것 모두 출력됨
pd.concat([df181002_2, df181002_1, df181002_1]).drop_duplicates(keep=False)
Out[11]: 
               거래처명          검색번호  전자\n여부
index                                    
8572   롯데제과(주)화성영업소  186-81-00924     1.0
8733           농가축산  123-91-43940     0.0
→ 전기만 남겨두고 당기자료 날리려면 [df181002_2, df181002_1, df181002_1] 해줘야함
'''
                       
df181002_4 = df181002_3[df181002_3['전자\n여부'].isin([0])]
if 0: #20200323
    if not df181002_4.empty:
        checkDic['9.5.2.0.직전기에만있는종이매입'] = ', '.join(df181002_4['거래처명'].tolist())
    else :
        checkDic['9.5.2.0.직전기에만있는종이매입'] = 'N/A'
    
df181002_6 = df6[df6['검색번호'].isin(df181002_4['검색번호'].tolist())][['거래처명','검색번호','전자\n여부']].drop_duplicates()
if not df181002_6.empty: #20200323
    #checkDic['9.5.2.1.직전기종이→당기전자'] = ', '.join(df181002_6['거래처명'].tolist()) #TypeError: sequence item 0: expected str instance, int found → [0]
    checkDic['9.5.2.1.직전기종이→당기전자'] = ', '.join([str(x) for x in df181002_6['거래처명'].tolist()]) #20200720 리스트컴프리헨션
    checkDic['9.5.2.1.직전기종이→당기전자'] = ', '.join(toStringList(df181002_6['거래처명'].tolist())) #20200720 함수로 간단히
    
else :
    checkDic['9.5.2.1.직전기종이→당기전자'] = 'N/A'

if 1 : #20200323
    직전기에만있는종이매입1 = df181002_4['거래처명'].tolist()
    당기에전자로발행된매입1 = df181002_6['거래처명'].tolist()
    진짜직전기에만있는종이매입2 = [s for s in 직전기에만있는종이매입1 if s not in 당기에전자로발행된매입1 ]
    new_b = 진짜직전기에만있는종이매입2
    
    if new_b != []:
        checkDic['9.5.2.2.Ⓡ직전기에만있는종이매입'] = ', '.join(new_b)
    else :
        checkDic['9.5.2.2.Ⓡ직전기에만있는종이매입'] = 'N/A'

if not df181002_5.empty:#20190720 직전기 불공업체리스트
    checkDic['9.5.3.직전기 불공업체'] = ', '.join(df181002_5['거래처명'].tolist())
else :
    checkDic['9.5.3.직전기 불공업체'] = 'N/A'

    
new_a = []
sublist = []
for sublist in df181002_3.values.tolist():
    for i in sublist:
        new_a.append(str(i))
if new_a != []:
    checkDic['9.5.1.직전기에만있는매입'] = ', '.join(new_a)
else :
    checkDic['9.5.1.직전기에만있는매입'] = 'N/A'

checkDF = pd.DataFrame(checkDic, index=['내용']).T
checkDF.index.name = '구분'


'''★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★'''
#★★★★★★★★★엑셀에 각종서식 적용하려했는데 시트이름이 한글이면 시트를 못 읽어서 일단 영문으로 변경하고 이후에 다시 한글로 수정★★★★★★
'''★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★'''

#sheetsList1 = ['검토사항', '중복체크', '합계표', '유형별_합계표', '종이세금검토(출력)', '발송용', '당기자료', '직전기합계표', '직전기유형별합계표', '직전기세액', '직전기자료','부가율','평균부가율등']
sheetsList1 = ['검토사항', '중복체크', '합계표', '유형별_합계표', '종이세금검토(출력)', '발송용', '당기자료', '직전기합계표', '직전기유형별합계표', '직전기세액', '직전기자료']
sheetsList2 = ['worksheet'+str(num) for num in range(len(sheetsList1))]
#['worksheet0', 'worksheet1', 'worksheet2', 'worksheet3', 'worksheet4', 'worksheet5', 'worksheet6', 'worksheet7', 'worksheet8', 'worksheet9']
sheetsDict1 = dict(zip(sheetsList1, sheetsList2))
#{'종이세금검토(출력)': 'worksheet4', '합계표': 'worksheet2', '당기자료': 'worksheet6', '직전기합계표': 'worksheet7', '검토사항': 'worksheet0', '직전기자료': 'worksheet9', '발송용': 'worksheet5', '중복체크': 'worksheet1', '유형별_합계표': 'worksheet3', '직전기세액': 'worksheet8'}
sheetsDict2 = dict(zip(sheetsList2, sheetsList1))
#{'worksheet1': '중복체크', 'worksheet3': '유형별_합계표', 'worksheet4': '종이세금검토(출력)', 'worksheet8': '직전기세액', 'worksheet6': '당기자료', 'worksheet7': '직전기합계표', 'worksheet5': '발송용', 'worksheet0': '검토사항', 'worksheet2': '합계표', 'worksheet9': '직전기자료'}
  
option = 3 # 1 : 한글시트이름 처음부터적용, 2 : 영문시트이름 처음부터적용, 3 : 영문 → 한글
  
if option == 3 :  #worksheet 객체 혼선 방지위해 직전에 생성된 객체(변수) 삭제함
    my_dir = dir() # 모든변수를 리스트로 저장
    for i in sheetsList2: #worksheet가 직전기 실행으로 변수로 정의되어있다면 변수를 삭제해야 알고리즘에 영향이 없다
        try :
            #print(i,'→ 변수 →',eval(i))
            pass
        except NameError as e:
            #print(e)     
            pass
            
        if i in my_dir:
            del locals()[i] #루프안에서 변수를 삭제할수 있다
            try :
                #print(i,':',eval(i))
                pass
            except NameError as e:
                #print(e)  
                pass
    #my_dir2 = dir()  
groupTotal1['공급대가'] = groupTotal1['공급가액'] + groupTotal1['부가세']  #20200725

if 10 : #20200714  asdf  
    
    if option == 1 : #'데이터프레임을 시트이름정해서 엑셀로 저장 to_excel': #엑셀에 포맷 적용하려고하면 시트이름이 한글이면 안되어서 영문으로
        
        checkDF.to_excel(writer, '검토사항') #삭제하지말것
        if not 중복체크그룹.empty:
            중복체크그룹.to_excel(writer, '중복체크')
        concated.to_excel(writer, '합계표')
        groupTotal1.to_excel(writer,'유형별_합계표')
        dff2.to_excel(writer, '종이세금검토(출력)')
        send1.to_excel(writer, '발송용')
        df6.to_excel(writer, '당기자료')
        if not priorYear1.empty:
            priorGroup.to_excel(writer, '직전기합계표')
            priorGrptotal.to_excel(writer, '직전기유형별합계표')
            gb.to_excel(writer, '직전기세액')
        priorYear1.to_excel(writer, '직전기자료')    
    
    elif option == 2 :
        
        checkDF.to_excel(writer, 'worksheet0') #삭제하지말것
        if not 중복체크그룹.empty:
            중복체크그룹.to_excel(writer, 'worksheet1')
        concated.to_excel(writer, 'worksheet2')
        groupTotal1.to_excel(writer,'worksheet3')
        dff2.to_excel(writer, 'worksheet4')
        send1.to_excel(writer, 'worksheet5')
        df6.to_excel(writer, 'worksheet6')
        if not priorYear1.empty:
            priorGroup.to_excel(writer, 'worksheet7')
            priorGrptotal.to_excel(writer, 'worksheet8')
            gb.to_excel(writer, 'worksheet9')
        priorYear1.to_excel(writer, 'worksheet10')    
        
        
    elif option == 3 :
             
        checkDF.to_excel(writer, sheetsDict1['검토사항'])
        if not 중복체크그룹.empty:
            중복체크그룹.to_excel(writer, sheetsDict1['중복체크'])
        concated.to_excel(writer, sheetsDict1['합계표'])
        groupTotal1.to_excel(writer,sheetsDict1['유형별_합계표'])
        dff2.to_excel(writer, sheetsDict1['종이세금검토(출력)'])
        send1.to_excel(writer, sheetsDict1['발송용'])
        df6.to_excel(writer, sheetsDict1['당기자료'])
        if not priorYear1.empty:
            priorGroup.to_excel(writer, sheetsDict1['직전기합계표'])
            priorGrptotal.to_excel(writer, sheetsDict1['직전기유형별합계표'])
            gb.T.to_excel(writer, sheetsDict1['직전기세액'])
        priorYear1.to_excel(writer, sheetsDict1['직전기자료'])
        #dff5.to_excel(writer, sheetsDict1['부가율'])
        #grouped.to_excel(writer, sheetsDict1['평균부가율등'])
        
    
    
            
        
    if option != 1 : #'엑셀시트에 서식적용 writer.sheets':
        
        if '포맷 설정':
            # Get the xlsxwriter workbook and worksheet objects.
            workbook  = writer.book
            
            # Add some cell formats.
            format1 = workbook.add_format({'num_format': '#,##'})
            format2 = workbook.add_format({'align':'right'})
            format3 = workbook.add_format({'align':'center'})
            REDformat = workbook.add_format({'color':'red'})
            red = workbook.add_format({'color': 'red'})
            format4 = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'color': 'blue'})    
            format5 = workbook.add_format({'text_wrap': True})    
            format6 = workbook.add_format({'shrink': True})   
            format7 = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
            format8 = workbook.add_format({'valign': 'vcenter'})
            
            numformat_vcenter = {'num_format': '#,##', 'valign': 'vcenter'}
            numformat_top = {'num_format': '#,##', 'valign': 'top'}
            numformat_top_shrink = {'num_format': '#,##', 'valign': 'top', 'shrink': True}
            numformat_shrink = {'num_format': '#,##', 'shrink': True}
            #REDformat = workbook.add_format({'color':'red'})
            #format4 = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'color': 'blue'})    
            #format5 = workbook.add_format({'text_wrap': True})
            border1 = workbook.add_format({'border':1})
            
        comName = 회사정보사전['회사명']; comNum = 회사정보사전['사업자등록번호']
        
        worksheet0 = writer.sheets[sheetsDict1['검토사항']]
        worksheet0.add_table(0, 0, len(checkDF), 1, {'style': 'Table Style Light 14','autofilter': False, 'header_row': False}) # 주황 밑줄
        worksheet0.set_column('A:A', 22, format4)
        worksheet0.set_column('B:B', 52, format5)
        worksheet0.write('B5',천단위콤마함수(int(납부세액))+'원',REDformat)
        vatRatio = checkDic['2.4.1.2.평균부가율등']
        #if vatRatio.find('평균값') != -1:
        #if checkDF.ix[12].str[:5].isin(['▲ 주업종']):
        if checkDF.ix[12,0][:5] in ['▲ 주업종'] :
            #worksheet0.write_rich_string('B14', vatRatio[:vatRatio.find('평균값')], red, vatRatio[vatRatio.find('평균값'):vatRatio.find('평균값')+12], vatRatio[vatRatio.find('평균값')+12:])
            worksheet0.write_rich_string(13, 1, vatRatio[:vatRatio.find('평균값')], red, vatRatio[vatRatio.find('평균값'):vatRatio.find('평균값')+12], vatRatio[vatRatio.find('평균값')+12:])
        
    
        worksheet5 = writer.sheets[sheetsDict1['발송용']]
        #worksheet5.add_table(0, 3, len(send1), 6, {'style': 'Table Style Light 15', 'autofilter': False, 'header_row': False})
        worksheet5.add_table(0, 3, len(send1), 6, {'style': 'Table Style Light 20', 'autofilter': False, 'header_row': False})
        #worksheet5.add_table(0, 3, len(send1), 6, {'header_row': False})
        worksheet5.set_column('G:G', 50, format5)
        worksheet5.set_column('E:E', 5, workbook.add_format(numformat_top))
        worksheet5.set_column('F:F', 10, workbook.add_format(numformat_top_shrink))
        worksheet5.set_column('C:C', 15)
        worksheet5.set_column('D:D', 13)
        worksheet5.set_landscape()
        worksheet5.set_column('A:B', 5)
        worksheet5.set_header('&L'+comName+'('+comNum+')'+'&C<<매입매출장>>&R &D &T')
        worksheet5.set_footer('&L※ 위 금액은 부가세가 제외된 금액입니다.  ※ 상기 매입매출장은 매출전체와 매입세금계산서·계산서 내역만 표시됩니다.&RPage &P of &N')
    
        worksheet2 = writer.sheets[sheetsDict1['합계표']]
        #worksheet = writer.sheets['Aggregate']
        worksheet2.set_column('I:J', 12, format1)
        worksheet2.set_column('E:E', 17)
        worksheet2.set_column('F:F', 14)
        worksheet2.set_column('B:B', None, None, {'hidden': True})
        worksheet2.set_column('G:H', 5)
        worksheet2.set_column('D:D', 7)
        
        worksheet3 = writer.sheets[sheetsDict1['유형별_합계표']]
        worksheet3.set_column('G:H', 12, format1)
        
        if not 중복체크그룹.empty:
            worksheet1 = writer.sheets[sheetsDict1['중복체크']]
            worksheet1.add_table(0, 1, len(중복체크그룹), 8, {'style': 'Table Style Light 20', 'autofilter': False, 'header_row': False})
            worksheet1.set_column('F:G', None, workbook.add_format(numformat_shrink))
            worksheet1.set_column('D:D', None, format3)
            worksheet1.set_column('H:H', None, format3)
            worksheet1.set_column('A:B', 19)
            worksheet1.set_column('E:E', 12)
            worksheet1.set_column('I:I', None, format2)
            worksheet1.set_column('F:F', 12)
    
        worksheet4 = writer.sheets[sheetsDict1['종이세금검토(출력)']]
        worksheet4.add_table(0, 1, len(dff2), 10, {'style': 'Table Style Light 15','autofilter': False, 'header_row': False})
        #worksheet4.set_column('B:K', None, border1)                                           
        worksheet4.set_column('A:A', 6, format3)                                           
        worksheet4.set_column('C:E', 6, format3)
        worksheet4.set_column('B:B', 8, format3)
        worksheet4.set_column('F:F', 12)
        worksheet4.set_column('G:G', 15, format6)
        worksheet4.set_column('H:J', 12, workbook.add_format(numformat_shrink))
        worksheet4.set_column('I:I', 8, workbook.add_format(numformat_shrink))
        worksheet4.set_column('K:K', None, format3)
        worksheet4.set_landscape()
        #worksheet4.set_row(0, None, format5)
        #worksheet4.set_footer('&C Page &P of &N')
        #worksheet4.set_footer('&R &D &T')
        worksheet4.set_footer('&CPage &P of &N &R &D &T')
        #comName = 회사정보사전['회사명']; comNum = 회사정보사전['사업자등록번호']
        worksheet4.set_header('&C'+comName+'('+comNum+')')
        worksheet4.write('K1', '의제')
        
    if option == 3 : #시트이름을 영문 → 한글로 수정함
        
        #직전기자료가 없으면 worksheet7이 없는데 막무가내로 worksheet7을 정의하면 에러가 발생하므로 아래와같이 엑셀에서 시트이름을 불러와야함
        for idx, val in enumerate(list(writer.sheets)): #enumerate는 순회시 문자열 따옴표가 없음
            #print(val)
            if not val in vars(): #본문중에 서식적용위해 worksheet를 정의했는지 그래서 변수가 되었는지 #  worksheet객체를 사용하기전에 미리 정의해둠
                #print(idx, val)
                #exec("worksheet%d = None"%idx)
                #print("%s = writer.sheets['%s']"% (val, val)) #enumerate는 순회시 문자열 따옴표가 없음; worksheet6 = writer.sheets['worksheet6']
                exec("%s = writer.sheets['%s']"% (val, val)) #변수로 만들어줘야(정의해줘야) 나중에 변수이름으로 접근해서 시트이름을 변경할수 있다.         
    
        for i in list(writer.sheets): #시트이름을 영문에서 한글로 수정함
            if i in vars(): # i가 변수라면, 정의되어 있다면
                #print(i, i in vars())
                eval(i).name = sheetsDict2[i] # worksheet1 = writer.sheets['worksheet1'] 에서 i는 첫번째 worksheet1임; 두번째 worksheet1은 엑셀시트의 이름; eval('worksheet1') → worksheet1
    
    #Mbox('leecta', '[매입매출검토]폴더에 저장했습니다', 1)
    
    #if True : # Pie plot; legend; 도표
    if False : # Pie plot; legend; 도표 #20190719 직원컴에서 에러발생함
    
        import matplotlib.pyplot as plt
         
        # Data to plot
        if 0 :
            labels = '매출', '매입세금계산서등', '매입신카등', '의제매입'
        else :
            labels = 'maeChul', 'maeIpT/I', 'maeIpCard', 'uiJe'
        #labels = 'mae_chul', 'mae_ip T/I', 'shin_ca', 'ui_je'
        sizes = [매출공급가액, 매입공급가액 - 의제류대상금액 - 매입신카등공급가액, 매입신카등공급가액, 의제류대상금액]
        colors = ['gold', 'yellowgreen', 'lightcoral', 'lightskyblue']
        '''explode = (0, 0, 0.1, 0)  # explode 1st slice
         
        # Plot
        plt.pie(sizes, explode=explode, labels=labels, colors=colors,
                autopct='%1.1f%%', shadow=True, startangle=140)
         
        plt.axis('equal')
        plt.show()
        '''
        매입매출 = 매출공급가액 + 매입공급가액
        비율 = [매출공급가액/매입매출, (매입공급가액 - 의제류대상금액 - 매입신카등공급가액)/매입매출, 매입신카등공급가액/매입매출, 의제류대상금액/매입매출]
        비율2 = [소숫점둘째반올림함수(x*100)+'%' for x in 비율]
        #print(비율2)       
        
        #labels = ['Cookies', 'Jellybean', 'Milkshake', 'Cheesecake']
        #labels = ['매출', '매입세금계산서등', '매입신카등', '의제매입']
        #sizes = [38.4, 40.6, 20.7, 10.3]
        #sizes = [매출공급가액, 매입공급가액 - 의제류대상금액 - 매입신카등공급가액, 매입신카등공급가액, 의제류대상금액]
        #colors = ['gold', 'yellowgreen', 'lightcoral', 'lightskyblue']
        #colors = ['yellowgreen', 'gold', 'lightskyblue', 'lightcoral']
        patches, texts = plt.pie(sizes, colors=colors, shadow=True, startangle=90)
        #plt.legend(patches, labels, loc="best") # 도표
        plt.legend(patches, labels, loc="upper left") # 도표
        plt.axis('equal')
        plt.tight_layout()
        #plt.show()
        
        plt.rc('font', family='NanumGothic') # For Windows
        #plt.rc('font', family='DejaVuSans') # For Windows
        #plt.rc('font', family='normal') # For Windows
        print(plt.rcParams['font.family'])
        #[출처] [Python] matplotlib 한글폰트 설정방법|작성자 똑똑이
    
    
        #import openpyxl
    
        # Your plot generation code here...
        plt.savefig("myplot.png")#, dpi = 150) 
        worksheet3.insert_image('B17', 'myplot.png',{'x_scale': 0.7, 'y_scale': 0.7, 'x_offset': 0, 'y_offset': 8,})
        #worksheet3.insert_textbox('A38', ' // '.join(비율2))
        options = {'width': 80, 'height': 80}
        worksheet3.insert_textbox('A18', '\n'.join(비율2), options)
        #wb = openpyxl.load_workbook('input.xlsx')
        #ws = wb.active
        
        
        #img = openpyxl.drawing.Image('myplot.png')
        #img.anchor(ws.cell('A1'))
        
        #ws.add_image(img)
            
    writer.save()
    
    print(convertingTime, 'converted from', base1)

    
if 0 :
    Mbox('leecta', '[매입매출검토]폴더에 저장했습니다', 1)
    
    

end = time.time()
endTime = time.strftime("%H_%M_%S")

print(startTime, '→', endTime, ':', str(round(end - start, 2))+'초')

# 20200911 python C:\Users\hana1602a\PythonPractice-201708\200714.py