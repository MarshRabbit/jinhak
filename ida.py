from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import time
import pyperclip

"""
사용 전 확인해야 할 것들
txt파일 저장할 경로, 크롬 드라이버 경로, 윈도우 맥 tag_.send_keys 변경, 엑셀파일 경로, 엑셀 시트 이름
"""

def create_txt(name, univ, score):
    filename = "/Users/qlover/Desktop/1/" + name + ".txt"  # txt파일 저장 경로 name 앞쪽것만 변경
    f = open(filename, "w")
    for i in range(len(univ)):
        univ_name = univ[i].string
        f.write(univ_name.split()[0])
        f.write(" ")
        f.write(univ_name.split()[1])
        f.write(" ")
        f.write(univ_name.split()[3])
        f.write("\n")
    f.close()

def crawling(name, id, pw):
    try:
        driver = webdriver.Chrome('/Users/qlover/Downloads/chromedriver')   # 드라이버 저장 경로

        url = 'https://member.jinhak.com/MemberV3/m/Login.aspx?ReturnSite=MJ&ReturnURL=http%3a%2f%2fm.jinhak.com%2fJ1Apply%2fJ1MyApplyList.aspx'
        driver.get(url)
        driver.implicitly_wait(10)
        # 경고문 제거
        # driver.find_element_by_xpath('//*[@id="idShareNotice"]/div/span/label').click()
        # driver.find_element_by_xpath('//*[@id="idShareNotice"]/div/div/a').click()
        time.sleep(1)
        
        if len(id.split()) > 1 and '네이버' in id.split()[1]: #네이버로그인
            driver.find_element_by_xpath('//*[@id="naverIdLogin_loginButton"]').click()
            driver.implicitly_wait(10)
            time.sleep(1)
            driver.find_element_by_xpath('//*[@id="log.login"]').click()
            time.sleep(1)
            
            tag_id = driver.find_element_by_name('id')
            tag_pw = driver.find_element_by_name('pw')
            tag_id.clear()
            time.sleep(1)
            
            tag_id.click()
            pyperclip.copy(id.split()[0])
            tag_id.send_keys(Keys.COMMAND, 'v') #윈도우는 Keys.CONTROL
            time.sleep(1)
            
            tag_pw.click()
            pyperclip.copy(pw)
            tag_pw.send_keys(Keys.COMMAND, 'v') #윈도우는 Keys.CONTROL
            time.sleep(1)
        
            driver.find_element_by_xpath('//*[@id="log.login"]').click()
            driver.implicitly_wait(10)
            driver.get('https://m.jinhak.com/J1Apply/J1MyApplyList.aspx')
    
        else: #진학사 로그인
            driver.find_element_by_name('txtMemID').send_keys(id)
            driver.find_element_by_name('txtMemPass').send_keys(pw)
            driver.implicitly_wait(1)
            driver.find_element_by_xpath('//*[@id="form1"]/div/div/div[1]/div[4]/button').click() 
        
        time.sleep(1)
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        table = soup.select_one('div.result_list')

        univ_list = table.select('div.result > div.bottom > div.info > p.name > a > b')
        score_list = table.select('div.bottom > div.info > p:nth-child(2)')
        spaces_list = table.select('div.top > span.t_pass')

        create_txt(name, univ_list, score_list)
        print("대학목록 생성완료")
        time.sleep(1)
        
        filename = "/Users/qlover/Desktop/1/" + name + "_more.txt" # txt 파일 저장 경로 name 앞쪽것만 변경
        f = open(filename, "w")
        for i in range(len(univ_list)):
            try:
                link = '//*[@id="form1"]/div[6]/div[' + str(i+1) + ']/div[3]/a[1]'
                driver.find_element_by_xpath(link).click()
                driver.implicitly_wait(7)

                html = driver.page_source
                soup = BeautifulSoup(html, 'html.parser')
                
                # 작년경쟁률 모집인원 이월인원 칸수 등수 지원자수 예상경쟁률 최초추가인원 
                last_Ratio = soup.select_one('#rWrapper > div.rUnivRate.atjungsi > ul > li:nth-child(4) > dl > dd')
                recruiting = soup.select_one('#rWrapper > div.rUnivRate.atjungsi > ul > li:nth-child(1) > dl > dd')
                carried_Over = soup.select_one('#rWrapper > div.rUnivRate.atjungsi > ul > li:nth-child(2) > dl > dd')
                rank = soup.select_one('#rWrapper > div.rBg3 > div:nth-child(3) > div > div.pass_safety2 > div.analysis > div > div:nth-child(2) > span > div > p > small')
                spaces = spaces_list[i].string.split(':')[1]
                competitive_Rate = soup.select_one('#rWrapper > div.rBg3 > div:nth-child(1) > div > div.summaryBox.jungsi.mt25 > div:nth-child(2) > span > div > p')
                admissions = soup.select_one('#rWrapper > div.rBg3 > div:nth-child(5) > div > div.detailBox > table:nth-child(3) > tbody > tr:nth-child(8) > td > p')
                last_admissions = soup.select_one('#rWrapper > div.rBg3 > div:nth-child(6) > div > div.detailBox > div:nth-child(2) > table > tbody > tr:nth-child(1) > td:nth-child(3)')
                last_over = soup.select_one('#rWrapper > div.rBg3 > div:nth-child(6) > div > div.detailBox > div:nth-child(2) > table > tbody > tr:nth-child(1) > td:nth-child(8)')
                
                f.write(univ_list[i].string.split()[0])
                f.write(" ")
                f.write(univ_list[i].string.split()[1])
                f.write(" ")
                f.write(univ_list[i].string.split()[3])
                f.write("\n")
                f.write("작년경쟁률:")
                f.write(last_Ratio.string)
                f.write("\n")
                f.write("모집인원: ")
                f.write(recruiting.string)
                f.write(" 이월인원: ")
                f.write(carried_Over.string)
                f.write("\n")
                f.write("작년모집: ") ##
                f.write(last_admissions.string)
                f.write(" 작년추합: ")
                f.write(last_over.string)
                f.write("\n") ##
                f.write(last_Ratio.string)
                f.write(" ")
                f.write(last_over.string)
                f.write(" ")
                f.write(last_admissions.string)
                f.write(" ")
                f.write(recruiting.string)
                f.write(" ")
                f.write(carried_Over.string)
                f.write("\n")
                f.write("칸수:")
                f.write(spaces)
                f.write(" 등수: ")
                f.write(rank.string.split()[1])
                f.write(" 지원자수: ")
                f.write(rank.string.split()[3].split(')')[0])
                f.write(" 예상경쟁률: ")
                f.write(competitive_Rate.string)
                f.write(" 최초인원: ")
                f.write(admissions.string.split('+')[0].split('최초')[1])
                f.write(" 추가인원: ")
                f.write(admissions.string.split('+')[1].split(')')[0].split('추가')[1])
                f.write("\n")
                f.write(spaces.split()[0])
                f.write(" ")
                f.write(rank.string.split()[1])
                f.write(" ")
                f.write(rank.string.split()[3].split(')')[0])
                f.write(" ")
                f.write(competitive_Rate.string)
                f.write(" ")
                f.write(admissions.string.split('+')[0].split('최초')[1])
                f.write(" ")
                f.write(admissions.string.split('+')[1].split(')')[0].split('추가')[1])
                f.write("\n\n")
                
                driver.back()
            except:
                f.write("\n현재 분석중\n\n")
                print("현재 분석중이므로 error")
                driver.back()
                continue
        
        f.close()
        driver.close()
        print("%s 종료" %name)
    except:
        print("error ", name)
    

load_wb = load_workbook("/Users/qlover/Desktop/알바/모음집.xlsx", data_only=True) #엑셀파일 저장 경로
load_ws = load_wb['Sheet1'] # 엑셀 시트명

names = [] 
ids = []
pws = []

get_names = load_ws['C1':'C10'] #이름 열
for row in get_names:
    for cell in row:
        names.append(cell.value)
        

get_ids = load_ws['D1':'D10']   #아이디 열
for row in get_ids:
    for cell in row:
        ids.append(cell.value)
        
get_pws = load_ws['E1':'E10']   #비밀번호 열
for row in get_pws:
    for cell in row:
        pws.append(cell.value)

for i in range(len(names)):
    if ids[i] is None:
        continue
    print("%s 시작" %names[i])
    crawling(names[i], ids[i], pws[i])
