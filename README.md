# 진학사 입시정보 스크래핑

대학 입시 알바를 하게되어 학생의 진학사의 모의지원 정보를 저장해 표로 만드는 업무를 하게되었다.

희망 대학 리스트가 각 학생마다 20개씩 되어 손아픔 + 귀차니즘으로 파이썬을 이용해 한번에 스크래핑하여 텍스트파일로 저장하는 프로그램을 만들어 보기로 했다

## 전체적 아이디어

우선 학생들의 아이디와 비밀번호를 저장한 엑셀파일을 openpyxl로 가져와 저장하고 selenium을 이용하여 브라우저를 자동으로 제어해 로그인하고 bs4로 데이터를 파싱하여 정보를 가져올 생각이다.

그리고 가져온 정보들을 엑셀에 복붙하기 편한 형태로 나열하여 txt 파일로 저장하기로 했다

## 세부적 코드

- 로그인시 필요한 정보
  

```python
load_wb = load_workbook("/Users/qlover/Desktop/알바/모음집.xlsx", data_only=True)
load_ws = load_wb['Sheet1']

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
```

엑셀파일을 불러와 내용을 저장하고 이름, 아이디, 비밀번호를 담을 배열을 만든다음에 집어넣었다.

- 로그인하기
  

```python
driver.find_element_by_name('txtMemID').send_keys(id)
driver.find_element_by_name('txtMemPass').send_keys(pw)
driver.implicitly_wait(1)
driver.find_element_by_xpath('//*[@id="form1"]/div/div/div[1]/div[4]/button').click()
```

기본적인 로그인은 아이디와 비밀번호를 입력하는 칸의 속성 이름을 따와 키값을 넣고 로그인 버튼은 속성 이름이 없어 xpath 경로를 이용하여 클릭하게 하였다.

그런데 세상일이 그렇게 단순하진 않은지라 진학사 아이디를 만들지 않고 네이버 아이디를 이용해 로그인 하는 학생이 있었다.

- 네이버로 로그인
  

```python
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
```

네이버 아이디를 제공한 학생은 엑셀에 '학생아이디 (네이버로그인)' 형태로 저장되어있었다

고로 id의 저장된 값이 기본적으론 단어 하나지만 네이버로 로그인시엔 2마디 이상이므로 id를 split으로 나눴을때 2개 이상이고 내용에 '네이버'가 들어가면 분기하도록 하였다

처음 로그인화면에서 네이버 버튼을 누르면 네이버 로그인화면이 나온다. 그런데 여기서 앞에 썻던것 처럼 로그인하면 자동 프로그램이 로그인 하는것을 감지하는 것인지 캡챠를 입력하라는 칸이 새로 나왔다

하지만 로그인 정보를 입력할때 복사 붙여넣기를 사용하면 캡챠가 뜨지 않는다는 것을 이용하여 pyperclip으로 복붙하면서 정상적으로 로그인할 수 있었다

추가적으로 카카오로 로그인하기도 있었지만 그걸로 로그인하는 학생이 없어서 따로 구현하진 않았다

- 희망 대학 목록들
  

div.result_list > div.result > div.bottom > div.info > p.name > a > b 안에 학생이 희망한 대학이름과 학과가 모여있었다

그래서 리스트에 저장하고 .string()으로 문자열로 바꾼다음에 spilt으로 나누어서 학교명과 학과를 출력하였다.

- 모의지원 정보들
  

모의지원 정보를 보려면 저장한 희망대학에서 모의지원 정보 보기 버튼을 누르면 해당 정보가 있는 페이지로 이동하게 된다. 그 버튼의 xpath는 //*[@id="form1"]/div[6]/div[n]/div[3]/a[1]
이었고 n은 1부터 1씩 증가하는 형태였다.

고로 반복문으로 i값을 증가시켜 가면서 link = '//*[@id="form1"]/div[6]/div[' + str(i+1) + ']/div[3]/a[1]'로 링크를 만들어서 버튼을 클릭해 모의지원 페이지로 이동하였다.

그후 필요한 정보들을 css 선택자로 뽑아내 메모장으로 출력하였다.

## 후기

웹 스크래핑은 처음 도전해보는거라 오류가 생각보다 많이 나왔다.
로그인 화면까진 잘 가다가 로그인 과정에서 먹통이나서 디버깅으로 한번 브레이크 잡고 이어서 실행해야 정상적으로 실행되는 경우가 있었다

몇몇 학생은 홈페이지에서 알람이 나와 똑같은 프로그램을 하나 더 만들어 확인 버튼을 따로 눌러주는 코드를 추가해 그 학생들만 따로 돌리게 하였는데 지금와서 생각해보니 그냥 간단한 조건문으로 한번에 할 수 있는걸 굳이 두개로 나눴다고 생각이 든다.

그리고 생각보다 코드가 지저분하게 짜여진것 같다. 특히 파일 쓰는 부분에서 출력해야할 변수가 많은데 정리를 하지 못해 너무 길어졌다거나 이부분을 따로 함수로 만들지 않아 통일성을 해친거 같다.

조금 야매로 짜여진 감이 있지만 그래도 처음 도전하는것 치고 생각보다 잘 돌아가서 다행이었다. 엑셀 파일을 직접 수정하는 부분은 시간이 그렇게 많지않아서 건들게 많아 만들지는 못했지만 그래도 다른 사람들과 비교해 작업 속도가 2배 정도 빨랐고 손도 그렇게 많이 쓰지 않아서 편안해 만족스러웠다.
