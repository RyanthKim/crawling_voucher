import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import datetime
import openpyxl as xl

print('')
print(' 안녕하세요. 비대면 바우처 결제관리 내역 크롤링 입니다. - ryan')
print(' 제가 한땀한땀 작성한 코드로 정상작동하지 않는다면 수정이 필요합니다.')
print('')
print(' ※  중요 ※')
print(' 자동으로 실행되는 크롬브라우저를 닫지 말아주세요.')
print(' 작업이 종료되면 자동실행된 브라우저는 알아서 꺼집니다.')
print(' 다른 업무를 보고 계셔도 괜찮습니다!')

wb = xl.Workbook()
sheet1 = wb.active
sheet1.append(['아이디','수요기업명','사업자번호','대표자명','대표자연락처','대표번호','담당자 성명','담당자 연락처','담당자 휴대전화','담당자 이메일','주문번호','결제일시','상품명','옵션타입','상품고유번호','서비스분야','결제수단','결제상태','최종 결제 금액'])

id = 'ID',
input_pass = '비밀번호'
print('')
print('  1) 크롬 실행중')
options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])
wd = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wd.set_window_position(0, 800)
wd.set_window_size(1300, 1000)

print('')
print('  2) 페이지 접속중')
wd.get('https://supply.k-voucher.kr/webadm/login')
wd.find_element(by=By.CLASS_NAME,value = 'input_id').send_keys(id)
wd.find_element(by=By.CLASS_NAME,value = 'input_pw').send_keys(input_pass)
wd.find_element(by=By.CLASS_NAME,value = 'input_pw').send_keys(Keys.ENTER)

wd.get('https://supply.k-voucher.kr/webadm/sup_payment_management?mode=list&numPerPage=100&startPage=1')
print('  3) 5초 후 실행됩니다. 잠시 기다려주세요')
time.sleep(5)

l = int(wd.find_element(by=By.XPATH, value='//*[@id="crudTotalCount"]').text)
m = (l+100) // 100
p = 1
print('  4) 크롤링이 진행됩니다.')
print('')
print('   총 {0}개의 raw가 존재합니다.'.format(l))
print('')
# print('\r    - {0}/{1}'.format(2,3),' 진행중',end='') #test 용
for n in range(1,m+1):
    q = 1
    while q <10:
        try :
            if n != 1:
                time.sleep(3)
                wd.get('https://supply.k-voucher.kr/webadm/sup_payment_management?mode=list&numPerPage=100&startPage={0}'.format(n))
            # print('1_try T ] q : ',q,' n : ',n) #크롤링 검수용 로그
            for o in range(1,100+1):
                if p == l+1:
                    break
                print('\r      - {0}/{1}'.format(p,l),' 진행중',end='')
                temp = []
                q = 1
                while q <10:
                    try:                                   
                        wd.find_element(by=By.XPATH, value='//*[@id="crudList"]/div/table/tbody/tr[{0}]/td[1]'.format(o)).click()
                        # print('2_try T ] q : ',q,' o : ',o) #크롤링 검수용 로그
                        q = 1
                        while q <10:
                            try:
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[1]/div[2]/div[1]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[1]/div[2]/div[2]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[1]/div[2]/div[3]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[1]/div[2]/div[4]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[1]/div[2]/div[5]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[1]/div[2]/div[6]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[1]/div[2]/div[7]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[1]/div[2]/div[8]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[1]/div[2]/div[9]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[1]/div[2]/div[10]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[2]/div[2]/div[1]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[2]/div[2]/div[2]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[2]/div[2]/div[3]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[2]/div[2]/div[4]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[2]/div[2]/div[5]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[2]/div[2]/div[6]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[2]/div[2]/div[7]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[2]/div[2]/div[8]/div[2]/div').text)
                                temp.append(wd.find_element(by=By.XPATH, value='//*[@id="sidebar_parent"]/div/div/div/div[1]/div[2]/div[2]/div[9]/div[2]/div').text)
                                wd.get('https://supply.k-voucher.kr/webadm/sup_payment_management?mode=list&numPerPage=100&startPage={0}'.format(n))
                                # wd.find_element(by=By.CLASS_NAME,value = 'bbs_m_btn_2').click()
                                # print('3_try T ] q : ',q) #크롤링 검수용 로그
                                break
                            except:
                                q = q+1
                                # print('3_try F ] q : ',q) #크롤링 검수용 로그
                                time.sleep(1)
                            if q == 9:
                                break
                        break
                    except:
                        q = q+1
                        # print('2_try F ] q : ',q,' o : ',o) #크롤링 검수용 로그
                        time.sleep(1)
                    if q == 9:
                        break
                p = p+1
                sheet1.append(temp)
            break
        except:
            q = q+1
            time.sleep(1)
            # print('1_try F] q : ',q,' n : ',n) #크롤링 검수용 로그
        if q == 9:
            break
now = datetime.datetime.now()
nowDatetime = str(now.strftime('%Y%m%d%H%M%S'))
wb.save('비대면바우처 결제관리_crawling({0}).xlsx'.format(nowDatetime))
print('')
print('')
print('  [내보내기 완료]')
print('')
print(' ※  모든 창을 종료해주셔도 됩니다.')
print(' ※  크롤링 파일이 있는 폴더에 엑셀 파일을 참고 부탁드립니다.')
wd.close()