from lib2to3.pgen2 import driver
from flask import appcontext_tearing_down
from numpy import NaN
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import pandas as pd
import time
import os
import glob
import shutil
from PyPDF2 import PdfFileMerger



CHROME_PATH = ""
URL_PATH = ""
ID = ""
PASS = ""
TARGET_Y = "2022/01"
TARGET_M = "2022/03/01"
DOWNLOAD_FILE_PATH = os.getenv('HOME') + '/Downloads/Works/dl'
OUTPUT_FILE_PATH = os.getenv('HOME') + '/Downloads/Works/out'

#指定要素までページをスクロールする
def pgScroll(br,target_elem):
    actions = ActionChains(br)
    actions.move_to_element(target_elem)
    actions.perform()
    time.sleep(1)
    return actions

#ローディングモーダルが消えるまで待機する。ただし60秒でタイムアウト
def waitLoading(wait, br):
    loadingCnt = 0
    if len(br.find_elements_by_id('loadingDiv')) == 0 :
        return True
    time.sleep(3)
    while True :
        if loadingCnt > 59 :
            print('Loading Timed out !')
            return False
        
        print('Loading ...')
        _loading = br.find_elements_by_id('loadingDiv')
        if len(_loading) > 0 :
            time.sleep(3)
            loadingCnt += 3
            continue
        else :
            break
    
    return True


#集計ボタンまで移動し、クリックして集計モーダルを表示させる
def showSyukeiModal(wait,br):
    wait.until(EC.element_to_be_clickable((By.ID,'aggregateBtn')))
    _syukeiBtn = br.find_element_by_id('aggregateBtn')
    # pgScroll(br, _syukeiBtn)
    br.execute_script('window.scrollTo(0,0)')
    time.sleep(1)
    _syukeiBtn.click()
    time.sleep(1)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,'#makeAndShowModal button.yearly-btn')))
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR,'#makeAndShowModal button.yearly-btn')))

#集計表画面を表示させ、集計表をダウンロードする
def download(wait, br):
    #集計画面表示完了まで待機
    wait.until(EC.presence_of_element_located((By.ID,'main_div')))
    time.sleep(1)
    isAlert = alertCheck(wait, br)
    #ダウンロードボタン
    _btns = br.find_elements_by_css_selector("#main_div button")
    for _btn in _btns :
        if '保存' in _btn.text:
            pgScroll(br, _btn)
            time.sleep(1)
            #集計表をダウンロード
            print('ダウンロード中...')
            _btn.click()
            time.sleep(3)
            waitLoading(wait, br)
            break
    
    return isAlert

#月報の各集計表モーダルを設定
def get_Geppou(browser, id_name, checkReq) :
    browser.execute_script('window.scrollTo(0,0)')
    #月報集計    
    browser.find_element_by_id('monthPaperBtn').click()
    #集計種別を選択
    browser.find_element_by_id(id_name).click()
    #対象月を入力
    _inputM = browser.find_element_by_id('chooseOneStartDateIM')
    if _inputM.get_attribute('value') != TARGET_M :
        _inputM.send_keys(TARGET_M)
    #請求にチェック
    # //*[@id="aggregatemenuComboDiv2"]/div/div[1]/label/input
    if checkReq and not browser.find_element_by_id('requestInput').is_selected :
        browser.find_element_by_id('requestInput').click()
    #グラフを表示ボタン
    _graphBtn = browser.find_element_by_id('graphButton')
    pgScroll(browser, _graphBtn)
    time.sleep(1)
    _graphBtn.click()
    time.sleep(3)

#年報の各集計表モーダルを設定
def get_Nenpou(browser) :
    browser.execute_script('window.scrollTo(0,0)')
    #年報集計    
    browser.find_element_by_css_selector('#makeAndShowModal button.yearly-btn').click()
    #集計種別を選択
    browser.find_element_by_id("saleAggreBtn").click()
    #対象年を入力
    _inputY = browser.find_element_by_id('chooseOneMonI')
    if _inputY.get_attribute('value') != TARGET_Y :
        _inputY.send_keys(TARGET_Y)
    #請求にチェック
    if not browser.find_element_by_id('requestInput').is_selected :
        browser.find_element_by_id('requestInput').click()
    #グラフを表示ボタン
    _graphBtn = browser.find_element_by_id('graphButton')
    pgScroll(browser, _graphBtn)
    time.sleep(1)
    _graphBtn.click()
    time.sleep(3)

# 集計表の処理を実行する
def doReport(wait, browser, parkNm = ""):
    isAlert = False 
            
    #集計モーダルを表示
    showSyukeiModal(wait, browser)
    #年報集計
    print('年報（売上）集計開始...')
    get_Nenpou(browser)
    #ダウンロード
    isAlert = download(wait, browser)

    #集計モーダルを表示
    showSyukeiModal(wait, browser)
    #月報集計
    print('月報（売上）集計開始...')
    get_Geppou(browser,"saleAggreBtn", True)
    #ダウンロード
    isAlert = download(wait, browser)
            
    #集計モーダルを表示
    showSyukeiModal(wait, browser)
    #月報集計
    print('月報（時間帯別売上）集計開始...')
    get_Geppou(browser,"timeSaleBtn", True)
    #ダウンロード
    isAlert = download(wait, browser)

    #集計モーダルを表示
    showSyukeiModal(wait, browser)
    #月報集計
    if '二日市' in parkNm or '中町中央' in parkNm:
        print('月報（サービス券集計）集計開始...')
        get_Geppou(browser,"serviceInfoBtn", False)
    elif "天道茂" in parkNm :
        print('月報（曜日別時間帯別占有集計）集計開始...')
        get_Geppou(browser,"weekdayPeriodBtn", False)
    else :
        print('月報（車室別売上集計）集計開始...')
        get_Geppou(browser,"roomSaleBtn", True)
    #ダウンロード
    isAlert = download(wait, browser)

    #集計モーダルを表示
    showSyukeiModal(wait, browser)
    #月報集計
    print('月報（曜日別売上集計）集計開始...')
    get_Geppou(browser,"dayOfTheWeekSalesBtn", True)
    #ダウンロード
    isAlert = download(wait, browser)

    #集計モーダルを表示
    showSyukeiModal(wait, browser)
    #月報集計
    print('月報（駐車時間別台数集計）集計開始...')
    get_Geppou(browser,"timeAmountBtn", True)
    #ダウンロード
    isAlert = download(wait, browser)

    #PDFを結合する
    pdf_file_merger = PdfFileMerger()
    l = glob.glob(DOWNLOAD_FILE_PATH + '/*.pdf')
    l.sort()
    for file_name in l :
        pdf_file_merger.append(file_name)
    
    if isAlert :
        pdf_file_merger.write(OUTPUT_FILE_PATH + '/err/' + parkNm + '.pdf')
    else:
        pdf_file_merger.write(OUTPUT_FILE_PATH + '/' + parkNm + '.pdf')
    pdf_file_merger.close()

def click_park_chooseBtn(wait, br):
    #対象リストから、各駐車場ページへ遷移する
    time.sleep(3)
    btns = br.find_elements_by_id('ParkChoose')
    if len(btns) > 0:
        _parkChooseBtn = btns[0]
        pgScroll(br, _parkChooseBtn)
        _parkChooseBtn.click()
        wait.until(EC.presence_of_element_located((By.ID, 'parkChooseTable')))
        time.sleep(1)
    else:
        do_reflesh(wait, br)
    
def do_reflesh(wait, br):
    br.refresh() #リフレッシュ
    time.sleep(1)
    waitLoading(wait, br)
    time.sleep(3)
    click_park_chooseBtn(wait, br) #駐車場一覧ボタンを押す


def alertCheck(w,br):
    _alerts = br.find_elements_by_css_selector(".alertBox")
    alertFlg = False
    for alert in _alerts:
        time.sleep(1)
        alertFlg = True
        alert.find_element_by_css_selector('.close').click()
    
    return alertFlg
    
# driver の設定
options = Options()
options.add_argument('--start-maximized')          # 起動時にウィンドウを最大化する
options.add_argument('--user-agent=hogehoge')
options.add_argument('--disable-extensions')
options.add_argument('--proxy-server="direct://"')
options.add_argument('--proxy-bypass-list=*')
# options.add_argument('--headless')
# options.add_argument('--no-sandbox')
prefs = {
    "plugins.always_open_pdf_externally": True,
    "profile.default_content_settings.popups": 1,
    "download.default_directory": DOWNLOAD_FILE_PATH, #IMPORTANT - ENDING SLASH V IMPORTANT
    "directory_upgrade": True
}
options.add_experimental_option("prefs", prefs)
browser = webdriver.Chrome(executable_path=CHROME_PATH, chrome_options=options) # Mac
browser.implicitly_wait(3)
wait = WebDriverWait(browser, 60)
browser.get(URL_PATH)
# ログイン画面 の設定
_id = browser.find_element_by_id("username")
_pass = browser.find_element_by_id("password")
_submit = browser.find_elements_by_tag_name("button")[0]

_id.send_keys(ID)
_pass.send_keys(PASS)
_submit.click() #ログイン


# 駐車場一覧画面 の設定
click_park_chooseBtn(wait, browser)

# 売上データ取得対象リストの取得
intPageCnt = 0
_targetDf = NaN
while True :
    if intPageCnt > 0 :
        wait.until(EC.presence_of_element_located((By.ID, 'parkChooseTable')))
        time.sleep(1)
    
    # 駐車場一覧のテーブルから読み込み対象（1.WS）の行を取得する。
    _tblHTML = browser.find_element_by_id('parkChooseTable').get_attribute('outerHTML')
    df = pd.read_html(_tblHTML)[0]
    df = df.query("属性 == ['1.WS']")
    df["page"] = intPageCnt #取得したページの番号
    df["isRead"] = False #読み込み完了フラグ（False = 未, True = 完）
    df = df[['属性','駐車場名','page','isRead']]
    # 
    _targetDf = df if intPageCnt == 0 else pd.concat([_targetDf, df],join='inner')

    # 次へボタンが活性であれば次々ページへ遷移する
    _nextBtn = browser.find_element_by_id('parkChooseTable_next')
    _next = _nextBtn.find_element_by_tag_name('a')
    if _nextBtn.get_attribute('class').find('disabled') > -1 :
        _tblHTML = NaN
        df = NaN
        break
    else :
        pgScroll(browser, _nextBtn)
        _next.click()
        intPageCnt += 1

# 取得したテーブルの整形
_targetDf = _targetDf.set_index(['page','駐車場名'])

#駐車場一覧画面に戻る
click_park_chooseBtn(wait, browser)
intReadCnt = 15
currentPgCnt = 0
retryCnt = 0
while len(_targetDf.loc[_targetDf['isRead'] == False]) > 0 :

    try :
        _targetDf2 = _targetDf.loc[currentPgCnt, :] #ページ内の駐車場情報
        currentParkDf = _targetDf2.iloc[intReadCnt] #これから読み込む駐車場情報
        parkNm  = currentParkDf.name #駐車場名
        isRetry = False #リトライフラグ

        #次ページを判定する
        for i in range(currentPgCnt):
            #画面の中の次ページを探してクリック
            _nextBtn = browser.find_element_by_link_text('次へ')
            pgScroll(browser, _nextBtn)
            _nextBtn.click()
            time.sleep(2)

        if 'クエストコート' in parkNm :
             #一時的に読み込み回避
            intReadCnt += 1
            _targetDf.loc[([currentPgCnt], [parkNm]), 'isRead'] = True
            continue

        if currentParkDf['isRead'] == False :
            #まだ読み込まれていない場合

            #駐車場の各行を取得
            target_link_elem = browser.find_element_by_link_text(parkNm)

            #駐車場ページへ遷移
            print('START : ' + parkNm)
            pgScroll(browser, target_link_elem)
            time.sleep(1)
            target_link_elem.click()

            if waitLoading(wait, browser) : #Loadingまち
                #アラートは無視してOK
                alertCheck(wait, browser)

                #PDFのDLディレクトリを空にする
                shutil.rmtree(DOWNLOAD_FILE_PATH)
                os.mkdir(DOWNLOAD_FILE_PATH)   
            
                #レポートを出力する
                doReport(wait, browser, parkNm)
                

                #内部データの読み込み完了フラグを立てる
                _targetDf.loc[([currentPgCnt], [parkNm]), 'isRead'] = True

                #駐車場一覧ページに戻る
                print('のこり:' + str(len(_targetDf.loc[_targetDf['isRead'] == False])) + '件')
                do_reflesh(wait, browser)

            else:
                #Loading Timeout
                #リトライする
                isRetry = True
                if retryCnt > 2 :
                    #リトライ回数が3回を超えると、エラーを吐いて次に進む
                    print(parkNm + "のリトライ回数が上限に達しました。")
                    _targetDf.loc[([currentPgCnt], [parkNm]), 'isRead'] = True
                    isRetry = False
                else:
                    raise TimeoutException
                

        if isRetry == False and len(_targetDf.loc[([currentPgCnt]), 'isRead']) == 0 :
            #全ての駐車場が読まれていた場合はページのカウンターをインクリメントする
            print('GO NEXT PAGE ->')
            currentPgCnt += 1
            intReadCnt = 0
            retryCnt = 0
        elif isRetry == False :
            #駐車場の読み込みが完了→次の駐車場へ
            intReadCnt += 1
            retryCnt = 0        
        
        wait.until(EC.presence_of_element_located((By.ID, 'parkChooseTable')))

    except TimeoutException :
        #リトライする場合
        print('DO RETRY ->')
        retryCnt += 1
        do_reflesh(wait, browser)
    except Exception as e:
        print(e)
        break

print('FINISH')

browser.quit()