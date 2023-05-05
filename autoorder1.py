# # 海期報價及下單範例
from flask import Flask, request
import json
import pythoncom
import asyncio
import datetime
import pandas as pd
import comtypes.client as cc
import os
# 只有第一次使用 api ，或是更新 api 版本時，才需要呼叫 GetModule
# 會將 SKCOM api 包裝成 python 可用的 package ，並存放在 comtypes.gen 資料夾下
# 更新 api 版本時，記得將 comtypes.gen 資料夾 SKCOMLib 相關檔案刪除，再重新呼叫 GetModule 
cc.GetModule(os.path.split(os.path.realpath(__file__))[0] + r"/SKCOM.dll")
import comtypes.gen.SKCOMLib as sk

app = Flask(__name__)

exchange_code = ""
stock_code = ""

@app.route('/webhook', methods=['POST'])
def webhook():
    data = request.json
    # 在這裡處理從TradingView發送的訊息
    exchange_code = data['exchange']
    stock_code = data['symbol']
    # fo.bstrExchangeNo = exchange_code               # 交易所代碼。
    # fo.bstrStockNo = stock_code                  # 海外期權代號。
    # fo.bstrYearMonth = month_code
    print(data)
# login ID and PW
# 身份證
ID = 'H122717637'
# 密碼
PW = 'gn01658934'
print(datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S,"), 'Set ID and PW')


# # 建立 event pump and event loop
# 新版的jupyterlab event pump 機制好像有改變，因此自行打造一個 event pump機制，目前在 jupyterlab 環境下使用，也有在 spyder IDE 下測試過，都可以正常運行

# working functions, async coruntime to pump events
async def pump_task():
    '''在背景裡定時 pump windows messages'''
    while True:
        pythoncom.PumpWaitingMessages()
        # 想要反應更快 可以將 0.1 取更小值
        await asyncio.sleep(0.1)

# get an event loop
loop = asyncio.get_event_loop()
pumping_loop = loop.create_task(pump_task())
print(datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S,"), "Event pumping is ready!")


# # 建立 event handler

# 建立物件，避免重複 createObject
# 登錄物件
if 'skC' not in globals(): skC = cc.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
# 下單物件
if 'skO' not in globals(): skO = cc.CreateObject(sk.SKOrderLib , interface=sk.ISKOrderLib)
# 海期報價物件
if 'skOSQ' not in globals(): skOSQ = cc.CreateObject(sk.SKOSQuoteLib , interface=sk.ISKOSQuoteLib)
# 回報物件
if 'skR' not in globals(): skR = cc.CreateObject(sk.SKReplyLib, interface=sk.ISKReplyLib)


# 建立事件處理類別
# SKOSQ event handler
class skOSQ_events:
    def __init__(self):
        self.OverseaProductsDetail = []

    def OnConnect(self, nKind, nCode):
        '''連線海期主機狀況回報'''
        print(f'skOSQ_OnConnect nCode={nCode}, nKind={nKind}')

    def OnOverseaProductsDetail(self, bstrValue):
        '''查詢海期/報價下單商品代號'''
        if "##" not in self.OverseaProductsDetail:
            self.OverseaProductsDetail.append(bstrValue.split(','))
        else:
            print("skOSQ_OverseaProductsDetail downloading is completed.")

    def OnNotifyQuoteLONG(self, sIndex):
        '''requestStock 報價回報'''
        # 儘量避免在這裡使用繁複的運算，這裡僅在 console 端印出報價
        ts = sk.SKFOREIGNLONG()
        nCode = skOSQ.SKOSQuoteLib_GetStockByIndexLONG(sIndex, ts)
        print(ts.bstrExchangeNo, ts.bstrStockNo, ts.nClose, ts.nTickQty)

# SKReplyLib event handler
class skR_events:
    def OnReplyMessage(self, bstrUserID, bstrMessage):
        '''API 2.13.17 以上一定要返回 sConfirmCode=-1'''
        sConfirmCode = -1
        print('skR_OnReplyMessage ok')
        return sConfirmCode

    def OnNewData(self, bstrUserID, bstrData):
        '''委託單回報'''
        print("skR_OnNewData", bstrData)


# SKOrderLib event handler
class skO_events:
    def __init__(self):
        self.TFAcc = []

    def OnAccount(self, bstrLogInID, bstrAccountData):
        strI = bstrAccountData.split(',')
        # 找出期貨帳號
        if len(strI) > 3 :
            if strI[0] == 'TF' :
                # 分公司代碼 + 期貨帳號
                self.TFAcc = strI[1] + strI[3]
                
    def OnOFOpenInterestGWReport(self, bstrData):
        # 回報海期的 OI 資料
        print(bstrData)


# # 建立 event 跟 event handler 的連結

# Event sink, 事件實體化
EventOSQ = skOSQ_events()
EventR = skR_events()
EventO = skO_events()

# 建立 event 跟 event handler 的連結
ConnOSQ = cc.GetEvents(skOSQ, EventOSQ) 
ConnR = cc.GetEvents(skR, EventR) 
ConnO = cc.GetEvents(skO, EventO) 


# # 登入及各項初始化作業

# login
print('Login', skC.SKCenterLib_GetReturnCodeMessage(skC.SKCenterLib_Login(ID,PW)))

# 海期商品初始化
nCode = skOSQ.SKOSQuoteLib_Initialize()
print("SKOSQuoteLib_Initialize", skC.SKCenterLib_GetReturnCodeMessage(nCode))

# 下單前置至步驟
# 1. 下單初始化
nCode = skO.SKOrderLib_Initialize()
print("Order Initialize", skC.SKCenterLib_GetReturnCodeMessage(nCode))

# 2. 讀取憑證
nCode = skO.ReadCertByID(ID)
print("ReadCertByID", skC.SKCenterLib_GetReturnCodeMessage(nCode))

# 3. 取得海期帳號 
nCode = skO.GetUserAccount()
print("GetUserAccount", skC.SKCenterLib_GetReturnCodeMessage(nCode), EventO.TFAcc)

# 4. 連線委託回報主機
nCode = skR.SKReplyLib_ConnectByID(ID)
print("Connect to ReplyLib server", skC.SKCenterLib_GetReturnCodeMessage(nCode))


# # 登入海期報價主機，確認 OnConnect 出現 3001 回報後始可進行下列步驟
# 以下皆以手動輸入

# 5. 登入海期報價主機
nCode = skOSQ.SKOSQuoteLib_EnterMonitorLONG()
print('SKOSQuoteLib_EnterMonitor()', skC.SKCenterLib_GetReturnCodeMessage(nCode))


# # 下單前需要下載海期商品，才能下單
# 不然會報 1035 錯誤碼

# 6. 登入海期報價主機後，等確認 OnConnect 出現 3001 後，再下載海期商品
nCode = skO.SKOrderLib_LoadOSCommodity()
print('SKOrderLib_LoadOSCommodity', skC.SKCenterLib_GetReturnCodeMessage(nCode))


# # 查詢海期交易所及商品報價與下單代碼

# 等 OnConnect 出現 3001 回報後，可以查詢海期交易所及交易商品代號
# 查詢詳細交易所及商品代號，注意海期下單與報價代號有些不同
# EventOSQ.OverseaProductsDetail = []
# nCode = skOSQ.SKOSQuoteLib_GetOverseaProductDetail(1)
# print("GetOverseaProductDetail", skC.SKCenterLib_GetReturnCodeMessage(nCode))
# print("交易所代碼, 交易所名稱, 商品報價代碼, 商品名稱, 交易所下單代碼, 商品下單代碼, 最後交易日")
# print(EventOSQ.OverseaProductsDetail[0])
# print(EventOSQ.OverseaProductsDetail[-2])
# 下單代碼


# 離開海期報價主機，有需要再使用
# nCode = skOSQ.SKOSQuoteLib_LeaveMonitor()
# print(nCode, skC.SKCenterLib_GetReturnCodeMessage(nCode))


# # 海期報價範例

# 登陸海期商品報價, 格式為 "交易所代碼,商品代碼"，不同商品用#隔開，請利用
# GetOverseaProductDetail 查詢
# 登陸海期商品報，接收 callback  為 EventOSQ 的 OnNotifyQuoteLONG
# 注意熱門商品報價頻率會很高，要手動清除，不然 jupterlab 頁面會愈來愈慢
# code = skOSQ.SKOSQuoteLib_RequestStocks(1, "CBOT,MYM0000#TCE,JCO2205")
# print(datetime.datetime.now)().strftime("%Y/%m/%d %H:%M:%S,"), "RequestStocks", skC.SKCenterLib_GetReturnCodeMessage(code[1])

# # 海期下單物件 OVERSEAFUTUREORDER
# 委託價分子，這是海期商品小數點的部位，可以參考 https://www.order-master.com/doc/topic/54/

# 建立海期委託單物件
# 詳細參數請參考 api 手冊，這裡僅示範可以下單所需的參數
# 以下參數我是先用 api 附的 SKCOMtester.exe 測試，直到可以送單所測出來需要的參數
# 注意 bstr開頭的參數都要以文字型態帶入，特別是 委託價 (bstrOrder),委託價分子(bstrOrderNumerator)
# 根據 GetOverseaProductDetail 取得的下單代碼，
# 如 芝加哥交易所的微型小道瓊期貨，交易所代碼是 CBT, 商品下單代碼是 MYM_202206，
# 但 OVERSEAFUTUREORDER 物件的參數，要再另外拆成 海外期權代號(bstrStockNo) 及 近月商品年月(bstrYearMonth)
# 要將 MYM_202206 拆成 MYM 及 202206

fo = sk.OVERSEAFUTUREORDER()
fo.bstrFullAccount = EventO.TFAcc       # 海期帳號，分公司代碼＋帳號7碼
fo.bstrExchangeNo = exchange_code               # 交易所代碼。
fo.bstrStockNo = stock_code                  # 海外期權代號。
fo.bstrYearMonth = "202306"             # 近月商品年月( YYYYMM) 6碼
# fo.bstrYearMonth2                     # 遠月商品年月( YYYYMM) 6碼 {價差下單使用}
fo.bstrOrder = "0"                      # 委託價。
fo.bstrOrderNumerator = "0"             # 委託價分子。也就是小數點的部位
# fo.bstrTrigger                        # 觸發價。
# fo.bstrTriggerNumerator               # 觸發價分子。
fo.sBuySell = 0                         # 0:買進 1:賣出
                                        # {價差商品，需留意是否為特殊商品－近遠月前的「+、-」符號}
fo.sNewClose = 0                        # 新/平倉，0:新倉  {目前海期僅新倉可選}
fo.sDayTrade = 0                        # 當沖 0:否, 1:是；{海期價差單不提供當沖}
                                        # 可當沖商品請參考交易所規定。
fo.sTradeType = 0					    # 0:ROD 當日有效單; 1:FOK 立即全部成交否則取消; 2:IOC 立即成交否則取消(可部分成交)
                                        # {限價單LMT可選ROD/IOC/FOK，其餘單別固定ROD}
fo.sSpecialTradeType = 1                # 0:LMT 限價單 1:MKT  2:STL  3.STP
fo.nQty = 1                             # 交易口數。


# 海期下單 SendOverSeaFutureOrder(bstrLogInID, bAsyncOrder, pOrder)
msg, nCode = skO.SendOverSeaFutureOrder(ID, 0, fo)
print(msg, skC.SKCenterLib_GetReturnCodeMessage(nCode))


## 測試 GetOverSeaFutureOpenInterestGW 
# 手冊寫 OnGetOverSeaFutureOpenInterestGW 來回報是錯誤的，應該是 OnOFOpenInterestGWReport 來回報的
ncode = skO.GetOverSeaFutureOpenInterestGW(ID, EventO.TFAcc, 2)
print('GetOverSeaFutureOpenInterestGW', skC.SKCenterLib_GetReturnCodeMessage(ncode))

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=False)