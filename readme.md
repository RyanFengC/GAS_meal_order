# Line bot訂餐統計機器人(使用Google sheet及Google Apps Script(GAS))
## 0.Google sheet
(https://docs.google.com/spreadsheets/u/1/d/1pxqlCbg4mwPqo3n3HNcbTe0uLwLrx-HUtrbyQlQ8Ruc/copy#gid=0)
歡迎大家下載

## 1.緣起:
本文分享如何使用Google Apps Scripts 搭配 Google sheets 連接 line 做聊天機器人。
會有這製作這個機器人的想法主要是因為，每次訂便當的時候都要一個一個統計不同便當的數量及總金額，需要花費較多時間，因此才有製作這個聊天機器人的想法來解決這個問題。
但筆者還未學過line bot的製作，所以筆者就上網自學製作出這個聊天機器人，因為在自學的過程中受惠於各位前輩的分享，所以這個專案完成的時候我也決定分享出來。

## 2.需使用之工具:
Line developer帳號 、Google Sheet、Google Apps Script(簡稱GAS)

## 3.開始製作:
3.1資料庫部分:
使用Google Sheet作為資料庫
優點: 1.方便且只要有Google帳號的人都可以使用。
缺點:僅小專案使用，大專案較不適合。

## 3.2程式碼部分:
請參考	meal.gs中之註解，其中index.html 為要將"訂單統計資料庫"google sheet表格中之資料，結果顯示於頁面所刻的網頁

## 3.3測試心得:
在寫程式的過程中由於需要測試所寫之函數功能是否符合需求，但如果要發布告line bot再進行測試將會較麻煩，因此筆者使用的測試方式為，直接指定函式所需資料，接著直接在GAS中執行函式確定函式所要處理之功能沒問題後再將原本直接指定的資料換成由line bot接收到的資料。

## 3.4本程式使用規定!!
一、可使用之指令:
輸入以下雙引號(" ")內之紅色文字 :
1. 清除資料庫之資料: "clear","/clear","清空"或"清除"
2. 開始統計(一般模式):"開始"
3. 列出結果:"/list","list","結單"或"結果"
4. 開始統計(加1模式):"加1模式",”加ㄧ模式”,”+1模式”或"+一模式"
5. 寄送email:"Email","email","mail","電子郵件"

二、輸入之資料格式(請遵守此格式輸入，否則程式會有問題):  非常重要!!!!!!!!!!

★一般模式:
1. 一行一個品項輸入之格式需如多個品項用換行
2. 每個品項之輸入格式須為: 餐點 價錢 x 數量
3. 輸入之品項中不要有小圖示(表情符號)

★加1模式:
1. 輸入之內容僅輸入 + 數量
2. 不要用貼圖取代加1
3. 不要有小圖示(表情符號)

三、操作說明:
1. po完餐點資訊後即可輸入"開始"或"加1模式"相關指令，開始進行統計，如未輸入則不會統計。
2. 一般正常留言並遵守第二節之規則即可，如有違反第二節之規則，則可能會導致統計錯誤。
3. 結單後輸入列出結果之相關指令即可。
4. 參數貼入如下說明
 ![image](https://github.com/RyanFengC/GAS_meal_order/blob/master/%E5%8F%83%E6%95%B8%E4%BB%8B%E7%B4%B9.JPG)

## 4.參考資料:
1. D4- 如何透過 Google Apps Script 來整合 Google Form / Google Sheet 並自動寄出客製的 Email？ - iT 邦幫忙::一起幫忙解決難題，拯救 IT 人的一天 (https://ithelp.ithome.com.tw/articles/10259650)
2. 做個 LINE 機器人記錄誰 +1！群組 LINE Bot 製作教學與分享 (https://jcshawn.com/addone-linebot/)
3. 將 Line 及 Telegram 群組的對話與檔案即時自動備份到 Google 雲端硬碟 2021(https://www.youtube.com/watch?v=JDtaY231MJk&list=PLLrJ9DEA0QKNVGni0VJGjUU5nlIzut3Yj)
4. GitHub - GarrettTW/LineBot_ReportMessage: 利用Line機器人來擷取群組內的特格式訊息，最後可整理並輸出成軍中回報的特定訊息
5.【網知 EP9】將Google 試算表的資料顯示在網頁上/Google Apps Script怎麼用呢？How to pull your data from Google Sheets to Web (https://www.youtube.com/watch?v=zzIY22dHqY8&list=LL&index=11)
