# Excel_VBA 固定產品自動發信功能(Outlook)
該貨件在5年前非常火熱，台灣私人貨件大量進口，所以針對此專案做了一套該貨件專屬的小工具，該貨件的需求有以下:
1. 該貨件有分兩種款式
    - WIFI
    - Bluetooth 
2. 該貨件有無線射頻管制，所以要提供NCC自用切結書給客戶填寫，資料如下:
    - NCC自用切結書-WIFI COOKER(PCB & PCW).doc
    - NCC自用切結書-WIFI COOKER(PCB-120US-K1).doc
    - NCC自用切結書-WIFI COOKER(PCW-120US-K1).doc
3. 私人貨件皆需提供身分證及個案委任書申報，資料如下:
    - 身分證正反面黏貼處.doc
    - 個案委任書.doc
4. 當時鼓勵開始使用EZWAY並順便夾帶操作方式pdf檔案
    - 智慧型手機 下載APP委任更方便.pdf
![image](https://user-images.githubusercontent.com/37874182/204196507-5d66070d-4ad0-4735-b122-1d93c64141f6.png)

# 第一次使用安裝
1. 點選右邊[Setup]按鈕，新增資料夾(未來使用免點選)
    - D:\AutoNCC
    - D:\AutoPOA
    - D:\AutoPasteID
    - D:\Document
2. 更改K5:K9儲存格詳細資料

# 使用步驟
1. 貼上整欄提單號碼在A2儲存格上。
2. 貼上收件人Email於B2儲存格上。
3. 檢查客戶進口的商品是否為WIFI & Bluetooth 版本或個別進口各一個(2個);單次進口最大上限為2個。
4. 檢查進口數量1 & 2; 單次進口最大上限為2個。
5. 點選右方按鈕[Produce POA]，製作個案委任書。
6. 點選右方按鈕[Produce NCC]，製作NCC自用切結書。
7. 點選右方按鈕[Produce PasteID]，製作身分證正反面黏貼處。
8. 點選右方按鈕[Search Data]，搜尋(步驟5-7自動產生文件)和(10碼提單號碼.pdf)文件。
9. 若搜尋文件無缺失，點選右方[Send Mail]，並啟動Outlook寄出信件。

# 功能介紹
1. 按鈕[Send Mail] - 寄出信件。
2. 按鈕[Search Data] - 搜尋自動產生檔案，如(個案委任書，NCC自用切結書，身分證正反面黏貼處，10提單號碼)檔案等。
3. 按鈕[Produce POA] - 自動產生個案委任書。
4. 按鈕[Produce NCC] - 自動產生NCC自用切結書。
5. 按鈕[Produce PasteID] - 自動產生身分證正反面黏貼處。
6. 按鈕[Clear Data] - 刪除自動產生的相關檔案(不含10碼提單號碼.pdf)。
7. 按鈕[Clear All] - 清除工作表Cooker_Mail的儲存格A2:F1000的所有內容。


# 進階技術說明
1. 自動產生的檔案(Produce POA,NCC,PasteID)的概念是一樣的，利用A欄所提供的提單號碼，自動產生有提單號碼的內文及個別賦予不同檔名，來區分搜尋及寄出的對應檔案。
2. 按鈕[Search Data]若搜尋發現短缺的文件，會在該對應的的儲存格出現紅底(POA,NCC,PasteID)三個不同的文件短缺提示
3. 自動產生檔案顯示(Visible) 為False，製作文件相對費時，提供使用者UI介面查看Loading的狀態，避免使用者誤判為系統當機，造成製作檔案中斷及錯誤。
* PasteID 等待畫面
* ![image](https://user-images.githubusercontent.com/37874182/204202214-13751cd4-5c56-42b0-a2c0-5de7e3b5878e.png)
* NCC 等待畫面
* ![image](https://user-images.githubusercontent.com/37874182/204202286-3a912fa1-b318-4f50-9080-ad2b1422c66a.png)
* 個案委任書 等待畫面
* ![image](https://user-images.githubusercontent.com/37874182/204202360-bb28d647-0c27-493a-990c-180fd6245b6d.png)
