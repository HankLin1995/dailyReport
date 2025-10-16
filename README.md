# 監造日報表系統 (Daily Report System)

## 專案簡介

這是一個基於 Excel VBA 開發的工程監造日報表管理系統，用於管理工程施工的每日記錄、進度追蹤、查驗管理、估驗計算等功能。系統整合了 PCCES 預算書介面，提供完整的工程監造管理解決方案。

**開發者：** Hank Lin  
**最後更新：** 2023/07/31

---

## 📁 專案架構

### 核心模組 (*.bas)

#### 1. **FunctionModel.bas**
主要功能模組，包含所有的命令處理函數（cmd 開頭的函數）
- `cmdGetReportIDByDate()` - 依日期選擇報表頁數
- `cmdMixComplete()` - 檢查組合工項完成度
- `cmdOutput()` / `cmdOutput_XLS()` / `cmdOutput_Paper()` - 報表輸出
- `cmdOpenCheck()` - 開啟查驗紀錄
- `cmdCreateProgress()` - 建立進度資料
- `cmdGetPayItems()` / `cmdExportToPAY()` - 估驗管理
- `getOverNumberFromLastDay()` - 處理剩餘零星數量（校正回歸）
- `dealOverNum()` - 校正數量調整

#### 2. **GIT.bas**
Git 版本控制相關功能
- 檔案版本追蹤與管理

#### 3. **NumAndSum.bas**
數量與加總計算模組
- 工程數量統計與計算功能

#### 4. **UnitTest.bas**
單元測試模組
- 系統功能測試與驗證

#### 5. **Normal.bas**
一般通用功能模組

#### 6. **CBudget.bas**
預算相關計算模組

#### 7. **FetchURL_NEW.bas**
網路資料擷取模組

#### 8. **progress_plot.bas**
進度繪圖模組
- 工程進度視覺化

#### 9. **tranFunction.bas**
資料轉換功能模組

#### 10. **test.bas** / **mail_test.bas** / **word_test.bas**
各類測試與功能驗證模組

---

### 類別模組 (*.cls)

#### 核心類別

##### **clsRecord.cls**
施工紀錄管理類別
- `Recording()` - 記錄施工資料
- `Recording_Mix()` - 記錄組合工項
- `getRecordsByDate()` - 取得指定日期的施工紀錄
- `getChecksInfByDate()` - 取得查驗資訊
- `getTmpData()` - 取得暫存資料
- `exportReport_Main()` / `exportReport_Sum()` - 匯出報表數據

##### **clsReport.cls**
報表產生類別
- `CollectItem()` - 收集報表項目
- `CollectRec()` - 收集施工記錄
- `WriteReport()` - 寫入報表
- `getInfo()` - 取得基本資訊
- `ResetReport()` - 重置報表
- `getCorrectPgs()` - 取得正確進度
- `KeyInPGS()` - 輸入進度資料

##### **clsPCCES.cls**
PCCES 預算書介面管理類別
- `getAllContents()` - 取得預算書所有內容
- `getFileName()` - 開啟預算書檔案
- `getData()` - 取得預算資料
- `getSumMoney()` - 計算契約總價
- `getPercentageItems()` - 取得百分比項目
- `getRecordingItemsByRecDate()` - 取得可施作項目
- `exportToMain()` - 匯出至主工作表
- `clearBudget()` - 清除預算資料
- `t_change_to_column()` - 變更設計次數轉欄位

##### **clsBasicData.cls**
基本資料管理類別
- `ReadData()` - 讀取工程項目、試驗項目、施作位置
- `Init()` - 初始化表單資料
- `DiaryReset()` - 重整日誌工作表
- `addNewDiaryDays()` - 新增工期日數
- `addStopDays()` - 新增停工日數
- `getProgByInter()` - 以內插法計算進度
- `ReturnUnit()` - 查詢項目單位與剩餘數量

##### **clsMixData.cls**
組合工項管理類別
- `ReadData()` - 讀取組合工項資料
- `CheckComplete()` - 檢查組合工項完成度
- `CheckUnfoundMixName()` - 檢查未找到的組合工項

##### **clsPay.cls**
估驗管理類別
- `getPayInfo()` - 取得估驗資訊
- `exportPayNumToReport()` - 匯出估驗數量至報表
- `storePayItems()` - 儲存估驗項目
- `clearPAY()` / `clearPAY_Report()` - 清除估驗資料
- `getPayDates()` / `getPayCounts()` - 取得估驗日期與次數

##### **clsCheck.cls**
查驗管理類別
- 施工查驗表管理
- 查驗紀錄追蹤

##### **clsPrintOut.cls**
列印輸出類別
- `BeforePrintCheck()` - 列印前檢查
- `ToPDF()` - 輸出為 PDF
- `ToXLS()` - 輸出為 Excel
- `ToPaper()` - 列印至紙張

##### **clsInformation.cls**
工程資訊管理類別
- `GetStartDate()` / `GetEndDate()` - 取得開工日與竣工日
- `getContractChangesByDate()` - 取得變更設計次數
- `IsEnlarged` - 是否有展延工期

##### **clsReportTest.cls** / **clsReportPhoto.cls**
試驗報告與照片管理類別

##### **clsMyfunction.cls**
通用工具函數類別
- `getRowsByUser()` - 取得符合條件的列
- `getUniqueItems()` - 取得唯一值集合
- `AppendData()` - 附加資料
- `BubbleSort_coll()` - 集合排序
- `ReverseColl()` - 反轉集合
- `IsFileExists()` - 檢查檔案是否存在
- `tranCharcter()` - 字元轉換

##### **clsBudgetDB.cls** / **clsBudgetItems.cls** / **clsCBudgetXLS.cls**
預算資料庫與項目管理類別

##### **clsFetchURL.cls** / **clsFetchURL_TEST.cls**
URL 資料擷取類別

##### **clsUserInformation.cls**
使用者資訊管理類別
- `hideCmd()` - 隱藏命令

##### **clsDetail.cls**
明細管理類別

---

### 使用者表單 (*.frm / *.frx)

#### 1. **frmData.frm**
單一工項資料輸入表單
- 施工日期、施作位置、工項名稱、數量輸入

#### 2. **MixData.frm** / **MixData_Main.frm**
組合工項資料輸入表單
- 管理複合式施工項目

#### 3. **frm_Check.frm**
查驗表單
- 查驗項目、日期、位置、照片管理

#### 4. **frm_Report.frm**
報表表單

#### 5. **frm_Detail.frm**
明細表單

#### 6. **frm_login.frm**
登入表單

#### 7. **frm_signup.frm**
註冊表單

#### 8. **frm_Info.frm**
資訊顯示表單

#### 9. **frm_Test.frm**
測試表單

#### 10. **frm_Photo_TMP.frm**
暫存照片表單

#### 11. **frm_EndQRCode.frm**
QR Code 結尾表單

#### 12. **ERRORForm.frm**
錯誤訊息表單

#### 13. **frmMSG.frm**
訊息顯示表單

---

### 文件類別模組 (*.doccls)

#### **ThisWorkbook.doccls**
Excel 活頁簿事件處理
- `Workbook_Open()` - 開啟活頁簿時的初始化
  - 顯示資訊表單
  - 執行 URL 擷取
  - 取得工作簿路徑
  - 檢查試驗完成度

#### **工作表模組**
- 工作表4、工作表5、工作表16 等工作表事件處理

---

## 📊 工作表結構

系統主要使用以下工作表：

1. **Main** - 主要資料工作表
2. **Report** - 報表工作表
3. **Records** - 施工紀錄工作表
4. **Budget** - 預算工作表
5. **Diary** - 工程日誌工作表
6. **Check** - 查驗紀錄工作表
7. **Mix** - 組合工項工作表
8. **Mix_Sum** - 組合工項統計工作表
9. **TMP** / **TMPTOTAL** - 暫存資料工作表
10. **PAY** / **PAY_EX** - 估驗資料工作表
11. **Report_Sum** - 報表統計工作表

---

## 🔄 系統流程

### 1. 初始化流程
```
開啟檔案 → 顯示資訊表單 → 載入使用者資料 → 擷取雲端資料 → 檢查試驗完成度
```

### 2. 日常施工記錄流程
```
選擇日期 → 選擇施作位置 → 選擇工項 → 輸入數量 → 記錄查驗項目 → 儲存記錄
```

### 3. 報表產生流程
```
選擇報表日期 → 收集施工項目 → 收集施工記錄 → 計算進度 → 產生報表 → 輸出
```

### 4. 估驗流程
```
輸入估驗日期 → 取得估驗項目 → 匯出至估驗表 → 計算估驗金額 → 產生估驗報表
```

### 5. 變更設計流程
```
匯入 PCCES 預算書 → 標記變更項目 → 匯出至報表 → 更新契約金額
```

---

## 🔧 主要功能

### 1. 施工記錄管理
- 單一工項記錄
- 組合工項記錄
- 施作位置管理
- 數量統計與校正

### 2. 報表產生
- 監造日報表自動產生
- 進度計算與追蹤
- 工程資訊整合
- 多格式輸出 (PDF, XLS, 紙本)

### 3. 查驗管理
- 查驗項目記錄
- 查驗照片管理
- 查驗表產生
- 抽查表管理

### 4. 估驗管理
- 估驗數量計算
- 估驗金額統計
- 估驗報表產生
- 歷史估驗查詢

### 5. 預算管理
- PCCES 預算書匯入
- 契約項目管理
- 變更設計追蹤
- 契約金額計算

### 6. 進度管理
- 預定進度設定
- 實際進度計算
- 進度落後分析
- 進度圖表產生

### 7. 工期管理
- 工期天數計算
- 展延工期記錄
- 停工日數管理
- 竣工日期追蹤

---

## 🚀 使用方式

### 首次使用設定

1. **匯入 PCCES 預算書**
   - 執行 `cmdAddItemName()` 匯入預算書
   - 系統會自動分析預算項目
   - 設定契約開工日期

2. **設定工期**
   - 執行 `cmdCreateProgress()` 建立工程日誌
   - 輸入工期天數
   - 系統自動產生每日記錄框架

3. **設定預定進度**
   - 在 Diary 工作表的 D 欄輸入關鍵進度點
   - 系統會自動內插計算每日預定進度

### 日常使用

1. **記錄施工資料**
   - 點選【單一工項】按鈕開啟 `frmData` 表單
   - 或點選【組合工項】按鈕開啟 `MixData_Main` 表單
   - 填寫日期、位置、工項、數量等資訊

2. **記錄查驗資料**
   - 點選【查驗紀錄】按鈕開啟 `frm_Check` 表單
   - 填寫查驗項目、日期、位置、照片等資訊

3. **產生日報表**
   - 選擇報表日期（使用 `cmdGetReportIDByDate()`）
   - 系統自動計算當日施工數量與累計進度
   - 輸出報表（PDF/XLS/紙本）

4. **估驗作業**
   - 執行 `cmdGetPayItems()` 輸入估驗日期
   - 系統計算估驗數量
   - 執行 `cmdExportToPAY()` 產生估驗報表

### 進階功能

1. **變更設計管理**
   - 匯入新版 PCCES 預算書
   - 系統自動辨識變更項目
   - 執行【匯出至報表】更新 Main 工作表

2. **數量校正**
   - 執行 `getOverNumberFromLastDay()` 進行數量校正
   - 系統自動調整微小誤差

3. **展延工期**
   - 執行 `cmdEnlargeDays()` 輸入展延天數
   - 系統自動更新竣工日期與工程日誌

---

## ⚠️ 注意事項

1. **資料備份**
   - 請定期備份 Excel 檔案
   - 重要操作前建議先複製一份

2. **PCCES 預算書格式**
   - 預算書需包含「詳細表」工作表
   - 確保預算項目編號與名稱格式正確

3. **日期連續性**
   - Diary 工作表的日期必須連續
   - 若有停工需使用 `addStopDays()` 功能

4. **數量單位**
   - 確保輸入數量單位與預算書一致
   - 系統會自動顯示對應單位

5. **檔案路徑**
   - 輸出檔案預設儲存於工作簿同目錄
   - 可在 Main 工作表 B8 儲存格修改路徑

---

## 📝 版本控制

專案使用 Git 進行版本控制，相關功能在 `GIT.bas` 中實作。

---

## 📞 技術支援

如有任何問題或建議，請聯絡開發者：Hank Lin

---

## 🔄 更新紀錄

- **2023/07/31**: 更新安全聲明與使用者驗證機制
- **2023/02/25**: 新增試驗完成度檢查功能 (`checkTestCompleted`)
- **2022/11/25**: 新增依日期選擇報表頁數功能
- **2022/11/22**: 新增數量校正回歸功能

---

**注意：** 本文件僅供參考，實際功能以程式碼為準。
