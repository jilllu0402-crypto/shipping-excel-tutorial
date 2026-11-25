# 航運行政人員 MS Excel 資料前處理實務教學與練習題庫專案
> **專為航運行政、海運承攬（Forwarder）、報關行與貨櫃堆場初學者量身打造的 Excel 資料前處理實務指引**

---

## 📌 專案簡介

本專案聚焦於國際航運行政人員進行基礎資料分析時必須具備的**「資料前處理」（Data Preprocessing）**基本操作技能。

在現代海運供應鏈中，行政人員每天面對來自船公司（Carrier）、貨主（Shipper）、集裝箱堆場（CY/CFS）以及海關通關網路的大量資料。這些資料往往充滿「首尾空格、小寫字母、連字號雜訊、夾帶文字單位（如 18500 KGS）、點號非標準日期、重複申報列與阻斷選取的空行」。

本專案透過完整的 **Word 教學手冊**、**PowerPoint 投影片簡報**、**網頁互動式簡報**、**30 筆真實髒資料示範檔**、以及 **300 筆五大核心情境練習題庫**，協助初學者一步步建立「Tidy Data」乾淨一維資料表架構，達到 100% 樞紐分析就緒（Pivot Table Ready）！

---

## 📂 專案檔案清單與目錄結構

```text
shipping_excel_tutorial/
├── 航運行政人員MS_Excel資料前處理實務教學手冊.docx     # 完整 MS Word 教學手冊 (含封面、目錄、圖表目錄、頁碼)
├── 航運行政人員MS_Excel資料前處理實務教學簡報.pptx     # 16:9 寬螢幕 PowerPoint 簡報 (16 頁完整講稿與截圖)
├── shipping_excel_presentation.html                # 互動式 PPT 簡報網頁 (含鍵盤翻頁、即時演練器、模擬器)
├── 航運行政Excel資料前處理_30筆示範檔.xlsx           # 30 筆真實髒資料示範檔 (含原始表、清洗後表、代碼對照表)
├── 航運行政Excel資料前處理_300筆綜合情境練習檔.xlsx     # 300 筆五大核心業務情境題庫 (附完整參考解答公式)
├── images/                                         # 15 張高解析度操作指令與介面圖解截圖
│   ├── fig_1_1_excel_interface.png                 # 圖 1-1 MS Excel 主畫面核心介面分區說明
│   ├── fig_1_2_shortcuts.png                       # 圖 1-2 航運行政高頻快捷鍵指南
│   ├── fig_2_1_tidy_data.png                       # 圖 2-1 混亂報表 vs Tidy Data 乾淨架構對比
│   ├── fig_2_2_data_types.png                      # 圖 2-2 Excel 四大資料型態與預設對齊檢驗法
│   ├── fig_3_1_raw_30_data.png                     # 圖 3-1 30 筆示範檔 10 大經典髒資料地雷地圖
│   ├── fig_4_1_remove_duplicates.png               # 圖 4-1 移除重複項操作剖析
│   ├── fig_4_2_text_to_columns.png                 # 圖 4-2 資料剖析精靈步驟圖解
│   ├── fig_4_3_flash_fill.png                      # 圖 4-3 快速填入 (Ctrl+E) 智慧識別
│   ├── fig_4_4_find_replace_goto.png               # 圖 4-4 尋找取代與特殊目標 (F5) 刪空行
│   ├── fig_4_5_string_functions.png                # 圖 4-5 核心字串清洗函數工具庫
│   ├── fig_4_6_date_number_functions.png           # 圖 4-6 日期與數值清洗函數 (DATE, VALUE)
│   ├── fig_4_7_xlookup_mapping.png                 # 圖 4-7 XLOOKUP vs VLOOKUP 代碼標準化
│   ├── fig_4_8_cleaned_30_data.png                 # 圖 4-8 清洗完成之 30 筆標準化資料全貌
│   ├── fig_5_1_scenarios_overview.png              # 圖 5-1 300 筆五大業務情境題庫全景
│   └── fig_5_2_demurrage_case.png                  # 圖 5-2 貨櫃超期延滯費 (Demurrage) 階梯計費公式
└── README.md                                       # 本說明文件
```

---

## 🛠️ 前處理標準作業程序 (SOP) 六部曲

1. **第一步：剔除重複與瞬殺空行**
   - 功能選項：【資料】➔【移除重複項】（`Alt ➔ A ➔ M`）。
   - 快捷鍵：`F5` ➔ 特殊目標選【空格】➔ `Ctrl + -` 刪除整列。
2. **第二步：字串雜訊淨化與大小寫標準化**
   - 提單號公式：`=TRIM(CLEAN(UPPER(B2)))`（除空格、去換行、轉大寫）。
   - 貨櫃號連字號清洗：`=UPPER(SUBSTITUTE(SUBSTITUTE(C2, "-", ""), " ", ""))`。
   - 免公式神技：【快速填入】（`Ctrl + E`）。
3. **第三步：文字型數字與計價單位清洗**
   - `Ctrl + H` 批次替換 `" KGS"` 與 `"USD "` 為空。
   - 函數法：`=VALUE(SUBSTITUTE(H2, " KGS", ""))`。
4. **第四步：非標準日期的徹底修復**
   - 【資料剖析】精靈指定【日期: YMD】修復緊湊 8 碼文字（`20260310`）。
   - 公式重組法：`=DATE(LEFT(A2,4), MID(A2,5,2), RIGHT(A2,2))`。
5. **第五步：港口名稱與代碼標準化對照**
   - `=XLOOKUP(F2, Ref_Lookup_Tables!$A$2:$A$20, Ref_Lookup_Tables!$B$2:$B$20, F2)`。
6. **第六步：資料驗證防呆與成果檢核**
   - 設定報關狀態下拉選單（`放行,審驗中,補件中,查驗`）。

---

## 📚 300 筆五大航運業務情境練習題庫

| 情境編號 | 業務情境名稱 | 練習筆數 | 核心清洗技能 |
| :---: | :--- | :---: | :--- |
| **S1** | 海運提單與訂艙清單 (Ocean B/L & Booking) | 60 筆 | 訂艙號空白去除、船名航次複合字串拆分 (`TEXTBEFORE`/`TEXTAFTER`)、目的港轉 UN/LOCODE |
| **S2** | 貨櫃動態與堆場進出追蹤 (Container Tracking) | 60 筆 | ISO 6346 11碼長度檢驗 (`LEN(B2)<>11` 紅底高亮)、停留天數動態計算 (`TODAY()-GateIn`) |
| **S3** | 進出口報關與放行進度 (Customs Clearance) | 60 筆 | 14 碼報單格式化補全、通關方式 C1/C2/C3 資料驗證、申報貨價純化 |
| **S4** | 運費帳單與滯留費核銷 (Freight & Demurrage) | 60 筆 | 幣別拆分、營業稅 5% 核算、超期階梯延滯費萬能公式 (`IFS`) |
| **S5** | 出貨裝箱單與材積重檢核 (Packing List) | 60 筆 | 尺寸解析、單箱 CBM 計算 (`除以 1,000,000`)、計費重量 RT 決策 (`MAX(重量噸, 體積噸)`) |

---

## 🌐 權威國際標準規範引用

1. **聯合國歐洲經濟委員會 (UNECE / UN/CEFACT)**: UN/LOCODE 國際貿易港口代碼規範。
2. **國際貨櫃局 (BIC) & 國際標準化組織 (ISO)**: ISO 6346 貨櫃編號與尺寸代碼規範、ISO 4217 貨幣代碼規範。
3. **統計學與資料科學**: Hadley Wickham, *Tidy Data* (Journal of Statistical Software)。
4. **中華民國財政部關務署**: 《進出口通關作業手冊》通關方式（C1/C2/C3）規範。
5. **國際海事組織 (IMO)**: SOLAS 公約貨櫃驗證總重量（VGM）計算標準。

---
*專案編撰時間：2026 年 11 月*
