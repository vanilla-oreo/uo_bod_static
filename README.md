# UO BOD 獎勵反查系統 v3（純 HTML + JSON）

## 已使用的 Excel 版本

- 鐵匠：遊戲實測修正版（四張礦圖 600 / 850 / 950 / 1050 點）
- 抄錄：遊戲實測修正版（新增三項 300 點書籍；375 點 Miner's 天然染料）
- 鉗工、裁縫、鍊金、木工、烹飪：原版（先前已與遊戲截圖對照）
- 製弓／製箭：原版，尚未做遊戲內比對

**注意：各職業訂單本身的得分仍採原始 Excel／官網點數，未經完整遊戲實測。**

## 1. 網站部署

最簡單方法：只把這兩個檔案依路徑放上網站：

```
index.html
data/bod_data.json
```

GitHub Pages：建立 repository，上傳上述檔案，至 Settings → Pages 啟用 Branch / root。也可放一般虛擬主機。網址開啟 `index.html` 即可。

**直接在 Windows 雙擊 `index.html` 使用 `file://` 時，瀏覽器可能封鎖 `fetch()` 讀取 JSON。** 如要本機測試，在專案根目錄執行：

```
python -m http.server 8000
```

瀏覽器打開 `http://localhost:8000/`。

## 2. Excel → JSON

原始八份 Excel 在 `source_excels/`；工作用 Excel 在 `excel/`。初次整理：

```
python tools/bod_manager.py init
```

之後如果你直接編輯了 **Python 維護版** Excel 的「獎勵點數／小訂單點數／大型訂單點數」三份標準化主表，只需：

```
python tools/bod_manager.py export
```

**原始官方版 Excel** 的小／大型訂單採不同矩陣格式，「直接兌換對照表」可能存有舊的公式快取，因此不要直接手改原始官方版的訂單得分再執行 `export`；請使用下列 Python 指令修改訂單，讓工具先轉成標準化主表。修改原始官方版的獎勵名稱／點數則可以直接由 `export` 讀取。

然後把 **`data/bod_data.json`** 覆蓋到網站對應位置，重新整理網頁即可。JSON 會依點數即時計算，不依賴 Excel 兌換對照表中的公式快取。

## 3. Python 圖形化管理（Windows 推薦）

在 Windows 解壓縮後，**雙擊根目錄的 `tools/bod_manager.py`**，或以 Python 執行 `tools/bod_manager.py`（不帶參數），即可開啟管理視窗。

1. 選擇職業、獎勵品或訂單。
2. 點選清單項目，修改名稱或直接點數，按「更新選取資料」。也可以新增、刪除。
3. 工具自動備份舊 Excel，再重新建立五個工作表並輸出 JSON。
4. 將新的 `data/bod_data.json` 上傳網站即可。

**不用安裝 `artifact_tool`、`openpyxl` 或其他第三方 Python 套件。** 只需 Python 3.10+ 與 Windows 通常內建的 Tkinter。也可以用指令列：

```
python tools/bod_manager.py rewards blacksmith
python tools/bod_manager.py set-reward blacksmith --id R027 --points 650
python tools/bod_manager.py add-reward blacksmith --name "新增獎勵" --points 700
python tools/bod_manager.py delete-reward blacksmith --id R041
python tools/bod_manager.py set-order blacksmith --id O0001 --points 225
python tools/bod_manager.py export
```

上述指令中的 ID 僅為示例，請使用清單確認正確 ID。每次修改都有 `backup/` 備份，`source_excels/` 保留初始原檔。

> Python 重新建立的工作 Excel 保留必要資料與五個工作表，但不保證原始 Excel 的特殊裝飾版面完全相同。

## 4. JSON 更新與 GitHub Pages

每次修改完 Excel，只需覆蓋 `data/bod_data.json` 並 commit / push。網頁會從 JSON 即時計算 `訂單點數 >= 獎勵門檻`，並按差值從小到大排序。

**請勿直接修改 JSON 再執行 Python 匯出，否則會被 Excel 內容覆蓋。** 唯一主要來源是 Excel。

## 5. 環境

- Python 3.10 以上
- `export`、`init`、清單查詢只需要 Python 標準函式庫
- Python 修改 Excel 使用內建標準函式庫建立 XLSX，不需要另外安裝試算表套件。
- 前台支援現代 Chrome、Edge、Firefox、Safari

## 6. v3 篩選邏輯

選定獎勵後，其他四個條件只列出**實際達到點數門檻的訂單**所具有的值。依序選「訂單類型 → 材料／類型 → 品質 → 每種件數」，下游清單會自動收窄；若切換上游選項與下游選擇衝突，下游條件會清空。

## 7. 運作原則

- 每張 BOD 只能直接換一項獎勵；不累加不同訂單點數。
- 「點數最接近」不是「材料成本最低」，也不包含大型訂單收集難度。
- 網頁不讀取 `.xlsx`，部署只需 HTML + JSON。


## v3 修正紀錄
- 前台篩選改以保留的條件狀態查詢 JSON，不再讓下拉選單重繪干擾查詢結果。
- 材料／類型依職業階級從低到高排序。
- 已移除非必要的 Windows BAT；執行 `python tools/bod_manager.py` 可開啟管理視窗。
- 壓縮檔內統一以 `uo_bod_static/` 為根目錄。
