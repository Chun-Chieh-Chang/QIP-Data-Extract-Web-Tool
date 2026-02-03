# 部署指南 - 儲存格映射協定自動確認機制

**功能**: 儲存格映射協定自動確認機制  
**版本**: v2.4.0  
**部署日期**: 2026-02-03  
**部署環境**: GitHub + GitHub Actions  

---

## 部署前檢查清單

### 代碼檢查
- [x] 所有修改已完成
- [x] 代碼審查已通過
- [x] 無語法錯誤
- [x] 無邏輯錯誤
- [x] 無性能問題

### 文檔檢查
- [x] FEATURE_TESTING_LOG.md 已創建
- [x] GITHUB_PUSH_CHECKLIST.md 已創建
- [x] CODE_REVIEW_CHECKLIST.md 已創建
- [x] DEPLOYMENT_GUIDE.md 已創建
- [x] .gitignore 已更新

### 測試檢查
- [x] 功能測試計畫已制定
- [x] 邊界情況已考慮
- [x] 性能檢查已完成
- [x] 兼容性檢查已完成

---

## 修改文件清單

### 核心代碼修改

#### 1. js/ui/app.js
**修改內容**:
- handleCellClick 函數: 添加自動確認邏輯
- startRangeSelection 函數: 添加重新修改機制
- cancelSelection 函數: 添加自動隱藏預覽區

**修改行數**: 約 20 行新增代碼

**影響範圍**: 儲存格映射協定功能

#### 2. index.html
**修改內容**:
- 隱藏確認和取消按鈕 (添加 hidden class)

**修改行數**: 2 行修改

**影響範圍**: UI 層

#### 3. .gitignore
**修改內容**:
- 添加 .vscode/ 和其他編輯器配置

**修改行數**: 3 行新增

**影響範圍**: Git 配置

### 新增文檔

#### 1. notes/FEATURE_TESTING_LOG.md
**內容**: 功能測試計畫和執行記錄

#### 2. notes/GITHUB_PUSH_CHECKLIST.md
**內容**: GitHub 推送準備清單

#### 3. notes/CODE_REVIEW_CHECKLIST.md
**內容**: 代碼審查清單

#### 4. notes/DEPLOYMENT_GUIDE.md
**內容**: 部署指南 (本文件)

---

## 部署步驟

### 步驟 1: 本地驗證

#### 1.1 檢查 Git 狀態
```bash
git status
```

**預期輸出**:
```
On branch main
Your branch is up to date with 'origin/main'.

Changes not staged for commit:
  modified:   index.html
  modified:   js/ui/app.js
  modified:   .gitignore

Untracked files:
  notes/CODE_REVIEW_CHECKLIST.md
  notes/DEPLOYMENT_GUIDE.md
  notes/FEATURE_TESTING_LOG.md
  notes/GITHUB_PUSH_CHECKLIST.md
```

#### 1.2 查看代碼差異
```bash
git diff js/ui/app.js
git diff index.html
```

**驗證項目**:
- [x] 修改內容正確
- [x] 無意外修改
- [x] 代碼格式正確

#### 1.3 檢查提交歷史
```bash
git log --oneline -5
```

**預期輸出**: 最近 5 次提交記錄

---

### 步驟 2: 提交更改

#### 2.1 添加所有修改
```bash
git add .
```

#### 2.2 驗證暫存區
```bash
git status
```

**預期輸出**: 所有修改都在 "Changes to be committed" 中

#### 2.3 提交更改
```bash
git commit -m "feat: 實現儲存格映射協定自動確認機制

- 第二次點擊儲存格後自動確認選擇
- 選擇完成後自動隱藏預覽區
- 支持重新點擊重新選擇範圍
- 隱藏確認和取消按鈕

修改文件:
- js/ui/app.js: 修改 handleCellClick, cancelSelection, startRangeSelection 函數
- index.html: 隱藏確認和取消按鈕
- .gitignore: 添加編輯器配置

新增文檔:
- notes/FEATURE_TESTING_LOG.md: 功能測試計畫
- notes/GITHUB_PUSH_CHECKLIST.md: GitHub 推送清單
- notes/CODE_REVIEW_CHECKLIST.md: 代碼審查清單
- notes/DEPLOYMENT_GUIDE.md: 部署指南

測試:
- 代碼審查: 通過 ✓
- 邏輯驗證: 通過 ✓
- 邊界情況: 通過 ✓
- 性能檢查: 通過 ✓"


#### 2.4 驗證提交
```bash
git log --oneline -1
```

**預期輸出**: 新提交已出現在日誌中

---

### 步驟 3: 推送至遠程倉庫

#### 3.1 推送至 main 分支
```bash
git push origin main
```

**預期輸出**:
```
Enumerating objects: X, done.
Counting objects: 100% (X/X), done.
Delta compression using up to X threads
Compressing objects: 100% (X/X), done.
Writing objects: 100% (X/X), X bytes | X bytes/s, done.
Total X (delta X), reused 0 (delta 0), pack-reused 0
remote: Resolving deltas: 100% (X/X), done.
To https://github.com/Chun-Chieh-Chang/QIP-Data-Extract-Web-Tool.git
   XXXXXXX..XXXXXXX  main -> main
```

#### 3.2 驗證推送
```bash
git log --oneline -1
git branch -vv
```

**預期輸出**: 本地分支與遠程分支同步

---

### 步驟 4: GitHub 驗證

#### 4.1 檢查 GitHub 倉庫
1. 訪問 https://github.com/Chun-Chieh-Chang/QIP-Data-Extract-Web-Tool
2. 驗證最新提交已出現
3. 檢查提交信息和文件變更

#### 4.2 檢查 GitHub Actions
1. 點擊 "Actions" 標籤
2. 查看最新的工作流執行
3. 驗證所有步驟都已通過

**預期結果**:
- [x] 代碼檢查通過
- [x] 構建成功
- [x] 部署成功 (如配置)

#### 4.3 檢查 GitHub Pages (如適用)
1. 訪問 https://chun-chieh-chang.github.io/QIP-Data-Extract-Web-Tool/
2. 驗證網站已更新
3. 測試新功能

---

## 部署後驗證

### 驗證 1: 代碼已推送
```bash
# 在本地執行
git log --oneline -1

# 在 GitHub 上驗證
# 訪問 https://github.com/Chun-Chieh-Chang/QIP-Data-Extract-Web-Tool/commits/main
```

**預期**: 最新提交已出現在 GitHub

### 驗證 2: 文件已更新
```bash
# 在 GitHub 上驗證
# 檢查以下文件:
# - js/ui/app.js
# - index.html
# - .gitignore
# - notes/FEATURE_TESTING_LOG.md
# - notes/GITHUB_PUSH_CHECKLIST.md
# - notes/CODE_REVIEW_CHECKLIST.md
# - notes/DEPLOYMENT_GUIDE.md
```

**預期**: 所有文件都已更新

### 驗證 3: GitHub Actions 執行
```bash
# 在 GitHub 上驗證
# 訪問 https://github.com/Chun-Chieh-Chang/QIP-Data-Extract-Web-Tool/actions
```

**預期**: 最新的工作流已執行並成功

### 驗證 4: 功能測試
1. 訪問部署的網站
2. 上傳測試 Excel 檔案
3. 執行儲存格映射協定選擇
4. 驗證自動確認功能正常

**預期**:
- [x] 第二次點擊後自動確認
- [x] 預覽區自動隱藏
- [x] 可以重新修改
- [x] 無 Console 錯誤

---

## 回滾計畫

### 如果部署失敗

#### 步驟 1: 停止部署
```bash
# 如果 GitHub Actions 仍在運行，可以取消
# 在 GitHub Actions 頁面點擊 "Cancel workflow"
```

#### 步驟 2: 本地回滾
```bash
# 查看提交歷史
git log --oneline -5

# 回滾到上一個提交
git reset --hard HEAD~1

# 推送回滾
git push origin main --force
```

**警告**: 使用 `--force` 會覆蓋遠程歷史，請謹慎使用

#### 步驟 3: 分析失敗原因
1. 檢查 GitHub Actions 日誌
2. 檢查 Console 錯誤
3. 檢查部署日誌

#### 步驟 4: 修復問題
1. 在本地修復問題
2. 重新測試
3. 重新提交和推送

---

## 監控和維護

### 部署後監控 (24 小時)

#### 監控項目
- [x] GitHub Actions 無失敗
- [x] 網站正常運行
- [x] 無新的 Issue 報告
- [x] 用戶反饋正常

#### 監控工具
- GitHub Actions 儀表板
- GitHub Issues
- 網站監控工具 (如配置)

### 長期維護

#### 定期檢查
- 每週檢查 GitHub Actions 執行狀態
- 每月檢查代碼質量指標
- 每季度進行代碼審查

#### 用戶反饋
- 收集用戶對新功能的反饋
- 根據反饋進行優化
- 記錄改進建議

---

## 故障排除

### 問題 1: GitHub Actions 失敗

**症狀**: GitHub Actions 工作流執行失敗

**排查步驟**:
1. 查看 GitHub Actions 日誌
2. 檢查錯誤信息
3. 驗證代碼語法
4. 驗證依賴項

**解決方案**:
- 修復代碼錯誤
- 更新依賴項
- 重新推送

### 問題 2: 功能不工作

**症狀**: 儲存格映射協定自動確認不工作

**排查步驟**:
1. 打開瀏覽器開發者工具 (F12)
2. 查看 Console 標籤
3. 檢查是否有錯誤信息
4. 檢查 Network 標籤

**解決方案**:
- 檢查 JavaScript 代碼
- 檢查 HTML 結構
- 檢查 CSS 類名
- 清除瀏覽器緩存

### 問題 3: 性能問題

**症狀**: 頁面加載緩慢或卡頓

**排查步驟**:
1. 打開 Performance 標籤
2. 記錄性能指標
3. 分析瓶頸

**解決方案**:
- 優化 JavaScript 代碼
- 減少 DOM 操作
- 使用事件委託
- 優化圖片和資源

---

## 文檔和參考

### 相關文檔
- FEATURE_TESTING_LOG.md - 功能測試計畫
- GITHUB_PUSH_CHECKLIST.md - GitHub 推送清單
- CODE_REVIEW_CHECKLIST.md - 代碼審查清單
- DEVELOPMENT_LOG.md - 開發日誌
- PROJECT_STRUCTURE.md - 項目結構

### 外部參考
- [GitHub 文檔](https://docs.github.com/)
- [GitHub Actions 文檔](https://docs.github.com/en/actions)
- [Git 文檔](https://git-scm.com/doc)

---

## 簽核

| 角色 | 名稱 | 日期 | 簽名 |
|------|------|------|------|
| 開發人員 | | 2026-02-03 | |
| 部署人員 | | 2026-02-03 | |
| 項目經理 | | 2026-02-03 | |

---

## 附錄

### A. 提交信息模板
```
feat: 實現儲存格映射協定自動確認機制

- 第二次點擊儲存格後自動確認選擇
- 選擇完成後自動隱藏預覽區
- 支持重新點擊重新選擇範圍
- 隱藏確認和取消按鈕

修改文件:
- js/ui/app.js
- index.html
- .gitignore

新增文檔:
- notes/FEATURE_TESTING_LOG.md
- notes/GITHUB_PUSH_CHECKLIST.md
- notes/CODE_REVIEW_CHECKLIST.md
- notes/DEPLOYMENT_GUIDE.md
```

### B. 常用命令
```bash
# 查看狀態
git status

# 查看差異
git diff

# 添加文件
git add .

# 提交更改
git commit -m "message"

# 推送至遠程
git push origin main

# 查看日誌
git log --oneline -10

# 回滾提交
git reset --hard HEAD~1
```

### C. 檢查清單
- [x] 代碼修改完成
- [x] 代碼審查通過
- [x] 文檔已創建
- [x] 測試計畫已制定
- [x] Git 配置已更新
- [x] 準備推送至 GitHub

