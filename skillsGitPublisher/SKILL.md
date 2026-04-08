---
name: skillsGitPublisher
description: 管理並發佈 C:\Users\NMMIS\.codex\skills 到指定 GitHub 倉庫。當使用者要求推送 skills、commit 後 push、同步 .codex/skills 到 GitHub、或在完成明確里程碑後發佈 skills 時使用。只管理 C:\Users\NMMIS\.codex\skills，不處理其他目錄，也不寫入任何 GitHub token 或密碼。
---

# Skills Git Publisher

此 skill 只管理 `C:\Users\NMMIS\.codex\skills` 這個目錄的 Git 版本控管與發佈。

## 何時使用

- 使用者要求把目前 skills 推送到 GitHub。
- 使用者要求先 commit 再 push。
- 使用者說某個 MMIS skill 里程碑完成，要發佈。
- 使用者要求同步 `.codex/skills` 更新到 GitHub。

## 管理範圍

- 只允許管理 `C:\Users\NMMIS\.codex\skills`
- 不處理其他 repo 或其他資料夾
- 不寫死 token、密碼、PAT、SSH 私鑰
- GitHub 驗證沿用本機既有憑證機制

## 固定設定

- repo root: `C:\Users\NMMIS\.codex\skills`
- remote name: `origin`
- remote url: `https://github.com/Fukuei-roc/mmisAutomationSkills.git`
- repo-local git user.name: `Fukuei-roc`
- repo-local git user.email: `f113097@yahoo.com.tw`

## 里程碑定義

以下任一項成立，視為可 commit 的明確里程碑：

- 新增一個可用的 skill
- 完成一個可測試的功能
- 完成一次重構且驗證通過
- 完成檔名規則升級
- 完成 formatter 擴充
- 完成查詢參數化
- 完成文件與腳本同步更新，且行為可追蹤

若只有暫時除錯、半成品、未驗證 patch、或混雜多個未完成方向，先不要 push。

## Commit Message 規範

優先使用一致格式：

```text
type(scope): milestone summary
```

建議 `type`：

- `feat`
- `fix`
- `refactor`
- `docs`
- `chore`

建議 `scope`：

- `skill`
- `excel`
- `query`
- `login`
- `git`
- `knowledge`

禁止使用模糊訊息，例如：

- `update`
- `fix`
- `change`
- `misc`

## 執行流程

1. 先執行 `git status --short`
2. 檢查目前是否為 `C:\Users\NMMIS\.codex\skills` repo
3. 若尚未初始化，建立 Git 倉庫
4. 設定 repo-local `user.name` 與 `user.email`
5. 檢查 `origin` 是否存在且 URL 正確，不正確則修正
6. 檢查 working tree 是否有變更
7. 若沒有變更，不要空 commit，直接回報
8. 依本次里程碑撰寫清楚的 commit message
9. 執行 `git add .`
10. 執行 `git commit -m "<message>"`
11. 檢查目前分支，預設推送目前分支；若無分支則建立 `main`
12. 執行 `git push -u origin <branch>` 或 `git push origin <branch>`

## 錯誤處理

- 尚未初始化 Git：先 `git init`
- remote 不存在：新增 `origin`
- remote URL 錯誤：改成指定 GitHub URL
- 尚未登入 GitHub：停止並回報「本機 GitHub 驗證尚未完成」
- push 被拒絕：回報 Git 錯誤內容，不要覆寫歷史
- working tree 為空：不要 commit，直接回報無變更
- 沒有可提交的變更：直接回報 `no changes`

## 執行方式

優先執行 helper script：

```powershell
python C:\Users\NMMIS\.codex\skills\skillsGitPublisher\scripts\publish_skills_git.py --message "feat(skill): add reusable Git publisher skill"
```

若只要初始化與檢查：

```powershell
python C:\Users\NMMIS\.codex\skills\skillsGitPublisher\scripts\publish_skills_git.py --check-only
```

若要只 commit 不 push：

```powershell
python C:\Users\NMMIS\.codex\skills\skillsGitPublisher\scripts\publish_skills_git.py --message "docs(skill): update Git publisher usage" --skip-push
```

## 回覆要求

回覆時至少包含：

- repo 是否已初始化
- remote 是否正確
- 目前分支
- 是否有變更
- commit message
- 是否 commit 成功
- 是否 push 成功
- 若失敗，具體 Git 錯誤
