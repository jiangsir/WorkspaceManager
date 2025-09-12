雲端文件:
https://docs.google.com/document/d/1UdZUPG6Ob7vi__DrhqKDIiipx6T_K828J4Ga8I7d5-Q/edit?usp=sharing

進入環境:
列出現有環境
conda env list

創建一個新的環境
conda create -n appsscript

切換進入 appsscript

conda activate appsscript

使用 clasp 進行本地開發
node -v
npm install -g @google/clasp
clasp login
https://script.google.com/home/usersettings
clasp create --type webapp --title "WorkspaceManager"

clasp push --force

https://script.google.com/

手動部屬成"網頁應用程式"
https://script.google.com/home/projects/1SyJdsOSKKdEElKYxfLOcYSrEw7Babnc9KtcC9go8xgrm33hZOtHiK_3R/edit

並取得 URL 觀看結果。

要維持網址不變的方式:

進入 AppsScript 專案
點擊 管理部署作業->編輯->建立新版本->部署

就可以創建新版本，但沿用 URL。
