# 币安 Alpha 多号管理系统

本项目是币安 Alpha 多号管理的本地单用户工具。（本地数据库、历史表格和浏览器草稿仍兼容旧前缀 `bm2` / `BM2记录_*.xlsx`，以免影响现有数据。）

## 架构

```text
本地数据：SQLite
网页/手机：Supabase
同步桥：SQLite <-> SyncService <-> Supabase
Excel：只做导入、导出和旧数据兼容
前端源码：static/
页面模板：templates/
```

SQLite 是本地正式数据来源。Excel 不能再作为主要保存路径。

## 启动

```bash
python app.py
```

或双击：

```text
start.bat
```

## 基础检查

只保留两个基础检查：

```bash
python scripts/smoke_check.py
python scripts/architecture_check.py
```

也可以一次运行：

```bash
python scripts/check_all.py
```

`check_all.py` 只会运行冒烟检查和架构检查。

## 不要打包或提交

```text
.env.local
bm2_local.db
BM2记录_*.xlsx
__pycache__/
.tmp_*/
public/
bm2_cloud/
_legacy_quarantine_*/
```
