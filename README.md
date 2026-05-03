# BM2 管理系统

本项目是本地单用户管理工具。

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
