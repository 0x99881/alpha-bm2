# BM2 管理系统

BM2 是一个 Flask 项目，用于管理成员、每日积分、排行展示、盈亏日历和线上同步。

## 当前架构

```text
本地使用:   Flask 路由 -> 应用服务 -> SQLite
线上使用:   Flask 路由 -> 应用服务 -> Supabase
同步:       SQLite <-> SyncService <-> Supabase
Excel:      历史导入 / 导出文件
```

核心规则：

- SQLite 是本地主数据源。
- Supabase 是线上和手机端同步源。
- Excel 只做导出、备份格式和历史导入。
- JSON 文件同步、旧 cloud/blob 路径、旧 store 门面都已退出主流程。

## 主要功能

- 每日积分提交
- 手机端 / 线上积分提交
- 成员新增、修改、停用、删除、排序
- 积分排行和积分总览
- 盈亏日历展示
- Supabase push / pull
- Excel 导入和导出

## 主路径

保存积分：

```text
web_score_routes -> DailyEntryService -> SQLiteEntryWriter -> SQLite
```

成员管理：

```text
web_member_routes -> MemberService -> SQLite
```

页面展示：

```text
web routes -> StoreQueryService -> Presenter -> SQLite/Supabase
```

Excel 导入：

```text
ExcelImportService -> SQLite
```

Excel 导出：

```text
ExcelExportService -> SQLiteToExcelExporter -> ExcelWorkbookWriter
```

## 运行方式

安装依赖：

```bash
pip install -r requirements.txt
```

启动：

```bash
python app.py
```

浏览器打开：

```text
http://127.0.0.1:5000
```

Windows 下也可以双击：

```text
start.bat
```

## Supabase 配置

如需启用线上或手机端同步，需要配置：

```text
SUPABASE_URL
SUPABASE_KEY
```

只读线上模式：

```text
BM2_READ_ONLY=1
```

## 检查

每次改动后运行：

```bash
python scripts/architecture_check.py
python scripts/smoke_check.py
```

这些检查会阻止旧 store、旧保存入口、Excel 主流程、副作用写入和魔法转发重新出现。

完整回归门禁：

```bash
python scripts/check_all.py
```

它会依次运行编译检查、冒烟检查、架构检查、pytest 测试和旧路径关键词扫描。测试只使用临时目录、临时 SQLite 和临时 Excel，不会访问真实本地数据或真实云端。

## 项目结构

```text
app.py                         Flask 启动入口
bm2/web*.py                    HTTP 路由
bm2/services/                  应用服务和业务流程
bm2/repositories/              SQLite / Supabase 数据访问
bm2/excel/                     Excel 导入、导出和工作簿结构
bm2/presenters/                页面展示数据组装
static/                        前端源码
templates/                     页面模板
scripts/                       架构检查和冒烟检查
docs/                          架构说明
AGENTS.md                      项目规则
```

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
