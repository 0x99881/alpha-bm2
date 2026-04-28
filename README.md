# BM2 管理系统

BM2 是一个 Python / Flask 项目，用来管理成员、每日积分、排名展示和线上同步。

当前项目已经完成架构收口：

- 本地数据主库：SQLite
- 线上 / 手机端同步源：Supabase
- Excel：只作为导出、备份、历史导入格式
- 旧 JSON / cloud blob / bm2_cloud 路线：已废弃

## 当前数据流

本地使用：

```text
页面 / web.py
  -> service
  -> SQLite repository / writer
  -> bm2_local.db
```

线上 / 手机端使用：

```text
页面 / web.py
  -> service
  -> Supabase client
  -> Supabase
```

本地和线上同步：

```text
SQLite
  <-> SyncService
  <-> Supabase
```

Excel：

```text
SQLite -> Excel 导出
Excel  -> 历史导入 / 迁移来源
```

## 主要功能

- 每日积分提交
- 手机 / 线上积分提交
- 成员新增、修改、停用、删除、排序
- 积分排名和积分总览
- 盈亏日历展示
- Supabase push / pull / sync
- Excel 导出

## 架构规则

后续开发必须遵守：

- SQLite 是唯一的本地主数据源。
- Supabase 是唯一的线上 / 手机端同步源。
- Excel 不能再作为业务主写入路径。
- JSON 文件同步、bm2_cloud、cloud_sync、blob_sync 等旧路线不能恢复。
- `web.py` 只负责请求和响应。
- `services/` 负责业务流程。
- `repositories/` 负责数据读写。
- `bm2/store.py` 只能是薄入口，不承载业务逻辑。
- `static/` 是前端源码。
- `public/` 如果存在，只能是构建产物，不手写维护。

更详细的后续开发规则见：

```text
AGENTS.md
docs/development_rules.md
docs/data_flow.md
```

## 运行方式

安装依赖：

```bash
pip install -r requirements.txt
```

本地启动：

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

如果要启用线上 / 手机端同步，需要配置 Supabase 环境变量。

常用配置项：

```text
SUPABASE_URL
SUPABASE_KEY
```

只读线上模式：

```text
BM2_READ_ONLY=1
```

在 Vercel 环境中，项目会按线上只读 / 手机端模式运行。

## 测试和检查

每次改动后建议运行：

```bash
python scripts/architecture_check.py
python scripts/smoke_check.py
```

`architecture_check.py` 会阻止旧架构路线重新出现。

## 项目结构

```text
app.py                 Flask 启动入口
bm2/web.py             HTTP 路由
bm2/store.py           薄 facade
bm2/services/          业务流程
bm2/repositories/      SQLite / Supabase 数据访问
bm2/excel/             Excel 导入导出和表结构处理
bm2/presenters/        页面展示数据组装
static/                前端源码
templates/             页面模板
scripts/               架构检查和冒烟检查
docs/                  架构说明和 Supabase SQL
AGENTS.md              给 AI/开发者看的项目规则
```

## 不要上传的文件

这些文件不要放进公开代码包：

```text
.env.local
bm2_local.db
BM2记录_*.xlsx
__pycache__/
.tmp_*/
public/
bm2_cloud/
```

## 当前状态

代码层面的桥梁已经建好：

```text
本地 SQLite <-> SyncService <-> Supabase
```

也就是说，后续功能应该继续围绕这条主线扩展，不要再新增 Excel 主写入、JSON 同步或旧 cloud blob 通道。
