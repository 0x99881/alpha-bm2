# CHANGELOG

## 2026-04-16 Stable Baseline

### 这轮主要完成的整理
- 收紧了 `ExcelStore`、repository、presenter、service 之间的边界。
- presenter 不再直接依赖 `store` 的私有小工具。
- `record_repository.py` 已接管部分核心数据访问，不再只是简单转发。
- 前端脚本已拆成多个小文件：
  - `theme.js`
  - `scores.js`
  - `members_sort.js`
  - `profit_calendar.js`
  - `utils.js`
  - `app.js` 仅保留总入口
- 新增最小自动检查脚本：`scripts/smoke_check.py`

### 当前版本适合做什么
- 作为后续继续维护和小步迭代的稳定基线。
- 作为回退点保存。
- 作为继续补测试、补文档、继续做小范围解耦的起点。

### 当前版本不适合做什么
- 不适合作为“彻底完成重构”的终点版本理解。
- 不适合在没有进一步检查的前提下继续做大规模架构迁移。

### 已知仍然存在的过渡问题
- 旧 mixin 体系仍然存在，项目仍处于“新结构 + 旧兼容层”并存状态。
- README 当前仍有乱码，不能作为可靠说明文档使用。
- 默认“下一个日期”依然取决于现有 Excel 表里的最新记录；如果表里已经有更靠后的日期，页面会继续顺延。
- smoke check 目前只覆盖最低限度检查，还没有覆盖真实浏览器交互链路。

### 冻结建议
- 可以冻结保存当前版本。
- 冻结后，后续优先做小步修改，每次改动前后都运行 `python scripts/smoke_check.py`。
