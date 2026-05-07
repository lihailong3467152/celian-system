# 策链系统（CeLink）使用说明与系统架构

## 1. 项目概述

策链系统（CeLink）是一个基于 Streamlit 的第三方评估业务管理平台，面向咨询生产、项目协调、第三方机构资料维护、项目过程审批、文件评估、消息通知、待办审批和数据导入导出等场景。

系统采用单体应用架构，核心代码集中在 `app1.py`，数据默认存储在同级 SQLite 数据库 `performance.db` 中。首次启动时，系统会自动初始化数据库表结构、默认超级管理员、默认指标库和必要索引。

## 2. 核心能力

- 管理端和机构端双端登录
- 用户、机构、项目、文件、指标、日志统一管理
- 机构主账号和机构子账号权限隔离
- 机构信息维护和子账号变更审批
- 主评人管理及账号关联
- 项目 Gate 阶段审批
- 项目文件上传、审批、预览、下载
- 项目智库和评估指标库
- 待办事项和消息通知
- 数据导入导出
- 可视化大屏
- 操作日志审计
- 新系统自动初始化和旧库字段自动迁移

## 3. 技术栈

| 层级 | 技术 |
|---|---|
| Web 框架 | Streamlit |
| 数据库 | SQLite |
| 数据分析 | pandas |
| 可视化 | Plotly |
| Excel 读写 | openpyxl, xlrd |
| 密码加密 | bcrypt |
| PDF 导出 | ReportLab |
| 文件存储 | 本地文件系统 |

## 4. 目录结构

推荐项目目录如下：

```text
04/
├─ app1.py                 # 系统主程序
├─ requirements.txt        # Python 依赖清单
├─ README.md               # 使用说明和架构文档
├─ performance.db          # SQLite 数据库，首次启动自动生成
├─ uploads/                # 上传文件目录，首次使用可为空
├─ exports/                # 导出文件目录，首次使用可为空
└─ __pycache__/            # Python 缓存目录，可忽略
```

迁移到新电脑时，至少需要复制：

```text
app1.py
requirements.txt
README.md
```

如果要保留历史数据，还需要复制：

```text
performance.db
uploads/
exports/
```

如果要启动一个全新的空系统，不要复制旧的 `performance.db`。

## 5. 安装与启动

### 5.1 环境要求

建议使用：

- Windows 10/11
- Python 3.11
- pip

检查 Python 版本：

```bash
python --version
```

### 5.2 安装依赖

进入 `app1.py` 所在目录：

```bash
cd d:\Python\第三方评估系统\第三方评估平台\04
```

安装依赖：

```bash
pip install -r requirements.txt
```

如果网络较慢，可以使用国内镜像：

```bash
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 5.3 启动系统

```bash
streamlit run app1.py
```

或指定端口：

```bash
streamlit run app1.py --server.port 8502
```

启动后浏览器访问：

```text
http://localhost:8501
```

如果指定了 8502：

```text
http://localhost:8502
```

## 6. 首次初始化

系统首次启动时会自动创建同级数据库：

```text
performance.db
```

并自动初始化：

- 用户表
- 机构表
- 项目表
- 项目阶段表
- 项目文件表
- 主评人表
- 业绩记录表
- 培训记录表
- 指标库表
- 政策文件表
- 待办事项表
- 机构信息变更申请表
- 消息通知表
- 操作日志表
- 文件评估表
- 应用元数据表

系统也会创建默认超级管理员：

```text
用户名：admin
密码：Admin@123456
```

首次上线后建议立即修改默认密码。

## 7. 角色与权限

系统包含三类角色：

| 角色 | 标识 | 说明 |
|---|---|---|
| 超级管理员 | `super_admin` | 管理端账号，负责全局管理 |
| 机构主账号 | `org_admin` | 机构端主账号，负责本机构管理和审批 |
| 机构子账号 | `org_user` | 机构端子账号，负责项目资料提交和日常协作 |

### 7.1 登录端权限

| 登录端 | 允许角色 |
|---|---|
| 管理端 | 超级管理员 |
| 机构端 | 机构主账号、机构子账号 |
| 自动识别 | 根据账号角色自动进入对应端 |

### 7.2 页面权限

| 页面 | 超级管理员 | 机构主账号 | 机构子账号 |
|---|---:|---:|---:|
| 数据大盘 | 是 | 否 | 否 |
| 机构管理 | 是 | 否 | 否 |
| 账号管理 | 是 | 否 | 否 |
| 项目审核 | 是 | 否 | 否 |
| 日志查看 | 是 | 否 | 否 |
| 数据导出 | 是 | 否 | 否 |
| 审批待办 | 是 | 否 | 否 |
| 消息通知 | 是 | 是 | 是 |
| 可视化大屏 | 是 | 否 | 否 |
| 工作台 | 否 | 是 | 是 |
| 信息维护 | 否 | 是 | 是 |
| 子账号管理 | 否 | 是 | 否 |
| 项目管理 | 否 | 是 | 是 |
| 项目智库 | 是 | 是 | 是 |
| 待办事项 | 否 | 是 | 是 |

权限集中配置在 `app1.py` 顶部：

- `ROLE_NAMES`
- `CLIENT_ALLOWED_ROLES`
- `PAGE_MENU`
- `PAGE_ACCESS`

页面入口和侧边栏菜单均依赖这些配置，避免通过修改 session 绕过菜单访问无权页面。

## 8. 系统架构

### 8.1 总体架构

```text
浏览器
  │
  │ HTTP
  ▼
Streamlit 应用 app1.py
  │
  ├─ 认证与权限
  ├─ 管理端页面
  ├─ 机构端页面
  ├─ 文件上传与预览
  ├─ 导入导出
  ├─ 消息与待办
  ├─ 可视化大屏
  │
  ├─ SQLite 数据库 performance.db
  │    ├─ users
  │    ├─ organizations
  │    ├─ projects
  │    ├─ project_steps
  │    ├─ project_files
  │    ├─ evaluators
  │    ├─ todos
  │    ├─ messages
  │    ├─ logs
  │    └─ ...
  │
  └─ 本地文件目录
       ├─ uploads/
       └─ exports/
```

### 8.2 应用层模块

| 模块 | 说明 |
|---|---|
| 数据库初始化 | 自动建表、补字段、建索引、初始化默认数据 |
| 认证系统 | 用户名、手机号、邮箱登录，bcrypt 密码校验 |
| 权限系统 | 角色、客户端、页面访问统一控制 |
| 管理端 | 机构、账号、项目审核、日志、导入导出、指标、大屏 |
| 机构端 | 工作台、信息维护、子账号、项目、智库、待办、消息 |
| 文件系统 | 上传文件保存、文件预览、下载 |
| 消息待办 | 系统消息、未读红点、待办红点、高优先级审批 |
| 数据导入导出 | Excel 导出、导入预览、冲突处理 |
| 可视化 | 项目、文件、评估指标趋势展示 |

### 8.3 数据层设计

系统使用 SQLite，默认数据库文件：

```text
performance.db
```

数据库连接通过 `get_connection()` 创建，查询统一通过 `execute_query()` 执行。只读查询会通过 Streamlit 缓存优化，写操作后会触发缓存失效。

## 9. 数据库表说明

### 9.1 users 用户表

保存所有账号。

关键字段：

- `username`：用户名，唯一
- `password_hash`：bcrypt 加密密码
- `role`：角色
- `org_id`：所属机构
- `phone`：手机号
- `email`：邮箱
- `real_name`：姓名
- `status`：账号状态，`active` 或 `inactive`

### 9.2 organizations 机构表

保存机构基本信息。

关键字段：

- `name`：机构名称
- `org_type`：机构类型
- `credit_code`：统一社会信用代码
- `legal_person`：法定代表人
- `contact_person`：联系人
- `contact_phone`：联系电话
- `contact_email`：联系邮箱
- `address`：机构地址
- `description`：机构简介

### 9.3 projects 项目表

保存项目主信息。

关键字段：

- `name`：项目名称
- `org_id`：所属机构
- `category`：项目分类
- `subcategory`：子分类
- `current_stage`：当前 Gate
- `status`：项目状态

### 9.4 project_steps 项目阶段表

保存项目 Gate 阶段审批状态。

系统当前按 G0 到 G8 共 9 个阶段维护。

关键字段：

- `project_id`
- `stage`
- `status`
- `submitted_by`
- `submitted_at`
- `reviewed_by`
- `reviewed_at`
- `review_comment`

### 9.5 project_files 项目文件表

保存项目上传文件记录。

关键字段：

- `project_id`
- `step_id`
- `title`
- `file_type`
- `category`
- `subcategory`
- `file_path`
- `file_name`
- `approval_status`
- `approved_by`
- `approved_at`

### 9.6 evaluators 主评人表

保存机构主评人。

关键字段：

- `org_id`
- `account_user_id`：关联机构主账号或机构子账号
- `name`
- `title`
- `specialty`
- `phone`
- `email`
- `status`

关系说明：

- 一个主评人最多关联一个机构账号
- 一个机构账号可以关联多个主评人
- 只有机构主账号可以维护主评人及账号关联
- 机构子账号只能查看与自己账号关联的主评人

### 9.7 org_info_update_requests 机构信息变更申请表

保存机构子账号提交的机构信息变更申请。

关键字段：

- `org_id`
- `submitted_by`
- `approver_id`
- `status`
- `old_data`
- `new_data`
- `review_comment`
- `submitted_at`
- `reviewed_at`

审批通过后才会更新 `organizations` 表。

### 9.8 todos 待办事项表

保存用户待办和审批任务。

关键字段：

- `user_id`
- `title`
- `content`
- `status`
- `priority`
- `related_type`
- `related_id`
- `due_date`
- `completed_at`

机构信息变更审批会以高优先级待办进入机构主账号的待办列表。

### 9.9 messages 消息通知表

保存系统消息。

关键字段：

- `user_id`
- `title`
- `content`
- `msg_type`
- `is_read`
- `created_at`

### 9.10 logs 操作日志表

保存关键操作记录。

关键字段：

- `user_id`
- `username`
- `org_name`
- `action`
- `module`
- `ip_address`
- `details`
- `created_at`

## 10. 管理端使用说明

### 10.1 登录管理端

默认账号：

```text
用户名：admin
密码：Admin@123456
```

登录时可以选择：

- 自动识别
- 管理端

超级管理员不能登录机构端。

### 10.2 数据大盘

用于查看系统总体概况，包括：

- 机构数量
- 用户数量
- 项目数量
- 待审批数量
- 最近项目
- 消息和统计概览

### 10.3 机构管理

功能：

- 新增机构
- 编辑机构
- 启用机构
- 停用机构
- 删除机构
- 重置机构主账号密码

新增机构时，系统会自动创建机构主账号。

默认机构主账号密码：

```text
Org@123456
```

### 10.4 账号管理

功能：

- 查看全部账号
- 按角色筛选
- 按状态筛选
- 冻结账号
- 解冻账号
- 重置密码
- 新增账号
- 删除账号

管理端账号重置后的默认密码：

```text
Reset@123456
```

### 10.5 项目审核

管理端可以审核机构提交的项目阶段和文件。

常见状态：

- `pending`：待审核
- `approved`：已通过
- `rejected`：已驳回

### 10.6 日志查看

可按用户、模块、时间范围筛选操作日志。

日志用于审计：

- 登录
- 退出
- 密码修改
- 机构管理
- 账号管理
- 文件上传
- 项目审批
- 数据导入导出

### 10.7 数据导出和导入

导出支持：

- 用户数据
- 机构数据
- 项目数据
- 项目文件数据
- 用户日志数据

导入支持：

- 用户数据
- 机构数据

导入策略：

- 仅新增，跳过已有数据
- 按唯一字段更新已有数据

默认建议使用：

```text
仅新增，跳过已有数据
```

系统会在导入前显示预览：

- 总行数
- 新增数
- 跳过数
- 无效数
- 每行处理方式和原因

### 10.8 项目智库

管理端可以维护指标库，用于文件评估和项目资料管理。

### 10.9 可视化大屏

展示文件、项目、评估数据的整体趋势和质量分布。

## 11. 机构端使用说明

### 11.1 机构主账号登录

机构主账号由管理端创建机构时自动生成。

默认密码：

```text
Org@123456
```

机构主账号可以：

- 修改机构信息
- 管理子账号
- 管理主评人
- 管理项目
- 上传项目资料
- 审批机构子账号提交的机构信息变更申请
- 查看待办和消息

### 11.2 机构子账号登录

机构子账号由机构主账号创建。

机构子账号可以：

- 查看工作台
- 查看和提交机构信息变更申请
- 查看与自己关联的主评人
- 管理项目资料
- 上传项目文件
- 查看待办和消息

机构子账号不能：

- 管理子账号
- 管理其他主评人
- 调整主评人与账号关联
- 直接修改机构信息

### 11.3 信息维护

包含四个页签：

- 机构信息
- 主评人管理
- 业绩记录
- 培训记录

#### 机构信息

机构主账号更新机构信息时直接生效。

机构子账号更新机构信息时：

1. 系统创建机构信息变更申请
2. 系统给机构主账号创建高优先级待办
3. 机构主账号在待办列表审批
4. 审批通过后才更新机构表
5. 审批驳回则不更新机构表

#### 主评人管理

机构主账号可以：

- 新增主评人
- 编辑主评人
- 停用主评人
- 启用主评人
- 删除主评人
- 维护主评人与机构账号的关联

机构子账号只能查看当前账号关联的主评人。

#### 业绩记录

用于维护机构业绩信息。

#### 培训记录

用于维护机构培训信息。

### 11.4 子账号管理

仅机构主账号可访问。

功能：

- 查看子账号
- 新增子账号
- 编辑子账号
- 冻结子账号
- 解冻子账号
- 重置子账号密码
- 删除子账号
- 查看子账号关联的主评人

机构主账号重置子账号密码后，默认密码为：

```text
Reset@123456
```

### 11.5 项目管理

机构账号可以创建项目并上传项目资料。

项目按 Gate 阶段推进，系统会自动维护项目阶段记录。

常见流程：

1. 创建项目
2. 上传项目文件
3. 提交审核
4. 管理端审核
5. 审核通过后进入下一阶段
6. 全部阶段通过后项目完成

### 11.6 项目智库

机构端可以上传和管理项目知识文件。

支持文件类型：

- PDF
- Word
- Excel
- TXT

### 11.7 待办事项

待办按状态和优先级排序：

1. 待处理优先
2. 高优先级优先
3. 最新创建优先

优先级颜色：

| 优先级 | 颜色 |
|---|---|
| high | 红色 |
| medium | 黄色 |
| low | 绿色 |

机构信息变更审批默认是高优先级红色待办。

审批完成后：

- 待办状态变为已完成
- 优先级降为低
- 文案由“请审批”变为“已审批”

### 11.8 消息通知

系统会发送消息给相关用户。

常见消息：

- 登录成功
- 密码修改成功
- 项目审批通知
- 文件审批通知
- 机构信息变更审批结果

侧边栏会显示未读消息红色角标。

## 12. 关键业务流程

### 12.1 创建新机构流程

```text
超级管理员登录管理端
  ↓
进入机构管理
  ↓
新增机构
  ↓
系统自动创建机构主账号
  ↓
机构主账号用默认密码登录机构端
  ↓
机构主账号修改密码
```

### 12.2 创建机构子账号流程

```text
机构主账号登录机构端
  ↓
进入子账号管理
  ↓
新增子账号
  ↓
设置初始密码
  ↓
子账号登录机构端
```

### 12.3 子账号申请修改机构信息流程

```text
机构子账号登录机构端
  ↓
进入信息维护
  ↓
修改机构信息并提交
  ↓
系统创建机构信息变更申请
  ↓
系统给机构主账号创建高优先级待办
  ↓
机构主账号进入待办事项
  ↓
查看变更前后字段
  ↓
通过或驳回
  ↓
通过后更新机构信息
```

### 12.4 主评人与账号关联流程

```text
机构主账号登录
  ↓
进入信息维护 -> 主评人管理
  ↓
新增或编辑主评人
  ↓
选择关联账号
  ↓
一个账号可关联多个主评人
```

### 12.5 密码修改和重置流程

用户自己修改密码：

```text
登录系统
  ↓
点击侧边栏修改密码
  ↓
输入原密码和新密码
  ↓
保存后用新密码登录
```

管理端重置机构主账号密码：

```text
管理端 -> 机构管理 -> 重置主账号密码
默认密码：Org@123456
```

管理端重置账号密码：

```text
管理端 -> 账号管理 -> 重置密码
默认密码：Reset@123456
```

机构端重置子账号密码：

```text
机构端 -> 子账号管理 -> 重置密码
默认密码：Reset@123456
```

## 13. 数据导入模板说明

### 13.1 用户数据导入字段

建议 Excel 字段：

```text
用户名
姓名
手机号
邮箱
角色
所属机构
状态
```

角色可用值：

```text
org_admin
org_user
```

状态可用值：

```text
active
inactive
```

### 13.2 机构数据导入字段

建议 Excel 字段：

```text
机构名称
机构类型
信用代码
法定代表人
联系人
联系电话
联系邮箱
机构地址
状态
```

### 13.3 导入冲突规则

用户数据按以下字段判断冲突：

- 用户名
- 手机号
- 邮箱

机构数据按以下字段判断冲突：

- 机构名称
- 信用代码

默认策略会跳过已有数据，不覆盖。

## 14. 文件上传和存储

上传文件默认保存在：

```text
uploads/
```

导出文件默认保存在：

```text
exports/
```

支持上传类型：

```text
pdf
docx
doc
xlsx
xls
txt
```

注意：

- 大文件预览可能自动切换为下载
- PDF、Excel、文本文件支持不同程度预览
- Word 文件建议下载后查看

## 15. 备份与恢复

### 15.1 备份

建议定期备份：

```text
performance.db
uploads/
exports/
```

Windows PowerShell 示例：

```powershell
Copy-Item .\performance.db .\backup\performance_$(Get-Date -Format yyyyMMdd_HHmmss).db
Copy-Item .\uploads .\backup\uploads_$(Get-Date -Format yyyyMMdd_HHmmss) -Recurse
Copy-Item .\exports .\backup\exports_$(Get-Date -Format yyyyMMdd_HHmmss) -Recurse
```

### 15.2 恢复

停止 Streamlit 服务后，将备份文件复制回项目目录：

```text
performance.db
uploads/
exports/
```

然后重新启动：

```bash
streamlit run app1.py
```

## 16. 迁移到新电脑

### 16.1 迁移完整系统

复制以下内容到新电脑同一目录：

```text
app1.py
requirements.txt
README.md
performance.db
uploads/
exports/
```

安装依赖：

```bash
pip install -r requirements.txt
```

启动：

```bash
streamlit run app1.py
```

### 16.2 启动全新空系统

只复制：

```text
app1.py
requirements.txt
README.md
```

不要复制：

```text
performance.db
uploads/
exports/
```

启动系统后会自动生成新数据库。

## 17. 局域网访问

如果需要局域网其他电脑访问，可以启动：

```bash
streamlit run app1.py --server.address 0.0.0.0 --server.port 8501
```

然后在其他电脑访问：

```text
http://服务器IP:8501
```

注意：

- Windows 防火墙需要放行端口
- 多人同时使用 SQLite 时不适合高并发
- 文件上传目录必须在服务器本机存在

## 18. 运维建议

### 18.1 上线前检查

- 修改默认超级管理员密码
- 确认 `requirements.txt` 已安装成功
- 确认 `uploads/` 可写
- 确认 `exports/` 可写
- 确认数据库文件可写
- 确认防火墙端口已放行
- 确认备份策略已建立

### 18.2 日常检查

- 定期备份 `performance.db`
- 定期备份 `uploads/`
- 检查未处理审批待办
- 检查异常操作日志
- 清理无用导出文件
- 检查磁盘空间

### 18.3 密码安全

系统使用 bcrypt 保存密码，不保存明文密码。

建议：

- 首次登录后立即修改默认密码
- 定期重置离职人员账号
- 冻结不再使用的账号
- 不共用超级管理员账号

## 19. 常见问题

### 19.1 启动时报缺少模块

执行：

```bash
pip install -r requirements.txt
```

如果仍缺包，确认 Python 环境和 pip 是否一致：

```bash
python -m pip install -r requirements.txt
```

### 19.2 端口被占用

指定其他端口：

```bash
streamlit run app1.py --server.port 8502
```

### 19.3 登录失败

检查：

- 账号是否被冻结
- 密码是否正确
- 是否选择了正确登录端
- 超级管理员是否登录管理端
- 机构账号是否登录机构端

自动识别模式会根据角色自动进入对应端。

### 19.4 页面没有权限

说明当前角色无权访问该页面。

当前权限由集中配置控制：

- `PAGE_MENU`
- `PAGE_ACCESS`
- `can_access_page()`

### 19.5 上传文件后无法预览

可能原因：

- 文件太大，系统自动切换为下载
- 文件路径不存在
- 浏览器不支持内联预览
- Word 文件建议下载后查看

### 19.6 导入 Excel 失败

检查：

- 文件是否为 `.xlsx` 或 `.xls`
- 字段名称是否符合模板
- 是否安装 `openpyxl`
- 是否安装 `xlrd`
- 是否存在唯一字段冲突

### 19.7 修改机构信息后没有立即生效

如果是机构子账号提交，则不会直接生效。

需要机构主账号进入：

```text
待办事项 -> 待办列表
```

审批通过后才会更新机构信息。

### 19.8 待办红点不显示

待办红点统计口径：

```sql
SELECT COUNT(*)
FROM todos
WHERE user_id = 当前登录用户
  AND status = 'pending'
```

如果没有红点，说明当前用户没有待处理待办。

## 20. 二次开发说明

### 20.1 新增页面

新增页面时应同时修改：

1. 新增渲染函数，例如 `render_org_xxx()`
2. 在 `PAGE_MENU` 中加入菜单项
3. 确认 `PAGE_ACCESS` 自动生成结果符合预期
4. 在主路由表中加入页面函数

不要只在侧边栏加按钮，否则可能出现路由无效或权限绕过。

### 20.2 新增权限

优先修改集中权限配置：

```python
PAGE_MENU
PAGE_ACCESS
CLIENT_ALLOWED_ROLES
```

不要在多个页面内重复写散落的角色判断。

### 20.3 新增数据库字段

需要同时处理：

1. 新建库建表语句
2. 旧库自动迁移逻辑
3. 索引
4. 读写逻辑
5. 缓存失效

示例模式：

```python
cursor.execute("PRAGMA table_info(table_name)")
columns = {row[1] for row in cursor.fetchall()}
if 'new_column' not in columns:
    cursor.execute("ALTER TABLE table_name ADD COLUMN new_column TEXT")
```

### 20.4 写操作注意事项

所有写操作建议使用：

```python
execute_query(sql, params, commit=True)
```

写操作后会触发查询缓存失效。

### 20.5 日志审计

关键操作应调用：

```python
add_log(...)
```

建议记录：

- 操作人
- 机构
- 模块
- 操作类型
- 关键业务 ID

## 21. 当前默认密码汇总

| 场景 | 默认密码 |
|---|---|
| 默认超级管理员 | `Admin@123456` |
| 新建机构自动生成主账号 | `Org@123456` |
| 管理端重置机构主账号 | `Org@123456` |
| 管理端账号管理重置密码 | `Reset@123456` |
| 机构端重置子账号密码 | `Reset@123456` |

## 22. 依赖安装清单

依赖文件：

```text
requirements.txt
```

安装命令：

```bash
pip install -r requirements.txt
```

核心依赖包括：

- streamlit
- pandas
- bcrypt
- plotly
- reportlab
- openpyxl
- xlrd

完整依赖以 `requirements.txt` 为准。

## 23. 版本和维护说明

当前项目是单文件 Streamlit 应用，适合：

- 单机部署
- 局域网小规模多人使用
- 快速上线和内部管理

如果后续用户量、并发量或数据量明显增长，建议升级为：

- FastAPI 或 Django 后端
- PostgreSQL 或 MySQL 数据库
- 独立前端
- 对象存储或文件服务器
- 更细粒度 RBAC 权限模型
- 后台任务队列

## 24. 快速启动摘要

全新部署最短步骤：

```bash
cd 项目目录
pip install -r requirements.txt
streamlit run app1.py
```

浏览器访问：

```text
http://localhost:8501
```

默认登录：

```text
用户名：admin
密码：Admin@123456
```

首次登录立即强制修改默认密码。
