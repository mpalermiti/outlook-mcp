<!-- mcp-name: io.github.mpalermiti/outlook-mcp -->

# outlook-mcp

[English](README.md) | 简体中文

基于 Microsoft Graph API 的 Outlook 个人账户 MCP 服务器。

[![PyPI](https://img.shields.io/pypi/v/outlook-graph-mcp.svg)](https://pypi.org/project/outlook-graph-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-graph-mcp.svg)](https://pypi.org/project/outlook-graph-mcp/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![MCP Registry](https://img.shields.io/badge/MCP_Registry-listed-green)](https://registry.modelcontextprotocol.io/v0/servers?search=mpalermiti)

> **仅支持个人微软账户** — `@outlook.com`、`@hotmail.com`、`@live.com`。v1 不支持工作/学校账户（Entra ID）。

> **免责声明：** 独立开源项目，与 Microsoft Corporation 无隶属、合作或受其支持关系。"Outlook" 和 "Microsoft Graph" 是微软商标。

---

## 适用人群

以下场景适合你：

- **Agent 构建者**：把 Outlook 接入自己的基础设施（OpenClaw、Claude Code、Cursor、自建 MCP 宿主），需要**带类型定义的工具接口**——而不是让 agent 解析 stdout
- 基于**个人微软账户**（Outlook.com / Hotmail / Live）开发，想要完全掌控：自带 Azure 应用（BYOID），没有企业授权流程，不共享 client ID
- 需要**真正的全覆盖**——邮件、日历、联系人、待办、草稿、文件夹、批量操作、会话线程——而不是只包了邮件或只包了日历的壳
- 注重安全：令牌存操作系统钥匙串（macOS Keychain、Linux libsecret——除非你主动选择否则绝不落明文）、细粒度 `allow_categories`、可选 `read_only` 模式、零遥测

**不适合你**的场景：需要工作/学校 M365 账户（请用微软官方工具——Entra ID 认证和管理员授权流程超出本项目范围），或者一个基础的纯邮件客户端就够用（本项目有 65 个工具——对"读一下收件箱"来说太多了）。

### 与其他 Outlook 工具的差异

个人 Outlook 领域唯一的**一等公民 MCP 服务器**——多数替代品是 bash 脚本或让 agent 外呼的 CLI 壳。这个区别很关键：agent 拿到的是带结构化参数/返回值的类型化工具 schema，不是要自己解析的 stdout。其他别处找不到的能力：`/$batch` 优化的批量整理（大批量操作快 10-20 倍）、带名称解析的递归文件夹操作、按类别的细粒度权限、多账号支持、完整附件写入路径（含草稿 >3MB 上传会话）。

---

## 能做什么

给你的 AI agent 完整的 Outlook 访问能力。这些提示词开箱即用：

- *"总结我过去 24 小时的未读邮件，标记出任何有时间敏感性的。"*
- *"我的重点收件箱现在有什么？Other 里有没有看着该置顶的？"*
- *"收件箱里有物流更新吗？追踪我在等的包裹和预计到达时间。"*
- *"扫一下我的邮件里即将到期的订阅续费——未来两周有什么会自动扣款？"*
- *"我下周要去西雅图——查一下我日历上的行程，然后建一个带打包清单的待办任务。"*
- *"给我姐姐最后一封邮件草拟个回复，说这周末给她打电话。"*
- *"把本周的所有订阅类和营销类邮件移到 'Read Later' 文件夹——每次批量 20 封。"*

服务器提供 65 个独立工具，agent 可以自行组合工作流——读、整理、写、排程、追任务——没有写死的宏。

## 兼容客户端

- **[OpenClaw](https://openclaw.ai)** — 原生 MCP 支持，可从 [ClawHub](https://clawhub.ai/skills?q=outlook-mcp) 获取
- **[Claude Code](https://claude.com/claude-code)** — 加到 `~/.claude/settings.json` 的 `mcpServers`
- **[Cursor](https://cursor.com)** — MCP 兼容
- **任何 MCP 客户端** — 标准 stdio MCP 服务器

已列入[官方 MCP Registry](https://registry.modelcontextprotocol.io/v0/servers?search=mpalermiti)，注册名 `io.github.mpalermiti/outlook-mcp`。

---

## 功能

**65 个工具**，13 大类：

- **认证（1）** -- 认证状态检查（登录走 CLI）
- **邮件读取（7）** -- 收件箱列表（含重点收件箱和未分类过滤）、读邮件、按 ID 经 `$batch` 批量读、KQL 搜索、文件夹列表、收件箱增量同步、跨邮件/日历/联系人的"自上次调用以来"组合摘要
- **邮件写入（3）** -- 发送、回复/回复全部、转发
- **邮件整理（9）** -- 移动、删除（默认软删）、标记、分类、已读/未读、重归类（重点收件箱）、按发件人的重点收件箱覆盖规则增删查
- **日历读取（3）** -- 事件列表（展开循环事件）、事件详情、事件增量同步
- **日历写入（4）** -- 创建、更新、删除、RSVP（接受/拒绝/暂定）
- **联系人（7）** -- 列表、搜索、详情、创建、更新、删除、增量同步
- **待办（6）** -- 清单列表、任务增删改查/完成
- **草稿（5）** -- 列表、创建、更新、发送、删除
- **附件（5）** -- 列表、下载、带附件发送、附加到草稿、移除草稿附件
- **文件夹管理（3）** -- 创建、重命名、删除邮件文件夹
- **线程与批量（3）** -- 会话线程、复制邮件、批量整理
- **用户与管理（6）** -- whoami、日历列表、类别列表、邮件提示、账号管理

**设计原则：**

- **BYOID** -- 自带应用 ID。你自己注册 Azure AD 应用，不共享 client ID。
- **零遥测** -- 无分析、无本地缓存、无第三方调用。
- **令牌存储** -- 经 `azure-identity` 存操作系统钥匙串（macOS Keychain、Windows 凭据存储、Linux Secret Service）。
- **输入校验** -- 所有输入（邮箱、Graph ID、OData、KQL、日期时间）在发往 API 前先校验。
- **只读模式** -- 配置 `read_only: true` 阻断所有写操作。注意它限制的是*工具*而非*令牌*——见 [read_only 能做什么与不能做什么](#read_only-能做什么与不能做什么)。
- **软删除** -- 删除默认移入已删除邮件。硬删除需要显式 `permanent: true`。
- **时区感知** -- 日历操作遵循配置的 IANA 时区。
- **相对日期** -- 每个日期时间参数接受 ISO 8601 或偏移量：`7d` 是七天前，`+7d` 是七天后，`now` 是此刻。单位：`m`、`h`、`d`、`w`。
- **附件沙箱** -- 附件读写限定在 `attachments_dir` 内，所以一封让 agent 把磁盘别处的文件发出去的邮件无法被服从。
- **增量游标沙箱** -- `delta_token` 是调用方持有的状态，因此是不可信输入。每个将携带 Graph 令牌的 URL 都会被解析并要求是 `graph.microsoft.com` 上的 https——这正是阻止被投毒的游标把你的邮箱令牌重定向给别人的一环。
- **工作流提示词** -- `morning_brief`、`triage_folder`、`catch_up` 作为 MCP prompts 内置，常用流程不必逐次调用重建。

### 对 agent 友好的返回形状（1.8.0）

两项纯代码升级，让同样这批工具对 AI agent 更省 token、更可恢复：

- **精简模式** — 给五个高流量读取工具（`outlook_list_inbox`、`outlook_read_message`、`outlook_search_mail`、`outlook_list_events`、`outlook_list_thread`）传 `concise=True` 可丢弃大字段：完整邮件正文、每个事件的与会人列表、线程里引用的上文、收件箱列表的正文预览/分类。典型负载缩减约 10 倍。默认 `concise=False` 保持既有返回形状——严格向后兼容。

- **结构化 Graph 错误** — 每个工具把 msgraph SDK 异常包装成带运维友好恢复提示的 `{code, message, action}`：401 提示重新认证、403/`ErrorAccessDenied` 附仓库 [ROADMAP 死路清单](https://github.com/mpalermiti/outlook-mcp/blob/main/ROADMAP.md#investigated-and-not-viable)链接、404/`ErrorItemNotFound` 提示重新列表、429 提示退避、503 提示重试。`OutlookMCPError` 子类和校验错误原样透传。

---

## Azure AD 应用注册

你需要注册一个免费的 Azure AD 应用来获取 client ID。

### 前置条件（个人微软账户）

微软已停止无 Azure AD 租户的个人账户直接注册应用。你需要先创建免费 Azure 账户：

1. 打开 [azure.microsoft.com/free](https://azure.microsoft.com/free)，用你的个人 `@outlook.com` 账户注册。需要信用卡做身份验证但**不会扣费**。这会创建一个正式的 Azure AD 租户。

### 注册应用

1. 进入 [App Registrations](https://go.microsoft.com/fwlink/?linkid=2083908)，用 `@outlook.com` 账户登录。

2. 点 **"+ New registration"** 填写：
   - **名称：** 任意非微软品牌词（如 `mp-outlook-mcp`——"Outlook MCP" 这类名称会被拒）
   - **受支持的帐户类型：** 选 **"Personal Microsoft accounts only"**
   - **重定向 URI：** 留空

3. 点 **Register**。在概览页复制 **Application (client) ID**。

4. 进入 **Authentication (Preview)** → **Settings** 标签 → 把 **"Allow public client flows"** 切到 **Yes** → **Save**。

5. 进入 **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**（委托权限）→ 添加：
   - `Mail.ReadWrite`、`Mail.Send`
   - `Calendars.ReadWrite`
   - `Contacts.ReadWrite`、`Tasks.ReadWrite`
   - `User.Read`、`offline_access`

不需要 client secret。设备码流程使用公共客户端认证。

---

## 快速开始

### 安装

**方式 A — 从 PyPI（推荐）：**

```bash
uv tool install outlook-graph-mcp
# 或: pipx install outlook-graph-mcp
# 或: pip install outlook-graph-mcp
```

**方式 B — 从源码：**

```bash
git clone https://github.com/mpalermiti/outlook-mcp.git
cd outlook-mcp
uv sync
```

### 配置

创建 `~/.outlook-mcp/config.json`：

```json
{
  "client_id": "你的应用CLIENT_ID",
  "tenant_id": "consumers",
  "timezone": "Asia/Shanghai",
  "read_only": true,
  "attachments_dir": "~/.outlook-mcp/attachments"
}
```

唯一必填项是 `client_id`，其余都有合理默认值。建议从 `read_only: true` 起步——放心后再改 `false`。

### 注册到 MCP 客户端

**PyPI 安装：**

```json
{
  "mcpServers": {
    "outlook": {
      "command": "outlook-mcp"
    }
  }
}
```

**源码安装：**

```json
{
  "mcpServers": {
    "outlook": {
      "command": "uv",
      "args": ["--directory", "/path/to/outlook-mcp", "run", "outlook-mcp"]
    }
  }
}
```

**OpenClaw** 用户用 `openclaw mcp` CLI——它会帮你写入 `~/.openclaw/openclaw.json` 的 `mcp.servers`：

```bash
# PyPI 安装:
openclaw mcp set outlook '{"command":"outlook-mcp"}'

# 源码安装:
openclaw mcp set outlook '{"command":"uv","args":["--directory","/path/to/outlook-mcp","run","outlook-mcp"]}'

# 验证:
openclaw mcp list
openclaw mcp show outlook --json
```

注册后重启 OpenClaw 网关。SSE/HTTP 传输变体见 [OpenClaw MCP 文档](https://docs.openclaw.ai/cli/mcp)。

### 认证

在将运行 MCP 服务器的机器上执行一次：

```bash
uv run outlook-mcp auth
```

你会得到一个 URL 和一个代码。用任意浏览器打开 URL、输入代码、用你的微软账户登录。令牌会缓存进操作系统钥匙串——MCP 服务器自动取用。

其他 CLI 命令：

```bash
uv run outlook-mcp status   # 查看认证状态
uv run outlook-mcp logout   # 清除凭据
uv run outlook-mcp serve    # 启动 MCP 服务器（默认命令，OpenClaw/Claude 使用）
```

### 多账号（按能力路由）

早期版本里 `accounts` 只是配置摆设——列得出、用不上。现在它把**能力**路由到账号，一个微软身份管一摊事：邮件走收通知的账号，待办走存任务清单的账号。

```json
{
  "accounts": [
    {"name": "net",  "client_id": "<可复用同一个或各配各的应用ID>"},
    {"name": "neko", "client_id": "<应用ID>"}
  ],
  "default_account": "net",
  "capability_accounts": {"mail": "net", "calendar": "net", "todo": "neko"},
  "allow_cross_account": false
}
```

每个账号单独认证——设备码流程登录的是浏览器里当前的身份，务必选对：

```bash
uv run outlook-mcp auth net    # 浏览器里以 net 的身份登录
uv run outlook-mcp auth neko   # 再以 neko 的身份登录
uv run outlook-mcp status      # 每账号状态 + 路由表
```

每个账号有独立的令牌缓存（`outlook-mcp-<name>`）和认证记录。单账号安装（顶层 `client_id`、无 `accounts`）行为与从前完全一致。

**路由**：每个工具服务于其能力所配置的账号——邮件系分组（mail、drafts、attachments、folders、admin）并入 `mail`；`calendar`、`contacts`、`todo` 一一对应；增量工具按名字拆分；身份工具（`outlook_whoami` 等）跟随当前活跃账号。`outlook_changes_since` 跨三个能力，三者不在同一账号时拒绝执行——此时请用各能力的增量工具。

**`allow_cross_account`** 是配置路由之外一切的总闸：

- `false`（默认）：agent 眼中是**一个合并账号**。`outlook_switch_account` 拒绝（先于校验任何名字，拒绝信息不泄露任何东西）、`outlook_list_accounts` 折叠为当前身份、没有任何工具接受账号参数——不存在通往其他账号非默认内容的路径。
- `true`：`outlook_switch_account("neko")` 切换活跃账号；`outlook_switch_account("neko", capability="todo")` 重路由单个能力。配置的路由始终优先于活跃账号——切换身份不会拖着已路由的能力走。

**`allow_aggregate`** 是聚合读取工具（`outlook_list_inbox_all`、`outlook_list_events_all`、`outlook_list_tasks_all`）的独立开关：它们并发扇出到*所有*已认证账号，返回一份合并的、逐条带账号标签的列表——邮件新→旧、日程近→远、全部任务清单展平。单个账号失败落入 `errors` 而不影响其他账号；未认证账号列入 `skipped_unauthenticated`。与 `allow_cross_account` 正交：cross 管刻意切换路由，aggregate 管批量跨账号读取。默认关闭——工具拒绝时错误信息里带修复方法。

一个值得知道的注意点：会话级切换（仅限会话级）在服务器重启后失效。

---

## 故障排查

### Linux 上的 `SSL: CERTIFICATE_VERIFY_FAILED`

如果认证时报 `[SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed: unable to get local issuer certificate`，说明你的 Python 环境找不到系统 CA 证书束。这在最小化/容器 Linux 镜像和 `uv tool install` 的隔离 venv 里很常见。

让 Python 指向系统 CA 束。**两个**变量都要设——认证路径（`azure-identity` → `requests`）读 `REQUESTS_CA_BUNDLE`，而 delta/`$batch` 路径（`httpx`）读 `SSL_CERT_FILE`：

```bash
export SSL_CERT_FILE=/etc/ssl/certs/ca-certificates.crt      # httpx + Python ssl
export REQUESTS_CA_BUNDLE=/etc/ssl/certs/ca-certificates.crt  # azure-identity 认证
```

路径因发行版而异：Debian/Ubuntu 用 `/etc/ssl/certs/ca-certificates.crt`；RHEL/Fedora 用 `/etc/pki/tls/certs/ca-bundle.crt`。文件缺失就装发行版的 CA 包（`ca-certificates`）。要在 MCP 客户端拉起服务器的同一环境里设置，这样运行时也生效，而不只作用于一次性的 `auth` 命令。

### 令牌缓存未加密存储（Linux）

启动时出现一次性的"令牌缓存回退为明文"警告，说明 `libsecret`/PyGObject 不可导入——修复方法见[隐私与安全](#隐私与安全)。

---

## 工具参考

**日期。** 每个日期时间参数（`after`、`before`、`start`、`end`、`due`、`deferred_send_datetime`）接受 ISO 8601——`2026-10-22` 或 `2026-10-22T14:30:00Z`——或相对偏移：`7d` 是七天**前**，`+7d` 是七天**后**，`now` 是此刻。单位 `m`、`h`、`d`、`w`。裸值表示*之前*（遵循常见 CLI 惯例），所以未来的截止日期要带 `+`。不带时区的输入按你配置的 `timezone` 解释；响应始终为 UTC。

### 认证

| 工具 | 说明 |
|------|------|
| `outlook_auth_status` | 检查是否已认证、只读模式是否开启。 |

> **注意：** 认证通过 CLI（`outlook-mcp auth`）完成，不通过 MCP 工具。见上文[认证](#认证)一节。

### 邮件读取

| 工具 | 说明 |
|------|------|
| `outlook_list_inbox` | 列出文件夹内的邮件。`folder` 接受显示名、熟知名或 Graph ID。可按已读状态、发件人、日期范围、重点收件箱分类过滤。`skip` 分页。 |
| `outlook_list_inbox_all` | 聚合所有已认证账号的收件箱（需 `allow_aggregate`）。每封邮件带 `account` 标签；新→旧排序；单账号失败隔离进 `errors`。 |
| `outlook_read_message` | 按 ID 读完整邮件。格式：`text`、`html` 或 `full`（两者）。传 `include_deferred_send=True` 可一并返回草稿的定时发送时间。 |
| `outlook_read_messages` | 经 Graph `$batch` 单次往返批量读最多 20 封。每封的形状与相同 `(format, concise, include_deferred_send)` 下的 `outlook_read_message` 逐字节一致。容忍部分失败：个别 ID 404 进 `failures[]` 而不拖垮整次调用。请勿用 N 次 `outlook_read_message` 代替。 |
| `outlook_search_mail` | 用 KQL 查询搜索邮件。可选按文件夹名或 ID 限定范围。 |
| `outlook_list_folders` | 列出邮件文件夹及计数、`parent_id`、`child_count`。传 `recursive=true` 遍历完整文件夹树（含子文件夹）。 |
| `outlook_list_inbox_delta` | 只列出上次调用以来的收件箱变更。首次调用返回全量快照加 `delta_token`；后续（回传 token）只返回增/改/删。删除以 `{id, is_deleted: True}` 返回。游标无状态——由 agent 持久化并回放。 |
| `outlook_changes_since` | 组合邮件/日历/联系人增量的"自上次调用以来"结构化摘要。返回计数 + `urgent_flagged` 邮件 + top-5 `by_sender` + 新增/取消的事件。每个资源有独立 `delta_token`；游标过期（HTTP 410）自动重新同步该资源并在 `_meta.resync` 标注。首次快照按 `fallback_window_hours`（默认 24）过滤。为周期性 agent 循环设计。 |

### 邮件写入

| 工具 | 说明 |
|------|------|
| `outlook_send_message` | 发邮件。支持 TO/CC/BCC、HTML 正文、重要性级别。 |
| `outlook_reply` | 回复或回复全部。 |
| `outlook_forward` | 转发给一个或多个收件人，可附评论。 |

### 邮件整理

| 工具 | 说明 |
|------|------|
| `outlook_move_message` | 按名称或 ID 把邮件移到文件夹。 |
| `outlook_delete_message` | 删除邮件。默认软删（进已删除邮件）。`permanent: true` 硬删。 |
| `outlook_flag_message` | 设跟随标记：`flagged`、`complete` 或 `notFlagged`。 |
| `outlook_categorize_message` | 给邮件设分类。 |
| `outlook_mark_read` | 标记已读或未读。 |
| `outlook_reclassify_message` | 在重点收件箱和 Other 之间移动邮件（`focused` / `other`）。 |
| `outlook_list_inbox_overrides` | 列出重点收件箱按发件人的覆盖规则。 |
| `outlook_set_inbox_override` | 增改一条按发件人的覆盖（`focused` / `other`）。发件人匹配不区分大小写；存在则 PATCH 否则 POST。 |
| `outlook_delete_inbox_override` | 按 ID 删除覆盖规则。 |

### 日历读取

| 工具 | 说明 |
|------|------|
| `outlook_list_events` | 列出日期范围内的事件。展开循环事件。每个事件带 `type`，可区分日程母本与单次。经 `days`、`after`、`before` 配置。 |
| `outlook_list_events_all` | 聚合所有已认证账号的事件（需 `allow_aggregate`），近→远排序，带 `account` 标签。 |
| `outlook_get_event` | 事件完整详情：与会人、正文、在线会议 URL、循环规则、`type`（`singleInstance` / `seriesMaster` / `occurrence` / `exception`）。 |
| `outlook_list_events_delta` | 只列出窗口内上次调用以来的事件变更。首次调用必须给 `start` 和 `end`（ISO 8601）（Graph 约束——不支持整日历同步）。删除以 `{id, is_deleted: True}` 返回。游标无状态。 |

### 日历写入

| 工具 | 说明 |
|------|------|
| `outlook_create_event` | 创建带地点和与会人的事件。（`is_online` 对个人账户无效——Graph 对消费者邮箱忽略 `isOnlineMeeting`。）传 `recurrence` 创建**系列**：速记（`daily`、`weekdays`、`weekly`、`monthly`、`yearly`，锚定 `start`）或完整 [Graph 循环对象](https://learn.microsoft.com/graph/api/resources/patternedrecurrence)。`range.startDate` 默认取事件开始日期。 |
| `outlook_update_event` | 更新事件字段（主题、时间、地点、正文、与会人、全天）。只 patch 变化的字段。传 `recurrence` 把单次变系列，或 `remove_recurrence=True` 把系列变回单次。`attendees` **整体替换**宾客名单并发出邀请/取消邮件；`is_all_day` 需同调用带 `start`+`end`。 |
| `outlook_delete_event` | 删除日历事件。 |
| `outlook_rsvp` | 回应事件：`accept`、`decline` 或 `tentative`。可附言。 |

### 联系人

| 工具 | 说明 |
|------|------|
| `outlook_list_contacts` | 游标分页列出联系人。 |
| `outlook_search_contacts` | 按姓名或邮箱搜索联系人。 |
| `outlook_get_contact` | 按 ID 读完整联系人。 |
| `outlook_create_contact` | 创建联系人。 |
| `outlook_update_contact` | 更新联系人字段。 |
| `outlook_delete_contact` | 删除联系人。 |
| `outlook_list_contacts_delta` | 只列出上次调用以来的联系人变更。删除以 `{id, is_deleted: True}` 返回。游标无状态。 |

### 待办

| 工具 | 说明 |
|------|------|
| `outlook_list_task_lists` | 列出待办清单。 |
| `outlook_list_tasks` | 按状态过滤、分页列出任务。 |
| `outlook_list_tasks_all` | 聚合所有已认证账号及全部清单的任务（需 `allow_aggregate`），带 `account` 和 `list` 标签。 |
| `outlook_create_task` | 创建带截止日、重要性、循环的任务。 |
| `outlook_update_task` | 更新任务字段。 |
| `outlook_complete_task` | 标记任务完成。 |
| `outlook_delete_task` | 删除任务。 |

### 草稿

| 工具 | 说明 |
|------|------|
| `outlook_list_drafts` | 分页列出草稿。 |
| `outlook_create_draft` | 创建草稿。支持 `deferred_send_datetime` 定时发送（服务端实现，兼容 Outlook 桌面版的"延迟传递"）。 |
| `outlook_update_draft` | 更新草稿字段。接受 `is_html=True` 写 HTML 正文、`deferred_send_datetime` 设置或清除定时发送。 |
| `outlook_send_draft` | 发送已有草稿。 |
| `outlook_delete_draft` | 删除草稿。 |

### 附件

> **自 1.20.0 起这些工具只触达 `attachments_dir`**（默认 `~/.outlook-mcp/attachments`）。
> 裸文件名在其中解析；指向外部的路径一律拒绝，包括经符号链接。要发送文件请先移进去——
> 或者放宽 `attachments_dir`，但要知道从它能到达的一切都可以被发出去。1.20.0 之前这些工具
> 能读服务器进程可读的任意文件，意味着一封让 agent 附上某文件的邮件是可以被服从的。

| 工具 | 说明 |
|------|------|
| `outlook_list_attachments` | 列出邮件的附件。 |
| `outlook_download_attachment` | 下载附件并把解码后的字节存入 `attachments_dir`。 |
| `outlook_send_with_attachments` | 发送带附件的邮件，附件从 `attachments_dir` 读取（>3MB 自动上传会话）。 |
| `outlook_attach_to_draft` | 把 `attachments_dir` 里的附件加到已有草稿（>3MB 自动上传会话）。 |
| `outlook_remove_draft_attachment` | 从草稿移除单个附件。 |

### 文件夹管理

| 工具 | 说明 |
|------|------|
| `outlook_create_folder` | 创建邮件文件夹（顶层或嵌套）。 |
| `outlook_rename_folder` | 重命名邮件文件夹。 |
| `outlook_delete_folder` | 删除邮件文件夹（拒绝对熟知文件夹操作）。 |

### 线程与批量

| 工具 | 说明 |
|------|------|
| `outlook_list_thread` | 取会话线程内的全部邮件。 |
| `outlook_copy_message` | 把邮件复制到另一文件夹。 |
| `outlook_batch_triage` | 批量移动/标记/分类/已读（每次最多 20 条）。单次 Graph `/$batch` 往返——大批量整理比逐条调用快 10-20 倍。 |

### 用户与管理

| 工具 | 说明 |
|------|------|
| `outlook_whoami` | 取当前用户资料。 |
| `outlook_list_calendars` | 列出可用日历。 |
| `outlook_list_categories` | 列出类别定义及颜色。 |
| `outlook_get_mail_tips` | 发送前检查（外出答复、投递限制）。 |
| `outlook_list_accounts` | 列出已配置账号。 |
| `outlook_switch_account` | 切换活跃账号。 |

---

## 提示词

三个工作流以内置 MCP prompts 提供，常用流程不必逐次调用重建。任何支持 prompts 的 MCP 客户端都会列出它们；多数客户端里以斜杠命令或提示词选择器的形式出现。

| 提示词 | 参数 | 作用 |
|--------|------|------|
| `morning_brief` | `folder`（默认 `inbox`） | 今日事件、未读邮件、到期任务，按最省的顺序——各扫一遍、`concise=True`、批量读。 |
| `triage_folder` | `folder`（默认 `inbox`）、`count`（默认 50） | 对一个文件夹做一次省流扫描，归入回复/归档/垃圾，用一次 `outlook_batch_triage` 应用而非每封一次调用。 |
| `catch_up` | `since`（默认 `24h`） | 邮件、日历、联系人各自的变化，走增量路径——比定时重扫省约十倍。 |

调用前零成本：`prompts/list` 里每条只有一个名字和一行说明，正文按需拉取。

---

## 配置

配置位于 `~/.outlook-mcp/config.json`（以 `0600` 权限创建）。

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `client_id` | `string` | `null` | Azure AD 应用（客户端）ID。认证必需。 |
| `tenant_id` | `string` | `"consumers"` | Azure AD 租户。个人微软账户用 `"consumers"`。 |
| `timezone` | `string` | `"UTC"` | IANA 时区（如 `"Asia/Shanghai"`）。用于日历工具的相对日期计算。 |
| `read_only` | `bool` | `false` | `true` 时所有写工具（发送、回复、移动、删除、创建、更新、RSVP）报错。管的是工具不是微软令牌——见下文。 |
| `attachments_dir` | `string` | `"~/.outlook-mcp/attachments"` | 附件工具唯一可读写的目录。agent 提供的每个路径都会被解析并要求落在其中——符号链接出界或 `..` 一律拒绝。放宽前要想清楚：从它能到达的一切都可以被邮件发出。 |
| `allow_categories` | `list[string]` | `[]` | 可选。把写工具限制到特定类别（见下文）。空列表 = `read_only: false` 时允许全部写操作。 |
| `allow_unencrypted_token_cache` | `bool` | `false` | 允许在平台没有加密存储（无 libsecret 的 Linux）时把 OAuth 令牌缓存写为明文。默认关闭：认证会停下并说明原因，而不是默默把可复用的 Graph 令牌以明文落盘。macOS 和 Windows 始终加密，不受影响。 |

### 工具组选择（可选）— `OUTLOOK_MCP_TOOLSETS`

全部 65 个工具 schema 每回合都加载进客户端上下文（约 8.6k token）。只需要部分能力的客户端可设 `OUTLOOK_MCP_TOOLSETS` 环境变量为逗号分隔的工具组列表，只加载这些。`account` 组（认证/身份）始终可用。

```bash
# 例如一个周期性邮件+日历 agent：约 30 个工具而非 62（每回合省约 52% 工具 token）
OUTLOOK_MCP_TOOLSETS="mail,calendar,digest,delta"
```

分组：`mail`、`drafts`、`attachments`、`calendar`、`contacts`、`todo`、`folders`、`digest`、`delta`、`admin`。不设（默认）全量加载——完全向后兼容。只影响广播哪些工具；启用的工具行为不变。

### read_only 能做什么与不能做什么

`read_only: true` 阻止 outlook-mcp 的写工具运行。让它发邮件，它拒绝。

**它不会让你的微软凭据变成只读。** outlook-mcp 登录时请求 `.default` scope——"这个 Azure 应用被授权的一切"。如果你给应用授权了 `Mail.ReadWrite` 和 `Mail.Send`（上面的安装步骤就是这么让你做的），那么无论 `read_only` 开不开，存下来的令牌都能发邮件。

两个值得理解的推论：

- `read_only` 是文本文件里的一行。任何能编辑 `~/.outlook-mcp/config.json` 的东西都能把它关掉并立刻获得写权限——无需重新认证，没有新的授权提示。
- 执行逻辑在本服务器的 Python 代码里。持有缓存令牌的任何其他进程不受它约束。

所以把 `read_only` 当作防止 agent 鲁莽行事的护栏，**而不是安全边界**。想要真正写不了的凭据，注册第二个只授权读 scope（`Mail.Read`、`Calendars.Read`、`Contacts.Read`、`Tasks.Read`、`User.Read`）的 Azure 应用，把 `client_id` 指向它——让微软来执行，而不是我们。

### 细粒度写权限（可选）

默认 `read_only: false` 解锁**全部**写工具。要更细的控制，用 `allow_categories` 把写权限限定到特定类别。读工具（列表、搜索、详情）始终允许——`allow_categories` 只收窄写面。

**可用类别：**

| 类别 | 工具 | 风险 |
|---|---|---|
| `mail_drafts` | 草稿增删改 | 安全——只有草稿，不发送 |
| `mail_triage` | 移动、删除（软删）、标记、分类、已读、复制、批量 | 中等——除硬删外可逆 |
| `mail_folders` | 文件夹增删改名 | 中等 |
| `mail_send` | 发送、回复、转发、发草稿、带附件发送 | **危险**——以你的名义发邮件 |
| `calendar_write` | 事件增删改、RSVP | 中等——创建日历条目 |
| `contacts_write` | 联系人增删改 | 中等 |
| `todo_write` | 任务增删改完成 | 安全——你自己的任务清单 |

**示例策略：**

**只写草稿的助理**（agent 起草，你审阅发送）：

```json
{ "read_only": false, "allow_categories": ["mail_drafts", "mail_triage", "todo_write"] }
```

**只管日历**（agent 只能管理日程）：

```json
{ "read_only": false, "allow_categories": ["calendar_write"] }
```

**完全写权限**（agent 无限制）：

```json
{ "read_only": false }
```

**只读**（最安全的默认，无任何写操作）：

```json
{ "read_only": true }
```

设置了 `allow_categories` 时，未允许类别的工具返回权限拒绝错误（`PermissionDeniedError`），并指明被拦的类别。`allow_categories` 为空（或未设）且 `read_only` 为 false 时全部写工具可用。`read_only: true` 始终优先——设置后无论 `allow_categories` 如何全部写操作被拦。未知类别名在配置加载时即报校验错误；只接受上表七个名称。

---

## 隐私与安全

- **零遥测。** 无分析、无追踪、不收集使用数据。
- **零本地缓存。** 每次调用直连 Microsoft Graph。本地不存邮件/日历。
- **零第三方调用。** 服务器只与 `graph.microsoft.com` 和 `login.microsoftonline.com` 通信。
- **令牌存储。** OAuth 令牌经 `azure-identity` 的 `TokenCachePersistenceOptions` 持久化。macOS 用系统 Keychain；Windows 用 DPAPI；Linux 有 PyGObject/libsecret 时用 gnome-keyring。Linux *无* libsecret 时（如 `uv tool install` 建的隔离 venv），令牌回退为 `~/.IdentityService/` 下 `0600` 的明文文件，MCP 启动时打一次警告。要在 Linux 上加密存储，安装 `python3-gi gnome-keyring libsecret-1-0` 并用 `--system-site-packages` 重建 venv。
- **不记录敏感数据。** 邮件正文、收件人地址、令牌绝不写日志。
- **配置权限。** 配置目录 `0700`、配置文件 `0600`。拒绝符号链接的配置。
- **输入校验。** 所有用户输入（邮箱地址、Graph ID、OData 过滤、KQL 查询、日期时间）在到达 Graph API 前校验和消毒。

---

## 开发

```bash
# 安装开发依赖
uv sync --extra dev

# 跑测试
uv run pytest

# Lint
uv run ruff check src/ tests/

# 格式化
uv run ruff format src/ tests/

# 本地运行服务器（stdio）
uv run outlook-mcp
```

**环境要求：** Python 3.10+

---

## 路线图

- **收件箱规则** -- 规则的列出、创建、删除
- **高级邮件** -- 原始 MIME 导出、互联网邮件头
- **日历** -- 取消事件（含通知与会人）
- **清单** -- 待办任务的清单项
- **企业版（Entra ID）** -- 工作/学校账户支持

---

## 许可证

MIT。见 [LICENSE](LICENSE)。
