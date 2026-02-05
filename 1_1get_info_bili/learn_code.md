# B站合集简介批量导出程序原理笔记

## 0. 补充接口作用和原理理解？

### 1. 你是怎么拿到“合集里所有集的 BV 号”的？

核心逻辑只有一句话：

> 先用**任意一集**调用“视频详情接口”，从返回里拿到 **mid（UP主）** 和 **season_id（合集ID）**；
>  再用 **mid + season_id** 调“合集列表接口”，它会直接返回**合集内每一集的 bvid 列表**（分页）。

也就是说：**BV 号不是你从 URL 推出来的，而是接口直接返回的。**

------

### 2. 程序里用到的所有接口（作用 + 获得的信息）

#### 接口 1：视频详情接口（用一集定位合集 + 拿这一集的标题/简介）

##### URL

```
GET https://api.bilibili.com/x/web-interface/view?bvid=<BV号>
```

##### 你用它做了两件事

##### A) 用“任意一集”定位整个合集

从返回 JSON 的 `data` 里取：

- `data.owner.mid`  → **UP 主 mid**
- `data.ugc_season.id` → **合集 season_id**（真正的“合集标识”）

> 这一步回答了：为什么给一集就能知道合集？
>  因为这集的详情数据里带着它属于哪个 `ugc_season`。

##### B) 每一集导出时，拿“这一集自己的信息”

同一个接口还给你：

- `data.title` → **视频标题**（用于文件命名 `【P1】标题.txt`）
- `data.desc` → **简介（常见字段）**
- `data.desc_v2` → **简介分段（兼容用）**

你导出 txt 时，写的就是 `desc / desc_v2` 合成后的文本。

------

#### 接口 2：合集列表接口（真正返回“所有集的 BV 号”）

这是你“拿到其他 136 集 BV 号”的关键接口。

##### URL（你优先用的）

```
GET https://api.bilibili.com/x/polymer/web-space/seasons_archives_list?mid=<mid>&season_id=<season_id>&page_num=<n>&page_size=<k>
```

##### 它的作用

- **列出这个合集里所有视频（分分页）**
- 每个 item 都带 `bvid`（也常带 `title` 等）

##### 你从它拿到的关键信息

在返回 JSON 的 `data` 里通常有：

###### 1) 合集视频列表（最关键）

- `data.archives`（或有时是 `data.items`） → **列表**
  - 列表里每一项是一个视频条目（item）
  - 你要的 BV 就在每个 item 的：
    - `item.bvid` → **该集的 BV 号**

✅ 你就是这样得到 “合集内所有集 BV 号”的：

- 第 1 页返回 30 个 `bvid`
- 第 2 页返回 30 个 `bvid`
- ...
- 直到取完

###### 2) 分页信息（用于知道有没有下一页）

- `data.page.total` → **总共有多少个视频**
- `page_num / page_size` → 你请求参数里给的
- 你程序的停止条件是：
  - 如果 `len(all_items) >= total` 就结束
     或
  - 如果当前页返回数量 `< page_size` 也结束（最后一页）

------

#### 接口 2’（兜底）：另一个同功能合集列表接口

你代码里写了两个 endpoint，第二个是“兜底”。

##### URL

```
GET https://api.bilibili.com/x/polymer/space/seasons_archives_list?mid=<mid>&season_id=<season_id>&page_num=<n>&page_size=<k>
```

##### 作用

- 和上面接口 2 **同作用**：返回合集的分页视频列表（同样包含 `bvid`）
- 在某些地区/时期/接口版本变动时，接口 2 可能不可用，这个作为备用。

------

### 3. 你拿 BV 的完整链路（从 1 个 BV 到 137 个 BV）

按程序的真实数据流：

1. 用户输入一个合集内的 BV（或 URL）
2. 调接口 1（view）得到：
   - `mid`
   - `season_id`
3. 调接口 2（seasons_archives_list）分页得到：
   - `archives/items[*].bvid` → 全部 BV 列表
4. 对每个 BV 再调接口 1（view）得到：
   - `title`
   - `desc/desc_v2`
5. 输出 `【P{i}】title.txt` 文件

------

### 4. 你可以直接对照理解的“字段映射表”

#### 你输入 1 集时（接口 1 返回）

| 你想要的东西         | 返回 JSON 路径               |
| -------------------- | ---------------------------- |
| UP 主 mid            | `data.owner.mid`             |
| 合集 ID（season_id） | `data.ugc_season.id`         |
| 本集标题             | `data.title`                 |
| 本集简介             | `data.desc` / `data.desc_v2` |

#### 你拿到所有集 BV 时（接口 2 返回）

| 你想要的东西  | 返回 JSON 路径                                     |
| ------------- | -------------------------------------------------- |
| 当前页所有 BV | `data.archives[*].bvid`（或 `data.items[*].bvid`） |
| 合集总视频数  | `data.page.total`                                  |

## 1. 为什么“给合集里任意一集”就能找到整个合集？

### 1.1 合集在 B 站的真实结构

- B站“合集”（UGC Season）不是靠 URL 参数来标识的（`spm_id_from` 这种只是埋点）。
- **合集有一个真正的唯一 ID**：`season_id`（接口字段名通常是 `ugc_season.id`）。
- 合集属于某个 UP 主，因此还需要 UP 主的 `mid`。

✅ 结论：

> 合集 =（UP 主 mid）+（合集 season_id）
>  只要能从任意一集拿到这两个值，就可以把整个合集的视频列表拉出来。

------

### 1.2 为什么从“一集”里能拿到 `season_id`？

因为合集里的每一集视频，在 B站后台都有“归属关系”：

- 每个视频（BV）都有一个“视频详情接口”：
   `https://api.bilibili.com/x/web-interface/view?bvid=BVxxxx`

这个接口返回的信息不仅包括：

- 视频标题 `title`
- 视频简介 `desc` / `desc_v2`

还包括（如果该视频属于某个合集）：

- `ugc_season`：合集信息（包含 `id`、`title`、`intro` 等）
- `owner.mid`：UP 主 mid

所以：

1. 你输入合集里任意一集的 BV
2. 调 `view` 接口
3. 直接得到 `owner.mid` + `ugc_season.id`
4. 用它们就能获取合集完整列表

------

## 2. 程序整体执行过程（从输入到输出）

> 下面按“GUI点按钮后”程序发生的事情，从头到尾走一遍。

------

### Step 0：用户在 GUI 里输入

你输入：

- “合集任意一集视频链接” 或 “BV号”
- 选择输出目录
- 点击「开始导出」

GUI 会启动一个后台线程（避免界面卡死）。

------

### Step 1：提取 BV 号 `extract_bvid()`

程序会从你输入的 URL 中用正则找到 `BVxxxx`：

- 如果你直接输入 BV：直接返回
- 如果你输入 URL：用正则匹配 `BV...`

目的：把一切输入统一成 `bvid`

------

### Step 2：创建 requests Session `build_session()`

程序用 `requests.Session()` 做网络请求，并设置浏览器请求头：

- `User-Agent`：伪装成浏览器
- `Referer` / `Origin`：让请求更像来自 B站网页
- 如果你填了 Cookie：会附带到请求头里

目的：

- 提高成功率
- 降低被风控概率
- 支持需要登录的内容

------

### Step 3：通过这一集确定合集 `get_collection_meta()`

程序调用：

- `get_view(session, bvid)` → 拉视频详情 JSON
- 从返回里读取：
  - `mid = data["owner"]["mid"]`
  - `season_id = data["ugc_season"]["id"]`

这一步就是你问的核心：

> “为什么一集就能定位整个合集？”

因为 B站把“这集属于哪个合集”写在这个视频的详情数据里。

------

### Step 4：获取合集全部视频列表 `list_collection_items()`

有了：

- `mid`
- `season_id`

就可以调用合集列表接口：

- `x/polymer/web-space/seasons_archives_list`

分页拉取：

- `page_num=1,2,3...`
- `page_size=30`

每一页返回：

- 一堆 item（包含 `bvid`、`title` 等）

程序持续翻页直到：

- 达到 `total`
   或
- 本页数量 < page_size（表示最后一页）

最后得到 “合集内所有视频 item 列表”。

------

### Step 5：按合集顺序去重并遍历导出

程序按拉到的顺序遍历每一个视频：

1. 取到该集 `bvid`
2. 调 `get_view()` 再拿一次详情
3. 从详情里提取：
   - `title`（标题）
   - `desc_text`（简介全文）

------

### Step 6：每集生成一个 txt 文件（哪怕简介为空）

对每个视频，生成文件名：

- `【P{idx}】{title}.txt`

并做三层安全处理：

1. `safe_filename()`：处理非法文件名字符
2. 路径太长时再截断标题
3. `ensure_unique_path()`：同名就加 `(1)(2)` 防止覆盖

写入内容：

- 只写简介文本
- 简介为空 → 写空字符串 → 空文件也会创建 ✅

------

### Step 7：进度条与日志更新（GUI不卡）

因为网络请求在后台线程跑：

- 主线程只负责 UI 更新
- 后台线程把日志/进度通过 `ui_queue` 发给主线程

主线程用 `after()` 定期取队列消息，更新：

- 日志窗口
- 进度条
- 完成/失败弹窗

------

## 3. 程序中涉及的全部重要方法解释（按模块）

------

## 3.1 输入解析与文件命名

### `extract_bvid(url_or_bvid)`

**作用**：从输入（URL 或 BV）提取 BV号
 **核心**：正则匹配 `BV[a-zA-Z0-9]+`

------

### `safe_filename(name, max_len=180)`

**作用**：把视频标题变成 Windows 可用的文件名
 做了什么：

- 去掉控制字符（不可见字符）
- 替换 `< > : " / \ | ? *`
- 压缩空格
- 去掉结尾空格/点
- 限制长度避免路径过长

------

### `ensure_unique_path(path)`

**作用**：防止同名文件覆盖
 如果 `【P1】xxx.txt` 已存在，会生成：

- `【P1】xxx (1).txt`
- `【P1】xxx (2).txt`

------

## 3.2 网络请求与数据获取

### `build_session(cookie=None)`

**作用**：创建一个可复用的 HTTP 会话
 优点：

- 保持连接（比每次 requests.get 更快）
- 统一 headers
- 可带 cookie

------

### `get_view(session, bvid)`

**作用**：获取单个视频的详情 JSON
 接口：

- `x/web-interface/view?bvid=...`

返回里关键字段：

- `title`
- `desc`
- `desc_v2`
- `owner.mid`
- `ugc_season.id`

------

### `get_desc_text(view_data)`

**作用**：稳健拿到简介全文
 逻辑：

1. 先取 `desc`（最常见）
2. 如果空，再拼 `desc_v2` 里每段的 `raw_text/text`

为什么要这样？

- 有些视频 `desc` 可能为空，但 `desc_v2` 有内容（少见但存在）
- 兼容性更强

------

### `get_collection_meta(session, any_bvid)`

**作用**：从合集任意一集获取合集定位信息
 输出：

- `mid`
- `season_id`

------

### `list_collection_items(session, mid, season_id)`

**作用**：分页拉取合集内所有视频
 接口（优先 + 兜底）：

- `x/polymer/web-space/seasons_archives_list`
- `x/polymer/space/seasons_archives_list`

返回里通常含：

- `archives`（视频列表）
- `page.total`（总数）

------

## 3.3 导出主逻辑（核心流水线）

### `export_descriptions_to_txt(...)`

这是主执行函数，做了：

1. 创建输出目录
2. 解析 BV
3. get_collection_meta 得到 mid/season_id
4. list_collection_items 拉取所有 bvid
5. 遍历每个 bvid：
   - get_view → title/desc
   - safe_filename → 文件名安全
   - ensure_unique_path → 防覆盖
   - 写 txt

同时：

- 每一集更新一次进度
- 每一集写一条日志

------

## 3.4 GUI 与多线程机制（保证不卡）

### 为什么要多线程？

如果你在主线程里做 137 个网络请求：

- GUI 会“未响应”
- 窗口卡住，体验很差

所以：

- 网络任务放到后台线程
- 主线程只做 UI

------

### `ui_queue`（线程通信队列）

后台线程不能直接改 Tkinter 控件（会崩或不稳定），所以用队列发消息：

- `("log", "xxx")`
- `("progress", cur, total)`
- `("done", out_dir)`
- `("error", msg)`

主线程定时取队列消息并更新 UI。

------

### `after(80, self._poll_queue)`

Tkinter 的定时器机制：

- 每 80ms 执行一次 `_poll_queue`
- 实现“不断刷新日志/进度条”

------

### `stop_event`

点击“停止”时：

- `stop_event.set()`
- 后台线程每处理完一集就检查一次
- 如果发现 set，就安全退出（不会写到一半中断文件）

------

## 4. 你可以记住的“最核心一句话”

> **B站合集不是靠 URL 规律，而是靠视频详情里返回的 `ugc_season.id`。**
>  只要拿到任意一集 BV → `view` 接口 → 得到 `mid + season_id` → 就能列出合集全部视频 → 再逐个拿 `desc` 导出。