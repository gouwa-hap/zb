# 招投标统计工作台

> 一个**纯前端单页 + 本地 Node 服务**的数据工作台，用于统计招投标项目清单、以工代赈项目调度、开标日历。
> 作者署名：招投标统计 —— design:Tye

## 包含什么

| 模块 | 说明 |
|---|---|
| 📋 招投标项目清单 | 建设工程 / 政府采购两张表，支持按年/按月统计、Excel 导入、双击编辑、开标日历 |
| 🏗️ 以工代赈项目调度 | 20 字段项目调度表，支持 Excel 导入/编辑 |
| 📅 开标日历 | 嵌在清单页内，月视图，点击日期看当天开标安排 |

- 数据持久化在服务端 `server-data/store.json`（刷新不丢）。
- 全局修改密码：**123456**（点顶部「🔧 修改」按钮后输入，才能增删改/导入）。
- 解析链接：仅支持**海南省公共资源交易中心** `https://ggzy.hainan.gov.cn/ggzyjy/` 下的公告页。

## 如何部署（WorkBuddy，推荐给非技术人员）

1. 把这个文件夹整体交给你的 WorkBuddy 助手。
2. 对助手说一句话：
   > 帮我发布这个应用：目录指向「招投标统计-git」，应用名「招投标统计」，语言 Node.js，端口 8080，启动命令 `node server.js`
3. 助手会返回一个线上链接，点开即用。**数据已经内置在 `server-data/store.json` 里，无需任何导入。**

## 如何备份数据

只需复制 `server-data/store.json` 这一个文件即可（和代码分开存更安全）。

## 目录结构

```
招投标统计-git/
├─ index.html              首页
├─ server.js               服务程序（零依赖，系统自带 Node 即可）
├─ package.json            启动配置（发布工具必需）
├─ .gitignore
├─ README.md
├─ assets/                 样式 + 脚本 + xlsx 依赖
│  ├─ app.js
│  ├─ style.css
│  └─ vendor/xlsx.full.min.js
├─ plugins/                三个功能模块
│  ├─ tender-list.js       招投标项目清单
│  ├─ yigong-daizhen.js   以工代赈项目调度
│  └─ tender-calendar.js  开标日历
└─ server-data/
   └─ store.json          全部数据（招投标 + 以工代赈）
```

## 本地直接运行（可选）

```bash
node server.js
# 浏览器打开 http://localhost:8080
```
