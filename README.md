# Personal Academic Website

这个版本保留旧工程的主页和 publication index 风格，不改变 `index.html` / `pub_index.html` 的渲染方式。

## 日常更新方式

1. 打开 `setting.html`。
2. 点击 `Import JSON`，导入当前的 `data/siteData.json`。
3. 在不同标签里编辑：
   - `Basic`: 个人信息。
   - `Biography`: 个人经历。
   - `Research`: research 文字，以及预留的 project 数据。
   - `News`: 新闻。
   - `Publications`: 论文、书籍章节、专利，按 Journal / Conference / Patent / Book Chapter 分组。
   - `Honors`: 奖项。
   - `Display`: 原有显示设置。
4. 默认是展示模式；每条右上角点 `Edit` 才进入编辑，点 `Confirm` 后回到展示模式。
5. 条目右上角可以用 `↑` / `↓` 调整顺序，也可以拖动条目排序。
6. 可以点 `Save Draft` 保存到当前浏览器本地草稿。
7. 确认后点 `Export siteData.json`。
8. 用导出的文件替换 `data/siteData.json`。
9. 如果新增了图片/GIF，把文件放到相应的 `images/` 子目录，并在 JSON 或 setting 页面里填写相对路径。
10. 提交并 push 更新后的 `data/siteData.json` 和新增图片。

## IEEE Citation 快速新增

在 `setting.html` 的 `Publications` 标签里，点击 `Add Pub`，再选择 `Use IEEE Citation`。把 IEEE Xplore 复制的 citation 粘贴进去，点击 `Parse and Add Publication`。

它会尝试解析：

- authors
- title
- venue
- volume / issue
- pages
- year / month
- DOI
- journal 或 conference 类型

解析后请检查并手动修正不准确的字段。

## 文件入口

- `index.html`: 公开主页，保持旧版渲染。
- `pub_index.html`: 内部 publication 快速筛选页，保持旧版风格。
- `setting.html`: 本地编辑和导出 JSON 的设置页。
- `data/siteData.json`: 网站内容数据。
