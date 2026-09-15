# html2pptx

[English](README.md) | **简体中文**

将基于约定的 HTML 幻灯片转换为可编辑的 PowerPoint（`.pptx`）演示文稿——无需浏览器、无需渲染引擎，只需 Python。

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)

## 特性

- 🎞️ 16:9 幻灯片（13.333 × 7.5 英寸），可在 PowerPoint / WPS 中完全编辑
- 🧩 自动布局：卡片、统计块、对比栏、流程图、编号方案、金句引用、高亮面板、封面与结尾页
- ✍️ 保留 `<b>/<strong>/<em>` 加粗与 `<br>` 换行
- 📝 自动页码、页脚与每页演讲者备注
- 😀 支持 emoji
- 🌏 自动识别 UTF-8 / GB18030 编码

## 环境要求

- Python 3.8+
- [python-pptx](https://python-pptx.readthedocs.io/) — `pip install python-pptx`
- 建议安装微软雅黑字体以获得最佳中文渲染效果（Windows 自带；如需更换可在脚本中修改 `FONT` 常量）

## 用法

```bash
python html2ppt.py 你的演示.html
python html2ppt.py 你的演示.html -o 输出文件.pptx
```

输出默认与输入文件同名，扩展名为 `.pptx`。

## 工作原理

转换器**基于约定**，而非通用的 HTML 渲染器：它解析文档树，按各元素的 `class` 属性映射到对应的版式，再用内置的浅色 / 深色设计体系进行绘制。

| HTML | 渲染结果 |
|---|---|
| `<section class="slide">` | 一页幻灯片 |
| `class="cover"` / `class="final"` | 深色封面 / 结尾页 |
| `.kicker` | 青色小字眉标 |
| `h1` / `h2` | 页面标题（`<br>` 自动分行） |
| `.lead` | 引导段落 |
| `.sub` | 灰色脚注 |
| `.badge` / `.foot` | 封面徽标 / 封面页脚 |
| `.quote` + `.who` | 金句引用（含署名） |
| `.mirror` | 通栏高亮面板 |
| `.cards` + `.card`（`.ic`、`h3`、`p`） | 卡片网格（加 `c2` / `c3` 控制 2 / 3 列） |
| `.bullets` + `.b`（+ `<small>`） | 要点列表；`<small>` 显示为灰色小字 |
| `.stat` + `.box`（`.num`、`.lbl`） | 大数字统计块（`.num up` / `.num down` → 绿色 / 粉色） |
| `.vs` + `.side` | 双栏对比（`side a` ✔ 青色 / `side b` ✖ 粉色） |
| `.chain` + `.link`（`.tag`、`.tt`） | 带 ▼ 箭头的流程图行 |
| `.num3` + `.row`（`.n`、`h4`、`p`） | 编号方案行 |

### 最小示例

```html
<section class="slide">
  <div class="kicker">01 · 引言</div>
  <h2>为什么重要</h2>
  <p class="lead">一段简短的引导语。</p>

  <div class="cards c2">
    <div class="card">
      <div class="ic">💡</div>
      <h3>想法</h3>
      <p>卡片说明写在这里。</p>
    </div>
    <div class="card">
      <div class="ic">🚀</div>
      <h3>行动</h3>
      <p>另一张卡片的说明。</p>
    </div>
  </div>

  <p class="sub">脚注文字。</p>
</section>
```

## 输出

- 16:9（13.333 × 7.5 英寸）
- 微软雅黑字体
- 每个内容页页脚含演示名称及 `NN / TT` 页码
- 每页自动附加一条演讲者备注

## 限制

- 不是通用的 HTML → PPTX 渲染器：颜色、字体与版式来自内置模板，而非你的 CSS。
- 不支持图片、图表、表格与动画。
- 无法识别的元素会被跳过，并在 stderr 输出 `[warn]` 提示。
- 文本提取为启发式处理——嵌套过深的标记可能需要少量手工调整。

## 许可证

MIT