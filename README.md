# 成绩可视化助手（score-trend）

一个纯前端网页工具：导入 Excel/CSV 成绩表，按班级筛选学生，查看**总分 / 各科分数 / 班级排名 / 年级排名**的时间趋势折线图。
适合老师/班主任快速做“学生成绩波动”可视化分析。

> 在线地址：`https://charleyz2021.github.io/score-trend/`

---

## 功能

- ✅ 多文件 / 多 Sheet 导入（Excel .xls/.xlsx + CSV）
- ✅ 自动识别表头（姓名、班级、总分、各科、班级排名、年级排名）
- ✅ 班级筛选 → 学生选择 → 指标选择（总分/各科/班排/年排）
- ✅ 图表缩放（底部缩放条可拖动放大/缩小时间范围）
- ✅ 同班重名检测提示（建议补充学生ID/学号以区分）
- ✅ 支持“说明”Sheet：上传时自动忽略（模板里可带使用说明）

---

## 使用方法

1. 点击「选择文件」导入成绩表（可一次选多个）
2. 选择班级 → 选择学生 → 选择指标
3. 通过图表底部缩放条，放大/缩小查看的考试范围

---

## 成绩模板（推荐）

点击页面「下载模板」获取 `template.xlsx`。

### 模板要求

- 一个 Sheet = 一次考试（Sheet 名建议为考试名，避免重复）
- 表头建议为单行（第 1 行）
- **必填列：姓名、总分**
- **推荐列：学生ID/学号、班级**
- 可选列：班级排名、年级排名、语文/数学/英语/物理/化学/生物/政治/历史/地理等科目分数

### 关于重名

当同班出现同名时，系统会提示“需检查”。
**推荐做法：**

- 在表中补充「学生ID/学号」列；或
- 在姓名中添加区分标识（如“张三-1 / 张三-2”）

---

## 隐私说明

- 本工具为纯前端：文件在浏览器本地解析，不会上传到服务器。
- 建议在可信设备使用，避免在公共设备导入敏感数据。

---

## 本地开发

```bash
npm install
npm run dev
```

---

## 构建与部署（GitHub Pages）

项目已配置 GitHub Actions 自动部署（push 到默认分支后自动发布 Pages）。

如遇到 Pages 404：

* 确认 `vite.config.ts` 中 `base` 设置为 `/score-trend/`
* 确认模板文件放在 `public/template.xlsx`，页面下载链接为 `/score-trend/template.xlsx`（或用 `import.meta.env.BASE_URL` 拼路径）

---

## 许可（License）

MIT License

This project is licensed under the MIT License.
You are free to use, modify, and distribute it,
as long as the original license and copyright notice are included.

本项目使用 MIT 开源协议。
你可以自由使用、修改和分发，但需保留原始版权声明和协议文本。

---

## 免责声明

本工具仅用于教学辅助数据可视化。导入数据的准确性取决于表格内容与表头命名规则。

