# 报告目录

存放每次调研的 LaTeX 阅读报告。目录名固定为 `<批次>-<序号>`（两位补零），
与元数据表 `metadata.csv` 的前两列一一对应，归档脚本据此自动配对。

```
reports/
├── metadata.csv            # 元数据表（从 tools/metadata.example.csv 复制而来）
├── 01-1/                   # 第 1 次调研 · 第 1 篇
│   ├── document.tex        # 报告源文件（从 latex/document.tex 复制而来）
│   ├── document.pdf        # 阅读报告（xelatex 编译产物）
│   ├── 论文原文.pdf         # 从出版商/数据库下载的原文
│   ├── 阅读报告.pdf         # document.pdf 重命名而来，便于归档时识别
│   └── pic/                # 报告用图
└── 01-2/
```

约定：

- `document.tex` 必须保持 `latex/document.tex` 的 8 个 `\section` 标题不变。
- `阅读报告.pdf` 与 `论文原文.pdf` 的名字不要改，脚本按这两个名字取文件
  （想改就用 `--pdf-original` / `--pdf-report` 传新模板）。
- PDF 不进 Git（`.gitignore` 已忽略，体积大且多含版权内容）；
