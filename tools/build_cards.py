#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调研汇报论文库 · 归档脚本
=========================

把「LaTeX 阅读报告 + 元数据表」批量转成「Obsidian 论文卡片 + 分类表 + 总索引」。

    元数据表(CSV)  ─┐
                    ├─►  论文卡片.md  +  论文研究问题种类.md  +  调研汇报论文库.md
    LaTeX 报告     ─┘

只依赖 Python 3 标准库，不需要 pandas / 第三方包。

用法
----
    # 全量归档：读书报告 + 复制 PDF + 生成卡片/分类表/总索引
    python3 tools/build_cards.py \
        --meta reports/metadata.csv \
        --src  reports \
        --vault /path/to/Obsidian仓库

    # 只重算分类表与总索引（改了标签、不碰 LaTeX，秒级完成）
    python3 tools/build_cards.py --meta reports/metadata.csv --vault /path/to/Obsidian仓库 --classify-only

    # 用示例数据试跑（仓库自带）
    python3 tools/build_cards.py --meta tools/metadata.example.csv --vault ./_build \
        --cards-dir 论文卡片 --index-dir . --classify-only

参数说明见 docs/06-自动化与脚本使用.md。
"""

import argparse
import csv
import io
import os
import re
import shutil
import sys

# ---------------------------------------------------------------------------
# 1. 卡片小节 ← LaTeX 节标题 的映射
#    改报告节标题时，改这里就够了。值为列表表示「多个 LaTeX 节合并成一个卡片小节」。
# ---------------------------------------------------------------------------
CARD_SECTIONS = [
    ("研究问题", ["论文待研究科学问题"]),
    ("研究动机", ["作者研究动机"]),
    ("方法要点", ["论文工作的主要内容"]),
    ("实验结果与指标", ["实验结果", "实验数据与评估指标"]),
    ("自评（优点、缺点、可改进方案）", ["自己对该论文的评价（优点、缺点、可改进方案）"]),
]

CITATION_SECTION = "阅读文献来源"   # 卡片「文献来源」取该节第一行

# ---------------------------------------------------------------------------
# 2. 词表兜底（tools/taxonomy.csv 不存在时使用；存在则以 csv 为准）
# ---------------------------------------------------------------------------
FALLBACK_VOCAB = {
    "种类": ["疾病诊断与分型", "疾病风险与转化预测", "医学图像分割",
             "影像配准与纵向分析", "脑网络与影像表征分析", "多任务与基础模型"],
    "挑战": ["多模态融合", "高阶关系建模", "时序与动态建模", "半监督学习",
             "少标注与小样本", "缺失模态与缺失视图", "可解释性与证据推理",
             "知识注入与LLM驱动", "不确定性建模", "生成式模型与未来预测",
             "通用基础模型与上下文学习", "难样本与类别不平衡", "多任务协同",
             "结构-功能耦合"],
    "技术": ["图神经网络与超图建模", "大语言模型与知识注入", "时空动态建模",
             "提示学习与上下文学习", "对比学习", "一致性学习与伪标签",
             "原型与度量学习", "混合专家MoE", "自监督与预训练",
             "生成模型-扩散与GAN", "不确定性建模", "记忆与检索增强",
             "知识蒸馏", "强化学习", "脉冲神经网络"],
}


# ===========================================================================
# 元数据
# ===========================================================================
META_COLUMNS = ["批次", "序号", "短标题", "研究内容", "论文题目", "期刊会议",
                "年份", "文献类型", "分区", "影响因子", "报告时间",
                "种类", "挑战", "技术"]


def read_meta(path):
    """读元数据 CSV（utf-8-sig，Excel 直接另存即可）。"""
    rows = []
    with io.open(path, encoding="utf-8-sig", newline="") as f:
        for raw in csv.DictReader(f):
            row = {(k or "").strip(): (v or "").strip() for k, v in raw.items()}
            if not row.get("批次") or not row.get("序号"):
                continue
            row["批次"] = int(row["批次"])
            row["序号"] = int(row["序号"])
            if not row.get("文献类型"):
                row["文献类型"] = "会议" if "CCF" in row.get("分区", "").upper() else "期刊"
            for axis in ("种类", "挑战", "技术"):
                row[axis] = [t.strip() for t in re.split(r"[;；|、]", row.get(axis, "")) if t.strip()]
            rows.append(row)
    rows.sort(key=lambda r: (r["批次"], r["序号"]))
    return rows


def read_taxonomy(path):
    """读词表 CSV：列名 axis,tag（axis 取 种类/挑战/技术）。返回 {axis: [tags 排序]}。"""
    vocab = {axis: [] for axis in ("种类", "挑战", "技术")}
    if not path or not os.path.exists(path):
        print("[warn] 未找到词表 %s，改用内置兜底词表（顺序可能不理想）" % path)
        return {k: list(v) for k, v in FALLBACK_VOCAB.items()}
    with io.open(path, encoding="utf-8-sig", newline="") as f:
        for raw in csv.DictReader(f):
            row = {(k or "").strip(): (v or "").strip() for k, v in raw.items()}
            axis = row.get("axis") or row.get("轴") or ""
            tag = row.get("tag") or row.get("标签") or ""
            if axis in vocab and tag:
                vocab[axis].append(tag)
    return vocab


def check_tags(rows, vocab):
    """标签必须在词表里，否则报警（不阻塞，便于先跑起来再补词表）。"""
    known = {axis: set(vocab[axis]) for axis in vocab}
    bad = []
    for r in rows:
        for axis in ("种类", "挑战", "技术"):
            for tag in r[axis]:
                if tag not in known[axis]:
                    bad.append("  %02d-%d  %s/%s" % (r["批次"], r["序号"], axis, tag))
    if bad:
        print("[warn] 以下标签不在 tools/taxonomy.csv 词表中（建议先补词表）：")
        print("\n".join(bad))
    # 数据里出现过、但词表里没有的，追加到末尾，保证表格不漏
    for axis in ("种类", "挑战", "技术"):
        for r in rows:
            for tag in r[axis]:
                if tag not in vocab[axis]:
                    vocab[axis].append(tag)


# ===========================================================================
# LaTeX → Markdown
# ===========================================================================
def protect_math(s):
    """把数学片段替换成占位符，避免后续清洗破坏公式。"""
    store = []

    def keep(m):
        store.append(m.group(0))
        return "\x00MATH%d\x00" % (len(store) - 1)

    # cases/array 整体保护（内部含 \\ 与 &，不可拆）
    s = re.sub(r"\\begin\{(cases|cases\*|array)\}(.*?)\\end\{\1\}", keep, s, flags=re.S)
    # \[..\] \(..\) 归一为 $$..$$ $..$
    s = re.sub(r"\\\[(.*?)\\\]", lambda m: "$$" + m.group(1) + "$$", s, flags=re.S)
    s = re.sub(r"\\\((.*?)\\\)", lambda m: "$" + m.group(1) + "$", s, flags=re.S)

    # equation/align/gather/eqnarray → $$\displaylines{...}$$（去 &，保留 \\）
    def disp(m):
        return "$$\\displaylines{" + m.group(2).replace("&", "") + "}$$"

    s = re.sub(r"\\begin\{(equation|align|gather|eqnarray)\*?\}(.*?)\\end\{\1\*?\}",
               disp, s, flags=re.S)
    s = re.sub(r"\$\$.*?\$\$", keep, s, flags=re.S)
    s = re.sub(r"\$[^$]*?\$", keep, s)
    return s, store


def restore_math(s, store):
    for _ in range(5):          # 占位符可能嵌套（cases 外再包一层公式）
        changed = False
        for i, v in enumerate(store):
            tok = "\x00MATH%d\x00" % i
            if tok in s:
                s = s.replace(tok, v)
                changed = True
        if not changed:
            break
    return s


def fix_aligned(s):
    """aligned/gathered → displaylines（Obsidian 的 MathJax 兼容性更好）。"""
    store = []

    def keep(m):
        store.append(m.group(0))
        return "\x00MATH%d\x00" % (len(store) - 1)

    s = re.sub(r"\\begin\{(cases|cases\*|array)\}(.*?)\\end\{\1\}", keep, s, flags=re.S)
    s = re.sub(r"\\begin\{(aligned|gathered)\}(.*?)\\end\{\1\}",
               lambda m: "\\displaylines{" + m.group(2).replace("&", "") + "}", s, flags=re.S)
    return restore_math(s, store)


LIST_ENVS = [("itemize", "- "), ("enumerate", "1. ")]


def convert_lists(s):
    """itemize/enumerate → Markdown 列表（默认实现会把整个列表内容丢掉，这里改为保留）。"""
    for env, bullet in LIST_ENVS:
        pat = re.compile(r"\\begin\{%s\}(.*?)\\end\{%s\}" % (env, env), flags=re.S)

        def repl(m):
            body = re.sub(r"\\item(\[[^\]]*\])?", "\n" + bullet, m.group(1))
            return "\n" + body.strip() + "\n"

        prev = None
        while prev != s:        # 处理嵌套列表
            prev = s
            s = pat.sub(repl, s)
    return s


SECTION_RE = re.compile(r"\\(sub){0,2}section\*?\{([^{}]*)\}")


def latex_to_md(s, math_fix=None):
    s = convert_lists(s)
    s, math = protect_math(s)

    # 整块丢弃：图、表、代码（卡片不放图，看图直接看 PDF 报告）
    s = re.sub(r"\\begin\{(figure|table|tabular|longtable|verbatim|lstlisting)\*?\}"
               r".*?\\end\{\1\*?\}", " ", s, flags=re.S)
    # 交叉引用与引用命令 → 删除；题注 → 删除
    s = re.sub(r"\\(label|ref|eqref|cref|Cref|cite|citep|citet|nocite|autoref)\{[^}]*\}", "", s)
    s = re.sub(r"\\caption\{([^{}]*)\}", "", s)
    # 各级标题 → 加粗小标题
    s = SECTION_RE.sub(lambda m: "\n\n**" + m.group(2).strip() + "**\n\n", s)
    # 行内格式命令 → 保留内容
    s = re.sub(r"\\(textbf|textit|emph|underline|text|mathrm|mathbf|mathcal|mathnormal|"
               r"textrm|texttt|operatorname|textsc)\{([^{}]*)\}", r"\2", s)
    # 纯排版命令 → 删除
    s = re.sub(r"\\(noindent|par|small|footnotesize|scriptsize|large|centering|"
               r"hline|toprule|midrule|bottomrule|newline|bigskip|medskip|clearpage|"
               r"raggedright|linewidth|columnwidth|textwidth|checkmark|qquad\b)", " ", s)
    s = re.sub(r"\\vspace\*?\{[^}]*\}|\\hspace\*?\{[^}]*\}", " ", s)
    # 残余的环境与命令
    s = re.sub(r"\\begin\{[^}]*\}|\\end\{[^}]*\}", " ", s)
    s = re.sub(r"\\[a-zA-Z]+\*?(\[[^\]]*\])?(\{[^{}]*\})?", " ", s)
    # 转义字符还原
    for a, b in (("\\%", "%"), ("\\&", "&"), ("\\_", "_"), ("\\#", "#"),
                 ("\\$", "$"), ("\\{", "{"), ("\\}", "}")):
        s = s.replace(a, b)
    s = s.replace("\\\\", "\n")          # LaTeX 换行
    s = re.sub(r"[ \t\u00a0]+", " ", s)
    s = re.sub(r"(?m)^[ \t]+(?![ \t]*(?:[-*+]\s|\d+\.\s|>|\||#))", "", s)   # 去段首缩进
    s = re.sub(r"\n\s*\n+", "\n\n", s)
    s = restore_math(s, math)

    # ---- 公式后处理（在还原后的正文上做）----
    s = re.sub(r"\\(label|nonumber|notag|displaystyle|limits)\b(\{[^}]*\})?", "", s)
    s = re.sub(r"\\(tag|eqno)\s*\{\s*\(?([^}]*)\)?\s*\}", r"\\qquad (\2)", s)
    s = s.replace("\\boldsymbol", "\\mathbf")
    s = fix_aligned(s)
    # 独立公式块前后留空行，保证 Obsidian 能正确渲染
    s = re.sub(r"\n[ \t]*(\$\$)", r"\n\n\1", s)
    s = re.sub(r"(\$\$)[ \t]*\n", r"\1\n\n", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    if math_fix:
        for a, b in math_fix:
            s = s.replace(a, b)
    return s.strip()


CJK_RE = re.compile(r"[\u4e00-\u9fff]")
MATH_BLOCK_RE = re.compile(r"\$\$.*?\$\$|\$[^$\n]*?\$", flags=re.S)


def check_math_language(text, where):
    """MathJax 下数学环境里不建议出现中文。这里只提示，由作者回 LaTeX 改。"""
    hits = []
    for m in MATH_BLOCK_RE.finditer(text):
        if CJK_RE.search(m.group(0)):
            snippet = re.sub(r"\s+", " ", m.group(0))[:60]
            hits.append(snippet)
    if hits:
        print("[warn] %s：公式内出现中文，建议改写成 \\text{english}（%d 处）" % (where, len(hits)))
        for h in hits[:5]:
            print("        " + h)
    return hits


def read_math_fix(path):
    if not path or not os.path.exists(path):
        return []
    pairs = []
    with io.open(path, encoding="utf-8-sig", newline="") as f:
        for raw in csv.reader(f):
            if len(raw) >= 2 and raw[0].strip():
                pairs.append((raw[0], raw[1]))
    return pairs


def split_sections(tex):
    body = tex.split(r"\begin{document}", 1)[-1]
    body = body.split(r"\end{document}", 1)[0]
    parts = re.split(r"\\section\*?\{([^{}]*)\}", body)
    return {parts[i].strip(): parts[i + 1] for i in range(1, len(parts) - 1, 2)}


# ===========================================================================
# 卡片 / 索引 / 分类表
# ===========================================================================
def yaml_str(v):
    return '"' + str(v).replace("\\", "\\\\").replace('"', '\\"') + '"'


def slug_name(row):
    return "%02d-%d %s" % (row["批次"], row["序号"], row["短标题"])


def card_name(row):
    return "%02d-%d-%s.pdf" % (row["批次"], row["序号"], "论文原文")


def build_card(row, secs, card_body, attach_original, attach_report, index_names):
    key = "%02d-%d" % (row["批次"], row["序号"])
    fm = ["---"]
    fm.append("论文题目: " + yaml_str(row["论文题目"]))
    fm.append("期刊会议: " + yaml_str(row["期刊会议"]))
    fm.append("年份: " + (row["年份"] or ""))
    fm.append("文献类型: " + yaml_str(row["文献类型"]))
    fm.append("分区: " + yaml_str(row["分区"]))
    fm.append("影响因子: " + yaml_str(row["影响因子"]))
    fm.append("研究内容: " + yaml_str(row["研究内容"]))
    fm.append("批次: %d" % row["批次"])
    fm.append("报告时间: " + (row["报告时间"] or ""))
    fm.append("tags:")
    fm.append("  - 调研汇报论文库")
    fm.append("  - 研究论文/论文卡片")
    for axis in ("种类", "挑战", "技术"):
        for tag in row[axis]:
            fm.append("  - %s/%s" % (axis, tag))
    fm.append("---")

    out = [""]
    out += ["# " + row["短标题"], ""]
    if row.get("_citation"):
        out += ["> [!info] 文献来源", "> " + row["_citation"], ""]
    for title, keys in CARD_SECTIONS:
        text = card_body[title]
        out += ["## " + title, "", text if text else "（待补充）", ""]
    out += ["## 同类论文", "",
            "```dataview",
            "TABLE 分区 AS 分区, 年份 AS 年份, 研究内容 AS 研究内容",
            'FROM "%s"' % args_cards_dir,
            "WHERE file.name != this.file.name AND",
            "  length(filter(file.etags, (t) => contains(this.file.etags, t)",
            '    AND (startswith(t, "#种类/") OR startswith(t, "#挑战/") OR startswith(t, "#技术/")))) > 0',
            "SORT 批次 ASC",
            "```", ""]
    out += ["## 附件与出处", ""]
    if attach_original:
        out.append("- 论文原文：[[%s]]" % attach_original)
    if attach_report:
        out.append("- 阅读报告：[[%s]]" % attach_report)
    out.append("- LaTeX 源：`reports/%s/`" % key)
    out.append("- 回到 [[%s]] · [[%s]]" % (index_names[0], index_names[1]))
    out.append("")
    return "\n".join(fm + out)


def build_index(rows, cards, index_names, batch_label):
    L = ["---", "tags:", "  - 调研汇报论文库", "  - 研究论文/MOC", "---", "",
         "# " + index_names[0], "",
         "每次调研汇报的论文沉淀成卡片：一篇论文一张卡片，卡片之间靠三轴标签互相指向，"
         "PDF 原文与阅读报告放在附件目录。", "",
         "- 分类总表：[[%s]]（任务维度 + 方法挑战维度 + 技术维度）" % index_names[1],
         "- 元数据表：`reports/metadata.csv`（分区、影响因子、报告时间、标签，全库唯一来源）",
         "- LaTeX 报告源：`reports/批次-序号/`", ""]

    counts = {}
    for r in rows:
        counts[r["分区"]] = counts.get(r["分区"], 0) + 1
    summary = "；".join("%s %d 篇" % (k, v) for k, v in sorted(counts.items()))
    L += ["> [!tip] 汇总", "> 共 %d 篇；%s。" % (len(rows), summary), "",
          "## 按批次", ""]

    for b in sorted({r["批次"] for r in rows}):
        group = [r for r in rows if r["批次"] == b]
        date = next((r["报告时间"] for r in group if r["报告时间"]), "")
        L += ["### %s%s" % (batch_label.replace("{n}", str(b)),
                              "（%s）" % date if date else ""),
              "",
              "| 论文卡片 | 论文题目 | 期刊/会议 | 分区 | 报告时间 |",
              "| --- | --- | --- | --- | --- |"]
        for r in group:
            L.append("| [[%s]] | %s | %s | %s | %s |" % (
                cards[r["批次"], r["序号"]], r["论文题目"], r["期刊会议"],
                r["分区"], r["报告时间"]))
        L.append("")

    L += ["## 全部卡片", "",
          "```dataview",
          'TABLE filter(file.etags, (t) => startswith(t, "#种类/")) AS 种类, 分区 AS 分区, 年份 AS 年份, 批次 AS 批次',
          'FROM "%s"' % args_cards_dir,
          "SORT file.name ASC",
          "```", ""]
    return "\n".join(L)


def build_taxonomy(rows, cards, vocab, index_names):
    L = ["---", "tags:", "  - 调研汇报论文库", "  - 研究论文/分类表", "---", "",
         "# " + index_names[1], "",
         "全库论文按**研究问题**的分类总表。分类以嵌套标签 `种类/…`、`挑战/…`、`技术/…` "
         "落到每张 [[%s|论文卡片]] 上，点击标签即可看到同类论文。" % index_names[0], "",
         "> [!note] 字段分工（关键词只存一处）",
         "> - **关键词只写在 tags**（`种类/…`、`挑战/…`、`技术/…`），不再重复进属性——单一来源，不会写飘",
         "> - **属性只放标量字段**：论文题目、期刊会议、年份、文献类型、分区、影响因子、研究内容、批次、报告时间",
         "> - 词表见 `tools/taxonomy.csv`，新增词先改词表再加标签", "",
         "> [!tip] 怎么一键看同类论文",
         "> 点击任意标签即弹出全库同关键词的论文；也可在搜索框输入 `tag:#挑战/多模态融合`，",
         "> 或用本页下方的 Dataview（查询读的是 `file.etags`，与标签同源）。", ""]

    axis_titles = [("种类", "一、按研究问题种类（任务维度）"),
                   ("挑战", "二、按问题挑战（方法维度）"),
                   ("技术", "三、按方法技术（技术维度）")]
    for axis, title in axis_titles:
        L += ["## " + title, ""]
        seen = set()
        for tag in vocab[axis]:
            if tag in seen:
                continue
            seen.add(tag)
            group = [r for r in rows if tag in r[axis]]
            if not group:
                continue
            L += ["### %s（%d 篇）" % (tag, len(group)), "",
                  "| 论文卡片 | 论文题目 | 期刊/会议 | 年份 | 分区 |",
                  "| --- | --- | --- | --- | --- |"]
            for r in group:
                L.append("| [[%s]] | %s | %s | %s | %s |" % (
                    cards[r["批次"], r["序号"]], r["论文题目"], r["期刊会议"],
                    r["年份"], r["分区"]))
            L += [""]
        # 数据里有、词表里没有的标签也列出来（放最后）
        for tag in sorted({t for r in rows for t in r[axis]} - set(vocab[axis])):
            group = [r for r in rows if tag in r[axis]]
            L += ["### %s（%d 篇）" % (tag, len(group)), "",
                  "| 论文卡片 | 论文题目 | 期刊/会议 | 年份 | 分区 |",
                  "| --- | --- | --- | --- | --- |"]
            for r in group:
                L.append("| [[%s]] | %s | %s | %s | %s |" % (
                    cards[r["批次"], r["序号"]], r["论文题目"], r["期刊会议"],
                    r["年份"], r["分区"]))
            L += [""]

    L += ["## 四、交叉视图（Dataview）", "",
          "### 全部卡片（按批次）", "",
          "```dataview",
          'TABLE filter(file.etags, (t) => startswith(t, "#种类/")) AS 种类, '
          'filter(file.etags, (t) => startswith(t, "#挑战/")) AS 挑战, '
          "分区 AS 分区, 影响因子 AS IF, 年份 AS 年份",
          'FROM "%s"' % args_cards_dir,
          "SORT 批次 ASC, file.name ASC",
          "```", "",
          "### 只看 1 区期刊论文", "",
          "```dataview",
          'TABLE 期刊会议 AS 期刊, 影响因子 AS IF, '
          'filter(file.etags, (t) => startswith(t, "#种类/")) AS 种类',
          'FROM "%s"' % args_cards_dir,
          'WHERE 分区 = "1区"',
          "SORT 影响因子 DESC",
          "```", "",
          "返回 [[%s]]" % index_names[0], ""]
    return "\n".join(L)


# ===========================================================================
# 主流程
# ===========================================================================
args_cards_dir = "论文卡片"       # build_card 里生成 dataview 查询要用


def main():
    global args_cards_dir
    p = argparse.ArgumentParser(description="LaTeX 阅读报告 → Obsidian 论文卡片库")
    p.add_argument("--meta", required=True, help="元数据 CSV")
    p.add_argument("--vault", required=True, help="Obsidian 仓库根目录（分类表/总索引写到这里）")
    p.add_argument("--src", default="reports", help="LaTeX 报告根目录")
    p.add_argument("--cards-dir", default="论文卡片",
                   help="卡片目录（相对 vault），同时也是卡片里 dataview 查询的 FROM 值")
    p.add_argument("--index-dir", default=".",
                   help="分类表与总索引写到哪里（相对 vault，默认 vault 根）")
    p.add_argument("--attach-dir", default="附件",
                   help="PDF 附件目录（相对 vault）")
    p.add_argument("--taxonomy", default="tools/taxonomy.csv", help="词表 CSV（axis,tag）")
    p.add_argument("--math-fix", default="tools/math_replacements.csv",
                   help="公式内中文/记号的批量替换表（可缺省）")
    p.add_argument("--tex-pattern", default="{b:02d}-{i}/document.tex",
                   help="LaTeX 报告相对 --src 的路径模板")
    p.add_argument("--pdf-original", default="{b:02d}-{i}/论文原文.pdf",
                   help="原文 PDF 相对 --src 的路径模板（留空则不复制）")
    p.add_argument("--pdf-report", default="{b:02d}-{i}/阅读报告.pdf",
                   help="报告 PDF 相对 --src 的路径模板（留空则不复制）")
    p.add_argument("--attach-pattern", default="{b:02d}-{i}-{kind}.pdf",
                   help="复制到附件目录后的文件名模板，{kind} 为 论文原文/阅读报告")
    p.add_argument("--index-names", default="调研汇报论文库,论文研究问题种类",
                   help="总索引名,分类表名（逗号分隔）")
    p.add_argument("--batch-label", default="第 {n} 次调研", help="批次标题模板")
    p.add_argument("--classify-only", action="store_true",
                   help="只重算分类表与总索引（不读 LaTeX、不复制 PDF）")
    p.add_argument("--no-pdf", action="store_true", help="不复制 PDF")
    a = p.parse_args()

    rows = read_meta(a.meta)
    if not rows:
        sys.exit("元数据表没有有效行：%s" % a.meta)
    vocab = read_taxonomy(a.taxonomy)
    check_tags(rows, vocab)

    index_names = tuple(x.strip() for x in a.index_names.split(","))
    args_cards_dir = a.cards_dir
    cards_dir = os.path.join(a.vault, a.cards_dir)
    index_dir = os.path.join(a.vault, a.index_dir)
    attach_dir = os.path.join(a.vault, a.attach_dir)
    os.makedirs(index_dir, exist_ok=True)
    if not a.classify_only:
        os.makedirs(cards_dir, exist_ok=True)
        if not a.no_pdf:
            os.makedirs(attach_dir, exist_ok=True)

    math_fix = read_math_fix(a.math_fix)
    cards = {}
    n_new = 0

    for r in rows:
        cards[r["批次"], r["序号"]] = slug_name(r)

        if a.classify_only:
            continue

        rel = a.tex_pattern.format(b=r["批次"], i=r["序号"])
        tex_path = os.path.join(a.src, rel)
        if not os.path.exists(tex_path):
            print("[warn] 缺少 LaTeX 源：%s（跳过该篇正文）" % tex_path)
            secs = {}
        else:
            with io.open(tex_path, encoding="utf-8") as f:
                secs = split_sections(f.read())

        # 文献来源：取该节第一行非空文本
        citation = ""
        raw_cite = secs.get(CITATION_SECTION, "")
        for line in latex_to_md(raw_cite, math_fix).split("\n"):
            line = line.strip()
            if line and not line.startswith("**"):
                citation = line
                break
        r["_citation"] = citation

        body = {}
        for title, keys in CARD_SECTIONS:
            chunks = [latex_to_md(secs.get(k, ""), math_fix) for k in keys]
            body[title] = "\n\n".join(c for c in chunks if c).strip()
        check_math_language("\n\n".join(body.values()), "%02d-%d" % (r["批次"], r["序号"]))

        # PDF 附件
        names = {}
        if not a.no_pdf:
            for kind, pat in (("论文原文", a.pdf_original), ("阅读报告", a.pdf_report)):
                if not pat:
                    continue
                s = os.path.join(a.src, pat.format(b=r["批次"], i=r["序号"]))
                if os.path.exists(s):
                    d = os.path.join(attach_dir,
                                     a.attach_pattern.format(b=r["批次"], i=r["序号"], kind=kind))
                    shutil.copy2(s, d)
                    names[kind] = os.path.basename(d)
                else:
                    print("[warn] 缺少 PDF：%s" % s)

        text = build_card(r, secs, body, names.get("论文原文"), names.get("阅读报告"), index_names)
        with io.open(os.path.join(cards_dir, slug_name(r) + ".md"), "w", encoding="utf-8") as f:
            f.write(text)
        n_new += 1

    with io.open(os.path.join(index_dir, index_names[1] + ".md"), "w", encoding="utf-8") as f:
        f.write(build_taxonomy(rows, cards, vocab, index_names))
    with io.open(os.path.join(index_dir, index_names[0] + ".md"), "w", encoding="utf-8") as f:
        f.write(build_index(rows, cards, index_names, a.batch_label))

    print("完成：%d 张卡片%s；分类表与总索引已写入 %s" % (
        len(rows), "（新生成 %d）" % n_new if not a.classify_only else "（classify-only，未重新生成）",
        index_dir))


if __name__ == "__main__":
    main()
