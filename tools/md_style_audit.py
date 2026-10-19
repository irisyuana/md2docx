#!/usr/bin/env python3
"""口语化表述排查与修正工具（配合 MD格式约束.md §2.2 正式文风）。

词表按三类组织，便于在同类交付文档中复用。

用法:
  python3 tools/md_style_audit.py [文件...]   # 只读扫描（缺省扫描 build/ 全部 md）
  python3 tools/md_style_audit.py --list      # 打印 A/C 两级词表，供人工修订参考
  python3 tools/md_style_audit.py --apply [文件...]  # 仅执行 A 类无歧义词级替换（先预览后加 --apply）

输出代码:
  [A] 词级替换      —— 有唯一书面对应词，--apply 可自动替换
  [B] 疑问式标题    —— 标题/表题/表头出现疑问式措辞，需人工改写为名词性陈述
  [C] 口语表述建议  —— 口语词或句式，需结合上下文人工判断

扫描会自动跳过围栏代码块、行内代码和行内/独立公式，避免误报。
"""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
from md_audit import list_mds  # noqa: E402  复用与转换管线一致的扫描范围

# ---------------------------------------------------------------------------
# A 类：词级替换（口语词 -> 唯一书面对应词，可安全自动替换）
# 条目为 (模式, 替换)；模式可为普通字符串或已编译正则（用于规避子串误匹配）。
# ---------------------------------------------------------------------------
A_RULES = [
    (re.compile(r"(?<!堵)塞入"), "并入"),
    ("搬家", "迁移"),
    ("偷看", "预读"),
    ("胜出", "占优"),
    ("变错", "失效"),
    ("分叉口", "岔路口"),
    ("题眼", "核心主张"),
    ("搬到", "迁移至"),
    (re.compile(r"(?<![替转])换成"), "替换为"),
    ("拿到", "获得"),
    ("用来", "用于"),
    ("反直觉", "与直觉相悖"),
    ("静默剔除", "不作说明地剔除"),
    ("望远镜消去", "相邻项依次相消"),
    ("完成平方", "配方"),
    ("边收边算", "接收与计算并行"),
    ("边通边算", "通信与计算并行"),
    ("先算后传", "先计算、后传输"),
    (re.compile(r"(?<!线)上链路"), "占用链路传输"),
    ("压成一个分数", "压缩为单一标量指标"),
    ("瞬间变成", "即时转化为"),
    ("才知道", "方可获知"),
    ("一起到达", "同时到达"),
    ("更准但更慢", "提高精度但增加时延"),
]

# ---------------------------------------------------------------------------
# B 类：疑问式标题/表题/表头（标题层级、整行加粗表题、表格表头行中出现即报）
# 标题与表头应使用名词性陈述（对象+维度），疑问式仅适用于正文设问。
# ---------------------------------------------------------------------------
B_INTERROGATIVES = [
    "什么", "如何", "哪些", "哪个", "为何", "怎样", "怎么",
    "多少", "多大", "何时", "何处",
    "是不是", "能不能", "要不要", "吗？", "呢？",
]

# ---------------------------------------------------------------------------
# C 类：口语词与口语句式（仅提示，需结合上下文人工改写）
# ---------------------------------------------------------------------------
C_RULES = [
    # 口语动词/副词
    ("伪装成", "建议改为：表述为／呈现为（多用于否定句）"),
    ("强行", "建议改为：勉强／强制（按语义）"),
    ("重发", "建议改为：重新分配／重新发送（注意：通信领域“重发”＝重传属正常术语，勿误改）"),
    ("变成", "建议改为：转换为／成为／转化为"),
    ("搞", "建议改为：执行／构建／处理（按语义）"),
    ("弄", "建议改为：处理／设置（按语义）"),
    ("咱们", "建议改为：我们（交付文档建议避免第一人称复数）"),
    ("一下", "建议改为：一次／稍后，或删除"),
    ("挺", "建议改为：较为／相当，或删除"),
    ("好多", "建议改为：较多／大量"),
    ("差不多", "建议改为：近似／大致"),
    ("啥", "建议改为：什么（并优先改为名词性陈述）"),
    ("咋", "建议改为：如何（并优先改为名词性陈述）"),
    # 口语名词
    ("东西", "建议改为：内容／对象／要素"),
    ("事情", "建议改为：事项／任务"),
    ("里面", "建议改为：内／中"),
    # 口语句式
    ("的话", "口语条件句尾，建议改为：若／如果（前置）"),
    ("来说", "建议改为：就……而言／对……而言"),
    ("比如说", "建议改为：例如"),
    ("比方说", "建议改为：例如"),
    ("说白了", "建议改为：实质上"),
    ("一般来说", "建议改为：一般而言／通常"),
    ("很多情况下", "建议改为：在多数情形下"),
    ("越来越", "建议改为：日益／逐渐，或给出量化比较"),
    ("一边……一边", "口语并列，建议改为：……的同时……"),
    ("就变成了", "建议改为：则转化为／即成为"),
    ("堆一套", "建议改为：预先构建完整的……系统"),
    ("拼成", "建议改为：拼接为／组合为（并核查是否真实存在）"),
    ("拆成", "建议改为：分解为"),
    ("拆解成", "建议改为：分解为"),
    ("算是", "建议改为：视为"),
    ("没用", "建议改为：无效／无益"),
    ("意义不大", "建议改为：作用有限"),
    ("白白", "建议改为：无谓地，或删除"),
]

# 句式级正则（C 类补充：跨词模式与子串误报规避）
C_PATTERNS = [
    (re.compile(r"一边[^，。；]{1,12}一边"), "口语并列结构，建议改为：……的同时……"),
    (re.compile(r"其实(?!际)"), "建议改为：实际上，或删除"),
    (re.compile(r"有用(?!性)"), "建议改为：具有作用／有效／产生效用"),
    (re.compile(r"等到(?!达)"), "建议改为：等待至／等待……后"),
]

# 句式级改写示例（来自修订实践，供人工参考，不参与自动匹配）：
#   "先把……的逻辑做实"                 -> "首先将……的机制逻辑完整建立"
#   "不能伪装成可部署方法"             -> "不得表述为可部署方法"
#   "协调器只有收到后才知道"           -> "协调器仅在收到该消息后方可获知"
#   "证明……有用，没有证明……一定更优"   -> "证明……具有实际作用，但未证明……必然更优"
#   "防止旧路径配新资源、旧计划配新地图" -> "防止新资源与旧路径、新地图与旧计划之间的失配"
#   "两个端点可以直接核算"             -> "两个端点配置可直接核算"


# ---------------------------------------------------------------------------
# 文本分段：跳过围栏代码块、行内代码与公式
# ---------------------------------------------------------------------------

def segment_line(line):
    """把一行切成 (文本, 是否参与检查) 段；行内代码与公式不参与检查。"""
    segs = []
    pat = re.compile(r"(`[^`]*`|\$\$[^$]*\$\$|\$[^$]+\$)")
    pos = 0
    for m in pat.finditer(line):
        if m.start() > pos:
            segs.append((line[pos:m.start()], True))
        segs.append((m.group(0), False))
        pos = m.end()
    if pos < len(line):
        segs.append((line[pos:], True))
    return segs


def is_table_header(lines, idx):
    """判断 lines[idx] 是否为表格表头行（下一非空行是 |---| 分隔行）。"""
    j = idx + 1
    while j < len(lines) and not lines[j].strip():
        j += 1
    return j < len(lines) and re.match(r"^\s*\|?[\s:|-]*-[\s:|-]*\|?\s*$", lines[j])


def is_bold_title(line):
    return re.match(r"^\s*\*\*[^*]+\*\*\s*$", line)


def scan(path):
    lines = Path(path).read_text(encoding="utf-8").splitlines()
    issues = []
    in_fence = False
    for i, raw in enumerate(lines, 1):
        if re.match(r"^ {0,3}`{3,}", raw):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        heading = re.match(r"^(#{1,6})\s+(.*)$", raw)
        header_row = "|" in raw and is_table_header(lines, i - 1)

        # B 类：标题 / 整行加粗表题 / 表头行
        if heading or is_bold_title(raw) or header_row:
            for w in B_INTERROGATIVES:
                if w in raw:
                    kind = "标题" if heading else ("表题" if is_bold_title(raw) else "表头")
                    issues.append((i, "B", f"{kind}疑问式措辞「{w}」: {raw.strip()[:40]}"))
                    break

        # A / C 类：逐段检查
        for text, check in segment_line(raw):
            if not check:
                continue
            for w, rep in A_RULES:
                if isinstance(w, re.Pattern):
                    m = w.search(text)
                    if m:
                        issues.append((i, "A", f"「{m.group(0)}」→「{rep}」: {raw.strip()[:40]}"))
                elif w in text:
                    issues.append((i, "A", f"「{w}」→「{rep}」: {raw.strip()[:40]}"))
            for w, sug in C_RULES:
                if w in text:
                    issues.append((i, "C", f"「{w}」{sug}: {raw.strip()[:40]}"))
            for pat, sug in C_PATTERNS:
                m = pat.search(text)
                if m:
                    issues.append((i, "C", f"「{m.group(0)}」{sug}: {raw.strip()[:40]}"))
    # 同一行重复报出去重（保持顺序）
    seen, out = set(), []
    for it in issues:
        key = (it[0], it[1], it[2])
        if key not in seen:
            seen.add(key)
            out.append(it)
    return out


def apply_a(path):
    """仅替换 A 类词（跳过代码与公式），返回替换计数。"""
    p = Path(path)
    text = p.read_text(encoding="utf-8")
    lines = text.splitlines(keepends=True)
    count = 0
    in_fence = False
    for idx, raw in enumerate(lines):
        if re.match(r"^ {0,3}`{3,}", raw):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        new_line, changed = [], False
        for text_seg, check in segment_line(raw.rstrip("\n")):
            if not check:
                new_line.append(text_seg)
                continue
            for w, rep in A_RULES:
                if isinstance(w, re.Pattern):
                    text_seg, n = w.subn(rep, text_seg)
                    count += n
                    changed = changed or n > 0
                elif w in text_seg:
                    count += text_seg.count(w)
                    text_seg = text_seg.replace(w, rep)
                    changed = True
            new_line.append(text_seg)
        if changed:
            lines[idx] = "".join(new_line) + ("\n" if raw.endswith("\n") else "")
    if count:
        p.write_text("".join(lines), encoding="utf-8")
    return count


def print_list():
    print("A 类（词级替换，--apply 可自动执行）:")
    for w, rep in A_RULES:
        label = w.pattern if isinstance(w, re.Pattern) else w
        print(f"  {label}  ->  {rep}")
    print("\nB 类（疑问式检测词，作用于标题/表题/表头）:")
    print("  " + "、".join(B_INTERROGATIVES))
    print("\nC 类（口语表述，仅提示）:")
    for w, sug in C_RULES:
        print(f"  {w}  --  {sug}")
    for pat, sug in C_PATTERNS:
        print(f"  /{pat.pattern}/  --  {sug}")


def main():
    argv = sys.argv[1:]
    do_apply = "--apply" in argv
    if "--list" in argv:
        print_list()
        return
    argv = [a for a in argv if a != "--apply"]
    files = list_mds(argv) if not argv else argv
    if do_apply:
        total = 0
        for f in files:
            n = apply_a(ROOT / f if not Path(f).is_absolute() else Path(f))
            if n:
                print(f"{f}: 自动替换 {n} 处")
                total += n
        print(f"共替换 {total} 处；B/C 类需人工修订，请重新扫描确认。")
        return
    total = 0
    for f in files:
        issues = scan(ROOT / f if not Path(f).is_absolute() else Path(f))
        if issues:
            print(f"\n=== {f} ({len(issues)}) ===")
            for ln, code, msg in issues:
                print(f"  L{ln} [{code}] {msg}")
            total += len(issues)
    print(f"\n扫描 {len(files)} 个文件，共 {total} 处口语化提示（A 可 --apply 自动替换，B/C 需人工改写）")


if __name__ == "__main__":
    main()
