# -*- coding: utf-8 -*-
"""客服聊天记录 → 贴纸定制汇总表（列与 Excel 模板一致）。支持 DeepSeek / MiniMax 与本地学习库。"""

from __future__ import annotations

import json
import os
import re
import sys
import time
from datetime import datetime
from io import BytesIO
from pathlib import Path

import httpx
import pandas as pd
import streamlit as st

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None  # type: ignore


def _app_dir() -> Path:
    """源码运行：脚本目录；PyInstaller 打包：exe 所在目录（便于 data/ 可写）。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = _app_dir()
DATA_DIR = APP_DIR / "data"
LEARNING_PATH = DATA_DIR / "learning_cases.jsonl"
SETTINGS_PATH = DATA_DIR / "app_settings.json"

# MiniMax 文本接口（与 A-SIR 项目文档一致）：https://api.minimaxi.com/v1/text/chatcompletion_v2
MINIMAX_NATIVE_URL = "https://api.minimaxi.com/v1/text/chatcompletion_v2"
MINIMAX_RETRY_HTTP = {529}
MINIMAX_RETRY_BASE_STATUS = {1000, 1001, 2064}

DEFAULT_SETTINGS: dict[str, str] = {
    "provider": "DeepSeek",
    "deepseek_key": "",
    "minimax_key": "",
    "minimax_model": "MiniMax-M2.7",
}

MINIMAX_NATIVE_MODELS = [
    "MiniMax-M2.7",
    "MiniMax-M2.5",
    "MiniMax-M2.1",
    "MiniMax-M2",
]


def load_settings() -> dict[str, str]:
    out = dict(DEFAULT_SETTINGS)
    if SETTINGS_PATH.is_file():
        try:
            raw = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            raw = {}
        if isinstance(raw, dict):
            for k, v in raw.items():
                if isinstance(v, str):
                    out[k] = v
    if not (out.get("deepseek_key") or "").strip():
        ev = (os.getenv("DEEPSEEK_API_KEY") or "").strip()
        if ev:
            out["deepseek_key"] = ev
    if not (out.get("minimax_key") or "").strip():
        ev = (os.getenv("MINIMAX_API_KEY") or "").strip()
        if ev:
            out["minimax_key"] = ev
    return out


def save_settings_from_form(
    *,
    provider: str,
    deepseek_key: str,
    minimax_key: str,
    minimax_model: str,
) -> None:
    """写入本机；密钥框留空则保留原值。"""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    cur = load_settings()
    cur["provider"] = provider
    cur["minimax_model"] = minimax_model
    dk = normalize_api_secret(deepseek_key)
    mk = normalize_api_secret(minimax_key)
    if dk:
        cur["deepseek_key"] = dk
    if mk:
        cur["minimax_key"] = mk
    SETTINGS_PATH.write_text(json.dumps(cur, ensure_ascii=False, indent=2), encoding="utf-8")

# 与 Excel 模板一致；「工艺」列：局部烫金、ab膜、直喷（顿号连接）
COLUMNS = [
    "已开单打1",
    "时间",
    "收货地址",
    "订单号",
    "路径",
    "尺寸数量",
    "金额",
    "图片",
    "设计要求",
    "工艺",
    "",  # 空列
    "备注",
]

EXTRACTION_PROMPT_BASE = """你是贴纸定制电商的客服助理。用户会粘贴一段与客户的聊天记录、订单摘要或电商卡片文本。
请只根据文本中的事实提取，不要编造。没有的信息用空字符串 ""。

请严格输出一个 JSON 对象（不要 markdown 代码块），键名如下：
{
  "时间": "文本中的下单/聊天时间；若只有日期如 05/12 也可提取",
  "订单号": "如 订单 260512-155089693041901 中的编号；订单号、单号等",
  "收货地址": "若出现收件人、电话、省市区、详细地址，请合并成一行；完全没有则空",
  "路径": "文件路径、网盘、链接等，没有则空",
  "尺寸数量": "必须与当前商品一致：优先取规格行里的尺寸与件数（如 11*16cm、x1）；勿用无关数字",
  "金额": "实收、实付、合计、¥价格等，如 实收：¥6.16 则填 ¥6.16 或 6.16；没有则空",
  "图片": "商品名、款式简述，没有则空",
  "设计要求": "排版、颜色、出血、模切、**局部烫金/烫金/烫银**等画面与印后要求（可含金色银色说明）",
  "工艺": "**仅工艺词**，多项用顿号、连接。允许：**局部烫金**、**ab膜**、**直喷**。规则：塑料/金属/强粘/高粘/UV贴高粘等→**ab膜**（但若文本已明确「局部烫金」且无「局部烫金ab膜」等组合，**不要**再自动加 ab膜，默认只写 **局部烫金**）；贴纸盒/乳胶漆/低粘→**直喷**；出现局部烫金→必须写 **局部烫金**（四字全写，不要只写「局部」）。**不要写长句**",
  "备注": "**只写快递与客户杂项**：默认/中通/顺丰到付寄付等；**禁止**写 ab膜、直喷、胶粘、烫金工艺说明（这些只能出现在「工艺」）"
}

注意：
- 「工艺」与「设计要求」「备注」不得重复同一类信息。
- 尺寸数量、金额尽量从同一商品块中取值。
"""

BUSINESS_RULES = """
## 门店规则（胶粘/印后 →「工艺」列短码；快递 →「备注」）

### 局部烫金（写入「工艺」列，必须写全名 **局部烫金**）
- 聊天或设计要求里出现 **局部烫金**（或「局部」与「烫金」同指该工艺）→ 工艺列写 **局部烫金**，**不要**缩写成「局部」。
- **默认**：仅有局部烫金需求时，**不要再自动加 ab膜**（高粘、UV贴高粘等也不叠加），避免与车间理解不一致。
- **例外**：客户或文案**明确**「局部烫金 + ab膜」等组合（如连写、同句写明要 ab 膜）时，工艺列可同时写 **局部烫金、ab膜**。

### 胶粘（无局部烫金、或已按上条例外同时需要 ab 时）
- 高粘场景 → **ab膜**；低粘场景 → **直喷**（仍只写词，不写长句）。

### 快递（只写入 JSON「备注」）
- 未提快递 → 备注写「快递：中通（默认）」。
- 顺丰 → 备注写顺丰及到付/寄付；未说明谁付 → **到付（默认）**。

### 烫金色系（写入「设计要求」为主；备注不写烫金工艺）
- 金色系：玫瑰金、哑金、亮金（默认）、镭射金；银：亮银（默认）、哑银、镭射银。写在设计要求里便于做图。

### 输出纪律
- 「工艺」字段只能是 **局部烫金、ab膜、直喷** 的组合（顿号、），默认局部烫金场景**不**附带 ab膜。
- 「备注」**禁止**出现「胶粘：」「AB膜」「直喷胶」「烫金：」等工艺长句；仅快递与杂项。
"""


def load_learning_block(max_cases: int = 12, excerpt_chars: int = 900) -> str:
    """读取本地学习库，作为 system 补充，让模型对齐你已确认过的写法。"""
    if not LEARNING_PATH.is_file():
        return ""
    lines = LEARNING_PATH.read_text(encoding="utf-8").strip().splitlines()
    if not lines:
        return ""
    picked = lines[-max_cases:]
    blocks: list[str] = []
    for i, line in enumerate(picked, 1):
        try:
            obj = json.loads(line)
        except json.JSONDecodeError:
            continue
        ex = (obj.get("chat_excerpt") or "")[:excerpt_chars]
        rem = (obj.get("remark") or "").strip()
        if not ex and not rem:
            continue
        blocks.append(f"### 已确认范例 {i}\n聊天摘录：\n{ex}\n当时采用的「备注」写法：\n{rem}\n")
    if not blocks:
        return ""
    return (
        "\n## 历史已确认范例（请模仿备注的格式、用词与颗粒度；与客户原话冲突时以原话为准）\n"
        + "\n".join(blocks)
    )


def build_system_prompt() -> str:
    return EXTRACTION_PROMPT_BASE + "\n" + BUSINESS_RULES + load_learning_block()


def append_learning_case(*, chat_excerpt: str, remark: str) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    rec = {
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "chat_excerpt": (chat_excerpt or "").strip()[:12000],
        "remark": (remark or "").strip()[:8000],
    }
    with LEARNING_PATH.open("a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False) + "\n")


def learning_case_count() -> int:
    if not LEARNING_PATH.is_file():
        return 0
    with LEARNING_PATH.open("r", encoding="utf-8") as f:
        return sum(1 for line in f if line.strip())


def clear_learning_cases() -> None:
    if LEARNING_PATH.is_file():
        LEARNING_PATH.unlink()


def _note_has(note: str, *needles: str) -> bool:
    n = note.lower()
    return any(s.lower() in n for s in needles)


PROCESS_ORDER = ("局部烫金", "ab膜", "直喷")


def _chat_need_ab(c: str) -> bool:
    if not c:
        return False
    low = c.lower()
    keys = (
        "塑料",
        "金属",
        "PVC",
        "pvc",
        "不锈钢",
        "玻璃",
        "强粘",
        "粘性强",
        "粘性要强",
        "高粘",
        "牢固",
        "怕掉",
        "贴不住",
        "容易掉",
    )
    if any(k in c for k in keys):
        return True
    if "uv贴高粘" in low or ("uv转印" in c and "高粘" in c):
        return True
    return False


def _chat_need_direct(c: str) -> bool:
    if not c:
        return False
    return any(
        k in c
        for k in ("贴纸盒", "纸盒", "乳胶漆", "白墙", "墙面", "低粘", "直喷胶", "直喷")
    )


def _text_has_partial_jubu(c: str) -> bool:
    if not c:
        return False
    return "局部烫金" in c or ("局部" in c and "烫金" in c) or bool(re.search(r"局部\s*烫", c))


def extract_amount_from_chat(c: str) -> str:
    if not c:
        return ""
    for pat in (
        r"实收[：:]\s*[¥￥]?\s*(\d+(?:\.\d+)?)",
        r"实付[款]?[：:]\s*[¥￥]?\s*(\d+(?:\.\d+)?)",
        r"合计[：:]\s*[¥￥]?\s*(\d+(?:\.\d+)?)",
        r"总价[：:]\s*[¥￥]?\s*(\d+(?:\.\d+)?)",
    ):
        m = re.search(pat, c)
        if m:
            return f"¥{m.group(1)}"
    m = re.search(r"[¥￥]\s*(\d+(?:\.\d+)?)", c)
    if m:
        return f"¥{m.group(1)}"
    return ""


def extract_address_from_chat(c: str) -> str:
    if not c:
        return ""
    m = re.search(r"收货地址[：:\s]*([^\n\r]+)", c)
    if m:
        t = m.group(1).strip()
        if len(t) >= 4:
            return t[:240]
    m = re.search(
        r"(?:收件人|收货人)[：:\s]*([^\n]{2,30})\s*[\n\r]+\s*(1[3-9]\d{9})\s*[\n\r]+\s*([^\n]{6,120})",
        c,
    )
    if m:
        return f"{m.group(1).strip()} {m.group(2).strip()} {m.group(3).strip()}"[:240]
    return ""


def scrub_process_from_remark(note: str) -> str:
    """去掉不应出现在备注里的胶粘/烫金长句（工艺已单列）。"""
    if not note:
        return ""
    s = note
    s = re.sub(r"胶粘[：:][^；;\n]+", "", s)
    s = re.sub(r"烫金[：:][^；;\n]+", "", s)
    s = re.sub(r"印后[：:][^；;\n]+", "", s)
    s = re.sub(r"AB膜[^；;\n]*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"局部烫金[^；;\n]*", "", s)  # 防备注重复写工艺词
    s = re.sub(r"[；;]{2,}", "；", s)
    return s.strip("；; \n\t").strip()


def _explicit_jubu_and_ab(blob: str) -> bool:
    """客户明确要「局部烫金 + ab膜」等组合时才为 True（默认不叠加 ab膜）。"""
    if not blob:
        return False
    compact = re.sub(r"\s+", "", blob)
    if re.search(r"局部烫金.{0,6}ab膜", compact, flags=re.IGNORECASE):
        return True
    if "局部烫金ab膜" in compact.lower():
        return True
    if "局部烫金" in blob and re.search(
        r"(?:同时|并且|另外|还要|\+|＋).{0,8}ab膜", blob, flags=re.IGNORECASE
    ):
        return True
    return False


def merge_process_tokens(chat: str, design: str, ai_gongyi: str) -> str:
    """输出 局部烫金 / ab膜 / 直喷 的顿号组合；有局部烫金时默认不加 ab膜。"""
    blob = f"{chat or ''}\n{design or ''}\n{ai_gongyi or ''}"
    has_jubu = (
        _text_has_partial_jubu(chat)
        or _text_has_partial_jubu(design)
        or _text_has_partial_jubu(ai_gongyi or "")
    )
    force_ab_with_jubu = _explicit_jubu_and_ab(blob)

    parts: list[str] = []
    if has_jubu:
        parts.append("局部烫金")

    ag = (ai_gongyi or "").strip()
    low_ag = ag.lower()
    need_ab = (
        _chat_need_ab(chat)
        or ("ab膜" in low_ag)
        or ("ab" in low_ag and "膜" in ag)
    )
    need_direct = _chat_need_direct(chat) or ("直喷" in ag)

    if has_jubu and not force_ab_with_jubu:
        need_ab = False
    elif has_jubu and force_ab_with_jubu:
        need_ab = True

    if need_ab and need_direct:
        need_direct = False
    if need_ab:
        parts.append("ab膜")
    elif need_direct:
        parts.append("直喷")

    ordered = [t for t in PROCESS_ORDER if t in parts]
    return "、".join(ordered)


def apply_note_business_rules(chat: str, row: dict[str, str]) -> dict[str, str]:
    """金额/地址启发式补全；工艺列短码；备注仅保留快递等并去掉工艺长句。"""
    out = dict(row)
    c = chat or ""
    design = (out.get("设计要求") or "").strip()
    ai_gy = (out.get("工艺") or "").strip()

    if not (out.get("金额") or "").strip():
        amt = extract_amount_from_chat(c)
        if amt:
            out["金额"] = amt
    if not (out.get("收货地址") or "").strip():
        addr = extract_address_from_chat(c)
        if addr:
            out["收货地址"] = addr

    out["工艺"] = merge_process_tokens(c, design, ai_gy)

    note = scrub_process_from_remark((out.get("备注") or "").strip())
    add: list[str] = []

    if "顺丰" in c or "sf快递" in c.replace(" ", "").lower():
        if not _note_has(note, "顺丰"):
            if "寄付" in c:
                add.append("快递：顺丰 寄付")
            elif "到付" in c:
                add.append("快递：顺丰 到付")
            else:
                add.append("快递：顺丰 到付（默认）")
    else:
        if not _note_has(note, "中通", "顺丰"):
            add.append("快递：中通（默认）")

    if add:
        sep = "；"
        tail = sep.join(add)
        note = f"{note}{sep}{tail}" if note else tail
    out["备注"] = note
    return out


def empty_row() -> dict[str, str]:
    return {c: "" for c in COLUMNS}


def merge_ai_into_row(base: dict[str, str], data: dict) -> dict[str, str]:
    row = dict(base)
    mapping = {
        "时间": "时间",
        "订单号": "订单号",
        "收货地址": "收货地址",
        "路径": "路径",
        "尺寸数量": "尺寸数量",
        "金额": "金额",
        "图片": "图片",
    }
    for k, col in mapping.items():
        v = data.get(k)
        if isinstance(v, str) and v.strip():
            row[col] = v.strip()

    req = (data.get("设计要求") or "").strip()
    if req:
        row["设计要求"] = req

    ai_gy = (data.get("工艺") or "").strip()
    if ai_gy:
        row["工艺"] = ai_gy

    note = (data.get("备注") or "").strip()
    if note:
        row["备注"] = note
    return row


def normalize_api_secret(raw: str) -> str:
    """去掉首尾空白、引号；若误粘贴了「Bearer xxx」只保留密钥部分（SDK 会自动加 Bearer）。"""
    s = (raw or "").strip()
    s = s.strip('"').strip("'")
    low = s.lower()
    if low.startswith("bearer "):
        s = s[7:].strip()
    return s


def parse_json_response(text: str) -> dict:
    text = text.strip()
    # 去掉折叠思考块（部分模型会把推理写在正文标签内）
    _open_think, _close_think = (
        "\u003credacted_thinking\u003e",
        "\u003c/redacted_thinking\u003e",
    )
    _open_reason, _close_reason = (
        "\u003credacted_reasoning\u003e",
        "\u003c/redacted_reasoning\u003e",
    )
    text = re.sub(
        _open_think + r"[\s\S]*?" + _close_think,
        "",
        text,
        flags=re.IGNORECASE,
    ).strip()
    text = re.sub(
        _open_reason + r"[\s\S]*?" + _close_reason,
        "",
        text,
        flags=re.IGNORECASE,
    ).strip()
    # 去掉可能的 ```json 包裹
    m = re.search(r"\{[\s\S]*\}", text)
    if not m:
        raise ValueError("未找到 JSON 对象")
    return json.loads(m.group(0))


def parse_minimax_message_content(content: str) -> dict:
    """MiniMax chatcompletion_v2 的 message.content 常为 JSON 字符串。"""
    text = (content or "").strip()
    try:
        obj = json.loads(text)
        if isinstance(obj, dict) and any(
            k in obj
            for k in (
                "尺寸数量",
                "备注",
                "时间",
                "订单号",
                "设计要求",
                "工艺",
                "金额",
                "收货地址",
            )
        ):
            return obj
    except json.JSONDecodeError:
        pass
    return parse_json_response(text)


def call_minimax_native(
    api_key: str,
    chat_text: str,
    system_prompt: str,
    model: str,
    *,
    max_retries: int = 4,
) -> dict:
    key = normalize_api_secret(api_key)
    if not key:
        raise ValueError("MiniMax API Key 未配置，请在「设置」中填写并保存")

    headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    user_message = "以下为聊天记录：\n\n" + (chat_text or "")[:32000]
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message},
        ],
        "temperature": 0.7,
        "max_tokens": 4096,
        "response_format": {"type": "json_object"},
    }

    last_err: Exception | None = None
    for attempt in range(max_retries):
        try:
            with httpx.Client(timeout=httpx.Timeout(90.0, connect=15.0)) as client:
                resp = client.post(MINIMAX_NATIVE_URL, json=payload, headers=headers)
        except httpx.RequestError as e:
            last_err = e
            time.sleep(1.0 * (attempt + 1))
            continue

        if resp.status_code in MINIMAX_RETRY_HTTP:
            last_err = RuntimeError(f"MiniMax HTTP {resp.status_code}（服务繁忙，已重试）")
            time.sleep(1.2 * (attempt + 1))
            continue

        try:
            resp.raise_for_status()
        except httpx.HTTPStatusError as e:
            last_err = e
            time.sleep(1.0 * (attempt + 1))
            continue

        try:
            data = resp.json()
        except json.JSONDecodeError as e:
            last_err = e
            continue

        base_resp = data.get("base_resp") or {}
        try:
            sc = int(base_resp.get("status_code", 0) or 0)
        except (TypeError, ValueError):
            sc = -1

        if sc in MINIMAX_RETRY_BASE_STATUS:
            last_err = RuntimeError(
                f"MiniMax 繁忙 {sc}: {base_resp.get('status_msg', '')}"
            )
            time.sleep(1.2 * (attempt + 1))
            continue

        if sc != 0:
            raise RuntimeError(
                f"MiniMax 业务错误 {sc}: {base_resp.get('status_msg', '')!r}"
            )

        msg = (data.get("choices") or [{}])[0].get("message") or {}
        content = msg.get("content") or ""
        if not str(content).strip():
            raise RuntimeError("MiniMax 返回空 content")
        return parse_minimax_message_content(str(content))

    if last_err:
        raise last_err
    raise RuntimeError("MiniMax 请求失败（已达重试上限）")


def call_openai_compatible(
    api_key: str,
    base_url: str,
    model: str,
    chat_text: str,
    system_prompt: str,
    *,
    temperature: float = 0.2,
    extra_body: dict | None = None,
    default_headers: dict[str, str] | None = None,
) -> dict:
    if OpenAI is None:
        raise RuntimeError("请安装 openai: pip install openai")
    client_kw: dict = {"api_key": api_key, "base_url": base_url}
    if default_headers:
        client_kw["default_headers"] = default_headers
    client = OpenAI(**client_kw)
    kwargs: dict = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": "以下为聊天记录：\n\n" + chat_text[:32000],
            },
        ],
        "temperature": temperature,
    }
    if extra_body:
        kwargs["extra_body"] = extra_body
    resp = client.chat.completions.create(**kwargs)
    raw = resp.choices[0].message.content or ""
    return parse_json_response(raw)


def call_deepseek(api_key: str, chat_text: str, system_prompt: str) -> dict:
    return call_openai_compatible(
        normalize_api_secret(api_key),
        "https://api.deepseek.com",
        "deepseek-chat",
        chat_text,
        system_prompt,
        temperature=0.2,
    )


def naive_extract(chat: str) -> dict:
    """无 API 时的极简启发式，仅作占位提示。"""
    out: dict[str, str] = {}
    # 常见数量
    q = re.findall(r"(\d+)\s*[张个份套盒包]", chat)
    sizes = re.findall(
        r"(\d+(?:\.\d+)?)\s*[×xX*]\s*(\d+(?:\.\d+)?)\s*(?:cm|CM|毫米|mm|MM)?",
        chat,
    )
    parts = []
    if sizes:
        parts.append("约 " + " / ".join(f"{a}×{b}" for a, b in sizes[:3]))
    if q:
        parts.append("数量线索: " + ",".join(q[:5]) + " …")
    out["尺寸数量"] = (
        "；".join(parts)
        if parts
        else "（无 API 时无法可靠解析；请在右上角「设置」中配置密钥后使用 AI 分析）"
    )
    out["设计要求"] = chat[:500] + ("…" if len(chat) > 500 else "")
    return out


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    # 表头第一行与模板一致
    export_df = df.copy()
    # 列名里 J 列为空字符串，Excel 会显示为空表头
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        export_df.to_excel(w, index=False, sheet_name="汇总")
    buf.seek(0)
    return buf.read()


def default_table_df() -> pd.DataFrame:
    r = empty_row()
    r["时间"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    return pd.DataFrame([r], columns=COLUMNS)


def main() -> None:
    st.set_page_config(page_title="客服聊天总结", layout="wide")

    head_l, head_r = st.columns([6, 1])
    with head_l:
        st.title("贴纸定制 · 客服聊天总结")
        st.caption(
            "粘贴聊天记录 → 自动提取尺寸/数量、做图要求等；收货地址与收件人请在表格中手填。"
            " **API 在右上角「设置」里配置**；密钥写入本机文件后**重启也会自动加载**，只有更换密钥时才需要再输入。"
        )
    with head_r:
        st.write("")  # 与标题行对齐
        st.write("")
        with st.popover("⚙️ 设置"):
            st.markdown("**API 配置**（保存在本机 `data/app_settings.json`，重启后自动读取）")
            cfg0 = load_settings()
            has_dk = bool(normalize_api_secret(cfg0.get("deepseek_key", "")))
            has_mk = bool(normalize_api_secret(cfg0.get("minimax_key", "")))
            st.caption(
                f"当前状态：DeepSeek **{'已保存密钥' if has_dk else '未配置'}**　"
                f"MiniMax **{'已保存密钥' if has_mk else '未配置'}**"
            )
            prov = st.radio(
                "默认 AI 服务",
                ("DeepSeek", "MiniMax"),
                horizontal=True,
                index=0 if cfg0.get("provider", "DeepSeek") == "DeepSeek" else 1,
            )
            st.text_input(
                "更新 DeepSeek 密钥（可选）",
                type="password",
                key="settings_deepseek_key",
                placeholder="不修改请留空；填写后点保存会覆盖磁盘上的密钥",
            )
            st.text_input(
                "更新 MiniMax 密钥（可选）",
                type="password",
                key="settings_minimax_key",
                placeholder="不修改请留空；填写后点保存会覆盖磁盘上的密钥",
            )
            mm_cur = cfg0.get("minimax_model", "MiniMax-M2.7")
            mm_idx = (
                MINIMAX_NATIVE_MODELS.index(mm_cur)
                if mm_cur in MINIMAX_NATIVE_MODELS
                else 0
            )
            mm_model = st.selectbox(
                "MiniMax 模型",
                MINIMAX_NATIVE_MODELS,
                index=mm_idx,
                help="接口：api.minimaxi.com /v1/text/chatcompletion_v2（与 A-SIR 文档一致）",
            )
            st.caption(
                "MiniMax 使用 `chatcompletion_v2` + `response_format: json_object`；"
                "密钥可写在本文件，或设置环境变量 **`MINIMAX_API_KEY`** / **`DEEPSEEK_API_KEY`**（未写文件时生效）。"
            )
            if st.button("保存设置", type="primary", use_container_width=True):
                save_settings_from_form(
                    provider=prov,
                    deepseek_key=str(st.session_state.get("settings_deepseek_key") or ""),
                    minimax_key=str(st.session_state.get("settings_minimax_key") or ""),
                    minimax_model=mm_model,
                )
                st.success("已保存。下次启动会自动加载，无需重复输入。")
                st.rerun()

    if "table_df" not in st.session_state:
        st.session_state.table_df = default_table_df()
    else:
        df = st.session_state.table_df.copy()
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = ""
        extra = [x for x in df.columns if x not in COLUMNS]
        if extra:
            df = df.drop(columns=extra, errors="ignore")
        st.session_state.table_df = df.reindex(columns=COLUMNS, fill_value="")

    cfg = load_settings()
    prov = cfg.get("provider", "DeepSeek")
    active_key = normalize_api_secret(
        cfg.get("deepseek_key", "") if prov == "DeepSeek" else cfg.get("minimax_key", "")
    )

    chat = st.text_area("粘贴聊天记录", height=280, placeholder="Ctrl+V 粘贴整段对话…")

    col_a, col_b, _ = st.columns([1, 1, 2])
    with col_a:
        do_ai = st.button("AI 分析", type="primary", disabled=not active_key)
    with col_b:
        do_naive = st.button("本地粗提取（无 API）")

    if do_ai and active_key:
        system_prompt = build_system_prompt()
        with st.spinner("正在分析…"):
            try:
                if prov == "DeepSeek":
                    extracted = call_deepseek(active_key, chat, system_prompt)
                else:
                    mm = cfg.get("minimax_model", "MiniMax-M2.7")
                    extracted = call_minimax_native(
                        active_key, chat, system_prompt, mm
                    )
                row = merge_ai_into_row(empty_row(), extracted)
                row = apply_note_business_rules(chat, row)
                if not row.get("时间"):
                    row["时间"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                st.session_state.table_df = pd.DataFrame([row], columns=COLUMNS)
                st.success("分析完成，可在下方表格中继续手改。")
            except Exception as e:
                err = str(e)
                st.error(f"调用失败：{e}")
                if prov == "MiniMax" and (
                    "401" in err
                    or "2049" in err
                    or "invalid" in err.lower()
                    or "鉴权" in err
                    or "api key" in err.lower()
                ):
                    with st.expander("MiniMax 密钥 / 鉴权问题排查", expanded=True):
                        st.markdown(
                            """
1. 本软件 MiniMax 走 **`https://api.minimaxi.com/v1/text/chatcompletion_v2`**（与 A-SIR 项目文档一致），**不是** `api.minimax.io` 的 OpenAI 兼容域名。

2. 使用控制台创建的 **API Key**；若误用网页登录 JWT（常见 **`ey-`** 开头），会报 **2049 / invalid api key**。

3. `Authorization` 头格式为 **`Bearer` + 空格 + Key**；若从别处复制了带 `Bearer` 前缀的整段，也可直接粘贴，保存时会自动去掉多余前缀。

4. 除 HTTP 状态码外，还须看返回 JSON 里的 **`base_resp.status_code`**；若为 **1000 / 1001 / 2064** 等繁忙码，软件会自动重试数次。

5. 仍失败请到 [MiniMax 开放平台](https://platform.minimax.io) 核对密钥、余额与接口权限。
                            """
                        )

    if do_naive:
        naive = naive_extract(chat)
        base = empty_row()
        base["时间"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        row = merge_ai_into_row(base, naive)
        row = apply_note_business_rules(chat, row)
        st.session_state.table_df = pd.DataFrame([row], columns=COLUMNS)
        st.info("已用本地规则生成草稿；需要 AI 时请在「设置」中配置密钥后点「AI 分析」。")

    st.subheader("汇总表（与 Excel 模板列一致）")
    edited = st.data_editor(
        st.session_state.table_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
    )
    st.session_state.table_df = edited.copy()

    st.subheader("学习库（越用越聪明）")
    st.caption(
        "把您改定后的表格连同聊天一起保存，下次 AI 会在系统提示里附带这些范例，优先模仿您的「备注」写法。"
        f" 当前已存 **{learning_case_count()}** 条（本机 `data/learning_cases.jsonl`）。"
    )
    c1, c2, c3 = st.columns([2, 2, 2])
    with c1:
        if st.button("保存当前聊天 + 表格首行到学习库"):
            if not chat.strip():
                st.warning("请先粘贴聊天记录。")
            elif len(edited) == 0:
                st.warning("表格为空。")
            else:
                remark = str(edited.iloc[0].get("备注", "") or "").strip()
                append_learning_case(chat_excerpt=chat, remark=remark)
                st.success("已保存。下次点「AI 分析」会自动参考。")
    with c2:
        wipe_ok = st.checkbox("确认清空学习库", key="learn_wipe_ok")
    with c3:
        if st.button("清空学习库", disabled=not wipe_ok):
            clear_learning_cases()
            st.success("学习库已清空。")

    xbytes = to_excel_bytes(edited)
    st.download_button(
        label="下载 Excel（当前表格全部行）",
        data=xbytes,
        file_name=f"客服汇总_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.caption(
        "列说明：A 已开单打1 | B 时间 | C 收货地址 | D 订单号 | E 路径 | F 尺寸数量 | G 金额 | H 图片 | "
        "I 设计要求 | J 工艺（局部烫金、ab膜、直喷）| K 空 | L 备注（仅快递等）"
    )


if __name__ == "__main__":
    main()
