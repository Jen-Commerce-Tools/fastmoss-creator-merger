from __future__ import annotations

import io
import re
from collections import OrderedDict, defaultdict
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import streamlit as st


# =========================
# 1. 渠道配置
# =========================
CHANNEL_CONFIG = OrderedDict(
    {
        "商品关联达人": {
            "metrics": ["销量", "销售额"],
            "priority": ["销售额", "销量"],
        },
        "商品关联视频": {
            "metrics": ["该商品销量", "该商品销售额"],
            "priority": ["该商品销售额", "该商品销量"],
        },
        "带货达人榜": {
            "metrics": [],
            "priority": [],
        },
        "黑马达人榜": {
            "metrics": [],
            "priority": [],
        },
        "竞品店铺带货达人": {
            "metrics": ["[天]带货销量", "[天]带货销售额"],
            "priority": ["[天]带货销售额", "[天]带货销量"],
        },
        "热门视频": {
            "metrics": [],
            "priority": [],
        },
    }
)

COMMON_COLUMN_ORDER = [
    "达人ID",
    "达人昵称",
    "来源渠道",
    "来源文件",
    "国家/地区",
    "达人分类",
    "粉丝数",
    "FastMoss达人详情页",
    "TikTok官网达人主页",
]


# =========================
# 2. 基础读取与字段标准化
# =========================
def clean_header(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", " ").replace("\r", " ").strip()
    text = re.sub(r"\s+", " ", text)

    # 兼容 FastMoss 未来可能出现 [7天] / [28天] / [30天] 等动态表头
    if re.fullmatch(r"\[(?:\d+)?天\]带货销量", text):
        return "[天]带货销量"
    if re.fullmatch(r"\[(?:\d+)?天\]带货销售额", text):
        return "[天]带货销售额"
    return text


def _drop_empty_rows_cols(df: pd.DataFrame) -> pd.DataFrame:
    # 不按“数据是否全空”删除有名称的列。FastMoss 可能保留字段名但整列
    # 数据为空（如 2026-10 的“开始带货时间”）；如果删掉该列，会改变
    # 用于识别渠道的表头结构。
    df = df.dropna(axis=0, how="all").copy()
    df.columns = [clean_header(c) for c in df.columns]
    # 只剔除空表头，以及真正无数据的 Excel 占位列 Unnamed: N。
    # 按位置索引，避免重名表头时 df[col] 返回多列。
    keep_positions = [
        idx
        for idx, col in enumerate(df.columns)
        if col and not (
            re.fullmatch(r"Unnamed:\s*\d+", col, flags=re.IGNORECASE)
            and df.iloc[:, idx].isna().all()
        )
    ]
    return df.iloc[:, keep_positions]


@st.cache_data(show_spinner=False)
def read_table(file_bytes: bytes, filename: str) -> pd.DataFrame:
    suffix = filename.lower().rsplit(".", 1)[-1]
    bio = io.BytesIO(file_bytes)

    if suffix == "xlsx":
        df = pd.read_excel(bio, sheet_name=0, dtype=object, engine="openpyxl")
    elif suffix == "xls":
        df = pd.read_excel(bio, sheet_name=0, dtype=object, engine="xlrd")
    elif suffix == "csv":
        last_error = None
        for enc in ("utf-8-sig", "utf-8", "gb18030"):
            try:
                bio.seek(0)
                df = pd.read_csv(bio, dtype=object, encoding=enc, sep=None, engine="python")
                break
            except Exception as exc:
                last_error = exc
        else:
            raise last_error
    else:
        raise ValueError(f"暂不支持的文件格式：{suffix}")

    return _drop_empty_rows_cols(df)


def coalesce_alias(df: pd.DataFrame, canonical: str, aliases: Iterable[str]) -> pd.DataFrame:
    """把同义字段合并成统一字段名，优先保留 canonical 已有值。"""
    aliases = [a for a in aliases if a in df.columns and a != canonical]
    if canonical not in df.columns and aliases:
        df = df.rename(columns={aliases[0]: canonical})
        aliases = aliases[1:]

    if canonical in df.columns:
        for alias in aliases:
            left = df[canonical]
            right = df[alias]
            empty = left.isna() | left.astype(str).str.strip().isin(["", "nan", "None"])
            df.loc[empty, canonical] = right.loc[empty]
            df = df.drop(columns=[alias])
    return df


def normalize_common_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [clean_header(c) for c in df.columns]

    alias_map = {
        "达人ID": ["发布达人ID"],
        "粉丝数": ["达人粉丝数"],
        "国家/地区": ["国家"],
        "FastMoss达人详情页": ["FastMoss达人详情页链接", "达人FastMoss详情页链接"],
        "TikTok官网达人主页": ["TikTok官网达人主页链接", "达人TikTok主页链接"],
    }
    for canonical, aliases in alias_map.items():
        df = coalesce_alias(df, canonical, aliases)

    if "达人ID" in df.columns:
        df["达人ID"] = df["达人ID"].map(normalize_id_display)
    return df


def normalize_id_display(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if text.lower() in {"nan", "none", "null", "nat", "<na>"}:
        return ""
    return text


def creator_key(value) -> str:
    """FastMoss 达人ID匹配键；去空格、@ 前缀，不区分大小写。"""
    text = normalize_id_display(value)
    if not text:
        return ""
    return text.lstrip("@").strip().lower()


# =========================
# 3. 文件渠道识别（以表头结构为主）
# =========================
def detect_channel(columns: Iterable[str]) -> Optional[str]:
    headers = set(columns)

    if {"该商品销量", "该商品销售额", "发布达人ID"}.issubset(headers):
        return "商品关联视频"

    if {"[天]带货销量", "[天]带货销售额", "此店铺商品数"}.issubset(headers):
        return "竞品店铺带货达人"

    if "带货潜力指数" in headers and {"带货总GMV", "视频带货GMV", "直播带货GMV"}.issubset(headers):
        return "黑马达人榜"

    if "带货商品数" in headers and {"带货总GMV", "视频带货GMV", "直播带货GMV"}.issubset(headers):
        return "带货达人榜"

    if {"评论数", "达人ID", "TikTok官网视频主页", "FastMoss视频详情页"}.issubset(headers):
        return "热门视频"

    # “开始带货时间”可能整列为空，甚至在新版导出中直接缺失。
    # 用达人身份列 + 两项渠道核心指标识别，避免依赖非必要字段。
    if {"达人ID", "销量", "销售额"}.issubset(headers) and (
        {"达人昵称", "达人分类", "粉丝数", "FastMoss达人详情页"} & headers
    ):
        return "商品关联达人"

    return None


# =========================
# 4. 数值筛选
# =========================
def parse_number(value) -> float:
    if pd.isna(value):
        return float("nan")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "-", "--"}:
        return float("nan")

    text = text.replace(",", "").replace("$", "").replace("¥", "").replace("￥", "")
    text = text.replace("%", "").strip()

    multiplier = 1.0
    upper = text.upper()
    units = [
        ("B", 1_000_000_000),
        ("M", 1_000_000),
        ("K", 1_000),
        ("亿", 100_000_000),
        ("万", 10_000),
    ]
    for suffix, factor in units:
        if upper.endswith(suffix) if suffix in {"B", "M", "K"} else text.endswith(suffix):
            multiplier = factor
            text = text[:-len(suffix)].strip()
            break

    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return float("nan")
    return float(match.group()) * multiplier


def apply_filters_to_source(df: pd.DataFrame, thresholds: Dict[str, float]) -> pd.DataFrame:
    """
    仅对当前源表实际拥有的核心指标字段进行筛选。
    某个源表若没有某指标字段，则该指标不会影响此源表。
    """
    result = df.copy()
    for metric, minimum in thresholds.items():
        if metric not in result.columns:
            continue
        numeric = result[metric].map(parse_number)
        # 启用阈值后：缺失/无法识别的值不满足筛选条件
        result = result[numeric.notna() & (numeric >= minimum)].copy()
    return result


# =========================
# 5. 达人 ID 去重
# =========================
def unique_join(values: Iterable) -> str:
    seen = []
    for value in values:
        if pd.isna(value):
            continue
        text = str(value).strip()
        if text and text not in seen:
            seen.append(text)
    return " | ".join(seen)


def dedupe_by_creator_id(df: pd.DataFrame, priority_metrics: List[str]) -> Tuple[pd.DataFrame, int]:
    if df.empty:
        return df.copy(), 0
    if "达人ID" not in df.columns:
        return df.copy(), len(df)

    work = df.copy()
    work["_达人ID_key"] = work["达人ID"].map(creator_key)
    missing_mask = work["_达人ID_key"].eq("")
    missing_count = int(missing_mask.sum())

    with_id = work.loc[~missing_mask].copy()
    without_id = work.loc[missing_mask].copy()

    # 有核心指标的渠道：优先保留表现更高的一条记录
    sort_cols = []
    temp_cols = []
    for i, metric in enumerate(priority_metrics):
        if metric in with_id.columns:
            tmp = f"__priority_{i}"
            with_id[tmp] = with_id[metric].map(parse_number)
            sort_cols.append(tmp)
            temp_cols.append(tmp)

    if sort_cols:
        with_id = with_id.sort_values(sort_cols, ascending=[False] * len(sort_cols), na_position="last")

    if not with_id.empty:
        file_map = with_id.groupby("_达人ID_key", sort=False)["来源文件"].agg(unique_join)
        channel_map = with_id.groupby("_达人ID_key", sort=False)["来源渠道"].agg(unique_join)
        winners = with_id.drop_duplicates("_达人ID_key", keep="first").copy()
        winners["来源文件"] = winners["_达人ID_key"].map(file_map)
        winners["来源渠道"] = winners["_达人ID_key"].map(channel_map)
    else:
        winners = with_id

    out = pd.concat([winners, without_id], ignore_index=True, sort=False)
    drop_cols = ["_达人ID_key"] + temp_cols
    out = out.drop(columns=[c for c in drop_cols if c in out.columns])
    out = reorder_columns(out)
    return out, missing_count


def reorder_columns(df: pd.DataFrame) -> pd.DataFrame:
    first = [c for c in COMMON_COLUMN_ORDER if c in df.columns]
    rest = [c for c in df.columns if c not in first and not c.startswith("__")]
    return df.loc[:, first + rest]


def first_non_empty(series: pd.Series):
    for value in series:
        if pd.isna(value):
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return value
    return ""


def combine_all_channels(channel_results: Dict[str, pd.DataFrame]) -> Tuple[pd.DataFrame, int]:
    frames = [df for df in channel_results.values() if df is not None and not df.empty]
    if not frames:
        # 即使用户把全部渠道行筛选掉，导出仍保留 Excel 表头。
        source_frames = [df.iloc[:0] for df in channel_results.values() if df is not None]
        if not source_frames:
            return pd.DataFrame(columns=["达人ID"]), 0
        return reorder_columns(pd.concat(source_frames, ignore_index=True, sort=False)), 0

    all_rows = pd.concat(frames, ignore_index=True, sort=False)
    if "达人ID" not in all_rows.columns:
        return reorder_columns(all_rows), len(all_rows)

    all_rows["_达人ID_key"] = all_rows["达人ID"].map(creator_key)
    missing_mask = all_rows["_达人ID_key"].eq("")
    missing_count = int(missing_mask.sum())

    with_id = all_rows.loc[~missing_mask].copy()
    without_id = all_rows.loc[missing_mask].copy()

    rows = []
    if not with_id.empty:
        cols = [c for c in with_id.columns if c != "_达人ID_key"]
        for _, group in with_id.groupby("_达人ID_key", sort=False):
            row = {}
            for col in cols:
                if col in {"来源渠道", "来源文件"}:
                    row[col] = unique_join(group[col])
                else:
                    row[col] = first_non_empty(group[col])
            rows.append(row)

    merged = pd.DataFrame(rows)
    if not without_id.empty:
        without_id = without_id.drop(columns=["_达人ID_key"])
        merged = pd.concat([merged, without_id], ignore_index=True, sort=False)

    return reorder_columns(merged), missing_count


# =========================
# 6. 可选排除名单 & Excel 导出
# =========================
def load_exclusion_ids(file_bytes: bytes, filename: str, history: bool = False) -> Tuple[set, int]:
    """提取外部文件第一张工作表中的达人ID，不读取历史文件的第二张工作表。"""
    suffix = filename.lower().rsplit(".", 1)[-1]
    if history:
        if suffix not in {"xlsx", "xls"}:
            raise ValueError("历史全渠道汇总必须是双 Sheet Excel 文件（.xlsx / .xls）")
        engine = "openpyxl" if suffix == "xlsx" else "xlrd"
        with pd.ExcelFile(io.BytesIO(file_bytes), engine=engine) as book:
            if len(book.sheet_names) < 2:
                raise ValueError("历史全渠道汇总需要至少两个 Sheet；仅提取第一个 Sheet 的达人ID")

    # read_table 固定读取第一张 Sheet（索引 0），因此绝不会把 Sheet 2
    # 的“排除后达人”反向混入排除名单。
    df = normalize_common_columns(read_table(file_bytes, filename))
    if "达人ID" not in df.columns:
        raise ValueError("第一张 Sheet 未找到「达人ID」列，请检查是否上传了正确的文件")

    keys = {key for key in df["达人ID"].map(creator_key) if key}
    return keys, len(df)


def exclude_known_creators(df: pd.DataFrame, excluded_ids: set) -> Tuple[pd.DataFrame, int]:
    """只从本次完整汇总移除匹配项，不增加或修改源数据列。"""
    if "达人ID" not in df.columns or df.empty or not excluded_ids:
        return df.copy(), 0
    matched = df["达人ID"].map(creator_key).isin(excluded_ids)
    # 无ID的数据无法判断是否已联系，沿用旧版行为：保留并提示。
    return df.loc[~matched].copy(), int(matched.sum())


def to_workbook_bytes(sheets: List[Tuple[str, pd.DataFrame]]) -> bytes:
    """写一个或多个 Sheet；保持旧版导出字段与样式。"""
    output = io.BytesIO()
    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_formulas": False, "strings_to_urls": False}},
    ) as writer:
        workbook = writer.book
        header_fmt = workbook.add_format(
            {"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F2937",
             "align": "center", "valign": "vcenter", "border": 0}
        )
        for raw_sheet_name, df in sheets:
            sheet_name = re.sub(r"[\\/*?:\[\]]", "_", raw_sheet_name)[:31] or "结果"
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            worksheet = writer.sheets[sheet_name]
            worksheet.set_row(0, 24, header_fmt)
            worksheet.freeze_panes(1, 0)
            if len(df.columns) > 0:
                worksheet.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)
            for idx, col in enumerate(df.columns):
                sample = df[col].head(200).astype(str).replace({"nan": "", "None": "", "<NA>": ""})
                max_len = max([len(str(col))] + [len(x) for x in sample])
                if "链接" in str(col) or "主页" in str(col):
                    width = min(max(max_len + 2, 24), 42)
                else:
                    width = min(max(max_len + 2, 10), 30)
                worksheet.set_column(idx, idx, width)
    return output.getvalue()


def to_excel_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    """兼容原来的分渠道单 Sheet 下载函数。"""
    return to_workbook_bytes([(sheet_name, df)])


# =========================
# 7. Streamlit UI
# =========================
st.set_page_config(page_title="FastMoss 达人合并", page_icon="◼", layout="wide")

st.markdown(
    """
    <style>
      .block-container {max-width: 1280px; padding-top: 2.2rem; padding-bottom: 4rem;}
      h1 {font-size: 2rem !important; letter-spacing: -0.02em;}
      h2, h3 {letter-spacing: -0.01em;}
      [data-testid="stMetricValue"] {font-size: 1.45rem;}
      .small-note {color:#6B7280; font-size:0.9rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("FastMoss 达人合并")
st.caption("渠道识别 · 可选筛选 · 达人去重 · 排除近期已处理达人")

st.markdown("#### 1 · 今日 FastMoss 表格")
uploaded_files = st.file_uploader(
    "上传渠道表格",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True,
    key="channel_uploads",
    help="支持六种渠道混合上传；按表头自动识别，无需手动选渠道。",
)

with st.expander("2 · 排除已有达人（可选）", expanded=False):
    left_upload, right_upload = st.columns(2, gap="large")
    with left_upload:
        contact_files = st.file_uploader(
            "单日重点建联达人",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key="contact_uploads",
            help="支持多份文件，读取第一张 Sheet 的「达人ID」。",
        )
    with right_upload:
        history_file = st.file_uploader(
            "近期全渠道汇总",
            type=["xlsx", "xls"],
            accept_multiple_files=False,
            key="history_upload",
            help="双 Sheet 版本；仅提取 Sheet 1 的「达人ID」，不读取 Sheet 2。",
        )
    st.caption("排除表格只用于生成新增候选名单，不影响原六渠道结果或完整全渠道汇总。")

if not uploaded_files:
    st.info("请先上传 FastMoss 渠道表格。")
    st.stop()

# 先检查排除文件；失败时明确提示，不能悄悄忽略而输出错误的候选名单。
using_exclusions = bool(contact_files) or history_file is not None
contact_ids = set()
history_ids = set()
exclusion_errors = []
contact_rows = 0
history_rows = 0

for item in contact_files or []:
    try:
        keys, row_count = load_exclusion_ids(item.getvalue(), item.name)
        contact_ids.update(keys)
        contact_rows += row_count
    except Exception as exc:
        exclusion_errors.append(f"重点建联文件 {item.name}：{exc}")

if history_file is not None:
    try:
        history_ids, history_rows = load_exclusion_ids(
            history_file.getvalue(), history_file.name, history=True
        )
    except Exception as exc:
        exclusion_errors.append(f"历史汇总文件 {history_file.name}：{exc}")

excluded_ids = contact_ids | history_ids
if using_exclusions:
    st.caption(
        f"排除名单：重点建联 {len(contact_ids):,} 人"
        f" · 近期汇总 {len(history_ids):,} 人"
        f" · 合并去重 {len(excluded_ids):,} 人"
    )
    if (contact_files and contact_rows and not contact_ids) or (history_file and history_rows and not history_ids):
        st.warning("某份排除文件没有有效达人ID，请检查内容是否正确。")
    for error in exclusion_errors:
        st.error(error)

channel_files = defaultdict(list)
recognition_rows = []
unrecognized = []
read_errors = []

for uploaded in uploaded_files:
    try:
        raw_df = read_table(uploaded.getvalue(), uploaded.name)
        channel = detect_channel(raw_df.columns)
        if channel is None:
            unrecognized.append(uploaded.name)
            recognition_rows.append({"文件": uploaded.name, "渠道": "未识别", "数据行": len(raw_df)})
            continue

        df = normalize_common_columns(raw_df)
        channel_files[channel].append({"name": uploaded.name, "df": df})
        recognition_rows.append({"文件": uploaded.name, "渠道": channel, "数据行": len(df)})
    except Exception as exc:
        read_errors.append((uploaded.name, str(exc)))
        recognition_rows.append({"文件": uploaded.name, "渠道": "读取失败", "数据行": "-"})

# 顶部概览
if not channel_files:
    st.warning("目前没有可识别的 FastMoss 渠道文件，请检查上传区和表头。")

c1, c2, c3, c4 = st.columns(4)
c1.metric("上传文件", len(uploaded_files))
c2.metric("识别渠道", len(channel_files))
c3.metric("成功识别", sum(len(v) for v in channel_files.values()))
c4.metric("未识别/失败", len(unrecognized) + len(read_errors))

with st.expander("查看文件识别结果", expanded=False):
    st.dataframe(pd.DataFrame(recognition_rows), use_container_width=True, hide_index=True)
    if unrecognized:
        st.warning("未识别文件：" + "、".join(unrecognized))
    for filename, err in read_errors:
        st.error(f"{filename}：{err}")

st.divider()

channel_results: Dict[str, pd.DataFrame] = {}
channel_stats = {}

for channel, config in CHANNEL_CONFIG.items():
    items = channel_files.get(channel, [])
    if not items:
        continue

    with st.container(border=True):
        title_col, count_col = st.columns([5, 1])
        title_col.subheader(channel)
        count_col.caption(f"{len(items)} 个文件")

        thresholds: Dict[str, float] = {}
        available_metrics = [
            metric
            for metric in config["metrics"]
            if any(metric in item["df"].columns for item in items)
        ]

        if available_metrics:
            st.caption("核心指标筛选（可选）")
            metric_cols = st.columns(len(available_metrics))
            for metric, col in zip(available_metrics, metric_cols):
                with col:
                    enabled = st.checkbox(f"启用 {metric}", key=f"enable::{channel}::{metric}")
                    if enabled:
                        minimum = st.number_input(
                            f"{metric} ≥",
                            min_value=0.0,
                            value=0.0,
                            step=1.0,
                            key=f"min::{channel}::{metric}",
                        )
                        thresholds[metric] = float(minimum)
        else:
            st.caption("当前渠道暂无核心指标筛选")

        raw_count = sum(len(item["df"]) for item in items)
        filtered_parts = []
        for item in items:
            part = apply_filters_to_source(item["df"], thresholds)
            part = part.copy()
            part["来源渠道"] = channel
            part["来源文件"] = item["name"]
            filtered_parts.append(part)

        merged = pd.concat(filtered_parts, ignore_index=True, sort=False) if filtered_parts else pd.DataFrame()
        filtered_count = len(merged)
        deduped, missing_id_count = dedupe_by_creator_id(merged, config["priority"])
        channel_results[channel] = deduped
        channel_stats[channel] = (raw_count, filtered_count, len(deduped))

        m1, m2, m3 = st.columns(3)
        m1.metric("原始数据", f"{raw_count:,}")
        m2.metric("筛选后", f"{filtered_count:,}")
        m3.metric("达人去重后", f"{len(deduped):,}")

        if missing_id_count:
            st.warning(f"有 {missing_id_count} 条记录缺少达人ID，无法按ID去重，已原样保留。")

        preview_cols = [
            c
            for c in ["达人ID", "达人昵称", "粉丝数"] + available_metrics + ["来源文件"]
            if c in deduped.columns
        ]
        st.dataframe(
            deduped.loc[:, preview_cols].head(80) if preview_cols else deduped.head(80),
            use_container_width=True,
            hide_index=True,
            height=300,
        )

        st.download_button(
            "下载渠道结果",
            data=to_excel_bytes(deduped, channel),
            file_name=f"{channel}_达人去重结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download::{channel}",
        )

# 全渠道结果
if channel_results:
    st.divider()
    with st.container(border=True):
        st.subheader("全渠道达人")
        all_result, all_missing_id = combine_all_channels(channel_results)
        before_cross_dedupe = sum(len(df) for df in channel_results.values())
        remaining, excluded_count = exclude_known_creators(all_result, excluded_ids)

        if using_exclusions:
            a1, a2, a3 = st.columns(3)
            a1.metric("本次全渠道去重", f"{len(all_result):,}")
            a2.metric("排除已处理", f"{excluded_count:,}")
            a3.metric("剩余候选达人", f"{len(remaining):,}")
            st.caption(
                f"已纳入 {len(channel_results)} 个渠道 · 渠道内去重合计 {before_cross_dedupe:,} 条"
            )
        else:
            a1, a2, a3 = st.columns(3)
            a1.metric("已纳入渠道", len(channel_results))
            a2.metric("渠道去重达人合计", f"{before_cross_dedupe:,}")
            a3.metric("全渠道去重后", f"{len(all_result):,}")

        if all_missing_id:
            st.warning(f"有 {all_missing_id} 条记录缺少达人ID，无法按ID去重或排除，已保留。")

        preview_cols = [
            c for c in ["达人ID", "达人昵称", "粉丝数", "来源渠道", "来源文件"]
            if c in all_result.columns
        ]
        if using_exclusions:
            preview_mode = st.radio(
                "数据预览", ["Sheet 2 · 排除后达人", "Sheet 1 · 全渠道达人"],
                horizontal=True, key="all_preview_mode",
            )
            preview_df = remaining if preview_mode.startswith("Sheet 2") else all_result
        else:
            preview_df = all_result

        st.dataframe(
            preview_df.loc[:, preview_cols].head(100) if preview_cols else preview_df.head(100),
            use_container_width=True,
            hide_index=True,
            height=360,
        )

        if exclusion_errors:
            st.error("排除文件读取失败：请修正或移除错误文件后再下载全渠道结果，避免漏排除。")
        else:
            if using_exclusions:
                data = to_workbook_bytes([
                    ("全渠道达人", all_result),
                    ("排除后达人", remaining),
                ])
                file_name = "全渠道达人_排除已处理.xlsx"
                button_label = "下载全渠道结果（双 Sheet）"
            else:
                data = to_excel_bytes(all_result, "全渠道达人")
                file_name = "全渠道达人_去重结果.xlsx"
                button_label = "下载全渠道结果"
            st.download_button(
                button_label,
                data=data,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="all_channels_download",
            )
