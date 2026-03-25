#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
plot_time_history.py

功能说明：
1. 自动检索 results 目录中的 .csv/.txt/.dat 文件。
2. 读取时程数据（首列时间，后续列为响应数据）。
3. 支持“不同参数组”对比：固定物理量/节点列，跨多个工况文件同图绘制。
4. 提供简洁 Tkinter 交互界面，可自定义坐标轴、标题、图例、配色模式。
5. 图片自动保存到 results 同级 plots 目录（支持 png / tif）。

Author: Codex
"""

from __future__ import annotations

import re
import shutil
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple

import matplotlib.pyplot as plt
import pandas as pd
import tkinter as tk
import tkinter.font as tkfont
from matplotlib import cycler
from tkinter import scrolledtext
from tkinter import filedialog, messagebox, ttk


# 默认结果目录（可在界面中修改）
DEFAULT_RESULTS_DIR = r"D:\pyansys\Claude_pyansys\results"
SUPPORTED_EXTS = (".csv", ".txt", ".dat")


@dataclass
class PlotConfig:
    """绘图参数配置。"""

    xlabel: str = "时间 / s"
    ylabel: str = "响应量 / 单位"
    title: str = "时程曲线对比"
    style_mode: str = "彩色"  # 可选：彩色 / 黑白
    dpi: int = 600
    linewidth: float = 1.8
    output_format: str = "png"  # 可选：png / tif


def discover_result_files(results_dir: Path) -> List[Path]:
    """扫描结果目录并返回支持的数据文件列表。"""
    if not results_dir.exists() or not results_dir.is_dir():
        raise FileNotFoundError(f"结果目录不存在：{results_dir}")
    files = [p for p in results_dir.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS]
    return sorted(files)


def _extract_last_name_group(file_stem: str) -> str:
    """
    提取文件名中的“最后一组名称”作为分类名。

    说明：
    - 命名规则常见形式：..._A-0.5_z 或 ..._A-0.5_Cable_Axial_Force
    - 本函数优先提取“参数段（如 A-0.5）之后的整段名称”。
    - 若无法匹配该规则，则回退为最后一个下划线后的片段。
    - 若仍无法提取英文名称，则归类为 Uncategorized。
    """
    # 优先匹配：<参数字母>-<数字>_<名称段>
    # 示例：C3000000_A-0.6_Girder_Disp -> Girder_Disp
    m = re.search(r"[A-Za-z]-?\d+(?:\.\d+)?_(?P<name>[A-Za-z][A-Za-z0-9_]*)$", file_stem)
    if m:
        return m.group("name")

    # 回退：最后一个下划线后的片段
    if "_" in file_stem:
        tail = file_stem.rsplit("_", 1)[-1]
        if re.search(r"[A-Za-z]", tail):
            return re.sub(r"[^A-Za-z0-9_]+", "", tail) or "Uncategorized"

    # 再回退：尾部连续英文/数字/下划线段
    m2 = re.search(r"([A-Za-z][A-Za-z0-9_]*)$", file_stem)
    if m2:
        return m2.group(1)

    return "Uncategorized"


def _get_unique_target_path(target_dir: Path, file_name: str) -> Path:
    """若目标文件已存在，则自动追加序号避免覆盖。"""
    candidate = target_dir / file_name
    if not candidate.exists():
        return candidate

    stem = Path(file_name).stem
    suffix = Path(file_name).suffix
    idx = 1
    while True:
        candidate = target_dir / f"{stem}_{idx}{suffix}"
        if not candidate.exists():
            return candidate
        idx += 1


def classify_files_by_last_word(results_dir: Path) -> Tuple[int, int, List[str]]:
    """
    按文件名最后一组名称分类移动文件。

    Parameters
    ----------
    results_dir : Path
        待分类目录（仅处理该目录顶层文件，不递归处理子目录）。

    Returns
    -------
    moved_count : int
        成功移动文件数。
    total_count : int
        参与分类的文件总数。
    touched_categories : List[str]
        本次涉及的分类目录名列表。
    """
    files = discover_result_files(results_dir)
    total = len(files)
    moved = 0
    categories = set()

    for src in files:
        category = _extract_last_name_group(src.stem)
        category_dir = results_dir / category
        category_dir.mkdir(parents=True, exist_ok=True)
        categories.add(category)

        target = _get_unique_target_path(category_dir, src.name)
        try:
            shutil.move(str(src), str(target))
            moved += 1
        except Exception:
            # 单文件失败时不中断整个批次
            continue

    return moved, total, sorted(categories)


def _score_numeric_first_col(df: pd.DataFrame) -> float:
    """评估 DataFrame 首列可转换为数值的比例，用于判别读取方式。"""
    if df.empty or df.shape[1] < 2:
        return 0.0
    first_col = pd.to_numeric(df.iloc[:, 0], errors="coerce")
    return float(first_col.notna().mean())


def _read_table_auto(file_path: Path) -> pd.DataFrame:
    """自动识别分隔符并读取表格，兼容含/不含表头情况。"""
    # 方案 A：按“含表头”读取
    try:
        df_header = pd.read_csv(file_path, sep=None, engine="python", comment="#", header=0)
    except Exception:
        df_header = pd.DataFrame()

    # 方案 B：按“无表头”读取
    try:
        df_noheader = pd.read_csv(file_path, sep=None, engine="python", comment="#", header=None)
    except Exception:
        df_noheader = pd.DataFrame()

    # 选择首列数值比例更高的版本
    score_h = _score_numeric_first_col(df_header)
    score_nh = _score_numeric_first_col(df_noheader)

    if score_h == 0 and score_nh == 0:
        # 回退：尝试空白分隔
        try:
            df_noheader = pd.read_csv(file_path, sep=r"\s+", engine="python", comment="#", header=None)
            score_nh = _score_numeric_first_col(df_noheader)
        except Exception as exc:
            raise ValueError(f"无法识别文件格式：{file_path.name}") from exc

    df = df_noheader if score_nh > score_h else df_header
    if df.empty or df.shape[1] < 2:
        raise ValueError(f"文件列数不足（至少2列）：{file_path.name}")

    # 删除全空列，避免分隔符异常导致末尾空列
    df = df.dropna(axis=1, how="all")

    # 统一尝试数值化
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # 去掉关键列为空的行
    df = df.dropna(subset=[df.columns[0]])
    if df.empty:
        raise ValueError(f"文件数据为空或无法数值化：{file_path.name}")

    # 若原为无表头，自动命名列名
    if isinstance(df.columns[0], int):
        names = ["Time"] + [f"Col_{i}" for i in range(1, len(df.columns))]
        df.columns = names
    else:
        cols = list(df.columns)
        cols[0] = str(cols[0]) or "Time"
        df.columns = cols

    return df


def load_time_history(file_path: Path, data_col_idx: int) -> Tuple[pd.Series, pd.Series, str]:
    """
    读取单个文件中的时程数据。

    Parameters
    ----------
    file_path : Path
        结果文件路径。
    data_col_idx : int
        数据列索引（从1开始，1表示第2列；首列固定为时间列）。

    Returns
    -------
    time : pd.Series
        时间序列。
    data : pd.Series
        响应序列。
    data_col_name : str
        选中数据列名。
    """
    if not file_path.exists():
        raise FileNotFoundError(f"文件不存在：{file_path}")

    if data_col_idx < 1:
        raise IndexError("数据列索引必须 >= 1（首列是时间列）")

    df = _read_table_auto(file_path)
    target_idx = data_col_idx
    if target_idx >= df.shape[1]:
        raise IndexError(
            f"列索引越界：文件 {file_path.name} 共有 {df.shape[1]} 列，"
            f"可选数据列索引范围为 1 ~ {df.shape[1] - 1}"
        )

    time = pd.to_numeric(df.iloc[:, 0], errors="coerce")
    data = pd.to_numeric(df.iloc[:, target_idx], errors="coerce")

    valid = time.notna() & data.notna()
    time = time[valid].reset_index(drop=True)
    data = data[valid].reset_index(drop=True)

    if len(time) == 0:
        raise ValueError(f"文件无有效数值数据：{file_path.name}")

    return time, data, str(df.columns[target_idx])


def get_data_columns(file_path: Path) -> List[Tuple[int, str]]:
    """
    获取可用数据列信息（不包含第0列时间列）。

    Returns
    -------
    List[Tuple[int, str]]
        每个元素为 (数据列索引, 列名)，其中数据列索引从1开始。
    """
    if not file_path.exists():
        raise FileNotFoundError(f"文件不存在：{file_path}")
    df = _read_table_auto(file_path)
    if df.shape[1] < 2:
        raise ValueError(f"文件列数不足（至少2列）：{file_path.name}")
    return [(i, str(df.columns[i])) for i in range(1, df.shape[1])]


def extract_param_label(file_path: Path, regex_pattern: str) -> str:
    """根据文件名提取参数组标签，优先显示 C/A 参数换算结果。"""
    stem = file_path.stem

    # 优先解析文件名中的 C 与 A 参数：C 需要 /1000，A 需要 +1
    m_c = re.search(r"(?:^|_)C(-?\d+(?:\.\d+)?)", stem, flags=re.IGNORECASE)
    m_a = re.search(r"(?:^|_)A(-?\d+(?:\.\d+)?)", stem, flags=re.IGNORECASE)

    parts = []
    if m_c:
        c_val = float(m_c.group(1)) / 1000.0
        parts.append(f"C={c_val:g}")
    if m_a:
        a_val = float(m_a.group(1)) + 1.0
        parts.append(f"A={a_val:g}")
    if parts:
        return ", ".join(parts)

    if not regex_pattern.strip():
        return stem

    try:
        m = re.search(regex_pattern, stem)
    except re.error:
        return stem

    if not m:
        return stem

    return m.group(0)


def apply_journal_style(style_mode: str) -> None:
    """设置适合学术图件的 matplotlib 样式。"""
    plt.rcParams.update(
        {
            # 中文优先使用常见中文字体，英文数字回退到 Times/DejaVu
            "font.family": "sans-serif",
            "font.sans-serif": ["Microsoft YaHei", "SimHei", "Noto Sans CJK SC", "Arial Unicode MS", "DejaVu Sans"],
            "font.serif": ["Times New Roman", "DejaVu Serif"],
            "font.size": 11,
            "axes.labelsize": 12,
            "axes.titlesize": 13,
            "legend.fontsize": 10,
            "xtick.labelsize": 10,
            "ytick.labelsize": 10,
            "axes.linewidth": 1.0,
            "axes.unicode_minus": False,
            "savefig.bbox": "tight",
            "savefig.transparent": False,
        }
    )

    if style_mode == "黑白":
        # 黑白论文风格：灰度+线型区分
        plt.rcParams["axes.prop_cycle"] = cycler(color=["0.10", "0.25", "0.40", "0.55", "0.70", "0.85"])
    else:
        # 彩色论文风格：色盲友好配色（Okabe-Ito）
        plt.rcParams["axes.prop_cycle"] = cycler(
            color=["#0072B2", "#D55E00", "#009E73", "#CC79A7", "#E69F00", "#56B4E9", "#000000", "#F0E442"]
        )


def plot_time_history(
    files: List[Path],
    data_col_idx: int,
    plot_cfg: PlotConfig,
    legend_names: Optional[List[str]] = None,
    param_regex: str = r"(?:A|C|zeta|xi|damp|alpha|beta|n|vel|v)[-_]?[0-9]+(?:\.[0-9]+)?",
) -> plt.Figure:
    """绘制多文件时程对比曲线并返回 Figure。"""
    if not files:
        raise ValueError("未选择任何数据文件")

    apply_journal_style(plot_cfg.style_mode)
    fig, ax = plt.subplots(figsize=(8.0, 4.8), dpi=120)
    # 线型与标记组合，保证多条曲线在彩色/黑白下均可清晰区分
    linestyle_pool = ["-", "--", "-.", ":", (0, (3, 1, 1, 1)), (0, (5, 1))]
    marker_pool = [None, "o", "s", "^", "D", "v", "x", "*"]

    used_col_name = None
    for i, fp in enumerate(files):
        time, data, col_name = load_time_history(fp, data_col_idx)
        if used_col_name is None:
            used_col_name = col_name

        if legend_names and i < len(legend_names) and legend_names[i].strip():
            label = legend_names[i].strip()
        else:
            label = extract_param_label(fp, param_regex)

        linestyle = linestyle_pool[i % len(linestyle_pool)]
        marker = marker_pool[i % len(marker_pool)]
        # 控制标记密度，避免点过密影响可读性
        markevery = max(1, len(time) // 25) if marker is not None else None
        ax.plot(
            time,
            data,
            label=label,
            linewidth=plot_cfg.linewidth,
            linestyle=linestyle,
            marker=marker,
            markersize=4.2 if marker is not None else 0.0,
            markevery=markevery,
        )

    ax.set_xlabel(plot_cfg.xlabel)
    ax.set_ylabel(plot_cfg.ylabel)
    ax.set_title(plot_cfg.title)
    ax.grid(True, which="major", linestyle="--", linewidth=0.6, alpha=0.5)
    legend = ax.legend(frameon=True, framealpha=1.0, edgecolor="black", loc="best")
    enable_legend_interaction(fig, legend)

    # 标注所用物理列名，方便核查“固定物理量/节点号”的一致性
    if used_col_name:
        ax.text(
            0.99,
            0.02,
            f"Data Column: {used_col_name}",
            transform=ax.transAxes,
            ha="right",
            va="bottom",
            fontsize=9,
            color="0.25",
        )

    fig.tight_layout()
    return fig


def enable_legend_interaction(fig: plt.Figure, legend) -> None:
    """
    启用图例交互：
    1) 鼠标拖动图例位置；
    2) 鼠标滚轮在图例区域缩放图例字号。
    """
    if legend is None:
        return

    legend.set_draggable(True)
    state = {"fontsize": float(plt.rcParams.get("legend.fontsize", 10))}

    def _on_scroll(event):
        if event.inaxes is None:
            return
        if event.x is None or event.y is None:
            return
        contains, _ = legend.contains(event)
        if not contains:
            return

        step = 0.8 if event.button == "up" else -0.8
        new_size = max(6.0, min(28.0, state["fontsize"] + step))
        if abs(new_size - state["fontsize"]) < 1e-9:
            return
        state["fontsize"] = new_size

        for txt in legend.get_texts():
            txt.set_fontsize(new_size)
        legend.get_title().set_fontsize(new_size)
        fig.canvas.draw_idle()

    fig.canvas.mpl_connect("scroll_event", _on_scroll)


def _pct_diff(ref: float, val: float) -> str:
    """计算相对百分比差异（相对 ref），返回字符串。"""
    if abs(ref) < 1e-12:
        return "N/A(ref=0)"
    return f"{(val - ref) / abs(ref) * 100.0:+.2f}%"


def build_output_path(results_dir: Path, title: str, ext: str, custom_output_dir: str = "") -> Path:
    """生成输出文件路径（默认 results 同级 plots，支持自定义目录）。"""
    if custom_output_dir.strip():
        plots_dir = Path(custom_output_dir.strip())
    else:
        plots_dir = results_dir.parent / "plots"
    plots_dir.mkdir(parents=True, exist_ok=True)

    safe_title = re.sub(r"[^\w\-\u4e00-\u9fff]+", "_", title).strip("_") or "time_history"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return plots_dir / f"{safe_title}_{stamp}.{ext.lower()}"


class TimeHistoryPlotApp:
    """Tkinter 交互式时程绘图界面。"""
    AUTO_BASELINE_LABEL = "自动（第一条选中曲线）"

    def __init__(self, master: tk.Tk) -> None:
        self.master = master
        self.master.title("PyAnsys 时程曲线绘图工具")
        self.master.geometry("900x650")
        # 提升中文界面可读性，避免部分系统默认字体导致显示异常
        self._set_ui_font()

        self.results_dir_var = tk.StringVar(value=DEFAULT_RESULTS_DIR)
        self.col_select_var = tk.StringVar(value="1: Col_1")
        self.xlabel_var = tk.StringVar(value="时间 / s")
        self.ylabel_var = tk.StringVar(value="响应量 / 单位")
        self.title_var = tk.StringVar(value="时程曲线对比")
        self.style_var = tk.StringVar(value="彩色")
        self.dpi_var = tk.StringVar(value="600")
        self.linewidth_var = tk.StringVar(value="1.8")
        self.format_var = tk.StringVar(value="png")
        self.output_dir_var = tk.StringVar(value="")
        self.baseline_file_var = tk.StringVar(value=self.AUTO_BASELINE_LABEL)
        self.regex_var = tk.StringVar(
            value=r"(?:A|C|zeta|xi|damp|alpha|beta|n|vel|v)[-_]?[0-9]+(?:\.[0-9]+)?"
        )
        self.manual_legend_var = tk.StringVar(value="")

        self.file_paths: List[Path] = []
        self.col_options: List[Tuple[int, str]] = []
        self._build_ui()
        self.refresh_file_list()

    def _set_ui_font(self) -> None:
        """设置 Tkinter/ttk 字体，减少中文乱码概率。"""
        try:
            available = set(tkfont.families(self.master))
            candidates = ["Microsoft YaHei UI", "Microsoft YaHei", "SimHei", "Segoe UI", "Arial"]
            family = next((f for f in candidates if f in available), None)
            if not family:
                return

            # 统一配置 Tk 命名字体，避免 ttk 对带空格字体名解析异常
            named_fonts = [
                "TkDefaultFont",
                "TkTextFont",
                "TkMenuFont",
                "TkHeadingFont",
                "TkCaptionFont",
                "TkSmallCaptionFont",
                "TkIconFont",
                "TkTooltipFont",
            ]
            for name in named_fonts:
                try:
                    f = tkfont.nametofont(name)
                    f.configure(family=family, size=10)
                except tk.TclError:
                    pass

            style = ttk.Style(self.master)
            # 使用 Tcl 字符串字体描述，确保含空格字体名被正确解析
            style.configure(".", font=f"{{{family}}} 10")
        except Exception:
            # 字体不可用时保持默认，不阻断程序运行
            pass

    def _build_ui(self) -> None:
        """构建界面控件。"""
        frm_top = ttk.Frame(self.master, padding=10)
        frm_top.pack(fill="x")

        ttk.Label(frm_top, text="results 目录:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm_top, textvariable=self.results_dir_var, width=80).grid(row=0, column=1, padx=6, sticky="ew")
        ttk.Button(frm_top, text="浏览", command=self.select_dir).grid(row=0, column=2, padx=4)
        ttk.Button(frm_top, text="刷新文件", command=self.refresh_file_list).grid(row=0, column=3, padx=4)
        ttk.Button(frm_top, text="文件分类", command=self.classify_files).grid(row=0, column=4, padx=4)
        frm_top.columnconfigure(1, weight=1)

        frm_files = ttk.LabelFrame(self.master, text="结果文件（可多选，用于参数组对比）", padding=10)
        frm_files.pack(fill="both", expand=True, padx=10, pady=6)

        self.file_listbox = tk.Listbox(frm_files, selectmode=tk.EXTENDED, height=12, exportselection=False)
        self.file_listbox.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(frm_files, orient="vertical", command=self.file_listbox.yview)
        scroll.pack(side="right", fill="y")
        self.file_listbox.config(yscrollcommand=scroll.set)
        self.file_listbox.bind("<<ListboxSelect>>", self._on_file_selection_changed)

        frm_cfg = ttk.LabelFrame(self.master, text="绘图设置", padding=10)
        frm_cfg.pack(fill="x", padx=10, pady=6)

        ttk.Label(frm_cfg, text="数据列切换:").grid(row=0, column=0, sticky="w")
        self.col_combo = ttk.Combobox(frm_cfg, textvariable=self.col_select_var, width=30, state="readonly")
        self.col_combo.grid(row=0, column=1, padx=4, sticky="w")
        ttk.Button(frm_cfg, text="刷新列", command=self.refresh_columns).grid(row=0, column=2, padx=4, sticky="w")
        ttk.Button(frm_cfg, text="上一列", command=self.prev_column).grid(row=0, column=3, padx=4, sticky="w")
        ttk.Button(frm_cfg, text="下一列", command=self.next_column).grid(row=0, column=4, padx=4, sticky="w")

        ttk.Label(frm_cfg, text="X 标签:").grid(row=1, column=0, sticky="e")
        ttk.Entry(frm_cfg, textvariable=self.xlabel_var, width=22).grid(row=1, column=1, padx=4)

        ttk.Label(frm_cfg, text="Y 标签:").grid(row=1, column=2, sticky="e")
        ttk.Entry(frm_cfg, textvariable=self.ylabel_var, width=22).grid(row=1, column=3, padx=4, columnspan=2, sticky="we")

        ttk.Label(frm_cfg, text="标题:").grid(row=2, column=0, sticky="w")
        ttk.Entry(frm_cfg, textvariable=self.title_var, width=28).grid(row=2, column=1, columnspan=2, padx=4, sticky="we")

        ttk.Label(frm_cfg, text="风格:").grid(row=2, column=3, sticky="e")
        ttk.Combobox(frm_cfg, textvariable=self.style_var, values=["彩色", "黑白"], width=8, state="readonly").grid(
            row=2, column=4, sticky="w"
        )

        ttk.Label(frm_cfg, text="线宽:").grid(row=2, column=5, sticky="e")
        ttk.Entry(frm_cfg, textvariable=self.linewidth_var, width=8).grid(row=2, column=6, padx=4, sticky="w")

        ttk.Label(frm_cfg, text="DPI:").grid(row=3, column=0, sticky="w")
        ttk.Entry(frm_cfg, textvariable=self.dpi_var, width=10).grid(row=3, column=1, padx=4, sticky="w")

        ttk.Label(frm_cfg, text="保存格式:").grid(row=3, column=2, sticky="e")
        ttk.Combobox(frm_cfg, textvariable=self.format_var, values=["png", "tif"], width=8, state="readonly").grid(
            row=3, column=3, sticky="w"
        )

        ttk.Label(frm_cfg, text="输出目录(可选):").grid(row=4, column=0, sticky="w")
        ttk.Entry(frm_cfg, textvariable=self.output_dir_var, width=60).grid(row=4, column=1, columnspan=4, padx=4, sticky="we")
        ttk.Button(frm_cfg, text="浏览输出目录", command=self.select_output_dir).grid(row=4, column=5, columnspan=2, padx=4, sticky="we")

        ttk.Label(frm_cfg, text="参数提取正则(备用):").grid(row=5, column=0, sticky="w")
        ttk.Entry(frm_cfg, textvariable=self.regex_var, width=80).grid(row=5, column=1, columnspan=6, padx=4, sticky="we")

        ttk.Label(frm_cfg, text="手动图例(逗号分隔，可选):").grid(row=6, column=0, sticky="w")
        ttk.Entry(frm_cfg, textvariable=self.manual_legend_var, width=80).grid(row=6, column=1, columnspan=6, padx=4, sticky="we")

        ttk.Label(frm_cfg, text="统计基准文件:").grid(row=7, column=0, sticky="w")
        self.baseline_combo = ttk.Combobox(frm_cfg, textvariable=self.baseline_file_var, width=45, state="readonly")
        self.baseline_combo.grid(row=7, column=1, columnspan=4, padx=4, sticky="we")
        ttk.Button(frm_cfg, text="刷新基准选项", command=self.refresh_baseline_options).grid(row=7, column=5, columnspan=2, padx=4, sticky="we")

        for c in range(7):
            frm_cfg.columnconfigure(c, weight=1 if c in (1, 2, 3, 4, 5, 6) else 0)

        frm_btn = ttk.Frame(self.master, padding=10)
        frm_btn.pack(fill="x")

        ttk.Button(frm_btn, text="预览绘图", command=self.preview_plot).pack(side="left", padx=6)
        ttk.Button(frm_btn, text="统计对比", command=self.show_statistics).pack(side="left", padx=6)
        ttk.Button(frm_btn, text="绘图并保存", command=self.save_plot).pack(side="left", padx=6)
        ttk.Button(frm_btn, text="退出", command=self.master.destroy).pack(side="right", padx=6)

    def select_dir(self) -> None:
        """选择 results 目录。"""
        folder = filedialog.askdirectory(initialdir=self.results_dir_var.get() or str(Path.cwd()))
        if folder:
            self.results_dir_var.set(folder)
            self.refresh_file_list()

    def select_output_dir(self) -> None:
        """选择绘图输出目录。"""
        initial = self.output_dir_var.get().strip() or self.results_dir_var.get().strip() or str(Path.cwd())
        folder = filedialog.askdirectory(initialdir=initial)
        if folder:
            self.output_dir_var.set(folder)

    def refresh_file_list(self) -> None:
        """刷新文件列表。"""
        self.file_listbox.delete(0, tk.END)
        self.file_paths.clear()

        try:
            files = discover_result_files(Path(self.results_dir_var.get()))
        except Exception as exc:
            messagebox.showerror("目录错误", str(exc))
            return

        if not files:
            messagebox.showwarning("提示", "目录中未找到 .csv/.txt/.dat 文件")
            return

        self.file_paths = files
        for fp in files:
            self.file_listbox.insert(tk.END, fp.name)
        self.refresh_baseline_options()

        # 默认选中第一个文件，便于初始化列下拉
        if self.file_paths:
            self.file_listbox.selection_set(0)
            self._on_file_selection_changed()

    def _on_file_selection_changed(self, _event=None) -> None:
        """文件列表选择变化时，同步刷新列与基准选项。"""
        self.refresh_columns()
        self.refresh_baseline_options()

    def refresh_baseline_options(self) -> None:
        """刷新“统计基准文件”下拉选项。"""
        values = [self.AUTO_BASELINE_LABEL] + [fp.name for fp in self.file_paths]
        self.baseline_combo["values"] = values
        cur = self.baseline_file_var.get().strip()
        if cur not in values:
            self.baseline_file_var.set(self.AUTO_BASELINE_LABEL)

    def classify_files(self) -> None:
        """按文件名最后一组名称建立文件夹分类并移动文件。"""
        try:
            results_dir = Path(self.results_dir_var.get())
            if not results_dir.exists() or not results_dir.is_dir():
                raise FileNotFoundError(f"目录不存在：{results_dir}")

            files = discover_result_files(results_dir)
            if not files:
                messagebox.showwarning("提示", "当前目录无可分类的 .csv/.txt/.dat 顶层文件")
                return

            do_it = messagebox.askyesno(
                "确认分类",
                "将按文件名最后一组名称建立子文件夹并移动文件，是否继续？",
            )
            if not do_it:
                return

            moved, total, cats = classify_files_by_last_word(results_dir)
            self.refresh_file_list()
            cats_text = ", ".join(cats[:10])
            if len(cats) > 10:
                cats_text += " ..."
            messagebox.showinfo(
                "分类完成",
                f"共处理 {total} 个文件，成功移动 {moved} 个。\n"
                f"分类目录数：{len(cats)}\n"
                f"示例分类：{cats_text if cats_text else '无'}",
            )
        except Exception as exc:
            messagebox.showerror("分类失败", str(exc))

    def _parse_col_idx_from_combo(self) -> int:
        """从下拉文本中提取数据列索引。"""
        text = self.col_select_var.get().strip()
        m = re.match(r"^\s*(\d+)\s*:", text)
        if m:
            return int(m.group(1))
        # 兼容用户手动输入纯数字
        if text.isdigit():
            return int(text)
        raise ValueError("请选择有效的数据列")

    def refresh_columns(self) -> None:
        """根据当前选中文件刷新可用列。"""
        try:
            files = self._get_selected_files()
        except Exception:
            # 未选择文件时不弹窗，避免干扰
            self.col_options = []
            self.col_combo["values"] = []
            self.col_select_var.set("")
            return

        try:
            self.col_options = get_data_columns(files[0])
        except Exception as exc:
            messagebox.showerror("列读取失败", str(exc))
            self.col_options = []
            self.col_combo["values"] = []
            self.col_select_var.set("")
            return

        combo_vals = [f"{idx}: {name}" for idx, name in self.col_options]
        self.col_combo["values"] = combo_vals

        current = self.col_select_var.get().strip()
        if current not in combo_vals and combo_vals:
            self.col_select_var.set(combo_vals[0])

    def prev_column(self) -> None:
        """切换到上一数据列。"""
        vals = list(self.col_combo["values"])
        if not vals:
            return
        cur = self.col_select_var.get().strip()
        try:
            i = vals.index(cur)
        except ValueError:
            i = 0
        self.col_select_var.set(vals[(i - 1) % len(vals)])

    def next_column(self) -> None:
        """切换到下一数据列。"""
        vals = list(self.col_combo["values"])
        if not vals:
            return
        cur = self.col_select_var.get().strip()
        try:
            i = vals.index(cur)
        except ValueError:
            i = 0
        self.col_select_var.set(vals[(i + 1) % len(vals)])

    def _get_selected_files(self) -> List[Path]:
        """获取用户在列表中选择的文件。"""
        idxs = list(self.file_listbox.curselection())
        if not idxs:
            raise ValueError("请至少选择一个文件")
        return [self.file_paths[i] for i in idxs]

    def _collect_plot_inputs(self) -> Tuple[List[Path], int, PlotConfig, List[str], str, str]:
        """收集并校验输入参数。"""
        files = self._get_selected_files()

        try:
            col_idx = self._parse_col_idx_from_combo()
        except ValueError as exc:
            raise ValueError("请选择有效的数据列") from exc

        try:
            dpi_val = int(self.dpi_var.get())
        except ValueError as exc:
            raise ValueError("DPI 必须为整数") from exc

        try:
            lw_val = float(self.linewidth_var.get())
        except ValueError as exc:
            raise ValueError("线宽必须为数字") from exc

        cfg = PlotConfig(
            xlabel=self.xlabel_var.get().strip() or "时间 / s",
            ylabel=self.ylabel_var.get().strip() or "响应量 / 单位",
            title=self.title_var.get().strip() or "时程曲线对比",
            style_mode=self.style_var.get().strip() or "彩色",
            dpi=dpi_val,
            linewidth=lw_val,
            output_format=self.format_var.get().strip().lower() or "png",
        )

        manual_legend = [x.strip() for x in self.manual_legend_var.get().split(",") if x.strip()]
        regex = self.regex_var.get().strip()
        output_dir = self.output_dir_var.get().strip()

        return files, col_idx, cfg, manual_legend, regex, output_dir

    def preview_plot(self) -> None:
        """预览绘图（不保存）。"""
        try:
            files, col_idx, cfg, manual_legend, regex, _output_dir = self._collect_plot_inputs()
            fig = plot_time_history(
                files=files,
                data_col_idx=col_idx,
                plot_cfg=cfg,
                legend_names=manual_legend,
                param_regex=regex,
            )
            plt.show()
            plt.close(fig)
        except Exception as exc:
            messagebox.showerror("绘图失败", str(exc))

    def show_statistics(self) -> None:
        """统计当前选定列在不同曲线中的最大值/最小值及百分比差异。"""
        try:
            files = self._get_selected_files()
            col_idx = self._parse_col_idx_from_combo()
            baseline_name = self.baseline_file_var.get().strip()

            # 自由基准文件：可选任意文件；默认第一条选中曲线
            if not baseline_name or baseline_name == self.AUTO_BASELINE_LABEL:
                baseline_fp = files[0]
            else:
                baseline_fp = next((p for p in self.file_paths if p.name == baseline_name), None)
                if baseline_fp is None:
                    raise ValueError(f"基准文件不存在：{baseline_name}")

            lines = []
            stats = []
            col_name = ""

            def _calc_stat(fp: Path):
                t, y, c_name = load_time_history(fp, col_idx)
                nonlocal col_name
                col_name = c_name
                i_max = int(y.idxmax())
                i_min = int(y.idxmin())
                y_max, t_max = float(y.iloc[i_max]), float(t.iloc[i_max])
                y_min, t_min = float(y.iloc[i_min]), float(t.iloc[i_min])
                p2p = y_max - y_min
                return {"name": fp.name, "col": c_name, "y_max": y_max, "t_max": t_max, "y_min": y_min, "t_min": t_min, "p2p": p2p}

            ref = _calc_stat(baseline_fp)
            for fp in files:
                stats.append(_calc_stat(fp))

            if not stats:
                raise ValueError("没有可统计的数据")

            lines.append("=== 时程统计结果 ===")
            lines.append(f"统计列: {col_name or ref['col']}")
            lines.append(f"基准曲线: {ref['name']}")
            lines.append("")

            for s in stats:
                lines.append(f"[{s['name']}]")
                lines.append(f"  最大值: {s['y_max']:.6g}  (t={s['t_max']:.6g}s)")
                lines.append(f"  最小值: {s['y_min']:.6g}  (t={s['t_min']:.6g}s)")
                lines.append(f"  峰-峰值: {s['p2p']:.6g}")
                if s["name"] != ref["name"]:
                    lines.append(f"  相对基准最大值差异: {_pct_diff(ref['y_max'], s['y_max'])}")
                    lines.append(f"  相对基准最小值差异: {_pct_diff(ref['y_min'], s['y_min'])}")
                    lines.append(f"  相对基准峰-峰值差异: {_pct_diff(ref['p2p'], s['p2p'])}")
                lines.append("")

            self._show_text_window("统计对比结果", "\n".join(lines))
        except Exception as exc:
            messagebox.showerror("统计失败", str(exc))

    def _show_text_window(self, title: str, content: str) -> None:
        """使用可滚动文本窗口展示统计结果。"""
        win = tk.Toplevel(self.master)
        win.title(title)
        win.geometry("820x560")
        txt = scrolledtext.ScrolledText(win, wrap=tk.WORD, font=("Consolas", 10))
        txt.pack(fill="both", expand=True, padx=8, pady=8)
        txt.insert("1.0", content)
        txt.configure(state="disabled")

    def save_plot(self) -> None:
        """绘图并保存到同级 plots 目录（支持交互调整图例后保存）。"""
        try:
            files, col_idx, cfg, manual_legend, regex, output_dir = self._collect_plot_inputs()
            fig = plot_time_history(
                files=files,
                data_col_idx=col_idx,
                plot_cfg=cfg,
                legend_names=manual_legend,
                param_regex=regex,
            )

            out_path = build_output_path(
                Path(self.results_dir_var.get()),
                cfg.title,
                cfg.output_format,
                custom_output_dir=output_dir,
            )
            messagebox.showinfo(
                "提示",
                "将在图窗中进行交互调整：\n"
                "1) 鼠标拖动图例可移动位置；\n"
                "2) 鼠标放在图例上滚轮可缩放图例大小；\n"
                "关闭图窗后将按当前状态保存图片。",
            )
            plt.show()
            fig.savefig(out_path, dpi=cfg.dpi, format=cfg.output_format)
            plt.close(fig)

            messagebox.showinfo("保存成功", f"图片已保存：\n{out_path}")
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc))


def main() -> None:
    """程序入口。"""
    root = tk.Tk()
    app = TimeHistoryPlotApp(root)
    root.mainloop()


if __name__ == "__main__":
    # Example:
    # 1) 直接运行图形界面：
    #    python plot_time_history.py
    #
    # 2) 如果你希望在脚本中复用核心函数，可参考：
    #    from pathlib import Path
    #    from plot_time_history import plot_time_history, PlotConfig
    #    files = [
    #        Path(r"D:\pyansys\Claude_pyansys\results\C3000000_A-0.6_Girder_Disp.csv"),
    #        Path(r"D:\pyansys\Claude_pyansys\results\opt_C1000000_A-0.5_Cable_Axial_Force.csv"),
    #    ]
    #    fig = plot_time_history(files, data_col_idx=1, plot_cfg=PlotConfig(title="参数组对比"))
    #    fig.show()
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(0)
