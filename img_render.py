"""
img_render.py — 圖像渲染引擎 v4
Medieval Fantasy / Classical European Isekai Style
羊皮紙底、哥德裝飾、金色紋章
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager, rcParams
from matplotlib.patches import FancyBboxPatch, Rectangle
from matplotlib.collections import LineCollection
from io import BytesIO
from typing import List, Optional, Tuple, Dict
from datetime import datetime
import textwrap, math
import numpy as np

# ========== 字體設定 ==========
import platform, os

_FONT_PATHS = [
    '/System/Library/Fonts/PingFang.ttc',
    '/System/Library/Fonts/Supplemental/PingFang.ttc',
    'C:/Windows/Fonts/mingliu.ttc',
    '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
]
CJK_FONT_PATH = None
for p in _FONT_PATHS:
    if os.path.exists(p):
        CJK_FONT_PATH = p; break

if CJK_FONT_PATH:
    font_manager.fontManager.addfont(CJK_FONT_PATH)
    _fp = font_manager.FontProperties(fname=CJK_FONT_PATH)
    CJK_FONT = _fp.get_name()
else:
    _FALLBACK = ['PingFang TC','PingFang SC','Heiti TC','Hiragino Sans',
                 'Noto Sans CJK TC','Microsoft JhengHei']
    available = [f.name for f in font_manager.fontManager.ttflist]
    CJK_FONT = next((f for f in _FALLBACK if f in available), 'sans-serif')

# 嘗試找襯線字體做標題
_SERIF = ['Georgia','Palatino','Garamond','Times New Roman',
          'DejaVu Serif','Liberation Serif','serif']
available = [f.name for f in font_manager.fontManager.ttflist]
SERIF_FONT = next((f for f in _SERIF if f in available), 'DejaVu Serif')
LAT_FONT = SERIF_FONT

rcParams['font.sans-serif'] = [CJK_FONT, SERIF_FONT, 'DejaVu Sans']
rcParams['font.serif'] = [CJK_FONT, SERIF_FONT, 'DejaVu Serif']
rcParams['font.family'] = 'sans-serif'
rcParams['axes.unicode_minus'] = False
print(f"[img_render] CJK: {CJK_FONT} | Serif: {SERIF_FONT}")


# ========== 中世紀色彩主題 ==========
class Theme:
    # 羊皮紙色系
    PARCHMENT     = '#F5E6C8'    # 主羊皮紙
    PARCHMENT_D   = '#E8D5A8'    # 深羊皮紙
    PARCHMENT_DD  = '#D4BC82'    # 更深
    PARCHMENT_L   = '#FBF3E4'    # 淺
    BG            = '#F5E6C8'

    # 中世紀墨色
    INK           = '#2C1810'    # 深棕墨
    INK_LIGHT     = '#5C3D2E'    # 淺棕
    INK_FADED     = '#8B7355'    # 褪色墨
    TEXT          = '#2C1810'
    TEXT_LIGHT    = '#8B7355'
    TEXT_SEC      = '#8B7355'
    TEXT_DIM      = '#A89070'

    # 金屬色
    GOLD          = '#C5973B'
    GOLD_BRIGHT   = '#D4A84B'
    GOLD_DARK     = '#8B6914'
    SILVER        = '#8A8A8A'
    BRONZE        = '#CD7F32'
    COPPER        = '#B87333'

    # 紋章色
    ROYAL_BLUE    = '#1E3A5F'
    ROYAL_RED     = '#8B1A1A'
    HERALDIC_RED  = '#A52A2A'
    FOREST_GREEN  = '#2D5A27'
    DEEP_PURPLE   = '#4A1A6B'

    # 功能色 (保持與 bot.py 兼容)
    HEADER_BG     = '#3B2415'
    HEADER_FG     = '#F5E6C8'
    ACCENT        = '#C5973B'
    ACCENT2       = '#1E3A5F'
    BLUE          = '#1E3A5F'
    PURPLE        = '#4A1A6B'
    GREEN         = '#2D5A27'
    RED           = '#8B1A1A'
    ORANGE        = '#B87333'
    PINK          = '#A0526B'
    CYAN          = '#2E6B6B'

    # 表格
    ROW_EVEN      = '#F5E6C8'
    ROW_ODD       = '#EDD9B5'
    BORDER        = '#B8956A'
    DIVIDER       = '#D4BC82'
    CARD_BG       = '#EDD9B5'
    SURFACE       = '#E0C898'


# ========== 裝飾繪圖工具 ==========

def _draw_parchment_bg(fig, ax):
    """繪製羊皮紙紋理背景"""
    fig.set_facecolor(Theme.PARCHMENT)

def _draw_ornate_border(ax, x=0.02, y=0.02, w=0.96, h=0.96, color=None):
    """繪製中世紀裝飾邊框 (雙線+角飾)"""
    if color is None: color = Theme.GOLD_DARK
    # 外框
    ax.add_patch(FancyBboxPatch((x, y), w, h,
        boxstyle="round,pad=0,rounding_size=0.008",
        facecolor='none', edgecolor=color, linewidth=2.0,
        transform=ax.transAxes, clip_on=False))
    # 內框
    m = 0.012
    ax.add_patch(FancyBboxPatch((x+m, y+m), w-2*m, h-2*m,
        boxstyle="round,pad=0,rounding_size=0.006",
        facecolor='none', edgecolor=color, linewidth=0.6, alpha=0.35,
        transform=ax.transAxes, clip_on=False))
    # 四角小菱形
    cs = 0.012
    for cx, cy in [(x+m+0.008, y+m+0.008), (x+w-m-0.008, y+m+0.008),
                   (x+m+0.008, y+h-m-0.008), (x+w-m-0.008, y+h-m-0.008)]:
        ax.fill([cx, cx+cs, cx, cx-cs, cx],
                [cy+cs, cy, cy-cs, cy, cy+cs],
                color=color, alpha=0.4, transform=ax.transAxes, clip_on=False)

def _draw_section_divider(ax, y, x1=0.08, x2=0.92, color=None):
    """繪製帶裝飾的分隔線"""
    if color is None: color = Theme.GOLD_DARK
    ax.plot([x1, x2], [y, y], color=color, linewidth=0.8, alpha=0.5,
            transform=ax.transAxes, clip_on=False)
    # 中央菱形裝飾
    cx = (x1 + x2) / 2
    s = 0.008
    ax.fill([cx, cx+s*1.5, cx, cx-s*1.5, cx],
            [y+s, y, y-s, y, y+s],
            color=color, alpha=0.6, transform=ax.transAxes, clip_on=False)

def _draw_title_banner(ax, y, title, fontsize=22, color=None):
    """繪製帶裝飾的標題"""
    if color is None: color = Theme.INK
    # 左右裝飾線
    ax.plot([0.06, 0.30], [y-0.015, y-0.015], color=Theme.GOLD_DARK,
            linewidth=1.2, alpha=0.5, transform=ax.transAxes, clip_on=False)
    ax.plot([0.70, 0.94], [y-0.015, y-0.015], color=Theme.GOLD_DARK,
            linewidth=1.2, alpha=0.5, transform=ax.transAxes, clip_on=False)
    # 標題文字
    ax.text(0.5, y, title, fontsize=fontsize, fontweight='bold',
            color=color, ha='center', va='top', transform=ax.transAxes,
            fontfamily=CJK_FONT)

def _watermark(ax, text="omega"):
    ax.text(0.94, 0.035, text, fontsize=8, color=Theme.INK_FADED,
            ha='right', va='bottom', transform=ax.transAxes,
            fontfamily=SERIF_FONT, alpha=0.3, style='italic')

def _timestamp(ax, y=0.015):
    ax.text(0.5, y, datetime.now().strftime('%Y-%m-%d %H:%M'),
            fontsize=7.5, color=Theme.TEXT_DIM, ha='center', transform=ax.transAxes,
            fontfamily=SERIF_FONT, alpha=0.5, style='italic')


# ========== 通用圖像表格 ==========
def render_table_image(
    title: str,
    subtitle: str,
    headers: List[str],
    rows: List[List[str]],
    col_widths: List[float] = None,
    col_colors: Dict[int, str] = None,
    header_colors: List[str] = None,
    row_highlights: Dict[int, str] = None,
    footer: str = None,
    figsize: Tuple[float, float] = None,
    dpi: int = 150,
) -> BytesIO:
    n_cols = len(headers)
    n_rows = len(rows)
    if figsize is None:
        figsize = (max(10, n_cols * 2.0), max(5, 3.0 + n_rows * 0.52))
    if col_widths is None:
        col_widths = [1.0 / n_cols] * n_cols

    fig, ax = plt.subplots(figsize=figsize)
    ax.set_xlim(0, 1); ax.set_ylim(0, 1)
    ax.axis('off')
    _draw_parchment_bg(fig, ax)
    _draw_ornate_border(ax)

    # 標題
    ax.text(0.5, 0.975, title, fontsize=21, fontweight='bold',
            color=Theme.INK, ha='center', va='top', transform=ax.transAxes,
            fontfamily=CJK_FONT)
    # 標題裝飾線
    ax.plot([0.08, 0.32], [0.96, 0.96], color=Theme.GOLD_DARK,
            linewidth=1, alpha=0.4, transform=ax.transAxes, clip_on=False)
    ax.plot([0.68, 0.92], [0.96, 0.96], color=Theme.GOLD_DARK,
            linewidth=1, alpha=0.4, transform=ax.transAxes, clip_on=False)
    sub_y = 0.952
    if subtitle:
        ax.text(0.5, sub_y, subtitle, fontsize=10, color=Theme.INK_FADED,
                ha='center', va='top', transform=ax.transAxes, fontfamily=CJK_FONT,
                style='italic')

    # 表格 — 用 bbox 定位在標題下方
    table_top = 0.93
    table_bot = 0.06
    table = ax.table(cellText=rows, colLabels=headers,
                     cellLoc='center', colWidths=col_widths,
                     bbox=[0.02, table_bot, 0.96, table_top - table_bot])
    table.auto_set_font_size(False)
    table.set_fontsize(9.5)

    # 表頭
    for j in range(n_cols):
        cell = table[(0, j)]
        bg = Theme.HEADER_BG
        if header_colors and j < len(header_colors):
            bg = header_colors[j]
        cell.set_facecolor(bg)
        cell.set_text_props(color=Theme.GOLD_BRIGHT, fontweight='bold',
                            fontsize=10, fontfamily=CJK_FONT)
        cell.set_edgecolor(Theme.GOLD_DARK)
        cell.set_linewidth(0.8)

    # 資料列
    for i in range(n_rows):
        for j in range(n_cols):
            cell = table[(i + 1, j)]
            if row_highlights and i in row_highlights:
                cell.set_facecolor(row_highlights[i])
            else:
                cell.set_facecolor(Theme.ROW_EVEN if i % 2 == 0 else Theme.ROW_ODD)
            color = Theme.INK
            if col_colors and j in col_colors:
                color = col_colors[j]
            cell.set_text_props(color=color, fontfamily=CJK_FONT, fontsize=9.5)
            cell.set_edgecolor(Theme.BORDER)
            cell.set_linewidth(0.3)

    if footer:
        ax.text(0.5, 0.025, footer, fontsize=8, color=Theme.INK_FADED,
                ha='center', va='bottom', transform=ax.transAxes,
                fontfamily=CJK_FONT, style='italic')

    _watermark(ax)
    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png', facecolor=fig.get_facecolor(),
                bbox_inches='tight', dpi=dpi)
    plt.close(fig)
    buf.seek(0)
    return buf


# ========== 資訊卡片 ==========
def render_info_card(
    title: str,
    fields: List[Tuple[str, str]],
    accent_color: str = None,
    footer: str = None,
    figsize: Tuple[float, float] = (8, None),
) -> BytesIO:
    if accent_color is None: accent_color = Theme.GOLD
    n = len(fields)
    h = figsize[1] if figsize[1] else max(3.5, 1.8 + n * 0.50)
    fig, ax = plt.subplots(figsize=(figsize[0], h))
    ax.axis('off')
    _draw_parchment_bg(fig, ax)
    _draw_ornate_border(ax)

    # 標題
    _draw_title_banner(ax, 0.90, title, fontsize=19, color=Theme.INK)

    # 欄位
    y = 0.78
    dy = min(0.10, 0.66 / max(n, 1))
    for label, value in fields:
        ax.text(0.12, y, label, fontsize=11.5, color=Theme.INK_FADED,
                transform=ax.transAxes, fontfamily=CJK_FONT)
        ax.text(0.50, y, str(value), fontsize=12, color=Theme.INK,
                transform=ax.transAxes, fontweight='bold', fontfamily=CJK_FONT)
        # 點線分隔
        ax.plot([0.10, 0.90], [y - dy*0.38, y - dy*0.38], color=Theme.GOLD_DARK,
                linewidth=0.4, transform=ax.transAxes, clip_on=False,
                linestyle=(0, (2, 4)), alpha=0.4)
        y -= dy

    if footer:
        ax.text(0.5, 0.06, footer, fontsize=8.5, color=Theme.INK_FADED,
                ha='center', transform=ax.transAxes, fontfamily=CJK_FONT,
                style='italic')

    _watermark(ax)
    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png', facecolor=fig.get_facecolor(),
                bbox_inches='tight', dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf


# ========== 簡訊息方塊 ==========
def render_message_box(
    title: str,
    lines: List[str],
    accent_color: str = None,
    figsize: Tuple[float, float] = None,
) -> BytesIO:
    if accent_color is None: accent_color = Theme.ROYAL_BLUE
    n = len(lines)
    if figsize is None: figsize = (9, max(2.8, 1.6 + n * 0.34))
    fig, ax = plt.subplots(figsize=figsize)
    ax.axis('off')
    _draw_parchment_bg(fig, ax)
    _draw_ornate_border(ax)

    # 左側裝飾條
    ax.add_patch(Rectangle((0.035, 0.05), 0.012, 0.90,
        facecolor=accent_color, edgecolor='none', alpha=0.6,
        transform=ax.transAxes, clip_on=False))

    # 標題
    ax.text(0.08, 0.88, title, fontsize=17, fontweight='bold',
            color=Theme.INK, va='top', transform=ax.transAxes,
            fontfamily=CJK_FONT)

    y = 0.74
    dy = min(0.085, 0.64 / max(n, 1))
    for line in lines:
        if not line:
            y -= dy * 0.35; continue
        ax.text(0.08, y, line, fontsize=11, color=Theme.INK_LIGHT,
                transform=ax.transAxes, fontfamily=CJK_FONT)
        y -= dy

    _watermark(ax)
    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png', facecolor=fig.get_facecolor(),
                bbox_inches='tight', dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf


# ========== Help 指令表格 ==========
def render_help_image(
    bot_name: str,
    sections: List[Tuple[str, List[Tuple[str, str]]]],
    link: str = "",
) -> BytesIO:
    total_cmds = sum(len(cmds) for _, cmds in sections)
    h = max(7, 3.0 + total_cmds * 0.34 + len(sections) * 0.6)
    fig, ax = plt.subplots(figsize=(11, h))
    ax.axis('off')
    _draw_parchment_bg(fig, ax)
    _draw_ornate_border(ax)

    # 標題
    _draw_title_banner(ax, 0.975, bot_name, fontsize=24, color=Theme.INK)
    ax.text(0.5, 0.947, "- Grimoire of Commands -", fontsize=10, color=Theme.INK_FADED,
            ha='center', va='top', transform=ax.transAxes,
            fontfamily=SERIF_FONT, style='italic')

    y = 0.925
    total_h = 0.925 - 0.04
    dy = total_h / (total_cmds + len(sections) * 2 + 2)
    dy_gap = dy * 1.6

    for sec_name, cmds in sections:
        y -= dy_gap
        # 分類標題 — 紋章風格
        _draw_section_divider(ax, y + dy*0.5, 0.06, 0.94)
        ax.text(0.5, y + dy*0.35, sec_name, fontsize=12, fontweight='bold',
                color=Theme.ROYAL_BLUE, ha='center', transform=ax.transAxes,
                fontfamily=CJK_FONT,
                bbox=dict(boxstyle='round,pad=0.3', facecolor=Theme.PARCHMENT,
                         edgecolor='none'))
        y -= dy * 0.5

        for cmd, desc in cmds:
            ax.text(0.08, y, cmd, fontsize=9.5, color=Theme.HERALDIC_RED,
                    fontweight='bold', transform=ax.transAxes, fontfamily=CJK_FONT)
            ax.text(0.44, y, desc, fontsize=9.5, color=Theme.INK_LIGHT,
                    transform=ax.transAxes, fontfamily=CJK_FONT)
            y -= dy

    if link:
        y -= dy_gap * 0.3
        ax.text(0.5, max(y, 0.04), link, fontsize=8.5, color=Theme.ROYAL_BLUE,
                ha='center', transform=ax.transAxes, fontfamily=CJK_FONT,
                style='italic', alpha=0.7)

    _watermark(ax)
    _timestamp(ax, 0.012)

    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png', facecolor=fig.get_facecolor(),
                bbox_inches='tight', dpi=120)
    plt.close(fig)
    buf.seek(0)
    return buf


# ========== 走勢圖表 ==========
def render_line_chart(
    title: str,
    subtitle: str,
    x_labels: List[str],
    datasets: List[Tuple[str, List[float], str]],
    y_label: str = "",
    y_formatter=None,
    annotate_last: bool = True,
) -> BytesIO:
    fig, ax = plt.subplots(figsize=(14, 7))
    fig.set_facecolor(Theme.PARCHMENT)
    ax.set_facecolor(Theme.PARCHMENT_L)

    for label, values, color in datasets:
        ax.plot(range(len(values)), values, '-', color=color, linewidth=2.2,
                label=label, solid_capstyle='round')
        ax.plot(range(len(values)), values, 'o', color=color,
                markersize=4, markerfacecolor=Theme.PARCHMENT_L, markeredgewidth=1.8,
                markeredgecolor=color, zorder=5)
        ax.fill_between(range(len(values)), values, alpha=0.08, color=color)
        if annotate_last and len(values) > 1:
            ax.annotate(f'{values[-1]:,.1f}',
                       xy=(len(values)-1, values[-1]),
                       fontsize=9, color=color, fontweight='bold',
                       textcoords="offset points", xytext=(10, 8),
                       fontfamily=CJK_FONT)

    n = len(x_labels)
    step = max(1, n // 12)
    ticks = list(range(0, n, step))
    ax.set_xticks(ticks)
    ax.set_xticklabels([x_labels[i] for i in ticks], rotation=45, fontsize=8,
                       fontfamily=CJK_FONT, color=Theme.INK_FADED)
    ax.tick_params(axis='y', colors=Theme.INK_FADED, labelsize=9)

    if y_formatter:
        ax.yaxis.set_major_formatter(plt.FuncFormatter(y_formatter))

    ax.set_title(title, fontsize=16, fontweight='bold', color=Theme.INK, pad=18,
                 fontfamily=CJK_FONT)
    if subtitle:
        ax.text(0.5, 1.02, subtitle, fontsize=10, color=Theme.INK_FADED,
                ha='center', transform=ax.transAxes, fontfamily=CJK_FONT,
                style='italic')
    if y_label:
        ax.set_ylabel(y_label, fontsize=11, color=Theme.INK_FADED, fontfamily=CJK_FONT)

    ax.legend(loc='upper left', fontsize=9, prop={'family': CJK_FONT},
              facecolor=Theme.PARCHMENT, edgecolor=Theme.GOLD_DARK,
              labelcolor=Theme.INK)
    ax.grid(True, alpha=0.2, linestyle='--', color=Theme.GOLD_DARK)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(Theme.BORDER)
    ax.spines['bottom'].set_color(Theme.BORDER)

    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png', facecolor=fig.get_facecolor(),
                bbox_inches='tight', dpi=140)
    plt.close(fig)
    buf.seek(0)
    return buf
