"""Generate academic diagram PNGs for PPT slides using matplotlib.

Covers three diagram types that scipilot-figure-skill explicitly excludes:
  - flowchart      : rectangular process boxes + arrows (+ diamond decisions)
  - layer_diagram  : stacked horizontal bars representing system/architecture layers
  - block_diagram  : free-positioned labeled blocks with connecting arrows

Usage:
    python scripts/generate_diagram.py --demo --out diagrams/

API:
    from scripts.generate_diagram import draw_flowchart, draw_layer_diagram, draw_block_diagram
"""
from __future__ import annotations

import argparse
import math
from pathlib import Path
from typing import Any

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

# ── palette ──────────────────────────────────────────────────────────────────
DEEP_BLUE    = "#1F497D"
MID_BLUE     = "#2E75B6"
LIGHT_BLUE   = "#BDD7EE"
GOLD         = "#FFC000"
GREY_BG      = "#F2F2F2"
WHITE        = "#FFFFFF"
DARK_TEXT    = "#1A1A1A"
ARROW_COLOR  = "#555555"


# ─────────────────────────────────────────────────────────────────────────────
# 1. FLOWCHART
# ─────────────────────────────────────────────────────────────────────────────

def draw_flowchart(
    nodes: list[dict[str, str]],
    edges: list[tuple],
    output_path: str | Path,
    figsize: tuple[float, float] = (9, 7),
    dpi: int = 160,
    title: str = "",
    lang: str = "zh",
) -> Path:
    """Render a vertical flowchart.

    Args:
        nodes: list of dicts with keys:
            id   : unique string key
            label: display text (may contain \\n for line breaks)
            type : "start" | "end" | "process" | "decision" | "io"
        edges: list of (from_id, to_id) or (from_id, to_id, label)
        output_path: PNG output path.
        figsize: matplotlib figure size in inches.
        dpi: output DPI.
        title: optional title drawn above the chart.
        lang: "zh" (黑体) or "en" (sans-serif).

    Returns: resolved output Path.
    """
    _configure_font(lang)
    n = len(nodes)
    if n == 0:
        raise ValueError("nodes list is empty")

    fig, ax = plt.subplots(figsize=figsize)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, n * 1.6 + 0.5)
    ax.set_aspect("equal")
    ax.axis("off")
    if title:
        ax.set_title(title, fontsize=13, color=DEEP_BLUE, pad=8)

    id2pos: dict[str, tuple[float, float]] = {}
    id2hw:  dict[str, tuple[float, float]] = {}

    y_cursor = n * 1.6
    cx = 5.0

    for node in nodes:
        nid   = node["id"]
        label = node.get("label", nid)
        ntype = node.get("type", "process")
        y = y_cursor
        y_cursor -= 1.6

        id2pos[nid] = (cx, y)

        if ntype in ("start", "end"):
            w, h = 3.0, 0.65
            box = FancyBboxPatch(
                (cx - w / 2, y - h / 2), w, h,
                boxstyle="round,pad=0.12",
                facecolor=GOLD, edgecolor=DEEP_BLUE, linewidth=1.5, zorder=3
            )
            ax.add_patch(box)
            ax.text(cx, y, label, ha="center", va="center",
                    fontsize=10, color=DARK_TEXT, fontweight="bold", zorder=4)
            id2hw[nid] = (w, h)

        elif ntype == "decision":
            w, h = 3.6, 0.9
            diamond = plt.Polygon(
                [(cx, y + h / 2), (cx + w / 2, y),
                 (cx, y - h / 2), (cx - w / 2, y)],
                facecolor=LIGHT_BLUE, edgecolor=DEEP_BLUE, linewidth=1.5, zorder=3
            )
            ax.add_patch(diamond)
            ax.text(cx, y, label, ha="center", va="center",
                    fontsize=9, color=DARK_TEXT, zorder=4)
            id2hw[nid] = (w, h)

        elif ntype == "io":
            w, h = 3.4, 0.65
            skew = 0.25
            pts = [(cx - w / 2 + skew, y + h / 2),
                   (cx + w / 2 + skew, y + h / 2),
                   (cx + w / 2 - skew, y - h / 2),
                   (cx - w / 2 - skew, y - h / 2)]
            para = plt.Polygon(pts, facecolor=GREY_BG,
                               edgecolor=MID_BLUE, linewidth=1.5, zorder=3)
            ax.add_patch(para)
            ax.text(cx, y, label, ha="center", va="center",
                    fontsize=9, color=DARK_TEXT, zorder=4)
            id2hw[nid] = (w, h)

        else:  # process
            w, h = 3.4, 0.65
            box = FancyBboxPatch(
                (cx - w / 2, y - h / 2), w, h,
                boxstyle="square,pad=0.05",
                facecolor=WHITE, edgecolor=MID_BLUE, linewidth=1.5, zorder=3
            )
            ax.add_patch(box)
            ax.text(cx, y, label, ha="center", va="center",
                    fontsize=9, color=DARK_TEXT, zorder=4)
            id2hw[nid] = (w, h)

    # draw edges
    for edge in edges:
        fid, tid = edge[0], edge[1]
        elabel = edge[2] if len(edge) > 2 else ""
        fx, fy = id2pos[fid]
        tx, ty = id2pos[tid]
        _, fh = id2hw[fid]
        _, th = id2hw[tid]

        y_start = fy - fh / 2
        y_end   = ty + th / 2

        ax.annotate(
            "", xy=(tx, y_end), xytext=(fx, y_start),
            arrowprops=dict(arrowstyle="->", color=ARROW_COLOR,
                            lw=1.4, connectionstyle="arc3,rad=0"),
            zorder=2
        )
        if elabel:
            mx = (fx + tx) / 2 + 0.25
            my = (y_start + y_end) / 2
            ax.text(mx, my, elabel, fontsize=8, color=ARROW_COLOR,
                    ha="left", va="center")

    fig.tight_layout(pad=0.4)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(str(out), dpi=dpi, bbox_inches="tight",
                facecolor=WHITE, edgecolor="none")
    plt.close(fig)
    print(f"  [flowchart] {out}")
    return out


# ─────────────────────────────────────────────────────────────────────────────
# 2. LAYER / ARCHITECTURE DIAGRAM
# ─────────────────────────────────────────────────────────────────────────────

def draw_layer_diagram(
    layers: list[dict[str, str]],
    output_path: str | Path,
    figsize: tuple[float, float] = (9, 5),
    dpi: int = 160,
    title: str = "",
    lang: str = "zh",
    show_arrows: bool = True,
) -> Path:
    """Render a stacked-layer (architecture) diagram.

    Args:
        layers: list of dicts (bottom to top), each with:
            label : display text
            color : hex fill color (optional, defaults to blue gradient)
            sublabel: optional smaller text on the right side
        output_path: PNG output path.
        figsize, dpi, title, lang: standard options.
        show_arrows: draw bidirectional arrows between adjacent layers.

    Returns: resolved output Path.
    """
    _configure_font(lang)
    n = len(layers)
    if n == 0:
        raise ValueError("layers list is empty")

    default_colors = [
        "#1F497D", "#2E5FA3", "#3672C4", "#4E86D4",
        "#6AA0E0", "#8BBAE8", "#AECFF0", "#D0E4F7",
    ]

    fig, ax = plt.subplots(figsize=figsize)
    ax.set_xlim(0, 10)
    ax.set_ylim(-0.3, n * 1.1 + 0.5)
    ax.axis("off")
    if title:
        ax.set_title(title, fontsize=13, color=DEEP_BLUE, pad=8)

    layer_h = 0.82
    gap = 0.28
    lx, rx = 0.5, 9.5

    for i, layer in enumerate(reversed(layers)):   # bottom layer drawn first
        color = layer.get("color", default_colors[i % len(default_colors)])
        label = layer.get("label", f"Layer {i+1}")
        sublabel = layer.get("sublabel", "")
        y = i * (layer_h + gap)

        # shadow / depth effect
        shadow = FancyBboxPatch(
            (lx + 0.07, y - 0.07), rx - lx, layer_h,
            boxstyle="round,pad=0.04",
            facecolor="#CCCCCC", edgecolor="none", zorder=1, alpha=0.5
        )
        ax.add_patch(shadow)

        box = FancyBboxPatch(
            (lx, y), rx - lx, layer_h,
            boxstyle="round,pad=0.04",
            facecolor=color, edgecolor=WHITE, linewidth=1.5, zorder=2
        )
        ax.add_patch(box)

        # main label
        ax.text((lx + rx) / 2, y + layer_h / 2, label,
                ha="center", va="center", fontsize=10,
                color=WHITE, fontweight="bold", zorder=3)

        # optional sublabel (right-aligned, smaller)
        if sublabel:
            ax.text(rx - 0.1, y + layer_h / 2, sublabel,
                    ha="right", va="center", fontsize=8,
                    color=LIGHT_BLUE, style="italic", zorder=3)

        # layer index badge on left
        ax.text(lx + 0.18, y + layer_h / 2,
                f"L{len(layers) - i}",
                ha="center", va="center", fontsize=8,
                color=WHITE, alpha=0.7, zorder=3)

    # inter-layer arrows
    if show_arrows and n > 1:
        ax_mid = (lx + rx) / 2
        for i in range(n - 1):
            y_bot = i * (layer_h + gap) + layer_h
            y_top = (i + 1) * (layer_h + gap)
            y_mid = (y_bot + y_top) / 2
            ax.annotate(
                "", xy=(ax_mid, y_top - 0.02), xytext=(ax_mid, y_bot + 0.02),
                arrowprops=dict(arrowstyle="<->", color=GREY_BG,
                                lw=1.2, mutation_scale=12),
                zorder=4
            )

    fig.tight_layout(pad=0.4)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(str(out), dpi=dpi, bbox_inches="tight",
                facecolor=WHITE, edgecolor="none")
    plt.close(fig)
    print(f"  [layer_diagram] {out}")
    return out


# ─────────────────────────────────────────────────────────────────────────────
# 3. BLOCK DIAGRAM (free layout)
# ─────────────────────────────────────────────────────────────────────────────

def draw_block_diagram(
    blocks: list[dict[str, Any]],
    connections: list[dict[str, Any]],
    output_path: str | Path,
    figsize: tuple[float, float] = (10, 6),
    dpi: int = 160,
    title: str = "",
    lang: str = "zh",
) -> Path:
    """Render a free-layout block diagram.

    Args:
        blocks: list of dicts:
            id    : unique key
            label : display text
            x, y  : center position (0-10 coordinate space)
            w, h  : box width/height (default 2.0 / 0.7)
            color : optional hex fill color
        connections: list of dicts:
            from, to : block ids
            label    : optional arrow label
            style    : "->", "<->", "-->" (default "->")
        output_path, figsize, dpi, title, lang: standard options.

    Returns: resolved output Path.
    """
    _configure_font(lang)

    fig, ax = plt.subplots(figsize=figsize)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.set_aspect("equal")
    ax.axis("off")
    if title:
        ax.set_title(title, fontsize=13, color=DEEP_BLUE, pad=8)

    id2block: dict[str, dict] = {b["id"]: b for b in blocks}

    for block in blocks:
        cx = block["x"]
        cy = block["y"]
        w  = block.get("w", 2.0)
        h  = block.get("h", 0.7)
        color = block.get("color", MID_BLUE)
        label = block.get("label", block["id"])

        box = FancyBboxPatch(
            (cx - w / 2, cy - h / 2), w, h,
            boxstyle="round,pad=0.08",
            facecolor=color, edgecolor=WHITE, linewidth=1.5, zorder=2
        )
        ax.add_patch(box)
        ax.text(cx, cy, label, ha="center", va="center",
                fontsize=9, color=WHITE, fontweight="bold", zorder=3)

    for conn in connections:
        fb = id2block[conn["from"]]
        tb = id2block[conn["to"]]
        fx, fy = fb["x"], fb["y"]
        tx, ty = tb["x"], tb["y"]
        style = conn.get("style", "->")
        clabel = conn.get("label", "")
        rad = conn.get("rad", 0.0)

        ax.annotate(
            "", xy=(tx, ty), xytext=(fx, fy),
            arrowprops=dict(
                arrowstyle=style, color=ARROW_COLOR, lw=1.4,
                connectionstyle=f"arc3,rad={rad}"
            ),
            zorder=1
        )
        if clabel:
            mx = (fx + tx) / 2 + 0.1
            my = (fy + ty) / 2 + 0.1
            ax.text(mx, my, clabel, fontsize=7.5, color=ARROW_COLOR,
                    ha="center", va="center",
                    bbox=dict(facecolor=WHITE, edgecolor="none", pad=1))

    fig.tight_layout(pad=0.4)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(str(out), dpi=dpi, bbox_inches="tight",
                facecolor=WHITE, edgecolor="none")
    plt.close(fig)
    print(f"  [block_diagram] {out}")
    return out


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _configure_font(lang: str) -> None:
    import matplotlib as mpl
    if lang == "zh":
        for family in ["Noto Sans CJK SC", "Source Han Sans SC",
                        "SimHei", "Microsoft YaHei"]:
            try:
                mpl.rcParams["font.family"] = family
                fig_test = plt.figure()
                plt.close(fig_test)
                break
            except Exception:
                continue
    mpl.rcParams["axes.unicode_minus"] = False


def fit_image_in_box(img_path: str | Path,
                     box_w_in: float, box_h_in: float) -> tuple[float, float]:
    """Return (w, h) in inches that fit img within box_w_in x box_h_in preserving ratio."""
    from PIL import Image
    im = Image.open(str(img_path))
    iw, ih = im.size
    ratio = ih / iw
    w = min(box_w_in, box_h_in / ratio)
    h = w * ratio
    if h > box_h_in:
        h = box_h_in
        w = h / ratio
    return w, h


# ─────────────────────────────────────────────────────────────────────────────
# Demo + CLI
# ─────────────────────────────────────────────────────────────────────────────

def run_demo(out_dir: Path) -> None:
    """Generate three sample diagrams in out_dir."""
    out_dir.mkdir(parents=True, exist_ok=True)

    # 1. Flowchart: nanobolometer detection process
    draw_flowchart(
        nodes=[
            {"id": "s",  "label": "LWIR 光子入射",       "type": "start"},
            {"id": "p1", "label": "MIM 谐振腔波长选择",   "type": "process"},
            {"id": "d1", "label": "波长匹配?",            "type": "decision"},
            {"id": "p2", "label": "局域电场增强 + 热化",  "type": "process"},
            {"id": "p3", "label": "热载流子激发超导态",   "type": "process"},
            {"id": "p4", "label": "电阻变化 → 光电压",    "type": "io"},
            {"id": "e",  "label": "输出信号",             "type": "end"},
        ],
        edges=[
            ("s", "p1"), ("p1", "d1"),
            ("d1", "p2", "是"), ("p2", "p3"), ("p3", "p4"), ("p4", "e"),
        ],
        output_path=out_dir / "diagram_flowchart.png",
        title="超导红外纳米测辐射热计工作流程",
    )

    # 2. Layer diagram: device layer structure
    draw_layer_diagram(
        layers=[
            {"label": "Si 基底 (衬底层)",          "color": "#1F497D", "sublabel": "厚度 ~500 μm"},
            {"label": "SiO₂ 隔热层",               "color": "#2E5FA3", "sublabel": "100 nm"},
            {"label": "超导 NbN 薄膜 (传感层)",     "color": "#3672C4", "sublabel": "10 nm"},
            {"label": "Au 纳米条阵列 (MIM 顶层)",   "color": "#FFC000", "sublabel": "80 nm"},
        ],
        output_path=out_dir / "diagram_device_layers.png",
        title="器件层次结构",
    )

    # 3. Block diagram: measurement system
    draw_block_diagram(
        blocks=[
            {"id": "src",  "label": "宽谱红外光源",       "x": 1.5, "y": 3.0, "color": DEEP_BLUE},
            {"id": "mono", "label": "光栅单色仪",          "x": 3.5, "y": 3.0, "color": MID_BLUE},
            {"id": "det",  "label": "NbN 探测器阵列",      "x": 5.8, "y": 3.0, "color": "#C00000"},
            {"id": "lna",  "label": "低噪放大器 (LNA)",    "x": 5.8, "y": 1.8, "color": MID_BLUE},
            {"id": "daq",  "label": "锁相放大器 / DAQ",    "x": 8.2, "y": 2.4, "color": DEEP_BLUE},
            {"id": "pc",   "label": "上位机分析",          "x": 8.2, "y": 4.0, "color": "#555555"},
        ],
        connections=[
            {"from": "src",  "to": "mono", "label": "宽谱"},
            {"from": "mono", "to": "det",  "label": "单色光"},
            {"from": "det",  "to": "lna",  "label": "光电压"},
            {"from": "lna",  "to": "daq"},
            {"from": "daq",  "to": "pc",   "label": "数字信号"},
        ],
        output_path=out_dir / "diagram_measurement_system.png",
        title="实验测量系统框图",
    )

    print(f"\nDemo diagrams saved to: {out_dir}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--demo", action="store_true",
                        help="Generate demo flowchart, layer, and block diagrams")
    parser.add_argument("--out", default="diagrams",
                        help="Output directory (default: diagrams/)")
    args = parser.parse_args()
    if args.demo:
        run_demo(Path(args.out))
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
