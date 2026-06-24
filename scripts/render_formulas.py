"""Batch formula renderer: matplotlib mathtext -> transparent PNG.

Usage:
    python scripts/render_formulas.py --out <output_dir>

No LaTeX installation required. Uses matplotlib's built-in mathtext engine.

Rules for formula strings:
  - Wrap math in a single $...$ block per render() call.
  - Do NOT write adjacent $A$ $B$ pairs; merge into $A \\quad B$.
  - Do NOT use \\& (unsupported); use \\text{and} or \\mathrm{and}.
  - \\dfrac, \\sqrt, \\times, \\mathcal, \\mathrm, \\Rightarrow all work.
  - Mix plain text and math inside one block using \\text{}.
"""
import argparse
import os
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

DEEP_BLUE = "#1F497D"
RED_KEY = "#C00000"


def render(
    name: str,
    tex: str,
    out_dir: Path,
    fig_width: float = 7.0,
    fig_height: float = 1.0,
    fontsize: int = 22,
    color: str = DEEP_BLUE,
    background: str | None = None,
    dpi: int = 160,
) -> Path:
    """Render one formula and save to out_dir/name."""
    fig = plt.figure(figsize=(fig_width, fig_height))
    if background:
        fig.patch.set_facecolor(background)
    else:
        fig.patch.set_alpha(0.0)

    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_axis_off()
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.text(
        0.5, 0.5, tex,
        fontsize=fontsize,
        ha="center", va="center",
        color=color,
        transform=ax.transAxes,
        usetex=False,
    )
    out_path = out_dir / name
    fig.savefig(str(out_path), dpi=dpi, bbox_inches="tight",
                transparent=(background is None), pad_inches=0.05)
    plt.close(fig)
    print(f"  {name}")
    return out_path


def render_defaults(out_dir: Path) -> None:
    """Render the default set of formulas used in the APL nanobolometer paper demo."""
    print(f"Output directory: {out_dir}")
    print("Rendering formulas...")

    render("formula_lambda.png",
           r"$\lambda_i \approx 2\,n_{eff}\,W_i + \Delta$"
           r"   ($n_{eff} \approx 3.6 \approx n_{Si}$)",
           out_dir, fig_width=7, fig_height=0.8, fontsize=22)

    render("formula_eta_sort.png",
           r"$\eta_{sort} = \dfrac{E_{\lambda_1,1}}{E_{\lambda_1,1}+E_{\lambda_1,2}}"
           r" + \dfrac{E_{\lambda_2,2}}{E_{\lambda_2,2}+E_{\lambda_2,1}} - 1$",
           out_dir, fig_width=8, fig_height=1.2, fontsize=22)

    render("formula_nep.png",
           r"$NEP = \dfrac{S_n(\omega)}{|\mathcal{R}|}"
           r"\quad \approx \dfrac{1.2\times10^{-7}\;\mathrm{V/\sqrt{Hz}}}{\mathcal{R}}$",
           out_dir, fig_width=8, fig_height=1.0, fontsize=20)

    render("formula_dstar.png",
           r"$D^* = \dfrac{\sqrt{A_d}}{NEP} \Rightarrow"
           r" 8.33\times10^{11}\ \mathrm{and}\ 3.12\times10^{12}\ \mathrm{Jones}$",
           out_dir, fig_width=9, fig_height=1.0, fontsize=20)

    render("formula_mim.png",
           r"$\lambda_i \approx 2\,n_{eff}\,W_i + \Delta,\quad"
           r"\quad n_{eff} \approx 3.6 \approx n_{Si}$",
           out_dir, fig_width=9, fig_height=0.9, fontsize=19)

    render("formula_resp.png",
           r"$\mathcal{R}_1\!=\!9.7\!\times\!10^{6}\ \mathrm{V/W}\ (1025\ \mathrm{cm}^{-1}),"
           r"\quad \mathcal{R}_2\!=\!3.7\!\times\!10^{7}\ \mathrm{V/W}\ (1377\ \mathrm{cm}^{-1})$",
           out_dir, fig_width=10, fig_height=0.85, fontsize=19, color=RED_KEY)

    print("Done.")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default="formulas", help="Output directory for PNG files")
    args = parser.parse_args()
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    render_defaults(out_dir)


if __name__ == "__main__":
    main()
