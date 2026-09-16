"""Reusable Matplotlib theme for presentation charts.

Usage:
    from plot_theme import CHART_COLORS, apply_chart_theme

    apply_chart_theme()
    ax.bar(..., color=CHART_COLORS['year'][2022])
"""

import matplotlib.pyplot as plt

__all__ = ['CHART_COLORS', 'apply_chart_theme']

CHART_COLORS = {
    'navy': '#003057',
    'blue': '#0072CE',
    'sky': '#7BA7C2',
    'ink': '#1C2429',
    'muted': '#5B6770',
    'grid': '#E6E8EB',
    'spine': '#C5CCD1',
    'paper': '#FFFFFF',
    'year': {2020: '#7BA7C2', 2021: '#0072CE', 2022: '#003057'},
}


def apply_chart_theme():
    """Apply restrained typography, y-grid only, and slide-sized labels."""
    plt.rcParams.update({
        'figure.figsize': (13.33, 7.5),
        'figure.dpi': 120,
        'figure.facecolor': CHART_COLORS['paper'],
        'axes.facecolor': CHART_COLORS['paper'],
        'axes.edgecolor': CHART_COLORS['spine'],
        'axes.linewidth': 0.6,
        'axes.spines.top': False,
        'axes.spines.right': False,
        'axes.spines.left': False,
        'axes.spines.bottom': True,
        'axes.grid': True,
        'axes.grid.axis': 'y',
        'axes.axisbelow': True,
        'grid.color': CHART_COLORS['grid'],
        'grid.linewidth': 0.6,
        'grid.linestyle': '-',
        'grid.alpha': 1,
        'font.family': 'sans-serif',
        'font.sans-serif': ['Helvetica Neue', 'Helvetica', 'Arial', 'DejaVu Sans'],
        'font.size': 12,
        'text.color': CHART_COLORS['ink'],
        'axes.labelcolor': CHART_COLORS['ink'],
        'axes.titlesize': 16,
        'axes.titleweight': 'medium',
        'axes.titlelocation': 'left',
        'axes.titlepad': 12,
        'axes.labelsize': 12,
        'axes.labelpad': 8,
        'xtick.labelsize': 11,
        'ytick.labelsize': 11,
        'xtick.color': CHART_COLORS['muted'],
        'ytick.color': CHART_COLORS['muted'],
        'xtick.major.size': 0,
        'ytick.major.size': 0,
        'xtick.major.pad': 6,
        'ytick.major.pad': 6,
        'legend.frameon': False,
        'legend.fontsize': 11,
        'legend.title_fontsize': 11,
        'pdf.fonttype': 42,
        'ps.fonttype': 42,
    })
    return CHART_COLORS
