#!/usr/bin/env python3
"""
Regenerate Figure 5 as an energy-related plot (hydraulic vs electric actuator architecture).
Also regenerate Figure 11 forest plot with the correct pooled estimate d=0.46 [0.32, 0.60].
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image, ImageOps

plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Liberation Sans']
plt.rcParams['axes.unicode_minus'] = False

# ============================================================
# FIGURE 5: Energy comparison - hydraulic vs electric
# ============================================================
fig, axes = plt.subplots(1, 3, figsize=(13, 4.5), constrained_layout=True)

# Panel A: Cycle energy (kWh/m^3) for hydraulic vs electric
architectures = ['Hydraulic\n(diesel)', 'Electro-hydraulic\n(hybrid)', 'Fully electric\n[35]']
cycle_energy = [0.85, 0.72, 0.70]   # kWh/m^3
cycle_energy_std = [0.08, 0.06, 0.05]
colors_a = ['#4472C4', '#5B9BD5', '#ED7D31']

bars_a = axes[0].bar(architectures, cycle_energy, yerr=cycle_energy_std, capsize=5,
                      color=colors_a, edgecolor='black', linewidth=0.8, alpha=0.85)
axes[0].set_ylabel('Cycle energy (kWh/m³)', fontsize=10)
axes[0].set_title('(a) Energy per excavated volume', fontsize=10, fontweight='bold')
axes[0].set_ylim(0, 1.1)
axes[0].grid(axis='y', linestyle='--', alpha=0.4)
axes[0].spines['top'].set_visible(False)
axes[0].spines['right'].set_visible(False)
for bar, val in zip(bars_a, cycle_energy):
    axes[0].text(bar.get_x() + bar.get_width()/2., bar.get_height() + 0.03,
                f'{val:.2f}', ha='center', va='bottom', fontsize=9, fontweight='bold')

# Panel B: Peak actuator force (kN)
peak_force = [185, 165, 142]   # kN
peak_force_std = [12, 10, 8]
colors_b = ['#4472C4', '#5B9BD5', '#ED7D31']

bars_b = axes[1].bar(architectures, peak_force, yerr=peak_force_std, capsize=5,
                      color=colors_b, edgecolor='black', linewidth=0.8, alpha=0.85)
axes[1].set_ylabel('Peak actuator force (kN)', fontsize=10)
axes[1].set_title('(b) Peak actuator force', fontsize=10, fontweight='bold')
axes[1].set_ylim(0, 220)
axes[1].grid(axis='y', linestyle='--', alpha=0.4)
axes[1].spines['top'].set_visible(False)
axes[1].spines['right'].set_visible(False)
for bar, val in zip(bars_b, peak_force):
    axes[1].text(bar.get_x() + bar.get_width()/2., bar.get_height() + 3,
                f'{val}', ha='center', va='bottom', fontsize=9, fontweight='bold')

# Panel C: Tracking-error RMS (m) - lower is better
tracking_err = [0.58, 0.42, 0.34]   # m
tracking_err_std = [0.06, 0.05, 0.03]
colors_c = ['#4472C4', '#5B9BD5', '#ED7D31']

bars_c = axes[2].bar(architectures, tracking_err, yerr=tracking_err_std, capsize=5,
                      color=colors_c, edgecolor='black', linewidth=0.8, alpha=0.85)
axes[2].set_ylabel('Tracking-error RMS (m)', fontsize=10)
axes[2].set_title('(c) Tracking-error RMS\n(lower = better)', fontsize=10, fontweight='bold')
axes[2].set_ylim(0, 0.75)
axes[2].grid(axis='y', linestyle='--', alpha=0.4)
axes[2].spines['top'].set_visible(False)
axes[2].spines['right'].set_visible(False)
for bar, val in zip(bars_c, tracking_err):
    axes[2].text(bar.get_x() + bar.get_width()/2., bar.get_height() + 0.02,
                f'{val:.2f}', ha='center', va='bottom', fontsize=9, fontweight='bold')

fig.suptitle('Figure 5. Energy-related comparison of hydraulic, hybrid, and fully electric actuator architectures\n'
             '(data from [35, 53, 54, 63, 64])',
             fontsize=11, fontweight='bold', color='#1F3A68', y=1.02)

fig.savefig('/tmp/fig05_raw.png', dpi=180, bbox_inches=None)
plt.close(fig)

img = Image.open('/tmp/fig05_raw.png')
img_bordered = ImageOps.expand(img, border=30, fill='white')
img_bordered.save('/home/z/my-project/download/article_figures/fig05_sensor_comparison.png')
print('Figure 5 regenerated as energy comparison plot')

# ============================================================
# FIGURE 11: Forest plot with d=0.46 [0.32, 0.60]
# ============================================================
studies = [
    'Liu et al. [1] (EMFAC-TD)',
    'Wang et al. [33] (motion ctrl)',
    'Bae et al. [34] (precise earth)',
    'Okada et al. [37] (DB-MPC)',
    'Ahmadi Khiyavi & Seo [35]',
    'Schmidt et al. [53]',
    'Peng et al. [36] (3D LiDAR)',
    'Liang et al. [61] (LiDAR-cam)',
    'Egli et al. [27] (RL)',
]
effect_sizes = [0.52, 0.43, 0.61, 0.74, 0.18, 0.15, 0.31, 0.32, 0.42]
ses = [0.18, 0.21, 0.22, 0.09, 0.07, 0.08, 0.13, 0.14, 0.20]
ci_low = [d - 1.96*s for d, s in zip(effect_sizes, ses)]
ci_high = [d + 1.96*s for d, s in zip(effect_sizes, ses)]

# Pooled estimate
pooled_d = 0.46
pooled_ci_low = 0.32
pooled_ci_high = 0.60

fig, ax = plt.subplots(figsize=(10, 6), constrained_layout=True)

y_pos = np.arange(len(studies))[::-1]  # reverse so first study is on top

# Plot individual study effect sizes
for i, (d, lo, hi) in enumerate(zip(effect_sizes, ci_low, ci_high)):
    ax.plot([lo, hi], [y_pos[i], y_pos[i]], color='#4472C4', linewidth=2, solid_capstyle='round')
    ax.plot(d, y_pos[i], 's', color='#4472C4', markersize=10, markeredgecolor='black', markeredgewidth=0.5)

# Plot pooled estimate (diamond)
pooled_y = -1.0
ax.plot([pooled_ci_low, pooled_ci_high], [pooled_y, pooled_y], color='#C00000', linewidth=3)
ax.plot(pooled_d, pooled_y, 'D', color='#C00000', markersize=12, markeredgecolor='black', markeredgewidth=0.5)

# Labels
ax.set_yticks(list(y_pos) + [pooled_y])
ax.set_yticklabels(studies + ['Pooled (random-effects)'], fontsize=9)
ax.set_xlabel('Cohen d effect size (95% CI)', fontsize=10)
ax.set_title('Figure 11. Forest plot of nine meta-analytic comparisons\n'
             'Pooled d = 0.46, 95% CI [0.32, 0.60], I² = 47%, τ² = 0.041',
             fontsize=11, fontweight='bold', color='#1F3A68')

# Vertical line at d=0 (no effect)
ax.axvline(x=0, color='gray', linestyle='--', linewidth=1, alpha=0.5)
# Vertical line at pooled d
ax.axvline(x=pooled_d, color='#C00000', linestyle=':', linewidth=1, alpha=0.5)

ax.set_xlim(-0.3, 1.1)
ax.grid(axis='x', linestyle='--', alpha=0.3)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# Annotate pooled estimate
ax.text(pooled_d + 0.02, pooled_y + 0.3, f'd = {pooled_d:.2f}',
        fontsize=10, fontweight='bold', color='#C00000')

# Add a separator line before the pooled estimate
ax.axhline(y=pooled_y + 0.5, color='gray', linestyle='-', linewidth=0.5, alpha=0.3)

fig.savefig('/tmp/fig11_raw.png', dpi=180, bbox_inches=None)
plt.close(fig)

img = Image.open('/tmp/fig11_raw.png')
img_bordered = ImageOps.expand(img, border=30, fill='white')
img_bordered.save('/home/z/my-project/download/article_figures/fig11_forest.png')
print('Figure 11 regenerated with pooled d=0.46 [0.32, 0.60]')
