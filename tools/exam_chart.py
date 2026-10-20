import sys, numpy as np, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt, matplotlib.ticker
scores = np.array([float(x) for x in open(sys.argv[1]).read().split()])
exam, out = sys.argv[2], sys.argv[3]
pct = scores / 120 * 100
edges = np.arange(25, 100.01, 2.5)
counts, _ = np.histogram(pct, bins=edges)
assert counts.sum() == len(pct)
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 11})
fig, ax = plt.subplots(figsize=(8, 5), dpi=200)
ax.hist(pct, bins=edges, color="#0072BD", alpha=0.6, edgecolor="black", linewidth=0.5)
ax.set_xlim(25, 100); ax.set_xticks(range(25, 101, 5))
ax.set_ylim(0, counts.max() + 3); ax.yaxis.set_major_locator(matplotlib.ticker.MultipleLocator(2))
ax.tick_params(direction="in", top=True, right=True)
for s in ax.spines.values(): s.set_linewidth(0.8)
ax.set_xlabel("Score (%)"); ax.set_ylabel("Number of students")
ax.set_title(f"Exam {exam} Score Distribution (n = {len(pct)})")
fig.tight_layout(); fig.savefig(out, facecolor="white")
for lo, c in zip(edges[:-1], counts): print(f"{lo:.1f}-{lo+2.5:.1f}: {c}")
