import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
for f in font_manager.findSystemFonts():
    if "Carlito" in f: font_manager.fontManager.addfont(f)
labels = ["Below 45%","45–54.9%","55–59.9%","60–64.9%","65–69.9%","70–74.9%","75–79.9%","80–84.9%","85–89.9%","90–100%"]
counts = [13,27,30,23,23,18,20,10,9,13]
assert sum(counts)==186
GRAY="#595959"
plt.rcParams.update({"font.family":"Carlito","font.size":11})
fig, ax = plt.subplots(figsize=(9,5.2), dpi=200)
ax.bar(range(10), counts, width=0.45, color="#4472C4", zorder=3)
for i,c in enumerate(counts): ax.text(i, c+0.4, str(c), ha="center", va="bottom", color=GRAY, fontsize=11)
ax.set_xticks(range(10)); ax.set_xticklabels(labels, color=GRAY, fontsize=10.5)
ax.set_ylim(0,35); ax.set_yticks(range(0,36,5))
ax.tick_params(colors=GRAY, length=0)
ax.grid(axis="y", color="#D9D9D9", linewidth=0.8, zorder=0)
for s in ["top","right","left"]: ax.spines[s].set_visible(False)
ax.spines["bottom"].set_color("#BFBFBF")
ax.set_title("Exam 2 Score Distribution", color=GRAY, fontsize=16, pad=14)
ax.set_xlabel("Score Range", color=GRAY, fontsize=12, labelpad=8)
ax.set_ylabel("Number of Students", color=GRAY, fontsize=12, labelpad=8)
fig.tight_layout()
fig.savefig("NEUR411_Exam2_distribution.png", facecolor="white")
