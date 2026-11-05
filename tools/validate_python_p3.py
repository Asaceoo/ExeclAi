# -*- coding: utf-8 -*-
"""P3 验收脚本（docs/需求-Python高级数据分析.md AC-P3 / AC-P4）。

AC-P4 时序：销售明细按月聚合金额 → Holt-Winters(Holt 阻尼趋势) + 小网格 ARIMA(p,d,q≤2)；
留出最后 3 个月回测（报告 MAPE 并断言 <50%），全量拟合输出未来 3 个月预测 + 95% 置信区间。
  说明：季节性分解（STL）与季节项 Holt-Winters 需 ≥2 个完整周期（≥24 个月度点），
  测试工作簿仅 12 个月，故本脚本验证趋势型模型；≥24 点时走季节模型（运行期由 AI 按
  python-advanced-analysis 技能契约自行选择）。
AC-P3 文本：产品名称列 jieba 分词 → 词频 Top15 → wordcloud 词云 PNG（中文字体探测链
  msyh.ttc → simhei.ttf → simsun.ttc，全缺失时放弃出图 fail-closed，禁止乱码图）。

用法：<python.exe> tools/validate_python_p3.py [工作簿路径]
产物：_dev_tmp/p3_forecast.png、_dev_tmp/p3_wordcloud.png；退出码 0=全部通过。
"""
import os
import sys
import warnings
from collections import Counter

import pandas as pd

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_XLSX = os.path.join(REPO, "2026-09-11-9accc9f9", "看板图表全功能测试数据.xlsx")
OUT_DIR = os.path.join(REPO, "_dev_tmp")
FONT_CHAIN = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\simsun.ttc",
]


def probe_font():
    for path in FONT_CHAIN:
        if os.path.exists(path):
            return path
    return None


def load_monthly(xlsx):
    df = pd.read_excel(xlsx, sheet_name="销售明细")
    for col in ("月份", "金额", "产品名称"):
        assert col in df.columns, f"缺列 {col}"
    df["月份"] = df["月份"].astype(str)
    monthly = df.groupby("月份")["金额"].sum().sort_index()
    assert len(monthly) >= 8, f"月度点不足：{len(monthly)}"
    return df, monthly


def run_forecast(monthly):
    import numpy as np
    from statsmodels.tsa.holtwinters import Holt, ExponentialSmoothing
    from statsmodels.tsa.arima.model import ARIMA

    values = monthly.values.astype(float)
    holdout = 3
    train, test = values[:-holdout], values[-holdout:]
    seasonal = len(values) >= 24  # 季节模型需 ≥2 个完整周期（文档 v1.3 边界承诺）

    # 回测：留出 3 个月，比较候选模型 MAPE。
    def mape(actual, pred):
        return float(np.mean(np.abs((actual - pred) / actual)))

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        hw = Holt(train, damped_trend=True, initialization_method="estimated").fit()
        hw_pred = hw.forecast(holdout)
        candidates = [("Holt", mape(test, hw_pred))]
        if seasonal:
            try:
                # 注意：heuristic 初始化要求训练集 ≥2 个完整季节周期；留出回测后常不满足，
                # 必须用 estimated（MLE 初始化）；候选仍失败则降级剔除，不阻塞主链。
                hw_s = ExponentialSmoothing(
                    train, trend="add", damped_trend=True,
                    seasonal="add", seasonal_periods=12,
                    initialization_method="estimated").fit()
                candidates.append(("HoltWinters(季节)", mape(test, hw_s.forecast(holdout))))
            except Exception as exc:
                print(f"[AC-P4] 季节模型回测候选失败，降级剔除：{type(exc).__name__}: {exc}")
        best_ar = None
        for p in (0, 1, 2):
            for d in (0, 1):
                for q in (0, 1, 2):
                    if p == d == q == 0:
                        continue
                    try:
                        res = ARIMA(train, order=(p, d, q)).fit()
                    except Exception:
                        continue
                    if best_ar is None or res.aic < best_ar[0]:
                        best_ar = (res.aic, (p, d, q), res)
    arima_pred = best_ar[2].get_forecast(holdout).predicted_mean
    candidates.append((f"ARIMA{best_ar[1]}", mape(test, arima_pred)))
    for name, m in candidates:
        print(f"[AC-P4] 回测 MAPE：{name}={m:.1%}")
    assert all(m < 0.5 for _, m in candidates), "回测 MAPE ≥50%，预测不可信"

    # 全量拟合 → 未来 3 个月；Holt-Winters(季节) 为首选，ARIMA 次之（原生 CI），Holt 兜底（近似区间）。
    chosen, point, ci = None, None, None
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        if seasonal:
            full = ExponentialSmoothing(
                values, trend="add", damped_trend=True,
                seasonal="add", seasonal_periods=12,
                initialization_method="estimated").fit()
            point = np.asarray(full.forecast(3))
            spread = np.abs(point).max() * 0.2
            ci = np.column_stack([point - spread, point + spread])
            chosen = "HoltWinters季节(参考区间±20%)"
        elif mape(test, arima_pred) <= mape(test, hw_pred):
            full = ARIMA(values, order=best_ar[1]).fit()
            fc = full.get_forecast(3)
            point = np.asarray(fc.predicted_mean)
            ci = np.asarray(fc.conf_int(alpha=0.05))
            chosen = f"ARIMA{best_ar[1]}"
        else:
            full = Holt(values, damped_trend=True, initialization_method="estimated").fit()
            point = np.asarray(full.forecast(3))
            spread = np.abs(point).max() * 0.2
            ci = np.column_stack([point - spread, point + spread])
            chosen = "Holt(参考区间±20%)"
    assert np.all(np.isfinite(point)) and np.all(ci[:, 1] > ci[:, 0]), "预测值/置信区间异常"
    print(f"[AC-P4] 选型 {chosen}；未来 3 个月预测：")
    for i in range(3):
        nxt = f"2026-{i+1:02d}"
        print(f"    {nxt}: {point[i]:,.0f}  区间 [{ci[i][0]:,.0f}, {ci[i][1]:,.0f}]")
    return chosen, point, ci


def plot_forecast(monthly, point, ci, chosen, font_path):
    import matplotlib
    matplotlib.use("Agg")
    from matplotlib import font_manager
    import matplotlib.pyplot as plt

    font_manager.fontManager.addfont(font_path)
    name = font_manager.FontProperties(fname=font_path).get_name()
    plt.rcParams["font.family"] = name
    plt.rcParams["axes.unicode_minus"] = False

    hist = list(monthly.index) + [f"2026-{i:02d}" for i in (1, 2, 3)]
    vals = list(monthly.values) + list(point)
    x = range(len(hist))
    fig, ax = plt.subplots(figsize=(9, 4.2))
    ax.plot(list(x)[: len(monthly)], vals[: len(monthly)], "o-", label="历史实际")
    ax.plot(list(x)[len(monthly) - 1:], vals[len(monthly) - 1:], "s--", color="#d62728", label=f"预测（{chosen}）")
    ci_label = "置信区间" if chosen.startswith("ARIMA") else "参考区间"
    ax.fill_between(list(x)[len(monthly):], ci[:, 0], ci[:, 1], alpha=0.2, color="#d62728", label=f"{ci_label}")
    ax.set_title("月度销售额预测（看板图表全功能测试数据·销售明细）")
    ax.set_ylabel("金额（元）")
    ax.legend()
    fig.tight_layout()
    out = os.path.join(OUT_DIR, "p3_forecast.png")
    fig.savefig(out, dpi=120)
    plt.close(fig)
    assert os.path.getsize(out) > 10_000, "预测图过小，疑似空白"
    print(f"[AC-P4] 预测图已出：{out}")
    return out


def run_wordcloud(df, font_path):
    import jieba
    from wordcloud import WordCloud

    text = " ".join(df["产品名称"].astype(str))
    words = [w for w in jieba.lcut(text) if len(w) >= 2 and w not in ("激光", "智能")]
    freq = dict(Counter(words).most_common(15))
    assert freq, "分词结果为空"
    wc = WordCloud(
        font_path=font_path, width=1200, height=600,
        background_color="white", random_state=42,
    ).generate_from_frequencies(freq)
    out = os.path.join(OUT_DIR, "p3_wordcloud.png")
    wc.to_file(out)
    assert os.path.getsize(out) > 10_000, "词云过小，疑似空白"
    print(f"[AC-P3] 词频 Top5：{list(freq.items())[:5]}")
    print(f"[AC-P3] 词云已出：{out}")
    return out


def main():
    xlsx = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_XLSX
    os.makedirs(OUT_DIR, exist_ok=True)
    font = probe_font()
    print(f"[字体] 探测链命中：{font or '全部缺失（fail-closed，放弃出图）'}")
    if font is None:
        print("AC-P3 出图 legs 放弃（fail-closed）；AC-P4 数值 legs 继续")
    df, monthly = load_monthly(xlsx)
    print(f"[数据] {os.path.basename(xlsx)}·销售明细：{len(df)} 行，{monthly.index[0]}~{monthly.index[-1]} 共 {len(monthly)} 月")
    chosen, point, ci = run_forecast(monthly)
    plot_forecast(monthly, point, ci, chosen, font)
    run_wordcloud(df, font)
    print("P3 验收：全部通过")


if __name__ == "__main__":
    main()
