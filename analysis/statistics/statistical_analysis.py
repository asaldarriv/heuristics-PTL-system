import os
import sys
import time
from openpyxl import Workbook
import statistics
from scipy.stats import ttest_1samp, shapiro, f_oneway, mannwhitneyu

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))

from shared.data_loader.data_loader import load_data
from shared.utils import evaluate_solution, verify_solution

from constructive_method.heuristics import nearest_neighbor_minimize_max_workload_time
from random_method.heuristics import evolutionary_one_plus_one

instances = [
    '40_homogeneous.xlsx', 
    '40_heterogeneous.xlsx', 
    '60_homogeneous.xlsx',
    '60_heterogeneous.xlsx',
    '80_homogeneous.xlsx',
    '80_heterogeneous.xlsx',
]

n_evolutionary_runs = 30
evolutionary_max_iterations = 1000

anova_groups = []

wb = Workbook()
ws = wb.active
ws.title = "hypothesis_test"
ws.append([
    'instance',
    'wmax_deterministic',
    'mean_wmax_evolutionary',
    'std_wmax_evolutionary',
    'normality_p',
    'test_used',
    'p_value',
    'result'
])

for instance in instances:
    print(f"\n🔎 Processing instance: {instance}")
    args = load_data(instance)

    # --- Deterministic method ---
    det_assignments, det_load_zones, _ = nearest_neighbor_minimize_max_workload_time(*args)
    wmax_det, _ = evaluate_solution(det_load_zones)
    print(f"Deterministic Wmax: {wmax_det}")

    # --- Evolutionary method (multiple runs) ---
    wmax_evo_list = []
    for _ in range(n_evolutionary_runs):
        assignments, load_zones, _ = evolutionary_one_plus_one(
            *args, max_iterations=evolutionary_max_iterations
        )
        wmax, _ = evaluate_solution(load_zones)
        if verify_solution(assignments, load_zones, *args[:3]):
            wmax_evo_list.append(wmax)

    if len(wmax_evo_list) < 3:
        print("⚠️ Not enough valid evolutionary runs.")
        continue

    mean_wmax = statistics.mean(wmax_evo_list)
    std_wmax = statistics.stdev(wmax_evo_list)

    # Test normality
    _, normality_p = shapiro(wmax_evo_list)
    print(f"Shapiro-Wilk p = {normality_p:.4f}")

    # Choose appropriate test
    if normality_p > 0.05:
        # Use t-test (one-sample)
        stat, p_value = ttest_1samp(wmax_evo_list, popmean=wmax_det)
        test_used = 't-test (1-sample)'
    else:
        # Use Mann-Whitney U (non-parametric)
        stat, p_value = mannwhitneyu(wmax_evo_list, [wmax_det] * len(wmax_evo_list), alternative='less')
        test_used = 'Mann-Whitney U'

    result = 'reject H0' if p_value < 0.05 else 'fail to reject H0'

    print(f"Method comparison p = {p_value:.4f} → {result}")

    # Save for ANOVA
    anova_groups.append(wmax_evo_list)

    # Write to Excel
    ws.append([
        instance,
        round(wmax_det, 2),
        round(mean_wmax, 2),
        round(std_wmax, 2),
        round(normality_p, 4),
        test_used,
        round(p_value, 4),
        result
    ])

# --- ANOVA across instances (only evolutionary method) ---
anova_ws = wb.create_sheet("anova")
anova_ws.append(['anova_across_instances'])

if len(anova_groups) >= 2:
    f_stat, p_anova = f_oneway(*anova_groups)
    anova_ws.append(['F-statistic', f_stat])
    anova_ws.append(['p-value', p_anova])
    anova_ws.append(['result', 'significant differences' if p_anova < 0.05 else 'no significant differences'])

    print(f"\n📊 ANOVA p = {p_anova:.4f}")
else:
    anova_ws.append(['Insufficient data for ANOVA'])

# Save results
wb.save("analysis/statistics/statistical_analysis.xlsx")
print("✅ Results saved to 'analysis/statistics/statistical_analysis.xlsx'")
