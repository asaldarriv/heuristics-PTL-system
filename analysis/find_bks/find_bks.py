import os
import sys
import time

from openpyxl import Workbook

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))

from shared.data_loader.data_loader import load_data
from shared.utils import evaluate_solution, verify_solution
from random_method.heuristics import evolutionary_one_plus_one

instances_list = [
    '40_homogeneous.xlsx', 
    '40_heterogeneous.xlsx', 
    '60_homogeneous.xlsx',
    '60_heterogeneous.xlsx',
    '80_homogeneous.xlsx',
    '80_heterogeneous.xlsx',
]

n_iterations = 1000  # Internal iterations for each randomized execution
n_runs = 500          # Total number of randomized executions per instance

def find_bks():
    wb = Workbook()
    ws = wb.active
    ws.title = "bks_results"
    ws.append(['instance', 'bks_wmax', 'wmax_wmin', 'execution_time_sec'])

    for instance in instances_list:
        print(f"Processing instance: {instance}")
        P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, classification_times = load_data(instance)

        best_wmax = float('inf')
        best_gap = float('inf')
        best_time = None

        for run in range(n_runs):
            start = time.time()
            assignments, load_zones, exec_time = evolutionary_one_plus_one(
                P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, classification_times, n_iterations
            )
            wmax, wmax_wmin = evaluate_solution(load_zones)
            total_time = time.time() - start

            if wmax < best_wmax or (wmax == best_wmax and wmax_wmin < best_gap):
                if verify_solution(assignments, load_zones, P_i, Z_j, S_k):
                    best_wmax = wmax
                    best_gap = wmax_wmin
                    best_time = total_time

        ws.append([
            instance,
            round(best_wmax, 2),
            round(best_gap, 2),
            round(best_time, 4)
        ])
        print(f"✓ BKS found for {instance}: {best_wmax} (gap: {best_gap})\n")

    # Save Excel file
    wb.save('analysis/find_bks/bks_results.xlsx')
    print("✅ Report saved to 'bks_results.xlsx'")

if __name__ == '__main__':
    find_bks()
