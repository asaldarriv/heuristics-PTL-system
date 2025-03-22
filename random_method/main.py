import os
import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))


from shared.data_loader.data_loader import load_data
from shared.reports_generation.generate_report import generate_report
from shared.utils import save_results, verify_solution

from heuristics import nearest_neighbor_minimize_max_workload_time, nearest_neighbor_minimize_max_workload_time_randomized, evolutionary_one_plus_one

INSTANCES_LIST = [
    '40_homogeneous.xlsx', 
    '40_heterogeneous.xlsx', 
    '60_homogeneous.xlsx',
    '60_heterogeneous.xlsx',
    '80_homogeneous.xlsx',
    '80_heterogeneous.xlsx',
]

def main():
    # Creates output directory
    output_directory = 'random_method/solutions'
    os.makedirs(output_directory, exist_ok=True)

    N = 1000  # Number of iterations for randomized method

    report_data_list = []

    for instance in INSTANCES_LIST:

        instance_name = instance.split('.')[0]
        base_route_file = f'{output_directory}/solution_{instance_name}'

        P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, classification_times = load_data(instance)

        args = (P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, classification_times)

        deterministic_solution = nearest_neighbor_minimize_max_workload_time(*args)
        deterministic_assignments = deterministic_solution[0]
        deterministic_load_zones = deterministic_solution[1]
        deterministic_execution_time = deterministic_solution[2]
        if verify_solution(deterministic_assignments, deterministic_load_zones, P_i, Z_j, S_k):
            save_results(deterministic_assignments, deterministic_load_zones, f'{base_route_file}_deterministic.xlsx', instance_name)
            report_data_list.append((deterministic_assignments, deterministic_load_zones, deterministic_execution_time, instance_name, 'deterministic'))

        evolutionary_solution = evolutionary_one_plus_one(*args, N)
        evolutionary_assignments = evolutionary_solution[0]
        evolutionary_load_zones = evolutionary_solution[1]
        evolutionary_execution_time = evolutionary_solution[2]
        if verify_solution(evolutionary_assignments, evolutionary_load_zones, P_i, Z_j, S_k):
            save_results(evolutionary_assignments, evolutionary_load_zones, f'{base_route_file}_evolutionary.xlsx', instance_name)
            report_data_list.append((evolutionary_assignments, evolutionary_load_zones, evolutionary_execution_time, instance_name, 'evolutionary'))

    # Generate report
    generate_report(report_data_list, 'random')

if __name__ == "__main__":
    main()