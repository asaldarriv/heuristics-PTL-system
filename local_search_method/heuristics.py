import time
import random
from typing import Dict, Tuple, List

from shared.utils import evaluate_solution
from constructive_method.heuristics import nearest_neighbor_minimize_max_workload_time

def generate_neighbor(
        assignments: Dict[str, Tuple[str, str, float]],
        load_zones: Dict[str, float],
        P_i: List[str],
        d_jk: Dict[Tuple[str, str], float],
        R_m: List[str], 
        rp_im: Dict[Tuple[str, str], int],
        classification_times: Dict[str, float],
        v: float
    ) -> Tuple[Dict[str, Tuple[str, str, float]], Dict[str, float]]:
    """
    Generates a neighbor solution by swapping the exits of two orders.
    Updates the load zones incrementally by subtracting the old classification times
    and adding the new ones for the affected orders.
    """
    mutated_assignments = assignments.copy()
    mutated_load_zones = load_zones.copy()

    # Select two different orders randomly
    order1, order2 = random.sample(P_i, 2)

    # Get their current zones and exits
    zone_1, exit1, classification_time1 = assignments[order1]
    zone_2, exit2, classification_time2 = assignments[order2]

    # Subtract the current classification times of the affected orders from their respective zones
    mutated_load_zones[zone_1] -= classification_time1
    mutated_load_zones[zone_2] -= classification_time2

    # Swap the exits
    mutated_assignments[order1] = (zone_2, exit2, 0)  # Set temporary classification time
    mutated_assignments[order2] = (zone_1, exit1, 0)  # Set temporary classification time

    # Recalculate classification times for the affected orders
    for order in [order1, order2]:
        zone, new_exit, _ = mutated_assignments[order]
        num_skus = sum(1 for sku in R_m if rp_im.get((order, sku), 0) == 1)
        new_classification_time = classification_times[order]
        new_classification_time += num_skus * 2 * (d_jk[zone, new_exit] / v)  # Travel time per SKU
        mutated_assignments[order] = (zone, new_exit, new_classification_time)

        # Add the new classification time to the corresponding zone
        mutated_load_zones[zone] += new_classification_time

    return mutated_assignments, mutated_load_zones

def local_search_vns(
        P_i: List[str], 
        Z_j: List[str], 
        S_k: List[str], 
        R_m: List[str], 
        v: float, 
        s_jk: Dict[Tuple[str, str], int], 
        rp_im: Dict[Tuple[str, str], int], 
        d_jk: Dict[Tuple[str, str], float], 
        classification_times: Dict[str, float],
        max_iterations: int = 100,
        max_no_improve: int = 10
    ) -> Tuple[
            Dict[str, Tuple[str, str, float]], 
            Dict[str, float],
            float
        ]:
    """
    Local Search using Variable Neighborhood Search (VNS):
    - Starts with an initial solution from nearest_neighbor_minimize_max_workload_time.
    - Explores neighborhoods by mutating the current solution.
    - Selects the best solution in the neighborhood (best improvement).
    - Stops after max_iterations or if no improvement occurs for max_no_improve iterations.
    """
    start_time = time.time()

    # Generate initial solution using nearest neighbor heuristic
    current_assignments, current_load_zones, _ = nearest_neighbor_minimize_max_workload_time(
        P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, classification_times
    )
    best_assignments = current_assignments
    best_load_zones = current_load_zones
    best_wmax, _ = evaluate_solution(best_load_zones)

    no_improve_count = 0

    for _ in range(max_iterations):
        if no_improve_count >= max_no_improve:
            break

        # Generate a neighborhood and evaluate solutions
        neighborhood = [
            generate_neighbor(current_assignments, current_load_zones, P_i, d_jk, R_m, rp_im, classification_times, v)
            for _ in range(5)
        ]

        # Select the best solution in the neighborhood
        for neighbor_assignments, neighbor_load_zones in neighborhood:
            neighbor_wmax, _ = evaluate_solution(neighbor_load_zones)
            if neighbor_wmax < best_wmax:
                best_assignments = neighbor_assignments
                best_load_zones = neighbor_load_zones
                best_wmax = neighbor_wmax
                no_improve_count = 0  # Reset no improvement counter
                break
        else:
            no_improve_count += 1  # Increment no improvement counter if no better solution is found

    execution_time = time.time() - start_time

    return best_assignments, best_load_zones, execution_time