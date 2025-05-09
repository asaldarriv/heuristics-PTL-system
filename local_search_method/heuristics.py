import time
import random
from typing import Dict, Tuple, List

from shared.utils import evaluate_solution
from constructive_method.heuristics import nearest_neighbor_minimize_max_workload_time_randomized

def generate_aggressive_neighbor(
        assignments: Dict[str, Tuple[str, str, float]],
        load_zones: Dict[str, float],
        P_i: List[str],
        d_jk: Dict[Tuple[str, str], float],
        R_m: List[str], 
        rp_im: Dict[Tuple[str, str], int],
        classification_times: Dict[str, float],
        v: float,
        num_changes: int = 3
    ) -> Tuple[Dict[str, Tuple[str, str, float]], Dict[str, float]]:
    """
    Generates a more aggressive neighbor solution by swapping the exits of multiple orders.
    Updates the load zones incrementally by subtracting the old classification times
    and adding the new ones for the affected orders.
    """
    mutated_assignments = assignments.copy()
    mutated_load_zones = load_zones.copy()

    # Select multiple orders randomly
    orders_to_change = random.sample(P_i, num_changes)

    # Store the original zones and exits
    original_data = {order: assignments[order] for order in orders_to_change}

    # Perform swaps among the selected orders
    for i, order in enumerate(orders_to_change):
        next_order = orders_to_change[(i + 1) % len(orders_to_change)]
        zone, exit_, classification_time = original_data[next_order]
        mutated_assignments[order] = (zone, exit_, 0)  # Set temporary classification time

        # Subtract the old classification time
        mutated_load_zones[assignments[order][0]] -= assignments[order][2]

    # Recalculate classification times for the affected orders
    for order in orders_to_change:
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
        max_no_improve: int = 10,
        initial_neighborhood_size: int = 5,
        max_neighborhood_size: int = 10,
        num_changes: int = 3
    ) -> Tuple[
            Dict[str, Tuple[str, str, float]], 
            Dict[str, float],
            float
        ]:
    """
    Improved VNS with hierarchical neighborhoods and adaptive parameters:
    - Starts with an initial solution from nearest_neighbor_minimize_max_workload_time_randomized.
    - Dynamically adjusts neighborhood size and explores hierarchical neighborhoods.
    - Selects the best solution in the neighborhood (best improvement).
    - Stops after max_iterations or if no improvement occurs for max_no_improve iterations.
    """
    start_time = time.time()

    # Generate initial solution using randomized nearest neighbor heuristic
    current_assignments, current_load_zones, _ = nearest_neighbor_minimize_max_workload_time_randomized(
        P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, classification_times, 1000
    )
    best_assignments = current_assignments
    best_load_zones = current_load_zones
    best_wmax, _ = evaluate_solution(best_load_zones)

    no_improve_count = 0
    neighborhood_size = initial_neighborhood_size

    for _ in range(max_iterations):
        if no_improve_count >= max_no_improve:
            break

        # Generate a neighborhood with the current size
        neighborhood = [
            generate_aggressive_neighbor(current_assignments, current_load_zones, P_i, d_jk, R_m, rp_im, classification_times, v, num_changes)
            for _ in range(neighborhood_size)
        ]

        # Explicitly find the best solution in the neighborhood
        best_neighbor_assignments = None
        best_neighbor_load_zones = None
        best_neighbor_wmax = float('inf')

        for neighbor_assignments, neighbor_load_zones in neighborhood:
            neighbor_wmax, _ = evaluate_solution(neighbor_load_zones)
            if neighbor_wmax < best_neighbor_wmax:
                best_neighbor_assignments = neighbor_assignments
                best_neighbor_load_zones = neighbor_load_zones
                best_neighbor_wmax = neighbor_wmax

        # If the best neighbor improves the current solution, accept it
        if best_neighbor_wmax < best_wmax:
            best_assignments = best_neighbor_assignments
            best_load_zones = best_neighbor_load_zones
            best_wmax = best_neighbor_wmax
            no_improve_count = 0  # Reset no improvement counter
            neighborhood_size = initial_neighborhood_size  # Reset neighborhood size
        else:
            no_improve_count += 1  # Increment no improvement counter if no better solution is found
            neighborhood_size = min(neighborhood_size + 1, max_neighborhood_size)  # Expand neighborhood size

    execution_time = time.time() - start_time

    return best_assignments, best_load_zones, execution_time