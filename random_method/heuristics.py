import time
import random
from typing import Dict, Tuple, List

from shared.utils import evaluate_solution

def nearest_neighbor_minimize_max_workload_time(
        P_i: List[str], 
        Z_j: List[str], 
        S_k: List[str], 
        R_m: List[str], 
        v: float, 
        s_jk: Dict[Tuple[str, str], int], 
        rp_im: Dict[Tuple[str, str], int], 
        d_jk: Dict[Tuple[str, str], float], 
        classification_times: Dict[str, float]
    ) -> Tuple[
            Dict[str, Tuple[str, str, float]], 
            Dict[str, float],
            float
        ]:
    """
    This heuristic minimizes Wmax.
    - Orders are sorted initially.
    - Orders are assigned to the zone with the least workload.
    - Uses the nearest neighbor rule for exit selection.
    - Ensures that all orders are assigned to an exit without skipping any.
    """

    start_time = time.time()

    assignments = {}
    load_zones = {zone: 0 for zone in Z_j}  # Initialize load per zone
    remaining_exits = S_k.copy()  # Exits to be assigned
    
    # Sort orders by the number of SKUs in descending order (highest first)
    sorted_orders = sorted(P_i, key=lambda order: sum(1 for sku in R_m if rp_im.get((order, sku), 0) == 1), reverse=True)
    
    for order in sorted_orders:
        available_zones = [z for z in Z_j if any(s_jk.get((z, k), 0) == 1 and k in remaining_exits for k in S_k)]
        if not available_zones:
            raise ValueError("No available zones for assignment. Check instance constraints.")
        
        # Select the zone with the least workload to minimize Wmax
        selected_zone = min(available_zones, key=lambda z: load_zones[z])
        
        # Find the list of available exits in the selected zone
        available_exits = [k for k in remaining_exits if s_jk.get((selected_zone, k), 0) == 1]
        
        # Find the nearest available exit in the selected zone (nearest neighbor rule)
        selected_exit = min(available_exits, key=lambda k: d_jk[selected_zone, k])
        remaining_exits.remove(selected_exit)  # Mark exit as assigned
        
        # Calculate classification time
        num_skus = sum(1 for sku in R_m if rp_im.get((order, sku), 0) == 1)
        classification_time = classification_times[order]
        classification_time += num_skus * 2 * (d_jk[selected_zone, selected_exit] / v)  # Apply travel time per SKU
        
        # Save assignment
        assignments[order] = (selected_zone, selected_exit, classification_time)
        load_zones[selected_zone] += classification_time
    
    execution_time = time.time() - start_time

    return assignments, load_zones, execution_time

def nearest_neighbor_minimize_max_workload_time_randomized(
        P_i: List[str], 
        Z_j: List[str], 
        S_k: List[str], 
        R_m: List[str], 
        v: float, 
        s_jk: Dict[Tuple[str, str], int], 
        rp_im: Dict[Tuple[str, str], int], 
        d_jk: Dict[Tuple[str, str], float], 
        classification_times: Dict[str, float],
        N: int
    ) -> Tuple[
            Dict[str, Tuple[str, str, float]], 
            Dict[str, float],
            float
        ]:
    """
    This heuristic minimizes the Wmax.
    - Orderds are randomized initially.
    - Orders are assigned to the zone with the least workload.
    - Uses the nearest neighbor rule for exit selection.
    - Ensures that all orders are assigned to an exit without skipping any.
    - Iterates N times to find the best solution.
    """

    start_time = time.time()

    best_assignments = None
    best_load_zones = None
    best_wmax = float('inf')
    best_wmax_wmin = float('inf')

    for _ in range(N):
        assignments = {}
        load_zones = {zone: 0 for zone in Z_j}  # Initialize load per zone
        remaining_exits = S_k.copy()  # Exits to be assigned
        
        # Shuffle orders randomly
        random.shuffle(P_i)
        
        for order in P_i:
            available_zones = [z for z in Z_j if any(s_jk.get((z, k), 0) == 1 and k in remaining_exits for k in S_k)]
            if not available_zones:
                raise ValueError("No available zones for assignment. Check instance constraints.")
            
            # Select the zone with the least workload to minimize Wmax
            selected_zone = min(available_zones, key=lambda z: load_zones[z])
            
            # Find the list of available exits in the selected zone
            available_exits = [k for k in remaining_exits if s_jk.get((selected_zone, k), 0) == 1]
            
            # Find the nearest available exit in the selected zone (nearest neighbor rule)
            selected_exit = min(available_exits, key=lambda k: d_jk[selected_zone, k])
            remaining_exits.remove(selected_exit)  # Mark exit as assigned
            
            # Calculate classification time
            num_skus = sum(1 for sku in R_m if rp_im.get((order, sku), 0) == 1)
            classification_time = classification_times[order]
            classification_time += num_skus * 2 * (d_jk[selected_zone, selected_exit] / v)  # Apply travel time per SKU
            
            # Save assignment
            assignments[order] = (selected_zone, selected_exit, classification_time)
            load_zones[selected_zone] += classification_time
        
        wmax, wmax_wmin = evaluate_solution(load_zones)
        
        if wmax < best_wmax or (wmax == best_wmax and wmax_wmin < best_wmax_wmin):
            best_assignments = assignments
            best_load_zones = load_zones
            best_wmax = wmax
            best_wmax_wmin = wmax_wmin
    
    
    execution_time = time.time() - start_time

    return best_assignments, best_load_zones, execution_time

def mutate_solution(
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
    Mutates the current solution by swapping the exits of two orders.
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

def evolutionary_one_plus_one(
        P_i: List[str], 
        Z_j: List[str], 
        S_k: List[str], 
        R_m: List[str], 
        v: float, 
        s_jk: Dict[Tuple[str, str], int], 
        rp_im: Dict[Tuple[str, str], int], 
        d_jk: Dict[Tuple[str, str], float], 
        classification_times: Dict[str, float],
        max_iterations: int = 100
    ) -> Tuple[
            Dict[str, Tuple[str, str, float]], 
            Dict[str, float],
            float
        ]:
    """
    Implements (1+1) Evolutionary Strategy:
    - Starts with a greedy solution.
    - Mutates the solution at each iteration.
    - Accepts mutations if they improve Wmax.
    - Stops after max_iterations or if no improvement occurs.
    """

    start_time = time.time()

    # Generate initial solution using nearest neighbor heuristic
    current_assignments, current_load_zones, _ = nearest_neighbor_minimize_max_workload_time(
        P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, classification_times
    )

    best_assignments = current_assignments
    best_load_zones = current_load_zones
    best_wmax, _ = evaluate_solution(best_load_zones)

    for _ in range(max_iterations):
        # Mutate the current solution
        new_assignments, new_load_zones = mutate_solution(
            best_assignments, best_load_zones, P_i, d_jk, R_m,
            rp_im, classification_times, v
        )

        # Evaluate new solution
        new_wmax, _ = evaluate_solution(new_load_zones)

        # Accept mutation if it improves Wmax
        if new_wmax < best_wmax:
            best_assignments = new_assignments
            best_load_zones = new_load_zones
            best_wmax = new_wmax

    execution_time = time.time() - start_time

    return best_assignments, best_load_zones, execution_time