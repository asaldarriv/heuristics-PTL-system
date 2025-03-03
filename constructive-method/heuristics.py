import time
import random
from typing import Dict, Tuple, List

from utils import evaluate_solution

def nearest_neighbor_minimize_avg_time(
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
    This heuristic minimizes the average classification time across all zones.
    - Orderds are sorted initially.
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
        
        # Select the zone with the least workload to balance the average time across zones
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

def nearest_neighbor_minimize_avg_time_randomized(
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
    This heuristic minimizes the average classification time across all zones.
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
            
            # Select the zone that minimizes with less workload
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