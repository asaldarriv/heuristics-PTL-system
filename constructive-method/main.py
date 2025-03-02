import pandas as pd
import random
from typing import Dict, Tuple, List

# Name of the file with the problem instance
INSTANCE_NAME = 'Data_40_Salidas_composición_zonas_heterogéneas.xlsx'

def load_data(instance_name: str) -> Tuple[List[str], List[str], List[str], List[str], float, Dict[Tuple[str, str], int], Dict[Tuple[str, str], int], Dict[Tuple[str, str], float], Dict[Tuple[str, str], float], Dict[str, int]]:
    # Load the Excel file
    excel_model = pd.ExcelFile(f'instances-ptl/{instance_name}')

    # Load the sets
    P_i = list(pd.read_excel(excel_model, 'Pedidos', index_col=0).index)
    Z_j = list(pd.read_excel(excel_model, 'Zonas', index_col=0).index)
    S_k = list(pd.read_excel(excel_model, 'Salidas', index_col=0).index)
    R_m = list(pd.read_excel(excel_model, 'SKU', index_col=0).index)

    # Load parameters
    parameters = pd.read_excel(excel_model, 'Parametros', index_col=0)
    v = parameters['v'].iloc[0]

    # Load relationships and parameters
    s_jk_dataframe = pd.read_excel(excel_model, 'Salidas_pertenece_zona', index_col=0)
    rp_im_dataframe = pd.read_excel(excel_model, 'SKU_pertenece_pedido', index_col=0)
    d_jk_dataframe = pd.read_excel(excel_model, 'Tiempo_salida', index_col=0)
    tr_im_dataframe = pd.read_excel(excel_model, 'Tiempo_SKU', index_col=0)
    ns_j_dataframe = pd.read_excel(excel_model, 'Salidas_en_cada_zona', index_col=0)

    # Convert to appropriate data structures
    s_jk: Dict[Tuple[str, str], int] = {(j, k): s_jk_dataframe.at[j, k] for k in S_k for j in Z_j}
    rp_im: Dict[Tuple[str, str], int] = {(i, m): rp_im_dataframe.at[i, m] for m in R_m for i in P_i}
    d_jk: Dict[Tuple[str, str], float] = {(j, k): d_jk_dataframe.at[j, k] for k in S_k for j in Z_j}
    tr_im: Dict[Tuple[str, str], float] = {(i, m): tr_im_dataframe.at[i, m] for m in R_m for i in P_i}
    ns_j: Dict[str, int] = ns_j_dataframe.Num_Salidas.to_dict()

    return P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, tr_im, ns_j

def deterministic_method_nearest_neighbor(P_i: List[str], Z_j: List[str], S_k: List[str], R_m: List[str], v: float, s_jk: Dict[Tuple[str, str], int], rp_im: Dict[Tuple[str, str], int], d_jk: Dict[Tuple[str, str], float], tr_im: Dict[Tuple[str, str], float]) -> Tuple[Dict[str, Tuple[str, str, float]], Dict[str, float]]:
    assignments = {}
    load_zones = {zone: 0 for zone in Z_j}  # Initialize load per zone
    remaining_exits = S_k.copy()  # Exits to be assigned
    
    # Sort orders by the number of SKUs in descending order (highest first)
    sorted_orders = sorted(P_i, key=lambda order: sum(1 for sku in R_m if rp_im.get((order, sku), 0) == 1), reverse=True)
    
    for order in P_i:
        # Find the nearest available exit, regardless of the zone
        available_exits = [(j, k) for j in Z_j for k in remaining_exits if s_jk.get((j, k), 0) == 1]
        if not available_exits:
            continue
        
        # Apply nearest neighbor rule
        selected_zone, selected_exit = min(available_exits, key=lambda jk: d_jk[jk])
        remaining_exits.remove(selected_exit)  # Mark exit as assigned
        
        # Calculate classification time
        num_skus = sum(1 for sku in R_m if rp_im.get((order, sku), 0) == 1)  # Count SKUs in the order
        classification_time = sum(tr_im[order, sku] for sku in R_m if rp_im.get((order, sku), 0) == 1)
        classification_time += num_skus * 2 * (d_jk[selected_zone, selected_exit] / v)  # Apply travel time per SKU
        
        # Save assignment
        assignments[order] = (selected_zone, selected_exit, classification_time)
        load_zones[selected_zone] += classification_time
    
    return assignments, load_zones

def deterministic_method_minimize_difference(P_i: List[str], Z_j: List[str], S_k: List[str], R_m: List[str], v: float, s_jk: Dict[Tuple[str, str], int], rp_im: Dict[Tuple[str, str], int], d_jk: Dict[Tuple[str, str], float], tr_im: Dict[Tuple[str, str], float]) -> Tuple[Dict[str, Tuple[str, str, float]], float]:
    assignments = {}
    load_zones = {zone: 0 for zone in Z_j}  # Initialize load per zone
    remaining_exits = S_k.copy()  # Exits to be assigned

    for order in P_i:
        # Select the zone that minimizes the difference between max and min workloads
        selected_zone = min(load_zones, key=lambda z: max(load_zones.values()) - load_zones[z])
        
        # Find the nearest available exit in the selected zone
        available_exits = [k for k in remaining_exits if s_jk.get((selected_zone, k), 0) == 1]
        if not available_exits:
            continue
        selected_exit = min(available_exits, key=lambda k: d_jk[selected_zone, k])
        remaining_exits.remove(selected_exit)  # Mark exit as assigned
        
        # Calculate classification time
        classification_time = sum(tr_im[order, sku] for sku in R_m if rp_im.get((order, sku), 0) == 1)
        classification_time += 2 * (d_jk[selected_zone, selected_exit] / v)
        
        # Save assignment
        assignments[order] = (selected_zone, selected_exit, classification_time)
        load_zones[selected_zone] += classification_time
    
    # Compute the difference between max and min workload
    time_difference = max(load_zones.values()) - min(load_zones.values())
    
    return assignments, load_zones

def deterministic_method_minimize_maximum_time(P_i: List[str], Z_j: List[str], S_k: List[str], R_m: List[str], v: float, s_jk: Dict[Tuple[str, str], int], rp_im: Dict[Tuple[str, str], int], d_jk: Dict[Tuple[str, str], float], tr_im: Dict[Tuple[str, str], float]) -> Tuple[Dict[str, Tuple[str, str, float]], float]:
    assignments = {}
    load_zones = {zone: 0 for zone in Z_j}  # Initialize load per zone
    remaining_exits = S_k.copy()  # Exits to be assigned
    unassigned_orders = set(P_i)  # Keep track of unassigned orders

    while unassigned_orders:
        order = unassigned_orders.pop()
        
        # Find the zone that has the lowest max workload impact
        selected_zone = min(load_zones, key=lambda z: max(load_zones.values()) - load_zones[z])
        
        # Find the nearest available exit in the selected zone
        available_exits = [k for k in remaining_exits if s_jk.get((selected_zone, k), 0) == 1]
        
        if not available_exits:
            # If no available exits in the selected zone, search in another zone
            alternative_zones = [z for z in Z_j if any(s_jk.get((z, k), 0) == 1 for k in remaining_exits)]
            if not alternative_zones:
                raise ValueError("No available exits remaining for some orders. Check input consistency.")
            selected_zone = min(alternative_zones, key=lambda z: load_zones[z])
            available_exits = [k for k in remaining_exits if s_jk.get((selected_zone, k), 0) == 1]
        
        selected_exit = min(available_exits, key=lambda k: d_jk[selected_zone, k])
        remaining_exits.remove(selected_exit)  # Mark exit as assigned
        
        # Calculate classification time
        classification_time = sum(tr_im[order, sku] for sku in R_m if rp_im.get((order, sku), 0) == 1)
        classification_time += 2 * (d_jk[selected_zone, selected_exit] / v)
        
        # Save assignment
        assignments[order] = (selected_zone, selected_exit, classification_time)
        load_zones[selected_zone] += classification_time
    
    return assignments, load_zones


def randomized_method(P_i: List[str], Z_j: List[str], S_k: List[str], R_m: List[str], v: float, s_jk: Dict[Tuple[str, str], int], rp_im: Dict[Tuple[str, str], int], d_jk: Dict[Tuple[str, str], float], tr_im: Dict[Tuple[str, str], float]) -> Tuple[Dict[str, Tuple[str, str, float]], Dict[str, float]]:
    assignments = {}
    load_zones = {zone: 0 for zone in Z_j}  # Initialize load per zone
    random_orders = random.sample(P_i, len(P_i))  # Random order of orders
    
    for order in random_orders:
        # Select a zone with inverse probability to its load
        zones_prob = sorted(Z_j, key=lambda z: load_zones[z])
        selected_zone = random.choices(zones_prob, weights=[1/load_zones[z] if load_zones[z] > 0 else 1 for z in zones_prob])[0]
        
        # Select a random exit in the zone that has not been assigned yet, prioritizing the nearest ones
        assigned_exits = {assignment[1] for assignment in assignments.values()}
        available_exits = [k for k in S_k if s_jk.get((selected_zone, k), 0) == 1 and k not in assigned_exits]
        if not available_exits:
            continue  # Skip if no available exits in the selected zone
        
        selected_exit = random.choices(available_exits, weights=[1/d_jk[selected_zone, k] for k in available_exits])[0]
        
        # Calculate classification time
        classification_time = sum(tr_im[order, sku] for sku in R_m if rp_im.get((order, sku), 0) == 1)
        classification_time += 2 * (d_jk[selected_zone, selected_exit] / v)
        
        # Save assignment
        assignments[order] = (selected_zone, selected_exit, classification_time)
        load_zones[selected_zone] += classification_time
    
    return assignments, load_zones

def save_results(assignments: Dict[str, Tuple[str, str, float]], load_zones: Dict[str, float], filename: str, instance_name: str):
    # Create DataFrame for the "Resumen" sheet
    max_load_zone = max(load_zones, key=load_zones.get)
    # Calculate average load time per zone
    avg_load_zone = sum(load_zones.values()) / len(load_zones)
    print(f"Results for {filename}: ")
    print(f"Max load zone: {max_load_zone}: {load_zones[max_load_zone]}")
    print(f"Average load time per zone: {avg_load_zone}")
    resumen_df = pd.DataFrame({
        'Instancia': [instance_name],
        'Zona': [max_load_zone],
        'Maximo': [load_zones[max_load_zone]]
    })

    # Create DataFrame for the "Solucion" sheet
    solucion_df = pd.DataFrame({
        'Pedido': list(assignments.keys()),
        'Salida': [assignments[order][1] for order in assignments]
    })

    # Create DataFrame for the "Metricas" sheet
    metricas_df = pd.DataFrame({
        'Zona': list(load_zones.keys()),
        'Tiempo': list(load_zones.values())
    })

    # Save to an Excel file with multiple sheets
    with pd.ExcelWriter(filename) as writer:
        resumen_df.to_excel(writer, sheet_name='Resumen', index=False)
        solucion_df.to_excel(writer, sheet_name='Solucion', index=False)
        metricas_df.to_excel(writer, sheet_name='Metricas', index=False)

def main():
    P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, tr_im, ns_j = load_data(INSTANCE_NAME)
    
    deterministic_assignments, deterministic_load_zones = deterministic_method_nearest_neighbor(P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, tr_im)
    save_results(deterministic_assignments, deterministic_load_zones, f'solution_not_sorted_{INSTANCE_NAME}', INSTANCE_NAME)

    # deterministic_assignments, deterministic_load_zones = randomized_method(P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, tr_im)
    # save_results(deterministic_assignments, deterministic_load_zones, 'solution_randomized_method.xlsx', INSTANCE_NAME)

    # deterministic_assignments, deterministic_load_zones = deterministic_method_minimize_difference(P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, tr_im)
    # save_results(deterministic_assignments, deterministic_load_zones, 'solution_constructive_method_minimize_difference.xlsx', INSTANCE_NAME)

    # deterministic_assignments, deterministic_load_zones = deterministic_method_minimize_maximum_time(P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, tr_im)
    # save_results(deterministic_assignments, deterministic_load_zones, 'solution_constructive_method_minimize_maximum_time.xlsx', INSTANCE_NAME)
    
    # randomized_assignments, randomized_load_zones = randomized_method(P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, tr_im)
    # save_results(randomized_assignments, randomized_load_zones, 'solution_randomized_method.xlsx', INSTANCE_NAME)

if __name__ == "__main__":
    main()