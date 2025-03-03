import os
import random
import pandas as pd
from typing import Dict, Tuple, List

# Name of the file with the problem instance
INSTANCES_LIST = [
    'Data_40_Salidas_composición_zonas_heterogéneas.xlsx', 
    'Data_40_Salidas_composición_zonas_homogéneas.xlsx', 
    # 'Data_60_Salidas_composición_zonas_heterogéneas.xlsx', -- Contains 64 exits and 60 orders
    'Data_60_Salidas_composición_zonas_homogéneas.xlsx',
    'Data_80_Salidas_composición_zonas_heterogéneas.xlsx',
    'Data_80_Salidas_composición_zonas_homogéneas.xlsx',
    ]
    
def load_data(instance_name: str) -> Tuple[List[str], List[str], List[str], List[str], float, Dict[Tuple[str, str], int], Dict[Tuple[str, str], int], Dict[Tuple[str, str], float], Dict[Tuple[str, str], float]]:
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

    # Convert to appropriate data structures
    s_jk: Dict[Tuple[str, str], int] = {(j, k): s_jk_dataframe.at[j, k] for k in S_k for j in Z_j}
    rp_im: Dict[Tuple[str, str], int] = {(i, m): rp_im_dataframe.at[i, m] for m in R_m for i in P_i}
    d_jk: Dict[Tuple[str, str], float] = {(j, k): d_jk_dataframe.at[j, k] for k in S_k for j in Z_j}
    tr_im: Dict[Tuple[str, str], float] = {(i, m): tr_im_dataframe.at[i, m] for m in R_m for i in P_i}

    return P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, tr_im

def nearest_neighbor_minimize_avg_time(P_i: List[str], Z_j: List[str], S_k: List[str], R_m: List[str], v: float, s_jk: Dict[Tuple[str, str], int], rp_im: Dict[Tuple[str, str], int], d_jk: Dict[Tuple[str, str], float], tr_im: Dict[Tuple[str, str], float]) -> Tuple[Dict[str, Tuple[str, str, float]], Dict[str, float]]:
    """
    This heuristic minimizes the average classification time across all zones.
    - Orders are assigned to the zone with the least workload.
    - Uses the nearest neighbor rule for exit selection.
    - Ensures that all orders are assigned to an exit without skipping any.
    """    
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
        classification_time = sum(tr_im[order, sku] for sku in R_m if rp_im.get((order, sku), 0) == 1)
        classification_time += num_skus * 2 * (d_jk[selected_zone, selected_exit] / v)  # Apply travel time per SKU
        
        # Save assignment
        assignments[order] = (selected_zone, selected_exit, classification_time)
        load_zones[selected_zone] += classification_time
    
    return assignments, load_zones

def nearest_neighbor_minimize_avg_time_randomized(P_i: List[str], Z_j: List[str], S_k: List[str], R_m: List[str], v: float, s_jk: Dict[Tuple[str, str], int], rp_im: Dict[Tuple[str, str], int], d_jk: Dict[Tuple[str, str], float], tr_im: Dict[Tuple[str, str], float]) -> Tuple[Dict[str, Tuple[str, str, float]], Dict[str, float]]:
    """
    This heuristic minimizes the difference between max and min classification time (W_max - W_min).
    - Orders are assigned to the zone that balances the max-min workload difference.
    - Uses the nearest neighbor rule for exit selection.
    - Ensures that all orders are assigned to an exit without skipping any.
    """     
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
        classification_time = sum(tr_im[order, sku] for sku in R_m if rp_im.get((order, sku), 0) == 1)
        classification_time += num_skus * 2 * (d_jk[selected_zone, selected_exit] / v)  # Apply travel time per SKU
        
        # Save assignment
        assignments[order] = (selected_zone, selected_exit, classification_time)
        load_zones[selected_zone] += classification_time
    
    return assignments, load_zones

def evaluate_solution(load_zones: Dict[str, float]) -> Tuple[float, float]:
    """
    Evaluates the solution by calculating Wmax and Wmax - Wmin.
    """
    max_load = max(load_zones.values())
    min_load = min(load_zones.values())
    return max_load, max_load - min_load

def save_results(assignments: Dict[str, Tuple[str, str, float]], load_zones: Dict[str, float], filename: str, instance_name: str):
    # Create DataFrame for the "Resumen" sheet
    max_load_zone = max(load_zones, key=load_zones.get)
    # Calculate average load time per zone
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

    # Ensure the directory exists
    os.makedirs(os.path.dirname(filename), exist_ok=True)

    # Save to an Excel file with multiple sheets
    with pd.ExcelWriter(filename) as writer:
        resumen_df.to_excel(writer, sheet_name='Resumen', index=False)
        solucion_df.to_excel(writer, sheet_name='Solucion', index=False)
        metricas_df.to_excel(writer, sheet_name='Metricas', index=False)

def save_comparative_results(writer: pd.ExcelWriter, assignments: Dict[str, Tuple[str, str, float]], load_zones: Dict[str, float], instance_name: str, method_name: str):
    # Create DataFrame for the "Resumen" sheet
    max_load_zone = max(load_zones, key=load_zones.get)
    avg_load = sum(load_zones.values()) / len(load_zones)
    min_load = min(load_zones.values())
    max_load = load_zones[max_load_zone]
    diff_load = max_load - min_load

    resumen_df = pd.DataFrame({
        'Instancia': [instance_name],
        'Metodo': [method_name],
        'Zona': [max_load_zone],
        'Maximo': [max_load],
        'Promedio': [avg_load],
        'Diferencia': [diff_load]
    })

    # Create DataFrame for the "Solucion" sheet
    solucion_df = pd.DataFrame({
        'Instancia': [instance_name] * len(assignments),
        'Metodo': [method_name] * len(assignments),
        'Pedido': list(assignments.keys()),
        'Salida': [assignments[order][1] for order in assignments]
    })

    # Create DataFrame for the "Metricas" sheet
    metricas_df = pd.DataFrame({
        'Instancia': [instance_name] * len(load_zones),
        'Metodo': [method_name] * len(load_zones),
        'Zona': list(load_zones.keys()),
        'Tiempo': list(load_zones.values())
    })

    # Append data to the Excel file
    resumen_df.to_excel(writer, sheet_name='Resumen', index=False, header=writer.sheets['Resumen'].max_row == 0, startrow=writer.sheets['Resumen'].max_row)
    solucion_df.to_excel(writer, sheet_name='Solucion', index=False, header=writer.sheets['Solucion'].max_row == 0, startrow=writer.sheets['Solucion'].max_row)
    metricas_df.to_excel(writer, sheet_name='Metricas', index=False, header=writer.sheets['Metricas'].max_row == 0, startrow=writer.sheets['Metricas'].max_row)

def verify_solution(assignments: Dict[str, Tuple[str, str, float]], load_zones: Dict[str, float], P_i: List[str], Z_j: List[str], S_k: List[str]) -> bool:
    """
    Verifies the solution of the heuristic methods.
    
    - For assignments, checks if the length is equal to the length of P_i and S_k.
    - Ensures no element in assignments[0] or assignments[1] is repeated.
    - For load_zones, checks if the length is equal to the length of Z_j.
    
    Returns True if all conditions are met, raises an error otherwise.
    """
    # Check if the length of assignments is equal to the length of P_i
    if len(assignments) != len(P_i):
        raise ValueError("The number of assignments does not match the number of orders.")
    
    # Check if the length of assignments is equal to the length of S_k
    if len(assignments) != len(S_k):
        raise ValueError("The number of assignments does not match the number of exits.")
    
    # Check for duplicate elements in orders and assignments[1]
    orders = list(assignments.keys())
    assigned_exits = [assignment[1] for assignment in assignments.values()]
     
    if len(orders) != len(set(orders)):
        raise ValueError("Duplicate zones found in assignments.")
    
    if len(assigned_exits) != len(set(assigned_exits)):
        raise ValueError("Duplicate exits found in assignments.")
    
    # Check if the length of load_zones is equal to the length of Z_j
    if len(load_zones) != len(Z_j):
        raise ValueError("The number of load zones does not match the number of zones.")
    
    return True

def main():
    output_dir = 'constructive-method/solutions'
    os.makedirs(output_dir, exist_ok=True)
    output_file = f'{output_dir}/comparative_results.xlsx'

    N = 10  # Number of iterations for randomized methods

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Create empty sheets with headers
        pd.DataFrame(columns=['Instancia', 'Metodo', 'Zona', 'Maximo', 'Promedio', 'Diferencia']).to_excel(writer, sheet_name='Resumen', index=False)
        pd.DataFrame(columns=['Instancia', 'Metodo', 'Pedido', 'Salida']).to_excel(writer, sheet_name='Solucion', index=False)
        pd.DataFrame(columns=['Instancia', 'Metodo', 'Zona', 'Tiempo']).to_excel(writer, sheet_name='Metricas', index=False)

        # Iterate over all instances
        for instance_name in INSTANCES_LIST:
            P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, tr_im = load_data(instance_name)

            best_assignments = None
            best_load_zones = None
            best_wmax = float('inf')
            best_wmax_wmin = float('inf')

            for _ in range(N):
                assignments, load_zones = nearest_neighbor_minimize_avg_time_randomized(P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, tr_im)
                wmax, wmax_wmin = evaluate_solution(load_zones)
                
                if wmax < best_wmax or (wmax == best_wmax and wmax_wmin < best_wmax_wmin):
                    best_assignments = assignments
                    best_load_zones = load_zones
                    best_wmax = wmax
                    best_wmax_wmin = wmax_wmin

            if verify_solution(best_assignments, best_load_zones, P_i, Z_j, S_k):
                save_results(best_assignments, best_load_zones, f'{output_dir}/{instance_name}_nearest_neighbor_minimize_avg_time_randomized.xlsx', instance_name)
                save_comparative_results(writer, best_assignments, best_load_zones, instance_name, 'nearest_neighbor_minimize_avg_time_randomized')

            # Save results for nearest_neighbor_minimize_difference_time
            deterministic_assignments, deterministic_load_zones = nearest_neighbor_minimize_avg_time(P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, tr_im)
            if verify_solution(deterministic_assignments, deterministic_load_zones, P_i, Z_j, S_k):
                save_results(deterministic_assignments, deterministic_load_zones, f'{output_dir}/{instance_name}_nearest_neighbor_minimize_avg_time.xlsx', instance_name)
                save_comparative_results(writer, deterministic_assignments, deterministic_load_zones, instance_name, 'nearest_neighbor_minimize_avg_time')


if __name__ == "__main__":
    main()