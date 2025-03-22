import os
import pandas as pd
from typing import Dict, Tuple, List

def evaluate_solution(load_zones: Dict[str, float]) -> Tuple[float, float]:
    max_load = max(load_zones.values())
    min_load = min(load_zones.values())
    return max_load, max_load - min_load

def verify_solution(assignments: Dict[str, Tuple[str, str, float]], load_zones: Dict[str, float], P_i: List[str], Z_j: List[str], S_k: List[str]) -> bool:
    if len(assignments) != len(P_i):
        raise ValueError("The number of assignments does not match the number of orders.")
    
    orders = list(assignments.keys())
    assigned_exits = [assignment[1] for assignment in assignments.values()]
     
    if len(orders) != len(set(orders)):
        raise ValueError("Duplicate zones found in assignments.")
    
    if len(assigned_exits) != len(set(assigned_exits)):
        raise ValueError("Duplicate exits found in assignments.")
    
    if len(load_zones) != len(Z_j):
        raise ValueError("The number of load zones does not match the number of zones.")
    
    return True

def save_results(assignments: Dict[str, Tuple[str, str, float]], load_zones: Dict[str, float], filename: str, instance_name: str):
    max_load_zone = max(load_zones, key=load_zones.get)
    resumen_df = pd.DataFrame({
        'Instancia': [instance_name],
        'Zona': [max_load_zone],
        'Maximo': [load_zones[max_load_zone]]
    })

    solucion_df = pd.DataFrame({
        'Pedido': list(assignments.keys()),
        'Salida': [assignments[order][1] for order in assignments]
    })

    metricas_df = pd.DataFrame({
        'Zona': list(load_zones.keys()),
        'Tiempo': list(load_zones.values())
    })

    os.makedirs(os.path.dirname(filename), exist_ok=True)

    with pd.ExcelWriter(filename) as writer:
        resumen_df.to_excel(writer, sheet_name='Resumen', index=False)
        solucion_df.to_excel(writer, sheet_name='Solucion', index=False)
        metricas_df.to_excel(writer, sheet_name='Metricas', index=False)