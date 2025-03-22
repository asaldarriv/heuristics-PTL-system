import pandas as pd
from typing import Dict, Tuple, List

def load_data(instance_name: str) -> Tuple[
        List[str], 
        List[str], 
        List[str], 
        List[str], 
        float, 
        Dict[Tuple[str, str], int], 
        Dict[Tuple[str, str], int], 
        Dict[Tuple[str, str], float], 
        Dict[Tuple[str, str], float],
        Dict[str, float]
    ]:
    """
    Load data from an Excel file containing the instance model.
    """
    excel_model = pd.ExcelFile(f'shared/instances_ptl/{instance_name}')

    P_i = list(pd.read_excel(excel_model, 'Pedidos', index_col=0).index)
    Z_j = list(pd.read_excel(excel_model, 'Zonas', index_col=0).index)
    S_k = list(pd.read_excel(excel_model, 'Salidas', index_col=0).index)
    R_m = list(pd.read_excel(excel_model, 'SKU', index_col=0).index)

    parameters = pd.read_excel(excel_model, 'Parametros', index_col=0)
    v = parameters['v'].iloc[0]

    s_jk_dataframe = pd.read_excel(excel_model, 'Salidas_pertenece_zona', index_col=0)
    rp_im_dataframe = pd.read_excel(excel_model, 'SKU_pertenece_pedido', index_col=0)
    d_jk_dataframe = pd.read_excel(excel_model, 'Tiempo_salida', index_col=0)
    tr_im_dataframe = pd.read_excel(excel_model, 'Tiempo_SKU', index_col=0)

    s_jk = {(j, k): s_jk_dataframe.at[j, k] for k in S_k for j in Z_j}
    rp_im = {(i, m): rp_im_dataframe.at[i, m] for m in R_m for i in P_i}
    d_jk = {(j, k): d_jk_dataframe.at[j, k] for k in S_k for j in Z_j}
    tr_im = {(i, m): tr_im_dataframe.at[i, m] for m in R_m for i in P_i}

    classification_times = {
        order: sum(tr_im[order, sku] for sku in R_m if rp_im.get((order, sku), 0) == 1)
        for order in P_i
    }

    return P_i, Z_j, S_k, R_m, v, s_jk, rp_im, d_jk, classification_times