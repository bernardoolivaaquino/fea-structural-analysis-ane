import pandas as pd
import numpy as np
import openpyxl
import os
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# 1. CONFIGURACIÓN DEL USUARIO
# ==========================================
E_val = 30e6    # Módulo de Young
A_val = 8       # Área de la sección transversal

nodes_list = [
    (1, 0, 0), (2, 55.15384614, 0), (3, 27.46153846, 27.46153846), (4, 41.30769231, 33.23076924),(5,27.46153846,66.46153846)
]

elements_list = [
    (1, 1, 3), (2, 3, 2), (3, 2, 4), (4, 3, 4), (5, 3, 5),(6,4,5)
]

loads_data = {
    3: (300, 0),
    5: (300, 0)
}

supports_data = {
    2: 'fijo',      
    1: 'fijo'  
}

base_filename = "Calculo_Armadura_Con_Reacciones.xlsx"

# ==========================================
# 2. PRE-PROCESAMIENTO
# ==========================================
df_nodes = pd.DataFrame(nodes_list, columns=['Nodo', 'x', 'y'])
df_elems = pd.DataFrame(elements_list, columns=['Elemento', 'Nodo_i', 'Nodo_j'])
n_nodes = len(nodes_list)
n_dof = n_nodes * 2 

fixed_dofs = []
for node, type_appui in supports_data.items():
    idx = node - 1
    if type_appui == 'fijo':
        fixed_dofs.extend([2*idx, 2*idx+1])
    elif type_appui == 'rodillo_x':
        fixed_dofs.append(2*idx)
    elif type_appui == 'rodillo_y':
        fixed_dofs.append(2*idx+1)

fixed_dofs = sorted(list(set(fixed_dofs)))
free_dofs = [i for i in range(n_dof) if i not in fixed_dofs]

F_vector = np.zeros(n_dof)
for node, (fx, fy) in loads_data.items():
    idx = node - 1
    F_vector[2*idx] = fx
    F_vector[2*idx+1] = fy

# ==========================================
# 3. GENERACIÓN DE EXCEL (UNA SOLA HOJA)
# ==========================================
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Cálculo Completo"

# --- ESTILOS ---
style_header = Font(bold=True, color="FFFFFF")
style_fill_header = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
style_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
style_bold = Font(bold=True)
fmt_sci = '0.0000E+00'
fmt_num = '0.0000'

def write_header(ws, row, headers):
    for i, h in enumerate(headers):
        cell = ws.cell(row, i+1, h)
        cell.font = style_header
        cell.fill = style_fill_header
        cell.alignment = Alignment(horizontal='center')

current_row = 1

# ---------------------------------------------------------
# BLOQUE 1: PARÁMETROS Y NODOS
# ---------------------------------------------------------
ws.cell(current_row, 1, "1. PARÁMETROS GENERALES").font = style_bold
current_row += 1
write_header(ws, current_row, ["Parámetro", "Valor", "Unidad"])
current_row += 1
ws.append(["Módulo Young (E)", E_val, "Psi/Pa"])
ws.cell(current_row, 2).number_format = fmt_sci
cell_E_ref = f"$B${current_row}" 
current_row += 1
ws.append(["Área Sección (A)", A_val, "in²/m²"])
ws.cell(current_row, 2).number_format = fmt_sci
cell_A_ref = f"$B${current_row}" 
current_row += 2

ws.cell(current_row, 1, "2. NODOS").font = style_bold
current_row += 1
write_header(ws, current_row, ["Nodo ID", "Coord X", "Coord Y"])
current_row += 1

node_row_map = {} 
for i, row in df_nodes.iterrows():
    ws.append([row['Nodo'], row['x'], row['y']])
    node_row_map[row['Nodo']] = current_row
    ws.cell(current_row, 2).number_format = fmt_num
    ws.cell(current_row, 3).number_format = fmt_num
    current_row += 1

# ---------------------------------------------------------
# BLOQUE 2: GEOMETRÍA DE ELEMENTOS
# ---------------------------------------------------------
current_row += 1
ws.cell(current_row, 1, "3. CÁLCULOS GEOMÉTRICOS (ELEMENTOS)").font = style_bold
current_row += 1
headers_geom = ["Elem", "Ni", "Nj", "xi", "yi", "xj", "yj", "dx", "dy", "L", "cx", "cy", "k"]
write_header(ws, current_row, headers_geom)
current_row += 1

elem_data_map = {} 

for i, row in df_elems.iterrows():
    el_id = row['Elemento']; ni, nj = row['Nodo_i'], row['Nodo_j']
    r = current_row
    
    row_ni = node_row_map[ni]
    row_nj = node_row_map[nj]
    
    ws.cell(r, 1, el_id)
    ws.cell(r, 2, ni)
    ws.cell(r, 3, nj)
    ws.cell(r, 4, f"=B{row_ni}") # xi
    ws.cell(r, 5, f"=C{row_ni}") # yi
    ws.cell(r, 6, f"=B{row_nj}") # xj
    ws.cell(r, 7, f"=C{row_nj}") # yj
    ws.cell(r, 8, f"=F{r}-D{r}") # dx
    ws.cell(r, 9, f"=G{r}-E{r}") # dy
    ws.cell(r, 10, f"=SQRT(H{r}^2 + I{r}^2)") # L
    ws.cell(r, 11, f"=H{r}/J{r}") # cx
    ws.cell(r, 12, f"=I{r}/J{r}") # cy
    ws.cell(r, 13, f"=({cell_E_ref}*{cell_A_ref})/J{r}") # k = EA/L
    
    for c in range(4, 14): ws.cell(r, c).number_format = fmt_sci
    
    elem_data_map[el_id] = {
        'k': f"M{r}", 'cx': f"K{r}", 'cy': f"L{r}", 'row': r
    }
    current_row += 1

# ---------------------------------------------------------
# BLOQUE 3: MATRICES LOCALES
# ---------------------------------------------------------
current_row += 1
ws.cell(current_row, 1, "4. MATRICES DE RIGIDEZ LOCAL (k_local)").font = style_bold
current_row += 1

local_matrices_map = {} 

for i, row in df_elems.iterrows():
    el_id = row['Elemento']
    ws.cell(current_row, 1, f"Elemento {el_id}").font = Font(italic=True, bold=True)
    
    k = elem_data_map[el_id]['k']
    cx = elem_data_map[el_id]['cx']
    cy = elem_data_map[el_id]['cy']
    
    start_row = current_row + 1
    forms = [[f"={k}*{cx}^2", f"={k}*{cx}*{cy}", f"=-{k}*{cx}^2", f"=-{k}*{cx}*{cy}"],
             [f"={k}*{cx}*{cy}", f"={k}*{cy}^2", f"=-{k}*{cx}*{cy}", f"=-{k}*{cy}^2"],
             [f"=-{k}*{cx}^2", f"=-{k}*{cx}*{cy}", f"={k}*{cx}^2", f"={k}*{cx}*{cy}"],
             [f"=-{k}*{cx}*{cy}", f"=-{k}*{cy}^2", f"={k}*{cx}*{cy}", f"={k}*{cy}^2"]]
    
    mat_refs = []
    for r_idx in range(4):
        row_refs = []
        for c_idx in range(4):
            c_letter = get_column_letter(c_idx + 2) 
            r_num = start_row + r_idx
            ws.cell(r_num, c_idx + 2, forms[r_idx][c_idx]).number_format = fmt_sci
            row_refs.append(f"{c_letter}{r_num}")
        mat_refs.append(row_refs)
    
    local_matrices_map[el_id] = mat_refs
    current_row += 5 

# ---------------------------------------------------------
# BLOQUE 4: MATRIZ GLOBAL
# ---------------------------------------------------------
current_row += 1
ws.cell(current_row, 1, "5. MATRIZ DE RIGIDEZ GLOBAL (K)").font = style_bold
current_row += 1

for i in range(n_dof):
    ws.cell(current_row, i+2, f"U{i+1}").font = style_bold
    ws.cell(current_row + i + 1, 1, f"U{i+1}").font = style_bold

start_row_global = current_row + 1
global_matrix_range = {} 

global_k_formulas = [['' for _ in range(n_dof)] for _ in range(n_dof)]

for _, row in df_elems.iterrows():
    el_id = row['Elemento']; ni, nj = row['Nodo_i'], row['Nodo_j']
    dofs = [2*(ni-1), 2*(ni-1)+1, 2*(nj-1), 2*(nj-1)+1]
    refs = local_matrices_map[el_id]
    
    for r in range(4):
        for c in range(4):
            g_row, g_col = dofs[r], dofs[c]
            val = f"+{refs[r][c]}" 
            if global_k_formulas[g_row][g_col] == '': 
                global_k_formulas[g_row][g_col] = f"={refs[r][c]}"
            else: 
                global_k_formulas[g_row][g_col] += val

for r in range(n_dof):
    for c in range(n_dof):
        f_str = global_k_formulas[r][c]
        cell_addr = f"{get_column_letter(c+2)}{start_row_global+r}"
        
        cell = ws.cell(start_row_global+r, c+2, 0 if f_str == '' else f_str)
        cell.number_format = fmt_sci
        if f_str == '': cell.font = Font(color="D9D9D9")
        
        global_matrix_range[(r, c)] = cell_addr

current_row = start_row_global + n_dof + 1

# ---------------------------------------------------------
# BLOQUE 5: SOLUCIÓN MATRICIAL
# ---------------------------------------------------------
ws.cell(current_row, 1, "6. SOLUCIÓN DEL SISTEMA").font = style_bold
current_row += 1

col_K_reduced = 2
col_F = col_K_reduced + len(free_dofs) + 2
col_K_inv = col_F + 2
col_U = col_K_inv + len(free_dofs) + 2

ws.cell(current_row, col_K_reduced, "K_ff (Reducida)").font = Font(bold=True)
ws.cell(current_row, col_F, "F_f (Cargas)").font = Font(bold=True)
ws.cell(current_row, col_K_inv, "K_inv").font = Font(bold=True)
ws.cell(current_row, col_U, "U_f (Calc)").font = Font(bold=True)
current_row += 1

start_sol_row = current_row
n_free = len(free_dofs)

for i, r_idx in enumerate(free_dofs):
    for j, c_idx in enumerate(free_dofs):
        ref_cell = global_matrix_range[(r_idx, c_idx)]
        ws.cell(current_row+i, col_K_reduced+j, f"={ref_cell}").number_format = fmt_sci

range_kff_str = f"{get_column_letter(col_K_reduced)}{start_sol_row}:{get_column_letter(col_K_reduced+n_free-1)}{start_sol_row+n_free-1}"

for i, idx in enumerate(free_dofs):
    ws.cell(current_row+i, col_F, F_vector[idx]).number_format = fmt_sci

range_ff_str = f"{get_column_letter(col_F)}{start_sol_row}:{get_column_letter(col_F)}{start_sol_row+n_free-1}"

ws.cell(start_sol_row, col_K_inv, f"=MINVERSE({range_kff_str})")
for i in range(n_free):
    for j in range(n_free):
        ws.cell(start_sol_row+i, col_K_inv+j).number_format = fmt_sci

range_kinv_str = f"{get_column_letter(col_K_inv)}{start_sol_row}:{get_column_letter(col_K_inv+n_free-1)}{start_sol_row+n_free-1}"

ws.cell(start_sol_row, col_U, f"=MMULT({range_kinv_str},{range_ff_str})")
for i in range(n_free):
    ws.cell(start_sol_row+i, col_U).number_format = fmt_sci

u_free_map = {} 
for i, dof_idx in enumerate(free_dofs):
    u_free_map[dof_idx] = f"{get_column_letter(col_U)}{start_sol_row+i}"

current_row += n_free + 2

# ---------------------------------------------------------
# BLOQUE 6: DESPLAZAMIENTOS TOTALES
# ---------------------------------------------------------
ws.cell(current_row, 1, "7. DESPLAZAMIENTOS NODALES (U)").font = style_bold
current_row += 1
write_header(ws, current_row, ["DOF", "Descripción", "Valor (m/in)"])
current_row += 1

u_global_map = {} # dof -> celda address
u_vector_start_row = current_row # Guardamos donde empieza el vector U global para las reacciones

for i in range(n_dof):
    ws.cell(current_row, 1, f"U{i+1}")
    
    # Descripción
    node_idx = (i // 2) + 1
    dir_str = "X" if i % 2 == 0 else "Y"
    ws.cell(current_row, 2, f"Nodo {node_idx} ({dir_str})")
    
    # Valor
    if i in fixed_dofs:
        ws.cell(current_row, 3, 0).number_format = fmt_sci
    else:
        ws.cell(current_row, 3, f"={u_free_map[i]}").number_format = fmt_sci
    
    u_global_map[i] = f"C{current_row}" # Columna 3 es Valor
    current_row += 1

# ---------------------------------------------------------
# BLOQUE 7: TABLA FINAL COMBINADA (RESULTADOS BARRAS)
# ---------------------------------------------------------
current_row += 2
ws.cell(current_row, 1, "8. RESULTADOS DETALLADOS: DEFORMACIONES Y ESFUERZOS").font = style_bold
current_row += 1

# Encabezados solicitados: Deformaciones X, Y y Esfuerzos
headers_res = ["Elem", "Nodo i", "Nodo j", "u_i", "v_i", "u_j", "v_j", "Fuerza Axial", "Esfuerzo (Stress)", "Estado"]
write_header(ws, current_row, headers_res)
current_row += 1

for i, row in df_elems.iterrows():
    el_id = row['Elemento']; ni, nj = row['Nodo_i'], row['Nodo_j']
    r = current_row
    
    # Datos geométricos previos
    row_geom = elem_data_map[el_id]['row']
    k_ref = f"M{row_geom}"
    cx_ref = f"K{row_geom}"
    cy_ref = f"L{row_geom}"
    
    # Celdas de desplazamiento (referencias a Bloque 6)
    cell_uix = u_global_map[2*(ni-1)]
    cell_uiy = u_global_map[2*(ni-1)+1]
    cell_ujx = u_global_map[2*(nj-1)]
    cell_ujy = u_global_map[2*(nj-1)+1]
    
    ws.cell(r, 1, el_id)
    ws.cell(r, 2, ni)
    ws.cell(r, 3, nj)
    
    # 1. Escribir las deformaciones (Referencia directa)
    ws.cell(r, 4, f"={cell_uix}").number_format = fmt_sci
    ws.cell(r, 5, f"={cell_uiy}").number_format = fmt_sci
    ws.cell(r, 6, f"={cell_ujx}").number_format = fmt_sci
    ws.cell(r, 7, f"={cell_ujy}").number_format = fmt_sci
    
    # 2. Fuerza Axial usando las columnas D, E, F, G de esta misma fila (r)
    # F = k * ( (uj-ui)*cx + (vj-vi)*cy )
    # ui=D, vi=E, uj=F, vj=G
    f_force = f"={k_ref} * ( (F{r}-D{r})*{cx_ref} + (G{r}-E{r})*{cy_ref} )"
    ws.cell(r, 8, f_force).number_format = fmt_sci
    
    # 3. Esfuerzo
    ws.cell(r, 9, f"=H{r}/{cell_A_ref}").number_format = fmt_sci
    
    # 4. Estado
    ws.cell(r, 10, f'=IF(ABS(H{r})<1E-9,"-", IF(H{r}>0,"Tracción","Compresión"))')
    
    current_row += 1

# ---------------------------------------------------------
# BLOQUE 9: REACCIONES EN APOYOS
# ---------------------------------------------------------
current_row += 2
ws.cell(current_row, 1, "9. REACCIONES EN LOS APOYOS").font = style_bold
current_row += 1
write_header(ws, current_row, ["Nodo", "Dirección", "K*U (Calc)", "F_aplicada", "Reacción Neta"])
current_row += 1

# Referencia al vector U completo (Columna C de Bloque 6)
range_U_vector = f"$C${u_vector_start_row}:$C${u_vector_start_row + n_dof - 1}"

for dof in fixed_dofs:
    # Info del nodo
    node_idx = (dof // 2) + 1
    dir_str = "X" if dof % 2 == 0 else "Y"
    
    # Construir referencia a la fila correspondiente de la Matriz Global K (Bloque 4)
    # start_row_global es donde empiezan los datos de la matriz
    k_row_excel = start_row_global + dof # la fila exacta en Excel
    
    # Rango de la fila K: Desde col B hasta la última columna de la matriz
    last_col_letter = get_column_letter(n_dof + 1) # +1 porque empieza en B
    range_K_row = f"$B${k_row_excel}:${last_col_letter}{k_row_excel}"
    
    ws.cell(current_row, 1, node_idx)
    ws.cell(current_row, 2, dir_str)
    
    # 1. Fuerza calculada = Fila_K * Vector_U (SUMPRODUCT)
    # SUMAPRODUCTO es robusto para fila * columna en Excel
    ws.cell(current_row, 3, f"=SUMPRODUCT({range_K_row}, {range_U_vector})").number_format = fmt_sci
    
    # 2. Carga Aplicada (por si acaso hay carga sobre un apoyo)
    ws.cell(current_row, 4, F_vector[dof]).number_format = fmt_sci
    
    # 3. Reacción = Fuerza Calculada - Carga Aplicada
    ws.cell(current_row, 5, f"=C{current_row}-D{current_row}").number_format = fmt_sci
    
    current_row += 1

# ==========================================
# GESTIÓN DEL ARCHIVO
# ==========================================
final_filename = base_filename
counter = 1
while os.path.exists(final_filename):
    name, ext = os.path.splitext(base_filename)
    final_filename = f"{name}_{counter}{ext}"
    counter += 1

wb.save(final_filename)
print(f"✅ Archivo generado: '{final_filename}'")