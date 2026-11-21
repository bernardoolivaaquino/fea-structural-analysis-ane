import pandas as pd
import numpy as np
import openpyxl
import os
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# 1. CONFIGURACIÓN DEL USUARIO
# ==========================================

# --- A. PROPIEDADES DEL MATERIAL ---
E_val = 30e6    # Módulo de Young
A_val = 8       # Área de la sección transversal

# --- B. DEFINICIÓN DE NODOS (ID, X, Y) ---
nodes_list = [
    (1, 0, 0), (2, 0, 36), (3, -24, 18), (4, -48, 0)
]

# --- C. DEFINICIÓN DE ELEMENTOS (ID, Nodo_Inicio, Nodo_Fin) ---
elements_list = [
    (1, 1, 2), (2, 2, 3), (3, 3, 1), (4, 1, 4), (5, 3, 4)
]

# --- D. CARGAS (NODO: (Fx, Fy)) ---
# NOTA: Las cargas en nodos fijos son reacciones y no aparecen en el vector F_f reducido.
loads_data = {
    4: (0, -300)
}

# --- E. APOYOS --- 
# Tipos soportados: 'fijo' (bloquea X y Y), 'rodillo_x' (bloquea X), 'rodillo_y' (bloquea Y)
supports_data = {
    2: 'fijo',      
    1: 'rodillo_x'  
}

# NOMBRE BASE DEL ARCHIVO
base_filename = "Calculo_Armadura_Resultados.xlsx"

# ==========================================
# 2. PRE-PROCESAMIENTO
# ==========================================
df_nodes = pd.DataFrame(nodes_list, columns=['Nodo', 'x', 'y'])
df_elems = pd.DataFrame(elements_list, columns=['Elemento', 'Nodo_i', 'Nodo_j'])
n_nodes = len(nodes_list)
n_dof = n_nodes * 2 

# Identificación de Grados de Libertad (DOF)
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

# Vector de Fuerza Global
F_vector = np.zeros(n_dof)
for node, (fx, fy) in loads_data.items():
    idx = node - 1
    F_vector[2*idx] = fx
    F_vector[2*idx+1] = fy

# ==========================================
# 3. GENERACIÓN DE EXCEL
# ==========================================
wb = openpyxl.Workbook()

# --- ESTILOS Y FORMATOS ---
style_header = Font(bold=True, color="FFFFFF")
style_fill_header = PatternFill(start_color="203764", end_color="203764", fill_type="solid") # Azul profesional
style_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# FORMATO CIENTÍFICO: Crucial para ver valores infinitesimales (e.g. 1.23E-20)
fmt_sci = '0.0000E+00'  
fmt_num = '0.0000'      

def format_ws_header(ws):
    for cell in ws[1]:
        cell.font = style_header
        cell.fill = style_fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center')

# --- HOJA 1: DATOS ---
ws1 = wb.active; ws1.title = "Datos"
ws1.append(["PARÁMETROS", "Valor", "Unidad"])
ws1.append(["Módulo Young (E)", E_val, "Psi/Pa"])
ws1.append(["Área Sección (A)", A_val, "in²/m²"])
ws1.cell(2,2).number_format = fmt_sci 

ws1.append([])
ws1.append(["NODOS", "Coord X", "Coord Y"])
row_node_start = 6
node_refs = {}
for i, row in df_nodes.iterrows():
    r = row_node_start + i
    ws1.append([row['Nodo'], row['x'], row['y']])
    ws1.cell(r, 2).number_format = fmt_num
    ws1.cell(r, 3).number_format = fmt_num
    # Referencias para usar en fórmulas de otras hojas (ej. Datos!B6)
    node_refs[row['Nodo']] = {'x': f"Datos!B{r}", 'y': f"Datos!C{r}"}
format_ws_header(ws1)

# --- HOJA 2: CÁLCULOS GEOMÉTRICOS ---
ws2 = wb.create_sheet("Calculos_Elem")
ws2.append(["Elemento", "Ni", "Nj", "xi", "yi", "xj", "yj", "dx", "dy", "L", "cx", "cy", "k"])
format_ws_header(ws2)
elem_row_map = {} 
for i, row in df_elems.iterrows():
    r = i + 2
    el_id = row['Elemento']; ni, nj = row['Nodo_i'], row['Nodo_j']
    elem_row_map[el_id] = r
    
    ws2.append([el_id, ni, nj]) 
    
    # Escritura de fórmulas referenciando la hoja 'Datos'
    ws2.cell(r, 4, f"={node_refs[ni]['x']}")
    ws2.cell(r, 5, f"={node_refs[ni]['y']}")
    ws2.cell(r, 6, f"={node_refs[nj]['x']}")
    ws2.cell(r, 7, f"={node_refs[nj]['y']}")
    ws2.cell(r, 8, f"=F{r}-D{r}") # dx
    ws2.cell(r, 9, f"=G{r}-E{r}") # dy
    ws2.cell(r, 10, f"=SQRT(H{r}^2 + I{r}^2)") # Longitud (L)
    ws2.cell(r, 11, f"=H{r}/J{r}") # Cosinus director x (cx)
    ws2.cell(r, 12, f"=I{r}/J{r}") # Cosinus director y (cy)
    ws2.cell(r, 13, f"=(Datos!$B$2 * Datos!$B$3)/J{r}") # Rigidez axial k = EA/L
    
    for c in range(4, 14): ws2.cell(r, c).number_format = fmt_sci

# --- HOJA 3: MATRICES LOCALES ---
ws3 = wb.create_sheet("Matrices_Locales")
local_matrix_refs = {}
curr_row = 1
for i, row in df_elems.iterrows():
    el_id = row['Elemento']
    calc_r = elem_row_map[el_id]
    ws3.cell(curr_row, 1, f"Elemento {el_id}").font = Font(bold=True)
    curr_row += 1
    
    # Referencias a k, cx, cy en la hoja de cálculos
    k = f"Calculos_Elem!M{calc_r}"; cx = f"Calculos_Elem!K{calc_r}"; cy = f"Calculos_Elem!L{calc_r}"
    
    # Matriz de rigidez local [k]
    forms = [[f"={k}*{cx}^2", f"={k}*{cx}*{cy}", f"=-{k}*{cx}^2", f"=-{k}*{cx}*{cy}"],
             [f"={k}*{cx}*{cy}", f"={k}*{cy}^2", f"=-{k}*{cx}*{cy}", f"=-{k}*{cy}^2"],
             [f"=-{k}*{cx}^2", f"=-{k}*{cx}*{cy}", f"={k}*{cx}^2", f"={k}*{cx}*{cy}"],
             [f"=-{k}*{cx}*{cy}", f"=-{k}*{cy}^2", f"={k}*{cx}*{cy}", f"={k}*{cy}^2"]]
    mat_refs = []
    for r_idx in range(4):
        row_refs = []
        for c_idx in range(4):
            cell = ws3.cell(curr_row + r_idx, c_idx + 1, forms[r_idx][c_idx])
            cell.number_format = fmt_sci 
            cell.border = style_border
            # Guardamos la referencia de Excel para ensamblar la global
            row_refs.append(f"'Matrices_Locales'!{get_column_letter(c_idx + 1)}{curr_row + r_idx}")
        mat_refs.append(row_refs)
    local_matrix_refs[el_id] = mat_refs
    curr_row += 5

# --- HOJA 4: MATRIZ GLOBAL ---
ws4 = wb.create_sheet("Matriz_Global")
ws4.cell(1, 1, f"K Global (Ensamblada)")
for i in range(n_dof): 
    ws4.cell(2, i+2, f"U{i+1}"); ws4.cell(i+3, 1, f"U{i+1}") 

# Ensamblaje mediante concatenación de strings para la fórmula
global_k_formulas = [['' for _ in range(n_dof)] for _ in range(n_dof)]

for _, row in df_elems.iterrows():
    el_id = row['Elemento']; ni, nj = row['Nodo_i'], row['Nodo_j']
    dofs = [2*(ni-1), 2*(ni-1)+1, 2*(nj-1), 2*(nj-1)+1]
    refs = local_matrix_refs[el_id]
    for r in range(4):
        for c in range(4):
            g_row, g_col = dofs[r], dofs[c]
            val = f"+{refs[r][c]}"
            if global_k_formulas[g_row][g_col] == '': global_k_formulas[g_row][g_col] = f"={refs[r][c]}"
            else: global_k_formulas[g_row][g_col] += val

for r in range(n_dof):
    for c in range(n_dof):
        f_str = global_k_formulas[r][c]
        cell = ws4.cell(3+r, 2+c, 0 if f_str == '' else f_str)
        cell.number_format = fmt_sci 
        if f_str == '': cell.font = Font(color="D9D9D9")

# --- HOJA 5: SOLUCIÓN ---
ws5 = wb.create_sheet("Solucion")
ws5.append(["Cálculo Matriciel"])
ws5.cell(3, 1, "K_ff (Reducida)")
free_dof_indices = sorted(free_dofs); n_free = len(free_dof_indices)
r_off, c_off = 4, 2

# K_ff: Submatriz de grados de libertad libres
for i, g_row in enumerate(free_dof_indices):
    for j, g_col in enumerate(free_dof_indices):
        cell = ws5.cell(r_off+i, c_off+j, f"='Matriz_Global'!{get_column_letter(2+g_col)}{3+g_row}")
        cell.number_format = fmt_sci
range_kff = f"{get_column_letter(c_off)}{r_off}:{get_column_letter(c_off+n_free-1)}{r_off+n_free-1}"

# F_f: Vector de fuerzas conocidas
f_col = c_off + n_free + 2
ws5.cell(3, f_col, "F_f (Cargas)")
for i, g_row in enumerate(free_dof_indices): 
    cell = ws5.cell(r_off+i, f_col, F_vector[g_row])
    cell.number_format = fmt_sci
range_ff = f"{get_column_letter(f_col)}{r_off}:{get_column_letter(f_col)}{r_off+n_free-1}"

# K_inv: Inversa
inv_col = f_col + 2
ws5.cell(3, inv_col, "K_inv")
ws5.cell(r_off, inv_col, f"=MINVERSE({range_kff})") # MINVERSE = M.INVERSA (Inglés)
range_kinv = f"{get_column_letter(inv_col)}{r_off}:{get_column_letter(inv_col+n_free-1)}{r_off+n_free-1}"
for r in range(n_free):
    for c in range(n_free):
        ws5.cell(r_off+r, inv_col+c).number_format = fmt_sci

# U_f: Desplazamientos calculados
u_col = inv_col + n_free + 2
ws5.cell(3, u_col, "U_f (Desplazamientos)")
ws5.cell(r_off, u_col, f"=MMULT({range_kinv}, {range_ff})") # MMULT = M.MULT (Inglés)
for r in range(n_free): ws5.cell(r_off+r, u_col).number_format = fmt_sci

# Reconstrucción del Vector Global U (incluyendo ceros en apoyos)
u_glob_r = r_off + n_free + 3
ws5.cell(u_glob_r, 1, "VECTOR GLOBAL U").font = Font(bold=True)
u_global_refs = {}
for i in range(n_dof):
    r = u_glob_r + 2 + i
    ws5.cell(r, 1, f"DOF {i}")
    cell_val = ws5.cell(r, 2)
    if i in fixed_dofs: 
        cell_val.value = 0
    else:
        idx = free_dof_indices.index(i)
        cell_val.value = f"={get_column_letter(u_col)}{r_off+idx}"
    cell_val.number_format = fmt_sci
    u_global_refs[i] = f"'Solucion'!B{r}"

# --- HOJA 6: RESULTADOS POR NODO ---
ws_nodes = wb.create_sheet("Resultados_Nodos")
headers_nodes = ["Nodo", "Desplazamiento X (Ux)", "Desplazamiento Y (Uy)", "Magnitud Resultante"]
ws_nodes.append(headers_nodes)
format_ws_header(ws_nodes)

for i, row in df_nodes.iterrows():
    nid = row['Nodo']
    r = i + 2
    idx_x = 2*(nid-1)
    idx_y = 2*(nid-1)+1
    
    ws_nodes.cell(r, 1, nid)
    ws_nodes.cell(r, 2, f"={u_global_refs[idx_x]}").number_format = fmt_sci # Ux
    ws_nodes.cell(r, 3, f"={u_global_refs[idx_y]}").number_format = fmt_sci # Uy
    ws_nodes.cell(r, 4, f"=SQRT(B{r}^2 + C{r}^2)").number_format = fmt_sci  # Magnitud

# --- HOJA 7: RESULTADOS POR ELEMENTO ---
ws6 = wb.create_sheet("Resultados_Elementos")
headers_res = ["Elemento", "Nodos", "Fuerza Axial", "Esfuerzo", "Estado", "Ux_inicio", "Uy_inicio", "Ux_fin", "Uy_fin", "Deformación (Unit.)"]
ws6.append(headers_res)
format_ws_header(ws6)

for i, row in df_elems.iterrows():
    r = i + 2
    el_id = row['Elemento']; ni, nj = row['Nodo_i'], row['Nodo_j']
    
    ws6.cell(r, 1, el_id)
    ws6.cell(r, 2, f"{ni}-{nj}")

    calc_r = elem_row_map[el_id]
    k_ref = f"Calculos_Elem!M{calc_r}"; cx_ref = f"Calculos_Elem!K{calc_r}"; cy_ref = f"Calculos_Elem!L{calc_r}"
    area_ref = "Datos!$B$3"; E_ref = "Datos!$B$2"
    
    uix = u_global_refs[2*(ni-1)]; uiy = u_global_refs[2*(ni-1)+1]
    ujx = u_global_refs[2*(nj-1)]; ujy = u_global_refs[2*(nj-1)+1]
    
    # 1. Fuerza Axial
    formula_force = f"={k_ref} * ( ({ujx}-{uix})*{cx_ref} + ({ujy}-{uiy})*{cy_ref} )"
    ws6.cell(r, 3, formula_force).number_format = fmt_sci
    
    # 2. Esfuerzo (Stress)
    ws6.cell(r, 4, f"=C{r}/{area_ref}").number_format = fmt_sci
    
    # 3. Estado (Tracción/Compresión)
    # ABS para ignorar errores numéricos minúsculos
    ws6.cell(r, 5, f'=IF(ABS(C{r})<1E-9, "Nulo", IF(C{r}>0, "Tracción", "Compresión"))')

    # 4. Desplazamientos Nudales
    ws6.cell(r, 6, f"={uix}").number_format = fmt_sci
    ws6.cell(r, 7, f"={uiy}").number_format = fmt_sci
    ws6.cell(r, 8, f"={ujx}").number_format = fmt_sci
    ws6.cell(r, 9, f"={ujy}").number_format = fmt_sci
    
    # 5. Deformación Unitaria (Strain)
    ws6.cell(r, 10, f"=D{r}/{E_ref}").number_format = fmt_sci

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