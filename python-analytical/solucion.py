import pandas as pd
import numpy as np
import openpyxl
import os
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# 1. CONFIGURATION UTILISATEUR
# ==========================================

# --- A. PROPRIÉTÉS DU MATÉRIAU ---
E_val = 30e6    # Modulo de Young
A_val = 8       # Aire de la section

# --- B. DÉFINITION DES NOEUDS (ID, X, Y) ---
nodes_list = [
    (1, 0, 0), (2, 0, 36), (3, -24, 18), (4, -48, 0)
]

# --- C. DÉFINITION DES BARRES (ID, Noeud_Départ, Noeud_Arrivée) ---
elements_list = [
    (1, 1, 2), (2, 2, 3), (3, 3, 1), (4, 1, 4), (5, 3, 4)
]

# --- D. CHARGES ---
loads_data = {
    4: (0, -300)
}

# --- E. APPUIS ---  Types supportés: 'fixe' (bloque X et Y), 'rouleau_x' (bloque X), 'rouleau_y' (bloque Y)
supports_data = {
    2: 'fixe',      
    1: 'rouleau_x'  
}

# NOM DE BASE DU FICHIER
base_filename = "Calcul_Armure_Resultats.xlsx"

# ==========================================
# FIN DE LA CONFIGURATION
# ==========================================

# --- 2. PRÉTRAITEMENT ---
df_nodes = pd.DataFrame(nodes_list, columns=['Nodo', 'x', 'y'])
df_elems = pd.DataFrame(elements_list, columns=['Elemento', 'Nodo_i', 'Nodo_j'])
n_nodes = len(nodes_list)
n_dof = n_nodes * 2 

fixed_dofs = []
for node, type_appui in supports_data.items():
    idx = node - 1
    if type_appui == 'fixe':
        fixed_dofs.extend([2*idx, 2*idx+1])
    elif type_appui == 'rouleau_x':
        fixed_dofs.append(2*idx)
    elif type_appui == 'rouleau_y':
        fixed_dofs.append(2*idx+1)

fixed_dofs = sorted(list(set(fixed_dofs)))
free_dofs = [i for i in range(n_dof) if i not in fixed_dofs]

F_vector = np.zeros(n_dof)
for node, (fx, fy) in loads_data.items():
    idx = node - 1
    F_vector[2*idx] = fx
    F_vector[2*idx+1] = fy

# --- 3. GÉNÉRATION EXCEL ---
wb = openpyxl.Workbook()

# Styles
style_header = Font(bold=True, color="FFFFFF")
style_fill_header = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
style_fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
style_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

def format_ws_header(ws):
    for cell in ws[1]:
        cell.font = style_header
        cell.fill = style_fill_header
        cell.alignment = Alignment(horizontal='center')

# FEUILLE 1: DONNÉES
ws1 = wb.active; ws1.title = "Données"
ws1.append(["PARAMETRES MATERIAU", "Valeur", "Unité"])
ws1.append(["Module Young (E)", E_val, "Force/Area"])
ws1.append(["Section (A)", A_val, "Area"])
ws1.append([])
ws1.append(["NOEUDS", "Coord X", "Coord Y"])
row_node_start = 6
node_refs = {}
for i, row in df_nodes.iterrows():
    r = row_node_start + i
    ws1.append([row['Nodo'], row['x'], row['y']])
    node_refs[row['Nodo']] = {'x': f"Données!B{r}", 'y': f"Données!C{r}"}
format_ws_header(ws1)
for c in range(1, 4): 
    ws1.cell(5, c).font = style_header
    ws1.cell(5, c).fill = style_fill_header

# FEUILLE 2: CALCULS GÉOMÉTRIQUES
ws2 = wb.create_sheet("Calculs_Elem")
ws2.append(["Elément", "Noeud i", "Noeud j", "xi", "yi", "xj", "yj", "dx", "dy", "L", "cx", "cy", "k"])
format_ws_header(ws2)
elem_row_map = {} 
for i, row in df_elems.iterrows():
    r = i + 2
    el_id = row['Elemento']
    ni, nj = row['Nodo_i'], row['Nodo_j']
    elem_row_map[el_id] = r
    ws2.cell(r, 1, el_id); ws2.cell(r, 2, ni); ws2.cell(r, 3, nj)
    ws2.cell(r, 4, f"={node_refs[ni]['x']}")
    ws2.cell(r, 5, f"={node_refs[ni]['y']}")
    ws2.cell(r, 6, f"={node_refs[nj]['x']}")
    ws2.cell(r, 7, f"={node_refs[nj]['y']}")
    ws2.cell(r, 8, f"=F{r}-D{r}")
    ws2.cell(r, 9, f"=G{r}-E{r}")
    ws2.cell(r, 10, f"=SQRT(H{r}^2 + I{r}^2)")
    ws2.cell(r, 11, f"=H{r}/J{r}")
    ws2.cell(r, 12, f"=I{r}/J{r}")
    ws2.cell(r, 13, f"=(Données!$B$2 * Données!$B$3)/J{r}")

# FEUILLE 3: MATRICES LOCALES
ws3 = wb.create_sheet("Matrices_Locales")
local_matrix_refs = {}
curr_row = 1
for i, row in df_elems.iterrows():
    el_id = row['Elemento']
    calc_r = elem_row_map[el_id]
    ws3.cell(curr_row, 1, f"Elément {el_id}").font = Font(bold=True)
    curr_row += 1
    k = f"Calculs_Elem!M{calc_r}"; cx = f"Calculs_Elem!K{calc_r}"; cy = f"Calculs_Elem!L{calc_r}"
    forms = [[f"={k}*{cx}^2", f"={k}*{cx}*{cy}", f"=-{k}*{cx}^2", f"=-{k}*{cx}*{cy}"],
             [f"={k}*{cx}*{cy}", f"={k}*{cy}^2", f"=-{k}*{cx}*{cy}", f"=-{k}*{cy}^2"],
             [f"=-{k}*{cx}^2", f"=-{k}*{cx}*{cy}", f"={k}*{cx}^2", f"={k}*{cx}*{cy}"],
             [f"=-{k}*{cx}*{cy}", f"=-{k}*{cy}^2", f"={k}*{cx}*{cy}", f"={k}*{cy}^2"]]
    mat_refs = []
    for r_idx in range(4):
        row_refs = []
        for c_idx in range(4):
            cell = ws3.cell(curr_row + r_idx, c_idx + 1, forms[r_idx][c_idx])
            cell.number_format = '0.00'; cell.border = style_border
            row_refs.append(f"'Matrices_Locales'!{get_column_letter(c_idx + 1)}{curr_row + r_idx}")
        mat_refs.append(row_refs)
    local_matrix_refs[el_id] = mat_refs
    curr_row += 5

# FEUILLE 4: MATRICE GLOBALE
ws4 = wb.create_sheet("Matrice_Globale")
ws4.cell(1, 1, f"Matrice de Rigidité Globale")
for i in range(n_dof): ws4.cell(2, i+2, f"DOF {i}"); ws4.cell(i+3, 1, f"DOF {i}")
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
        cell.number_format = '0.00'; 
        if f_str == '': cell.font = Font(color="D9D9D9")

# FEUILLE 5: RÉSOLUTION
ws5 = wb.create_sheet("Solution")
ws5.append(["Calcul des Déplacements"])
ws5.cell(3, 1, "Matrice Réduite K_ff")
free_dof_indices = sorted(free_dofs); n_free = len(free_dof_indices)
r_off, c_off = 4, 2
for i, g_row in enumerate(free_dof_indices):
    for j, g_col in enumerate(free_dof_indices):
        ws5.cell(r_off+i, c_off+j, f"='Matrice_Globale'!{get_column_letter(2+g_col)}{3+g_row}")
range_kff = f"{get_column_letter(c_off)}{r_off}:{get_column_letter(c_off+n_free-1)}{r_off+n_free-1}"

f_col = c_off + n_free + 2
ws5.cell(3, f_col, "Vecteur F_f")
for i, g_row in enumerate(free_dof_indices): ws5.cell(r_off+i, f_col, F_vector[g_row])
range_ff = f"{get_column_letter(f_col)}{r_off}:{get_column_letter(f_col)}{r_off+n_free-1}"

inv_col = f_col + 2
ws5.cell(3, inv_col, "Inverse K_inv")
ws5.cell(r_off, inv_col, f"=MINVERSE({range_kff})")
range_kinv = f"{get_column_letter(inv_col)}{r_off}:{get_column_letter(inv_col+n_free-1)}{r_off+n_free-1}"

u_col = inv_col + n_free + 2
ws5.cell(3, u_col, "Déplacements U_f")
ws5.cell(r_off, u_col, f"=MMULT({range_kinv}, {range_ff})")

u_glob_r = r_off + n_free + 3
ws5.cell(u_glob_r, 1, "VECTEUR GLOBAL U").font = Font(bold=True)
u_global_refs = {}
for i in range(n_dof):
    r = u_glob_r + 2 + i
    ws5.cell(r, 1, i)
    if i in fixed_dofs: ws5.cell(r, 2, 0)
    else:
        idx = free_dof_indices.index(i)
        ws5.cell(r, 2, f"={get_column_letter(u_col)}{r_off+idx}")
    u_global_refs[i] = f"'Solution'!B{r}"

# FEUILLE 6: EFFORTS & CONTRAINTES & DÉFORMATIONS
ws6 = wb.create_sheet("Resultats_Finaux")
# Ajout des en-têtes demandés
headers_res = ["Elément", "Noeuds", "Force Axiale", "Contrainte", "Etat", "Delta Ux", "Delta Uy", "Déformation"]
ws6.append(headers_res)
format_ws_header(ws6)

for i, row in df_elems.iterrows():
    r = i + 2
    el_id = row['Elemento']
    ni, nj = row['Nodo_i'], row['Nodo_j']
    
    ws6.cell(r, 1, el_id)
    ws6.cell(r, 2, f"{ni}-{nj}")

    calc_r = elem_row_map[el_id]
    k_ref = f"Calculs_Elem!M{calc_r}"
    cx_ref = f"Calculs_Elem!K{calc_r}"
    cy_ref = f"Calculs_Elem!L{calc_r}"
    area_ref = "Données!$B$3"
    E_ref = "Données!$B$2" # Référence pour le module de Young
    
    uix = u_global_refs[2*(ni-1)]; uiy = u_global_refs[2*(ni-1)+1]
    ujx = u_global_refs[2*(nj-1)]; ujy = u_global_refs[2*(nj-1)+1]
    
    # Force Axiale
    formula_force = f"={k_ref} * ( ({ujx}-{uix})*{cx_ref} + ({ujy}-{uiy})*{cy_ref} )"
    ws6.cell(r, 3, formula_force).number_format = '0.00'
    
    # Contrainte (Stress)
    ws6.cell(r, 4, f"=C{r}/{area_ref}").number_format = '0.00'
    
    # État
    ws6.cell(r, 5, f'=IF(C{r}>0, "Traction", "Compression")')

    # --- NOUVELLES COLONNES ---
    
    # Delta Ux (ujx - uix) : Déplacement relatif en X
    ws6.cell(r, 6, f"=({ujx}-{uix})").number_format = '0.0000'
    
    # Delta Uy (ujy - uiy) : Déplacement relatif en Y
    ws6.cell(r, 7, f"=({ujy}-{uiy})").number_format = '0.0000'
    
    # Déformation (Strain) = Contrainte / E (Loi de Hooke)
    # Contrainte est en colonne 4 (D), E est dans Données!B2
    ws6.cell(r, 8, f"=D{r}/{E_ref}").number_format = '0.00E+00' # Notation scientifique pour les petits nombres

# ==========================================
# 4. GESTION DES NOMS DE FICHIERS (BOUCLE)
# ==========================================
final_filename = base_filename
counter = 1

while os.path.exists(final_filename):
    name, ext = os.path.splitext(base_filename)
    final_filename = f"{name}_{counter}{ext}"
    counter += 1

wb.save(final_filename)
print(f"✅ Fichier généré avec succès : '{final_filename}'")