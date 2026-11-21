import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# --- 1. PARAMÈTRES ET CALCULS (Invariant) ---
E = 200e6
A_exact = 0.0052
nodes_data = {
    'Nodo': range(1, 11),
    'x': [0, 4, 8, 12, 16, 0, 4, 8, 12, 16],
    'y': [4, 4, 4, 4, 4, 0, 0, 0, 0, 0]
}
df_nodes = pd.DataFrame(nodes_data)
# Liste corrigée (N6 bas gauche, N5 haut droite)
elements_list = [
    (1, 1, 2), (2, 2, 3), (3, 3, 4), (4, 4, 5),
    (5, 6, 7), (6, 7, 8), (7, 8, 9), (8, 9, 10),
    (9, 1, 6), (10, 2, 7), (11, 3, 8), (12, 4, 9), (13, 5, 10),
    (14, 2, 6), (15, 3, 7), (16, 3, 9), (17, 4, 10)
]
df_elems = pd.DataFrame(elements_list, columns=['Elemento', 'Nodo_i', 'Nodo_j'])
df_elems['E'] = E
df_elems['A'] = A_exact

# --- CALCUL NUMÉRIQUE (Nécessaire pour les valeurs) ---
n_dof = 20; F = np.zeros(n_dof)
cargas = {6: -30, 7: -20, 8: -20, 9: -40}
for n, f in cargas.items(): F[2*(n-1)+1] = f
fixed_dofs = [2*(6-1), 2*(6-1)+1, 2*(10-1)+1]
free_dofs = [i for i in range(n_dof) if i not in fixed_dofs]
K_global = np.zeros((n_dof, n_dof))
local_matrices = {}
results_calc = []

for index, row in df_elems.iterrows():
    el_id = int(row['Elemento']); ni, nj = int(row['Nodo_i']), int(row['Nodo_j'])
    xi = df_nodes.loc[df_nodes['Nodo'] == ni, 'x'].values[0]
    yi = df_nodes.loc[df_nodes['Nodo'] == ni, 'y'].values[0]
    xj = df_nodes.loc[df_nodes['Nodo'] == nj, 'x'].values[0]
    yj = df_nodes.loc[df_nodes['Nodo'] == nj, 'y'].values[0]
    dx = xj - xi; dy = yj - yi; L = np.sqrt(dx**2 + dy**2)
    cx = dx / L; cy = dy / L
    k_val = (row['E'] * row['A']) / L
    k_mat = k_val * np.array([[cx**2, cx*cy, -cx**2, -cx*cy], [cx*cy, cy**2, -cx*cy, -cy**2], [-cx**2, -cx*cy, cx**2, cx*cy], [-cx*cy, -cy**2, cx*cy, cy**2]])
    local_matrices[el_id] = k_mat
    dofs = [2*(ni-1), 2*(ni-1)+1, 2*(nj-1), 2*(nj-1)+1]
    for r in range(4):
        for c in range(4): K_global[dofs[r], dofs[c]] += k_mat[r, c]
    results_calc.append({'Elemento': el_id, 'L': L, 'cx': cx, 'cy': cy, 'k_val': k_val, 'dx': dx, 'dy': dy, 'xi': xi, 'yi': yi, 'xj': xj, 'yj': yj})

u_f = np.linalg.solve(K_global[np.ix_(free_dofs, free_dofs)], F[free_dofs])
U = np.zeros(n_dof); U[free_dofs] = u_f
Reactions = K_global @ U - F

# --- GÉNÉRATION EXCEL AVEC FORMULES ---
wb = openpyxl.Workbook()
styles = {
    'header': Font(bold=True, color="FFFFFF"),
    'header_fill': PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid"),
    'formula': Font(color="FF0000", italic=True), # Rouge pour bien voir
    'yellow': PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
}

# 1. FEUILLE CALCULS (Avec formules prêtes à copier)
ws1 = wb.active; ws1.title = "Calculos"
# Colonnes: A=Elem, B=Ni, C=Nj, D=xi, E=yi, F=xj, G=yj, H=dx, I=dy, J=L, K=cx, L=cy, M=k
headers = ["Elemento", "Nodo_i", "Nodo_j", "xi", "yi", "xj", "yj", "dx", "dy", "L", "cx", "cy", "k=EA/L", "FORMULES A COPIER (Ligne 2)"]
ws1.append(headers)
for c, h in enumerate(headers, 1): 
    cell = ws1.cell(1, c); cell.font = styles['header']; cell.fill = styles['header_fill']

# Remplissage Données + Formules Texte
for i, row in df_elems.iterrows():
    r = i + 2
    el_id = row['Elemento']
    c = next(item for item in results_calc if item['Elemento'] == el_id)
    
    # Valeurs
    data = [el_id, row['Nodo_i'], row['Nodo_j'], c['xi'], c['yi'], c['xj'], c['yj'], c['dx'], c['dy'], c['L'], c['cx'], c['cy'], c['k_val']]
    for col_idx, val in enumerate(data, 1): ws1.cell(r, col_idx, val)
    
    # Formules EXCEL TEXTE (Si tu devais le faire à la main)
    # Supposons que les coords sont entrées manuellement ou par VLOOKUP.
    # Ici je donne les formules géométriques.
    formulas_text = (
        f"dx (H{r}): =F{r}-D{r} | "
        f"dy (I{r}): =G{r}-E{r} | "
        f"L (J{r}): =RAIZ(H{r}^2 + I{r}^2) | "
        f"cx (K{r}): =H{r}/J{r} | "
        f"cy (L{r}): =I{r}/J{r} | "
        f"k (M{r}): =({E}*{A_exact})/J{r}"
    )
    ws1.cell(r, 14, formulas_text).font = styles['formula']

# 2. FEUILLE ESFUERZOS (Stress)
ws4 = wb.create_sheet("Esfuerzos")
# Colonnes: A=Elem, B=Ni, C=Nj, D=Fuerza, E=Esfuerzo, F=Formule_Excel
headers_stress = ["Elemento", "Ni", "Nj", "Fuerza Axial (kN)", "Esfuerzo (kPa)", "FORMULE EXCEL A COPIER"]
ws4.append(headers_stress)
for c, h in enumerate(headers_stress, 1): 
    cell = ws4.cell(1, c); cell.font = styles['header']; cell.fill = styles['header_fill']

stress_results = []
for index, row in df_elems.iterrows():
    r = index + 2
    el_id = int(row['Elemento']); ni, nj = int(row['Nodo_i']), int(row['Nodo_j'])
    
    u_idx = [2*(ni-1), 2*(ni-1)+1, 2*(nj-1), 2*(nj-1)+1]
    u_vec = U[u_idx]
    c = next(item for item in results_calc if item['Elemento'] == el_id)
    T_vec = np.array([-c['cx'], -c['cy'], c['cx'], c['cy']])
    Axial_Force = c['k_val'] * c['L'] * (1/c['L']) * np.dot(T_vec, u_vec) # Recalcul simple
    Stress = Axial_Force / row['A']
    
    ws4.cell(r, 1, el_id)
    ws4.cell(r, 2, ni)
    ws4.cell(r, 3, nj)
    ws4.cell(r, 4, Axial_Force)
    ws4.cell(r, 5, Stress)
    
    # FORMULE MAGIQUE EXCEL
    # Je vais référencer les colonnes de la feuille 'Calculos' pour cx, cy, k, L
    # Et supposer que les U sont nommés u1x, u1y... 
    # Pour simplifier, je donne la formule avec les valeurs numériques de cx/cy/k pour cette ligne
    # C'est ce que ANSYS fait en interne.
    
    # Si on voulait lier aux cellules de l'onglet Calculos (ex: Calculos!K2 pour cx)
    ref_k = f"Calculos!M{r}"
    ref_cx = f"Calculos!K{r}"
    ref_cy = f"Calculos!L{r}"
    # U doit être cherché. Pour l'exercice, je mets les valeurs de U en dur dans la formule texte pour qu'elle marche "seule"
    # ou mieux, une formule générique :
    excel_formula = f"={ref_k} * ( (U_{nj}_x - U_{ni}_x)*{ref_cx} + (U_{nj}_y - U_{ni}_y)*{ref_cy} )"
    ws4.cell(r, 6, excel_formula).font = styles['formula']
    
    if ni == 3 and nj == 4: # Highlight CD
         for col in range(1, 7): ws4.cell(r, col).fill = styles['yellow']

# 3. FEUILLE RESULTATS (U)
ws3 = wb.create_sheet("Resultados_U")
ws3.append(["Nodo", "Ux (m)", "Uy (m)"])
for i in range(10):
    ws3.append([i+1, U[2*i], U[2*i+1]])

wb.save('Armadura_Formulas_Excel.xlsx')
print("Fichier généré avec les formules textuelles prêtes à être analysées.")