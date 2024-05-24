import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image


valeurs = np.array([[10,  9,  8],
                    [ 6,  1,  4],
                    [ 5,  7,  2]])


df = pd.DataFrame(valeurs, columns=["Colonne 1", "Colonne 2", "Colonne 3"], index=["Ligne 1", "Ligne 2", "Ligne 3"])


excel_path = 'tableau_valeurs_fournis.xlsx'
df.to_excel(excel_path, index=True)


plt.figure(figsize=(8, 6))
for col in df.columns:
    plt.plot(df.index, df[col], marker='o', label=col)

plt.title("Tableau de valeurs")
plt.xlabel("Lignes")
plt.ylabel("Valeurs")
plt.legend(bbox_to_anchor=(.5,1), loc='upper left')
plt.grid(True)


image_path = 'graphique_fournis.png'
plt.savefig(image_path)


wb = load_workbook(excel_path)
ws = wb.active


img = Image(image_path)
img.anchor = 'E5'  
ws.add_image(img)


wb.save(excel_path)

print(f"Le fichier Excel avec le tableau de valeurs et le graphique a été créé : {excel_path}")
