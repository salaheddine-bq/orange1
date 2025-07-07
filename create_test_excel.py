import pandas as pd

# Créer des données de test
data = {
    'code site': ['SITE001', 'SITE002', 'SITE001', 'SITE003', 'SITE002', 'SITE001', 'SITE003', 'SITE004'],
    'DR IAm': ['DR_NORD', 'DR_SUD', 'DR_NORD', 'DR_EST', 'DR_SUD', 'DR_NORD', 'DR_EST', 'DR_OUEST'],
    'ville': ['Paris', 'Lyon', 'Lille', 'Strasbourg', 'Marseille', 'Roubaix', 'Nancy', 'Nantes'],
    'ST FO': ['ST_A', 'ST_B', 'ST_A', 'ST_C', 'ST_B', 'ST_A', 'ST_C', 'ST_D']
}

# Créer le DataFrame
df = pd.DataFrame(data)

# Sauvegarder en Excel
df.to_excel('test_data.xlsx', index=False)
print("Fichier test_data.xlsx créé avec succès!")
print("\nContenu du fichier:")
print(df)
