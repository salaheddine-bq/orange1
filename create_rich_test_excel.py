import pandas as pd
import numpy as np

# Créer des données de test plus riches
data = {
    'code site': [
        'SITE001', 'SITE002', 'SITE001', 'SITE003', 'SITE002', 'SITE001', 'SITE003', 'SITE004',
        'SITE005', 'SITE001', 'SITE002', 'SITE006', 'SITE003', 'SITE004', 'SITE005', 'SITE006',
        'SITE007', 'SITE001', 'SITE002', 'SITE008', 'SITE003', 'SITE009', 'SITE010', 'SITE001'
    ],
    'DR IAm': [
        'DR_NORD', 'DR_SUD', 'DR_NORD', 'DR_EST', 'DR_SUD', 'DR_NORD', 'DR_EST', 'DR_OUEST',
        'DR_CENTRE', 'DR_NORD', 'DR_SUD', 'DR_EST', 'DR_EST', 'DR_OUEST', 'DR_CENTRE', 'DR_EST',
        'DR_SUD', 'DR_NORD', 'DR_SUD', 'DR_OUEST', 'DR_EST', 'DR_CENTRE', 'DR_NORD', 'DR_NORD'
    ],
    'ville': [
        'Paris', 'Lyon', 'Lille', 'Strasbourg', 'Marseille', 'Roubaix', 'Nancy', 'Nantes',
        'Clermont-Ferrand', 'Amiens', 'Toulon', 'Metz', 'Mulhouse', 'Brest', 'Limoges', 'Reims',
        'Montpellier', 'Créteil', 'Nice', 'Bordeaux', 'Colmar', 'Orléans', 'Rouen', 'Versailles'
    ],
    'ST FO': [
        'ST_A', 'ST_B', 'ST_A', 'ST_C', 'ST_B', 'ST_A', 'ST_C', 'ST_D',
        'ST_E', 'ST_A', 'ST_B', 'ST_C', 'ST_C', 'ST_D', 'ST_E', 'ST_C',
        'ST_B', 'ST_A', 'ST_B', 'ST_D', 'ST_C', 'ST_E', 'ST_A', 'ST_A'
    ]
}

# Créer le DataFrame
df = pd.DataFrame(data)

# Ajouter quelques lignes avec des valeurs manquantes pour tester
df.loc[len(df)] = ['SITE011', np.nan, 'Toulouse', 'ST_F']  # DR manquant
df.loc[len(df)] = ['SITE012', 'DR_SUD', np.nan, 'ST_B']    # Ville manquante
df.loc[len(df)] = [np.nan, 'DR_NORD', 'Dijon', 'ST_A']     # Code site manquant

# Sauvegarder en Excel
df.to_excel('rich_test_data.xlsx', index=False)
print("Fichier rich_test_data.xlsx créé avec succès!")
print(f"\nStatistiques du fichier:")
print(f"- Total de lignes: {len(df)}")
print(f"- Lignes complètes: {len(df.dropna())}")
print(f"- Lignes avec valeurs manquantes: {len(df) - len(df.dropna())}")

print(f"\nRépartition par DR IAm:")
dr_counts = df['DR IAm'].value_counts()
for dr, count in dr_counts.items():
    print(f"- {dr}: {count} lignes")

print(f"\nRépartition par code site:")
site_counts = df['code site'].value_counts()
for site, count in site_counts.head(10).items():  # Top 10
    print(f"- {site}: {count} lignes")

print(f"\nAperçu des données:")
print(df.head(10))
