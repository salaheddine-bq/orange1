import pandas as pd
import os

# Vérifier si le fichier existe
if os.path.exists('test_data.xlsx'):
    print("✅ Fichier test_data.xlsx trouvé")
    
    # Lire le fichier Excel
    df = pd.read_excel('test_data.xlsx')
    
    print(f"\n📊 Informations sur le fichier:")
    print(f"   - Nombre de lignes: {len(df)}")
    print(f"   - Nombre de colonnes: {len(df.columns)}")
    
    print(f"\n📋 Colonnes disponibles:")
    for i, col in enumerate(df.columns):
        print(f"   {i+1}. '{col}' (type: {type(col).__name__})")
    
    print(f"\n🔍 Colonnes requises:")
    required = ['code site', 'DR IAm', 'ville', 'ST FO']
    for col in required:
        if col in df.columns:
            print(f"   ✅ '{col}' - TROUVÉE")
        else:
            print(f"   ❌ '{col}' - MANQUANTE")
    
    print(f"\n📄 Aperçu des données:")
    print(df.head())
    
    print(f"\n🔤 Types de données:")
    print(df.dtypes)
    
else:
    print("❌ Fichier test_data.xlsx non trouvé")
    print("Fichiers disponibles dans le répertoire:")
    for file in os.listdir('.'):
        if file.endswith(('.xlsx', '.xls')):
            print(f"   - {file}")
