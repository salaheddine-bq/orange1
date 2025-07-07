import requests
import os
import json

# URL de votre serveur Django
url = 'http://127.0.0.1:8080/upload/'

# Colonnes de tri à tester
sort_columns = ['code site', 'DR IAm', 'ville', 'ST FO']

if os.path.exists('rich_test_data.xlsx'):
    print("📁 Test avec rich_test_data.xlsx (27 lignes)...")
    
    for sort_col in sort_columns:
        print(f"\n🔄 Test avec tri par '{sort_col}':")
        print("-" * 50)
        
        # Préparer les données
        files = {'file': open('rich_test_data.xlsx', 'rb')}
        data = {'sort_column': sort_col}
        
        try:
            # Envoyer la requête
            response = requests.post(url, files=files, data=data)
            
            if response.status_code == 200:
                result = response.json()
                if result.get('success'):
                    stats = result.get('stats', {})
                    
                    print(f"✅ Tri par {sort_col} réussi!")
                    print(f"📊 Total lignes: {stats.get('total_rows', 0)}")
                    print(f"📋 {stats.get('primary_label', 'Items')}: {stats.get('primary_count', 0)}")
                    print(f"📁 Groupes générés: {stats.get('group_count', 0)}")
                    print(f"🏷️  Titre: {stats.get('stats_title', 'N/A')}")
                    
                    # Afficher les top 5 des statistiques principales
                    primary_stats = stats.get('primary_stats', {})
                    if primary_stats:
                        print(f"🔝 Top 5 {stats.get('primary_label', 'items')}:")
                        sorted_items = sorted(primary_stats.items(), key=lambda x: x[1], reverse=True)[:5]
                        for name, count in sorted_items:
                            print(f"   - {name}: {count} ligne{'s' if count > 1 else ''}")
                else:
                    print(f"❌ Erreur: {result.get('error', 'Erreur inconnue')}")
            else:
                print(f"❌ Erreur HTTP {response.status_code}")
                
        except Exception as e:
            print(f"❌ Erreur de connexion: {e}")
        
        finally:
            files['file'].close()
            
else:
    print("❌ Fichier rich_test_data.xlsx non trouvé")
    print("Exécutez d'abord: python create_rich_test_excel.py")
