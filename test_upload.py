import requests
import os

# URL de votre serveur Django
url = 'http://127.0.0.1:8080/upload/'

# Vérifier si le fichier existe
if os.path.exists('test_data.xlsx'):
    print("📁 Upload du fichier test_data.xlsx...")
    
    # Préparer les données
    files = {'file': open('test_data.xlsx', 'rb')}
    data = {'sort_column': 'code site'}
    
    try:
        # Envoyer la requête
        response = requests.post(url, files=files, data=data)
        
        print(f"📊 Status Code: {response.status_code}")
        print(f"📄 Response: {response.text}")
        
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                print("✅ Upload réussi!")
                print(f"📁 Fichiers générés: {result.get('files', [])}")
            else:
                print("❌ Erreur dans la réponse:")
                print(result.get('error', 'Erreur inconnue'))
        else:
            print("❌ Erreur HTTP:")
            try:
                error_data = response.json()
                print(error_data.get('error', 'Erreur inconnue'))
                if 'available_columns' in error_data:
                    print(f"Colonnes disponibles: {error_data['available_columns']}")
            except:
                print(response.text)
                
    except Exception as e:
        print(f"❌ Erreur de connexion: {e}")
    
    finally:
        files['file'].close()
        
else:
    print("❌ Fichier test_data.xlsx non trouvé")
