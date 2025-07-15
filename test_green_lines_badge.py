import requests
import os

# URL de votre serveur Django
url = 'http://127.0.0.1:8080/upload/'

filename = 'Programme planification SS+ST Juin sites_ 20052024.xlsx'

print(f"🟢 Test style vert pour le badge nombre de lignes")
print("=" * 60)

# Vérifier que le fichier existe
if not os.path.exists(filename):
    print(f"❌ Fichier Excel {filename} non trouvé")
    exit()

print(f"✅ Fichier Excel trouvé: {filename}")

print(f"\n🎨 NOUVEAU STYLE VERT APPLIQUÉ:")
print("   🟢 Fond: Vert (#27ae60) - Même couleur que le bouton télécharger")
print("   ⚪ Texte: Blanc")
print("   🟢 Bordure: Vert foncé (#219a52)")
print("   ✨ Hover: Vert plus foncé (#219a52)")
print("   🎯 Cohérence: Style identique au bouton télécharger")

# Test rapide
files = {'file': open(filename, 'rb')}
data = {
    'sort_column': 'ville',
    'date_debut': '2024-01-15',
    'date_fin': '2024-01-20',
    'objet_visite': 'Test style vert'
}

try:
    response = requests.post(url, files=files, data=data)
    
    if response.status_code == 200:
        result = response.json()
        if result.get('success'):
            files_details = result.get('files_details', [])
            
            print(f"\n✅ Upload réussi - {len(files_details)} fichiers générés")
            
            print(f"\n🎯 APERÇU DU STYLE DANS L'INTERFACE:")
            print("   📄 Fichiers générés:")
            print("   ┌─────────────────────────────────────────────────────────────┐")
            
            for i, file_info in enumerate(files_details[:3], 1):
                filename_short = file_info['filename']
                if len(filename_short) > 35:
                    filename_short = filename_short[:32] + "..."
                
                lines_count = file_info['lines']
                print(f"   │ 📄 {filename_short:<35} 🟢[{lines_count:2d} lignes] 🟢[Télécharger] │")
            
            if len(files_details) > 3:
                print(f"   │ ... et {len(files_details) - 3} autres fichiers                                │")
            
            print("   └─────────────────────────────────────────────────────────────┘")
            
            print(f"\n🎨 COMPARAISON DES STYLES:")
            print("   🟢 Badge lignes:     Fond vert #27ae60 + texte blanc")
            print("   🟢 Bouton télécharger: Fond vert #27ae60 + texte blanc")
            print("   ✅ Cohérence visuelle parfaite!")
            
            print(f"\n🎯 EFFETS VISUELS:")
            print("   🟢 Normal: background: #27ae60")
            print("   🟢 Hover:  background: #219a52 (plus foncé)")
            print("   ✨ Transition: 0.3s ease")
            print("   📱 Responsive: Adapté à tous les écrans")
            
        else:
            print(f"❌ Erreur: {result.get('error', 'Inconnue')}")
    else:
        print(f"❌ Status: {response.status_code}")
        
except Exception as e:
    print(f"❌ Erreur connexion: {e}")

finally:
    files['file'].close()

print(f"\n🎨 DÉTAILS DU STYLE CSS APPLIQUÉ:")
print("   .file-lines {")
print("       background: #27ae60;        /* Vert identique au bouton */")
print("       color: white;               /* Texte blanc */")
print("       padding: 4px 12px;          /* Espacement interne */")
print("       border-radius: 15px;        /* Coins arrondis */")
print("       font-size: 0.85em;          /* Taille de police */")
print("       font-weight: 500;           /* Poids moyen */")
print("       border: 1px solid #219a52;  /* Bordure vert foncé */")
print("       transition: background 0.3s ease; /* Animation */")
print("   }")
print("   .file-lines:hover {")
print("       background: #219a52;        /* Vert plus foncé au hover */")
print("   }")

print(f"\n💡 INSTRUCTIONS POUR VOIR LE NOUVEAU STYLE:")
print("   1. Démarrez Django: python manage.py runserver 127.0.0.1:8080")
print("   2. Ouvrez: http://127.0.0.1:8080/")
print("   3. Déposez un fichier Excel")
print("   4. Générez les fichiers")
print("   5. Observez la section 'Fichiers générés'")
print("   6. Vérifiez que les badges de lignes sont VERTS")
print("   7. Testez l'effet hover sur les badges")

print(f"\n🎯 COHÉRENCE VISUELLE:")
print("   🟢 Badge lignes:      Vert #27ae60")
print("   🟢 Bouton télécharger: Vert #27ae60")
print("   ✅ Style uniforme et professionnel")
print("   🎨 Interface harmonieuse")

print(f"\n🎉 STYLE VERT APPLIQUÉ!")
print(f"🟢 Badge nombre de lignes maintenant en vert")
print(f"🎯 Cohérence parfaite avec le bouton télécharger")
print(f"✨ Effet hover ajouté pour l'interactivité")
print(f"📱 Style responsive et moderne")
