import os
import pandas as pd
from .views import create_powerpoint

# Create your tests here.

def test_create_powerpoint_formatting():
    # Création d'un DataFrame simulé avec toutes les colonnes
    data = {
        'code site': ['A1'],
        'ST FO': ['ST1'],
        'contact ERPT': ['contact1@erpt.fr'],
        'Contact IAM': ['iam1'],
        'DR IAM': ['DR1'],
        'ville': ['Ville1'],
        'Date TSS': ['2024-06-01'],
        'X Départ ERPT - Y Départ ERPT': ['-13.207345678 17.14526789'],
        'X Arrivée ERPT Proposition1 - Y Arrivée': ['27.14829876 -13.20555321']
    }
    df = pd.DataFrame(data)
    output_path = 'test_ppt.pptx'
    group_name = 'TestGroup'
    sort_column = 'code site'
    # Appel de la fonction
    files = create_powerpoint(df, output_path, group_name, sort_column)
    # Vérification du fichier généré
    assert os.path.exists(files[0]['filename']), 'Le fichier PowerPoint n\'a pas été généré.'
    # Nettoyage
    os.remove(files[0]['filename'])
    print('Test create_powerpoint_formatting OK')
