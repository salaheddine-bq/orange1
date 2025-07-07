from django.shortcuts import render
from django.http import JsonResponse, HttpResponse, Http404
from django.views.decorators.csrf import csrf_exempt
from django.conf import settings
import pandas as pd
import os
from pptx import Presentation
from pptx.util import Inches

def index(request):
    """Page principale avec l'interface de upload"""
    return render(request, 'file_processor/index.html')

@csrf_exempt
def upload_file(request):
    """Traite le fichier Excel uploadé et génère les PowerPoints"""
    if request.method != 'POST':
        return JsonResponse({'error': 'Méthode non autorisée'}, status=405)

    if 'file' not in request.FILES:
        return JsonResponse({'error': 'Aucun fichier fourni'}, status=400)

    file = request.FILES['file']
    sort_column = request.POST.get('sort_column', 'code site')

    # Vérifier l'extension du fichier
    if not file.name.endswith(('.xlsx', '.xls')):
        return JsonResponse({'error': 'Le fichier doit être un fichier Excel (.xlsx ou .xls)'}, status=400)

    try:
        # Lire le fichier Excel
        df = pd.read_excel(file)

        # Debug: afficher les colonnes disponibles
        print(f"Colonnes disponibles dans le fichier: {list(df.columns)}")
        print(f"Forme du DataFrame: {df.shape}")
        print(f"Premières lignes:\n{df.head()}")

        # Normaliser les noms de colonnes (supprimer espaces en début/fin et convertir en minuscules pour comparaison)
        df.columns = df.columns.str.strip()

        # Vérifier que les colonnes requises existent
        required_columns = ['code site', 'DR IAm', 'ville', 'ST FO']

        # Créer un mapping flexible pour les noms de colonnes
        column_mapping = {}
        available_cols_lower = [col.lower().strip() for col in df.columns]

        for req_col in required_columns:
            req_col_lower = req_col.lower().strip()
            found = False
            for i, avail_col in enumerate(available_cols_lower):
                if req_col_lower == avail_col or req_col_lower.replace(' ', '') == avail_col.replace(' ', ''):
                    column_mapping[req_col] = df.columns[i]
                    found = True
                    break
            if not found:
                # Chercher des correspondances partielles
                for i, avail_col in enumerate(available_cols_lower):
                    if req_col_lower.replace(' ', '') in avail_col.replace(' ', '') or avail_col.replace(' ', '') in req_col_lower.replace(' ', ''):
                        column_mapping[req_col] = df.columns[i]
                        found = True
                        break

        missing_columns = [col for col in required_columns if col not in column_mapping]

        if missing_columns:
            return JsonResponse({
                'error': f'Colonnes manquantes: {", ".join(missing_columns)}. Colonnes disponibles: {", ".join(df.columns)}',
                'available_columns': list(df.columns),
                'required_columns': required_columns
            }, status=400)

        # Renommer les colonnes pour correspondre aux noms attendus
        df = df.rename(columns={v: k for k, v in column_mapping.items()})

        # Trier les données selon la colonne choisie
        if sort_column not in df.columns:
            return JsonResponse({'error': f'Colonne de tri "{sort_column}" non trouvée'}, status=400)

        # Calculer les statistiques selon la colonne de tri
        total_rows = len(df)

        # Statistiques par colonne de tri
        sort_stats = df[sort_column].value_counts().to_dict()

        # Adapter les statistiques selon la colonne choisie
        if sort_column == 'DR IAm':
            # Pour DR IAm : afficher DR stats
            primary_stats = df['DR IAm'].value_counts().to_dict()
            primary_count = len(primary_stats)
            primary_label = 'DR'
            stats_title = 'Répartition par DR IAm'
            show_details = True
        elif sort_column == 'ville':
            # Pour ville : afficher ville stats
            primary_stats = df['ville'].value_counts().to_dict()
            primary_count = len(primary_stats)
            primary_label = 'ville'
            stats_title = 'Répartition par ville'
            show_details = True
        elif sort_column == 'ST FO':
            # Pour ST FO : afficher ST FO stats
            primary_stats = df['ST FO'].value_counts().to_dict()
            primary_count = len(primary_stats)
            primary_label = 'ST FO'
            stats_title = 'Répartition par ST FO'
            show_details = True
        else:  # code site - simple (pas de détails)
            # Pour code site : pas de statistiques détaillées
            primary_stats = {}
            primary_count = len(df['code site'].value_counts())
            primary_label = 'code site'
            stats_title = ''
            show_details = False

        # Statistiques générales
        stats = {
            'total_rows': total_rows,
            'primary_count': primary_count,
            'primary_stats': primary_stats,
            'primary_label': primary_label,
            'stats_title': stats_title,
            'sort_column': sort_column,
            'sort_stats': sort_stats,
            'group_count': len(sort_stats),
            'show_details': show_details
        }

        # Grouper les données par la colonne de tri
        grouped_data = df.groupby(sort_column)

        # Créer un dossier temporaire pour les fichiers PowerPoint
        output_dir = os.path.join(settings.MEDIA_ROOT, 'generated_ppts')
        os.makedirs(output_dir, exist_ok=True)

        generated_files = []

        # Générer un PowerPoint pour chaque groupe
        for group_name, group_df in grouped_data:
            ppt_filename = f"{sort_column}_{group_name}.pptx"
            ppt_path = os.path.join(output_dir, ppt_filename)

            # Créer la présentation PowerPoint
            create_powerpoint(group_df, ppt_path, group_name, sort_column)
            generated_files.append(ppt_filename)

        return JsonResponse({
            'success': True,
            'message': f'{len(generated_files)} fichiers PowerPoint générés',
            'files': generated_files,
            'stats': stats
        })

    except Exception as e:
        return JsonResponse({'error': f'Erreur lors du traitement: {str(e)}'}, status=500)

def create_powerpoint(df, output_path, group_name, sort_column):
    """Crée une présentation PowerPoint à partir des données"""
    prs = Presentation()

    # Slide de titre
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = f"Données pour {sort_column}: {group_name}"
    subtitle.text = f"Nombre d'enregistrements: {len(df)}"

    # Slide avec tableau des données
    bullet_slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(bullet_slide_layout)
    title = slide.shapes.title
    title.text = f"Détails - {group_name}"

    # Ajouter un tableau
    rows = len(df) + 1  # +1 pour l'en-tête
    cols = len(df.columns)

    left = Inches(0.5)
    top = Inches(2)
    width = Inches(9)
    height = Inches(5)

    table = slide.shapes.add_table(rows, cols, left, top, width, height).table

    # En-têtes
    for i, column in enumerate(df.columns):
        table.cell(0, i).text = str(column)

    # Données
    for i, (_, row) in enumerate(df.iterrows(), 1):
        for j, value in enumerate(row):
            table.cell(i, j).text = str(value) if pd.notna(value) else ""

    # Sauvegarder
    prs.save(output_path)

def download_file(request, filename):
    """Permet de télécharger un fichier PowerPoint généré"""
    file_path = os.path.join(settings.MEDIA_ROOT, 'generated_ppts', filename)

    if not os.path.exists(file_path):
        raise Http404("Fichier non trouvé")

    with open(file_path, 'rb') as f:
        response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.presentationml.presentation')
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response
