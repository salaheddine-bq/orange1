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

        # Créer un mapping flexible pour les noms de colonnes
        column_mapping = {}

        # Mapping des colonnes avec recherche très flexible
        def normalize_column_name(name):
            """Normalise un nom de colonne pour la comparaison"""
            return ''.join(name.lower().split())  # Supprime tous les espaces et met en minuscules

        column_searches = {
            'code site': ['codesite'],
            'ST FO': ['stfo'],
            'contact ERPT': ['contacterpt'],
            'DR IAm': ['driam'],
            'ville': ['ville'],
            'X Départ ERPT - Y Départ ERPT': ['xdéparterpt-ydéparterpt', 'xdéparterptydepart'],
            'X Arrivée ERPT Proposition1 - Y Arrivée ERPT Proposition1': ['xarrivéeerptproposition1-yarrivéeerptproposition1']
        }

        # Recherche très flexible des colonnes
        for target_col, search_terms in column_searches.items():
            found = False
            for actual_col in df.columns:
                actual_col_normalized = normalize_column_name(actual_col)
                for search_term in search_terms:
                    if search_term in actual_col_normalized or actual_col_normalized in search_term:
                        column_mapping[target_col] = actual_col
                        found = True
                        break
                if found:
                    break

        # Vérifier les colonnes essentielles
        essential_columns = ['code site', 'ST FO', 'DR IAm', 'ville']
        missing_columns = [col for col in essential_columns if col not in column_mapping]

        if missing_columns:
            return JsonResponse({
                'error': f'Colonnes essentielles manquantes: {", ".join(missing_columns)}. Colonnes disponibles: {", ".join(df.columns)}',
                'available_columns': list(df.columns),
                'column_mapping': column_mapping
            }, status=400)

        # Renommer les colonnes pour correspondre aux noms attendus
        rename_dict = {v: k for k, v in column_mapping.items()}
        df = df.rename(columns=rename_dict)

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
    from pptx.dml.color import RGBColor
    from pptx.util import Pt

    # Colonnes finales à afficher dans le PowerPoint
    final_columns = [
        'code site',
        'ST FO',
        'contact ERPT',
        'DR IAm',
        'ville',
        'X Départ ERPT - Y Départ ERPT',
        'X Arrivée ERPT Proposition1 - Y Arrivée ERPT Proposition1'
    ]

    # Créer une copie du DataFrame pour éviter de modifier l'original
    df_work = df.copy().reset_index(drop=True)  # Reset index pour éviter les problèmes d'ordre

    # Créer les colonnes manquantes avec des valeurs par défaut SEULEMENT si elles n'existent pas
    if 'contact ERPT' not in df_work.columns:
        # Générer des contacts ERPT basés sur l'index SEULEMENT si la colonne n'existe pas
        df_work['contact ERPT'] = [f"contact{i+1}@erpt.fr" for i in range(len(df_work))]

    # Gérer les colonnes X-Y (soit séparées, soit déjà combinées)
    # Vérifier si les colonnes sont déjà combinées dans le fichier Excel
    if 'X Départ ERPT - Y Départ ERPT' not in df_work.columns:
        # Les colonnes ne sont pas encore combinées, essayer de les combiner
        if 'X Départ ERPT' in df_work.columns and 'Y Départ ERPT' in df_work.columns:
            combined_depart = []
            for _, row in df_work.iterrows():
                x_val = row['X Départ ERPT']
                y_val = row['Y Départ ERPT']
                if pd.notna(x_val) and pd.notna(y_val) and str(x_val).strip() != '' and str(y_val).strip() != '':
                    combined_depart.append(f"{x_val} - {y_val}")
                else:
                    combined_depart.append("")
            df_work['X Départ ERPT - Y Départ ERPT'] = combined_depart
        else:
            df_work['X Départ ERPT - Y Départ ERPT'] = [""] * len(df_work)
    # Si la colonne existe déjà, s'assurer qu'elle contient des données valides
    else:
        # Nettoyer les valeurs NaN et vides
        df_work['X Départ ERPT - Y Départ ERPT'] = df_work['X Départ ERPT - Y Départ ERPT'].fillna("").astype(str)

    if 'X Arrivée ERPT Proposition1 - Y Arrivée ERPT Proposition1' not in df_work.columns:
        # Les colonnes ne sont pas encore combinées, essayer de les combiner
        if 'X Arrivée ERPT Proposition1' in df_work.columns and 'Y Arrivée ERPT Proposition1' in df_work.columns:
            combined_arrivee = []
            for _, row in df_work.iterrows():
                x_val = row['X Arrivée ERPT Proposition1']
                y_val = row['Y Arrivée ERPT Proposition1']
                if pd.notna(x_val) and pd.notna(y_val) and str(x_val).strip() != '' and str(y_val).strip() != '':
                    combined_arrivee.append(f"{x_val} - {y_val}")
                else:
                    combined_arrivee.append("")
            df_work['X Arrivée ERPT Proposition1 - Y Arrivée ERPT Proposition1'] = combined_arrivee
        else:
            df_work['X Arrivée ERPT Proposition1 - Y Arrivée ERPT Proposition1'] = [""] * len(df_work)
    # Si la colonne existe déjà, s'assurer qu'elle contient des données valides
    else:
        # Nettoyer les valeurs NaN et vides
        df_work['X Arrivée ERPT Proposition1 - Y Arrivée ERPT Proposition1'] = df_work['X Arrivée ERPT Proposition1 - Y Arrivée ERPT Proposition1'].fillna("").astype(str)

    # Créer les autres colonnes manquantes si nécessaire
    for col in final_columns:
        if col not in df_work.columns:
            df_work[col] = [""] * len(df_work)

    # Sélectionner seulement les colonnes finales à afficher dans l'ordre exact
    df_filtered = df_work[final_columns].copy()



    prs = Presentation()

    # Slide de titre
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = f"Données pour {sort_column}: {group_name}"
    subtitle.text = f"Nombre d'enregistrements: {len(df_filtered)}"

    # Appliquer la police Arial Narrow au titre
    title_paragraph = title.text_frame.paragraphs[0]
    title_paragraph.font.name = 'Arial Narrow'
    title_paragraph.font.size = Pt(32)

    subtitle_paragraph = subtitle.text_frame.paragraphs[0]
    subtitle_paragraph.font.name = 'Arial Narrow'
    subtitle_paragraph.font.size = Pt(18)

    # Slide avec tableau des données
    bullet_slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(bullet_slide_layout)
    title = slide.shapes.title
    title.text = f"Détails - {group_name}"

    # Appliquer la police Arial Narrow au titre de la slide
    title_paragraph = title.text_frame.paragraphs[0]
    title_paragraph.font.name = 'Arial Narrow'
    title_paragraph.font.size = Pt(24)

    # Ajouter un tableau
    rows = len(df_filtered) + 1  # +1 pour l'en-tête
    cols = len(final_columns)

    # Calculer la largeur totale nécessaire (somme des largeurs de colonnes)
    from pptx.util import Cm
    total_width_cm = 1.8 + 2.4 + 3.7 + 2.6 + 2.8 + 6.2 + 4.8  # = 24.3 cm

    # Centrer le tableau sur la slide
    left = Cm(1.0)  # Marge à gauche
    top = Inches(1.8)
    width = Cm(total_width_cm)  # Largeur exacte du tableau
    height = Inches(6)

    table = slide.shapes.add_table(rows, cols, left, top, width, height).table

    # En-têtes avec police Arial Narrow
    for i, column in enumerate(final_columns):
        cell = table.cell(0, i)
        cell.text = str(column)

        # Appliquer la police Arial Narrow et le style aux en-têtes
        paragraph = cell.text_frame.paragraphs[0]
        paragraph.font.name = 'Arial Narrow'
        paragraph.font.size = Pt(7)  # Taille réduite à 7pt
        paragraph.font.bold = True
        paragraph.font.color.rgb = RGBColor(255, 255, 255)  # Blanc

        # Réduire l'espacement dans la cellule
        cell.text_frame.margin_top = Inches(0.02)
        cell.text_frame.margin_bottom = Inches(0.02)
        cell.text_frame.margin_left = Inches(0.05)
        cell.text_frame.margin_right = Inches(0.05)

        # Couleur de fond pour l'en-tête
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(68, 114, 196)  # Bleu

    # Données avec police Arial Narrow
    for i, (_, row) in enumerate(df_filtered.iterrows(), 1):
        for j, value in enumerate(row):
            cell = table.cell(i, j)
            cell.text = str(value) if pd.notna(value) and str(value) != "" else ""

            # Appliquer la police Arial Narrow aux données
            paragraph = cell.text_frame.paragraphs[0]
            paragraph.font.name = 'Arial Narrow'
            paragraph.font.size = Pt(7)  # Taille réduite à 7pt

            # Réduire l'espacement dans la cellule de données
            cell.text_frame.margin_top = Inches(0.02)
            cell.text_frame.margin_bottom = Inches(0.02)
            cell.text_frame.margin_left = Inches(0.05)
            cell.text_frame.margin_right = Inches(0.05)

            # Couleur alternée pour les lignes
            if i % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(242, 242, 242)  # Gris clair

    # Largeurs de colonnes spécifiées en centimètres (nouvelles spécifications)
    from pptx.util import Cm

    column_widths = [
        Cm(1.8),   # colonne 1: code site - 1,8 cm
        Cm(2.4),   # colonne 2: ST FO - 2,4 cm
        Cm(3.7),   # colonne 3: contact ERPT - 3,7 cm
        Cm(2.6),   # colonne 4: DR IAm - 2,6 cm
        Cm(2.8),   # colonne 5: ville - 2,8 cm
        Cm(6.2),   # colonne 6: X Départ ERPT - Y Départ ERPT - 6,2 cm
        Cm(4.8)    # colonne 7: X Arrivée ERPT Proposition1 - Y Arrivée ERPT Proposition1 - 4,8 cm
    ]

    # Appliquer les largeurs aux colonnes
    for i, width in enumerate(column_widths):
        if i < len(table.columns):
            table.columns[i].width = width

    # Ajuster la hauteur des lignes selon le type (en-tête vs données)
    from pptx.util import Cm

    for i, row in enumerate(table.rows):
        if i == 0:  # En-tête
            row.height = Cm(0.35)  # En-tête: 0.35 cm
        else:  # Lignes de données
            row.height = Cm(0.26)  # Données: 0.26 cm

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
