from django.shortcuts import render
from django.http import JsonResponse, HttpResponse, Http404
from django.views.decorators.csrf import csrf_exempt
from django.conf import settings
import pandas as pd
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

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
        # Calculer la taille du fichier
        file.seek(0, 2)  # Aller à la fin du fichier
        file_size_bytes = file.tell()
        file.seek(0)  # Revenir au début

        # Convertir en format lisible
        if file_size_bytes < 1024:
            file_size_display = f"{file_size_bytes} B"
        elif file_size_bytes < 1024 * 1024:
            file_size_display = f"{file_size_bytes / 1024:.1f} KB"
        else:
            file_size_display = f"{file_size_bytes / (1024 * 1024):.1f} MB"

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
            'DR IAM': ['driam'],
            'ville': ['ville'],
            'X Départ ERPT - Y Départ ERPT': ['xdéparterpt-ydéparterpt', 'xdéparterptydepart'],
            'X Arrivée ERPT Proposition1 - Y Arrivée': ['xarrivéeerptproposition1-yarrivéeerptproposition1']
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
        essential_columns = ['code site', 'ST FO', 'DR IAM', 'ville']
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
        if sort_column == 'DR IAM':
            # Pour DR IAM : afficher DR stats
            primary_stats = df['DR IAM'].value_counts().to_dict()
            primary_count = len(primary_stats)
            primary_label = 'DR'
            stats_title = 'Répartition par DR IAM'
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

        # Éviter la duplication des statistiques si on trie par DR IAM
        if sort_column == 'DR IAM':
            # Ne pas envoyer sort_stats pour éviter la duplication
            final_sort_stats = {}
        else:
            final_sort_stats = sort_stats

        # Statistiques générales
        stats = {
            'total_rows': total_rows,
            'primary_count': primary_count,
            'primary_stats': primary_stats,
            'primary_label': primary_label,
            'stats_title': stats_title,
            'sort_column': sort_column,
            'sort_stats': final_sort_stats,
            'group_count': len(final_sort_stats),
            'show_details': show_details,
            'file_size': file_size_display
        }

        # Grouper les données par la colonne de tri
        grouped_data = df.groupby(sort_column)

        # Créer un dossier temporaire pour les fichiers PowerPoint
        output_dir = os.path.join(settings.MEDIA_ROOT, 'generated_ppts')
        os.makedirs(output_dir, exist_ok=True)

        # Nettoyer les anciens fichiers du même type de tri
        sort_prefix = sort_column.replace(' ', ' ')  # Garder les espaces comme dans les noms de fichiers
        cleaned_count = 0
        for existing_file in os.listdir(output_dir):
            if existing_file.startswith(f"{sort_prefix}_") and existing_file.endswith('.pptx'):
                os.remove(os.path.join(output_dir, existing_file))
                cleaned_count += 1


        generated_files = []

        # Générer un PowerPoint pour chaque groupe
        for group_name, group_df in grouped_data:
            ppt_filename = f"{sort_column}_{group_name}.pptx"
            ppt_path = os.path.join(output_dir, ppt_filename)

            # Créer les présentations PowerPoint (peut créer plusieurs fichiers)
            created_files = create_powerpoint(group_df, ppt_path, group_name, sort_column)
            generated_files.extend(created_files)

        # Ajouter le nombre de fichiers générés aux statistiques
        stats['file_count'] = len(generated_files)

        return JsonResponse({
            'success': True,
            'message': f'{len(generated_files)} fichiers PowerPoint générés',
            'files': generated_files,
            'stats': stats
        })

    except Exception as e:
        return JsonResponse({'error': f'Erreur lors du traitement: {str(e)}'}, status=500)

def load_prestataire_sheets():
    """Charge toutes les feuilles du fichier Excel 'Liste prestataire'"""
    try:
        # Chercher le fichier Liste prestataire dans le répertoire courant
        prestataire_files = [
            'Liste prestataire.xlsx',
            'Liste prestataire.xls',
            'liste prestataire.xlsx',
            'liste prestataire.xls'
        ]

        for filename in prestataire_files:
            if os.path.exists(filename):
                # Lire toutes les feuilles du fichier Excel
                all_sheets = pd.read_excel(filename, sheet_name=None)
                print(f"📋 Fichier prestataires trouvé: {filename}")
                print(f"📊 Feuilles disponibles: {list(all_sheets.keys())}")

                # Afficher un aperçu de chaque feuille
                for sheet_name, df in all_sheets.items():
                    print(f"   📄 Feuille '{sheet_name}': {len(df)} lignes, colonnes: {list(df.columns)}")

                return all_sheets

        print("⚠️ Aucun fichier 'Liste prestataire' trouvé")
        return None
    except Exception as e:
        print(f"❌ Erreur lors du chargement des prestataires: {e}")
        return None

def get_prestataire_sheets_for_st_fo(all_sheets, st_fo_list):
    """Récupère les feuilles correspondantes aux ST FO donnés"""
    if all_sheets is None:
        return {}

    try:
        matching_sheets = {}

        print(f"🔍 Recherche de feuilles pour ST FO: {st_fo_list}")
        print(f"📋 Feuilles disponibles: {list(all_sheets.keys())}")

        for st_fo in st_fo_list:
            if pd.notna(st_fo) and str(st_fo).strip():
                st_fo_clean = str(st_fo).strip()
                found = False

                # Chercher une feuille qui correspond EXACTEMENT au nom du ST FO
                for sheet_name, df_sheet in all_sheets.items():
                    sheet_name_clean = str(sheet_name).strip()

                    # Correspondance exacte (insensible à la casse)
                    if st_fo_clean.lower() == sheet_name_clean.lower():
                        print(f"✅ Correspondance exacte trouvée: '{st_fo_clean}' → feuille '{sheet_name_clean}'")
                        matching_sheets[st_fo_clean] = {
                            'sheet_name': sheet_name_clean,
                            'dataframe': df_sheet
                        }
                        found = True
                        break

                    # Correspondance partielle (ST FO contenu dans le nom de la feuille)
                    elif st_fo_clean.lower() in sheet_name_clean.lower():
                        print(f"✅ Correspondance partielle trouvée: '{st_fo_clean}' → feuille '{sheet_name_clean}'")
                        matching_sheets[st_fo_clean] = {
                            'sheet_name': sheet_name_clean,
                            'dataframe': df_sheet
                        }
                        found = True
                        break

                if not found:
                    print(f"⚠️ Aucune feuille trouvée pour ST FO: '{st_fo_clean}' - IGNORÉ")

        print(f"📊 Résultat: {len(matching_sheets)} feuilles trouvées sur {len(st_fo_list)} ST FO")
        return matching_sheets
    except Exception as e:
        print(f"❌ Erreur lors de la récupération des feuilles: {e}")
        return {}

def create_powerpoint(df, output_path, group_name, sort_column):
    """Crée des présentations PowerPoint à partir des données avec division en fichiers (max 19 lignes par fichier)"""
    from pptx.dml.color import RGBColor
    from pptx.util import Pt, Cm
    import random

    # Colonnes finales à afficher dans le PowerPoint
    final_columns = [
        'code site',
        'ST FO',
        'contact ERPT',
        'DR IAM',
        'ville',
        'X Départ ERPT - Y Départ ERPT',
        'X Arrivée ERPT Proposition1 - Y Arrivée'
    ]

    # Chemin vers l'image AA
    image_path = 'AA.jpeg'
    image_exists = os.path.exists(image_path)

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

    if 'X Arrivée ERPT Proposition1 - Y Arrivée' not in df_work.columns:
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
            df_work['X Arrivée ERPT Proposition1 - Y Arrivée'] = combined_arrivee
        else:
            df_work['X Arrivée ERPT Proposition1 - Y Arrivée'] = [""] * len(df_work)
    # Si la colonne existe déjà, s'assurer qu'elle contient des données valides
    else:
        # Nettoyer les valeurs NaN et vides
        df_work['X Arrivée ERPT Proposition1 - Y Arrivée'] = df_work['X Arrivée ERPT Proposition1 - Y Arrivée'].fillna("").astype(str)

    # Créer les autres colonnes manquantes si nécessaire
    for col in final_columns:
        if col not in df_work.columns:
            df_work[col] = [""] * len(df_work)

    # Sélectionner seulement les colonnes finales à afficher dans l'ordre exact
    df_filtered = df_work[final_columns].copy()

    # Nettoyer les données
    for col in df_filtered.columns:
        df_filtered[col] = df_filtered[col].fillna("").astype(str)

    # Appliquer le deuxième tri par ST FO pour diviser en sous-groupes
    df_filtered = df_filtered.sort_values('ST FO', na_position='last').reset_index(drop=True)

    # Grouper par ST FO (deuxième niveau de tri)
    st_fo_groups = df_filtered.groupby('ST FO', sort=False)

    # Optimiser la distribution des ST FO dans les chunks
    chunks = []
    chunk_names = []

    # Collecter tous les groupes ST FO avec leurs tailles
    st_fo_data_list = []
    for st_fo_name, st_fo_group in st_fo_groups:
        st_fo_data_list.append({
            'name': st_fo_name,
            'data': st_fo_group.reset_index(drop=True),
            'size': len(st_fo_group)
        })

    # Algorithme d'optimisation pour remplir les chunks
    current_chunk = pd.DataFrame()
    current_chunk_st_fo = []
    remaining_st_fo = st_fo_data_list.copy()

    while remaining_st_fo:
        current_space = 19 - len(current_chunk)

        if current_space <= 0:
            # Chunk plein, le sauvegarder
            chunks.append(current_chunk.copy())
            if len(current_chunk_st_fo) == 1:
                chunk_names.append(current_chunk_st_fo[0])
            else:
                chunk_names.append(" + ".join(current_chunk_st_fo))

            # Commencer un nouveau chunk
            current_chunk = pd.DataFrame()
            current_chunk_st_fo = []
            continue

        # Chercher le meilleur ST FO à ajouter
        best_fit = None
        best_index = -1

        for i, st_fo_item in enumerate(remaining_st_fo):
            if st_fo_item['size'] <= current_space:
                # Ce ST FO peut entrer entièrement
                if best_fit is None or st_fo_item['size'] > best_fit['size']:
                    best_fit = st_fo_item
                    best_index = i

        if best_fit:
            # Ajouter le ST FO entier au chunk
            if len(current_chunk) == 0:
                current_chunk = best_fit['data'].copy()
            else:
                current_chunk = pd.concat([current_chunk, best_fit['data']], ignore_index=True)
            current_chunk_st_fo.append(best_fit['name'])
            remaining_st_fo.pop(best_index)
        else:
            # Aucun ST FO ne peut entrer entièrement, prendre une partie du plus grand
            if remaining_st_fo:
                largest_st_fo = max(remaining_st_fo, key=lambda x: x['size'])
                largest_index = remaining_st_fo.index(largest_st_fo)

                # Prendre autant de lignes que possible
                lines_to_take = min(current_space, largest_st_fo['size'])
                partial_data = largest_st_fo['data'].iloc[:lines_to_take].copy()

                if len(current_chunk) == 0:
                    current_chunk = partial_data
                else:
                    current_chunk = pd.concat([current_chunk, partial_data], ignore_index=True)

                # Mettre à jour le nom du chunk
                if lines_to_take == largest_st_fo['size']:
                    # ST FO entier utilisé
                    current_chunk_st_fo.append(largest_st_fo['name'])
                    remaining_st_fo.pop(largest_index)
                else:
                    # ST FO partiellement utilisé
                    current_chunk_st_fo.append(f"{largest_st_fo['name']} (partiel)")
                    # Mettre à jour le ST FO restant
                    remaining_st_fo[largest_index]['data'] = largest_st_fo['data'].iloc[lines_to_take:].reset_index(drop=True)
                    remaining_st_fo[largest_index]['size'] = len(remaining_st_fo[largest_index]['data'])

    # Ajouter le dernier chunk s'il n'est pas vide
    if len(current_chunk) > 0:
        chunks.append(current_chunk)
        if len(current_chunk_st_fo) == 1:
            chunk_names.append(current_chunk_st_fo[0])
        else:
            chunk_names.append(" + ".join(current_chunk_st_fo))

    # Créer plusieurs fichiers PowerPoint si nécessaire
    created_files = []

    for chunk_index, chunk_data in enumerate(chunks):
        # Créer un nouveau PowerPoint pour chaque chunk
        prs = Presentation()

        # Déterminer le nom du fichier basé sur ST FO
        chunk_st_fo_name = chunk_names[chunk_index]
        chunk_lines = len(chunk_data)

        # Slide de titre
        title_slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(title_slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]

        title.text = f"Données pour {sort_column}: {group_name} - {chunk_st_fo_name}"
        subtitle.text = f"Lignes: {chunk_lines} | ST FO: {chunk_st_fo_name}"

        # Appliquer la police Arial Narrow au titre
        title_paragraph = title.text_frame.paragraphs[0]
        title_paragraph.font.name = 'Arial Narrow'
        title_paragraph.font.size = Pt(32)

        subtitle_paragraph = subtitle.text_frame.paragraphs[0]
        subtitle_paragraph.font.name = 'Arial Narrow'
        subtitle_paragraph.font.size = Pt(18)
        # Slide avec l'image AA et le tableau en dessous (sans titre)
        blank_slide_layout = prs.slide_layouts[6]  # Layout vide sans titre
        data_slide = prs.slides.add_slide(blank_slide_layout)

        # Ajouter l'image AA si elle existe
        if image_exists:
            try:
                # Positionner l'image tout en haut de la slide (sans titre)
                left = Cm(1.0)
                top = Cm(0.5)  # Tout en haut avec petite marge
                width = Cm(26.0)  # Largeur adaptée à la slide
                height = Cm(7.0)  # Hauteur réduite pour laisser plus de place au tableau

                data_slide.shapes.add_picture(image_path, left, top, width, height)
            except Exception as e:
                # Si erreur avec l'image, continuer sans l'image
                pass

        # Ajouter le tableau en dessous de l'image pour ce chunk
        rows = len(chunk_data) + 1  # +1 pour l'en-tête
        cols = len(final_columns)

        # Calculer la largeur totale nécessaire (somme des largeurs de colonnes)
        total_width_cm = 1.8 + 2.4 + 3.7 + 2.6 + 2.8 + 6.2 + 4.8  # = 24.3 cm

        # Positionner le tableau DIRECTEMENT collé à l'image (aucun espace)
        left = Cm(1.0)  # Marge à gauche
        if image_exists:
            top = Cm(7.5)  # DIRECTEMENT collé à l'image (0.5 + 7.0 + 0.0 espace)
        else:
            top = Cm(1.0)   # Plus haut si pas d'image
        width = Cm(total_width_cm)  # Largeur exacte du tableau
        height = Inches(6)

        table = data_slide.shapes.add_table(rows, cols, left, top, width, height).table

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

        # Données avec police Arial Narrow (utiliser chunk_data au lieu de df_filtered)
        for i, (_, row) in enumerate(chunk_data.iterrows(), 1):
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
            Cm(2.6),   # colonne 4: DR IAM - 2,6 cm
            Cm(2.8),   # colonne 5: ville - 2,8 cm
            Cm(6.2),   # colonne 6: X Départ ERPT - Y Départ ERPT - 6,2 cm
            Cm(4.8)    # colonne 7: X Arrivée ERPT Proposition1 - Y Arrivée - 4,8 cm
        ]

        # Appliquer les largeurs aux colonnes
        for i, width in enumerate(column_widths):
            if i < len(table.columns):
                table.columns[i].width = width

        # Ajuster la hauteur des lignes - format uniforme ultra-compact

        for i, row in enumerate(table.rows):
            if i == 0:  # En-tête
                row.height = Cm(0.10)  # En-tête: 0.10 cm
            else:  # Toutes les lignes de données
                row.height = Cm(0.05)  # Toutes les données: 0.05 cm

        # Ajouter le 3ème slide avec la liste des prestataires
        try:
            # Charger toutes les feuilles du fichier prestataires
            all_sheets = load_prestataire_sheets()

            if all_sheets is not None:
                # Récupérer les ST FO uniques de ce chunk
                st_fo_list = chunk_data['ST FO'].dropna().unique().tolist()

                # Récupérer les feuilles correspondantes aux ST FO
                matching_sheets = get_prestataire_sheets_for_st_fo(all_sheets, st_fo_list)

                if matching_sheets:
                    # Créer le 3ème slide pour chaque ST FO trouvé
                    for st_fo, sheet_info in matching_sheets.items():
                        prestataire_slide_layout = prs.slide_layouts[6]  # Layout vide
                        prestataire_slide = prs.slides.add_slide(prestataire_slide_layout)

                        # Ajouter le titre manuellement
                        title_shape = prestataire_slide.shapes.add_textbox(
                            Cm(1), Cm(1), Cm(23), Cm(2)
                        )
                        title_frame = title_shape.text_frame
                        title_frame.text = f"Liste des Prestataires - {st_fo}"
                        title_paragraph = title_frame.paragraphs[0]
                        title_paragraph.font.name = 'Arial Narrow'
                        title_paragraph.font.size = Pt(24)
                        title_paragraph.font.bold = True

                        # Récupérer le DataFrame de la feuille
                        df_sheet = sheet_info['dataframe']

                        # Nettoyer le DataFrame (supprimer les lignes/colonnes vides)
                        df_sheet = df_sheet.dropna(how='all').dropna(axis=1, how='all')

                        # Filtrer pour ne garder que les colonnes "nom et prenom" et "CIN"
                        df_filtered = pd.DataFrame()

                        if not df_sheet.empty:
                            print(f"📋 Analyse de la feuille '{st_fo}': {len(df_sheet)} lignes")
                            print(f"📊 Colonnes disponibles: {list(df_sheet.columns)}")

                            # Normaliser les noms de colonnes pour la recherche
                            df_sheet.columns = df_sheet.columns.astype(str)

                            # Chercher la colonne "nom et prenom" (variations possibles)
                            nom_col = None
                            for col in df_sheet.columns:
                                col_lower = str(col).lower().strip()
                                # Recherche plus précise pour "nom et prenom"
                                if ('nom et prenom' in col_lower or
                                    'nom et prénom' in col_lower or
                                    'nom_et_prenom' in col_lower or
                                    col_lower == 'nom' or
                                    'nom complet' in col_lower):
                                    nom_col = col
                                    print(f"✅ Colonne nom trouvée: '{col}'")
                                    break

                            # Chercher la colonne "CIN" (variations possibles)
                            cin_col = None
                            for col in df_sheet.columns:
                                col_lower = str(col).lower().strip()
                                # Recherche plus précise pour "CIN"
                                if (col_lower == 'cin' or
                                    'c.i.n' in col_lower or
                                    'carte identite' in col_lower or
                                    'carte d\'identite' in col_lower or
                                    'numero cin' in col_lower):
                                    cin_col = col
                                    print(f"✅ Colonne CIN trouvée: '{col}'")
                                    break

                            # Créer le DataFrame filtré avec les colonnes trouvées
                            if nom_col is not None and cin_col is not None:
                                # Filtrer les lignes vides
                                df_temp = df_sheet[[nom_col, cin_col]].copy()
                                df_temp = df_temp.dropna(how='all')  # Supprimer les lignes complètement vides

                                if not df_temp.empty:
                                    df_filtered = df_temp.copy()
                                    df_filtered.columns = ['Nom et Prénom', 'CIN']  # Renommer pour uniformiser
                                    print(f"✅ Tableau créé pour {st_fo}: {len(df_filtered)} personnes avec nom et CIN")
                                else:
                                    print(f"⚠️ Feuille {st_fo} vide après filtrage")
                            elif nom_col is not None:
                                df_temp = df_sheet[[nom_col]].copy().dropna(how='all')
                                if not df_temp.empty:
                                    df_filtered = df_temp.copy()
                                    df_filtered.columns = ['Nom et Prénom']
                                    print(f"⚠️ Seule la colonne nom disponible pour {st_fo}: {len(df_filtered)} personnes")
                            elif cin_col is not None:
                                df_temp = df_sheet[[cin_col]].copy().dropna(how='all')
                                if not df_temp.empty:
                                    df_filtered = df_temp.copy()
                                    df_filtered.columns = ['CIN']
                                    print(f"⚠️ Seule la colonne CIN disponible pour {st_fo}: {len(df_filtered)} personnes")
                            else:
                                print(f"❌ Aucune colonne 'nom et prenom' ou 'CIN' trouvée pour {st_fo}")
                                print(f"   Colonnes disponibles: {list(df_sheet.columns)}")
                                print(f"   → Slide ignoré pour ce ST FO")
                                continue  # Passer au ST FO suivant

                        if not df_filtered.empty:
                            # Créer un tableau avec les données filtrées (nom et prenom + CIN)
                            rows, cols = len(df_filtered) + 1, len(df_filtered.columns)  # +1 pour l'en-tête

                            # Ajouter le tableau
                            table_shape = prestataire_slide.shapes.add_table(
                                rows, cols, Cm(1), Cm(4), Cm(23), Cm(12)
                            )
                            table = table_shape.table

                            # Définir les largeurs de colonnes optimisées pour 2 colonnes
                            if cols == 2:  # Nom et CIN
                                table.columns[0].width = Cm(15)  # Nom et Prénom (plus large)
                                table.columns[1].width = Cm(8)   # CIN (plus étroit)
                            else:  # Une seule colonne
                                table.columns[0].width = Cm(23)

                            # Remplir l'en-tête
                            for col_idx, column_name in enumerate(df_filtered.columns):
                                cell = table.cell(0, col_idx)
                                cell.text = str(column_name)

                                # Style de l'en-tête
                                paragraph = cell.text_frame.paragraphs[0]
                                paragraph.font.name = 'Arial Narrow'
                                paragraph.font.size = Pt(12)
                                paragraph.font.bold = True

                                # Couleur de fond de l'en-tête
                                fill = cell.fill
                                fill.solid()
                                fill.fore_color.rgb = RGBColor(220, 220, 220)

                            # Remplir les données
                            for row_idx, (_, row_data) in enumerate(df_filtered.iterrows(), 1):
                                if row_idx >= rows:  # Sécurité
                                    break

                                for col_idx, value in enumerate(row_data):
                                    if col_idx >= cols:  # Sécurité
                                        break

                                    cell = table.cell(row_idx, col_idx)
                                    cell.text = str(value) if pd.notna(value) else ""

                                    # Style des données
                                    paragraph = cell.text_frame.paragraphs[0]
                                    paragraph.font.name = 'Arial Narrow'
                                    paragraph.font.size = Pt(10)

                            print(f"✅ 3ème slide ajouté pour {st_fo} avec tableau {rows-1}x{cols} (Nom et Prénom + CIN)")
                        else:
                            # Si la feuille est vide, ajouter un message
                            text_shape = prestataire_slide.shapes.add_textbox(
                                Cm(1), Cm(4), Cm(23), Cm(5)
                            )
                            text_frame = text_shape.text_frame
                            text_frame.text = f"Aucune donnée disponible pour {st_fo}"
                            paragraph = text_frame.paragraphs[0]
                            paragraph.font.name = 'Arial Narrow'
                            paragraph.font.size = Pt(16)

                    print(f"✅ {len(matching_sheets)} slides de prestataires ajoutés")
                else:
                    print(f"⚠️ Aucune feuille correspondante trouvée pour les ST FO: {st_fo_list}")
                    print(f"   ST FO dans le PowerPoint: {st_fo_list}")
                    print(f"   Feuilles disponibles: {list(all_sheets.keys()) if all_sheets else 'Aucune'}")
            else:
                print("⚠️ Fichier 'Liste prestataire' non trouvé dans le répertoire courant")
                print("   Vérifiez que le fichier 'Liste prestataire.xlsx' est présent")
                print("   3ème slide non ajouté")
        except Exception as e:
            print(f"❌ Erreur lors de la création du 3ème slide: {e}")
            import traceback
            print(f"   Détails: {traceback.format_exc()}")

        # Sauvegarder ce fichier PowerPoint avec nom basé sur ST FO
        base_name = output_path.replace('.pptx', '')
        # Nettoyer le nom ST FO pour le nom de fichier (enlever caractères spéciaux)
        safe_st_fo_name = chunk_st_fo_name.replace('/', '_').replace('\\', '_').replace(':', '_')
        chunk_output_path = f"{base_name} {safe_st_fo_name}.pptx"

        prs.save(chunk_output_path)
        created_files.append(os.path.basename(chunk_output_path))

    # Retourner la liste des fichiers créés
    return created_files

def download_file(request, filename):
    """Permet de télécharger un fichier PowerPoint généré"""
    file_path = os.path.join(settings.MEDIA_ROOT, 'generated_ppts', filename)

    if not os.path.exists(file_path):
        raise Http404("Fichier non trouvé")

    with open(file_path, 'rb') as f:
        response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.presentationml.presentation')
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response
