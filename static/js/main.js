let selectedFile = null;

// Éléments DOM
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileSelected = document.getElementById('fileSelected');
const fileName = document.getElementById('fileName');
const sortSection = document.getElementById('sortSection');
const statsSection = document.getElementById('statsSection');
const loading = document.getElementById('loading');
const results = document.getElementById('results');
const error = document.getElementById('error');
const filesList = document.getElementById('filesList');
const errorMessage = document.getElementById('errorMessage');

// Gestion du drag & drop
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFileSelection(files[0]);
    }
});

// Gestion de la sélection de fichier
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFileSelection(e.target.files[0]);
    }
});

// Clic sur la zone de drop
dropZone.addEventListener('click', () => {
    fileInput.click();
});

function handleFileSelection(file) {
    // Vérifier le type de fichier
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        showError('Veuillez sélectionner un fichier Excel (.xlsx ou .xls)');
        return;
    }
    
    selectedFile = file;
    fileName.textContent = file.name;
    
    // Masquer les autres sections
    hideAllSections();
    
    // Afficher le fichier sélectionné et les options de tri
    fileSelected.style.display = 'block';
    sortSection.style.display = 'block';
}

function removeFile() {
    selectedFile = null;
    fileInput.value = '';
    hideAllSections();
    fileSelected.style.display = 'none';
}

function processFile() {
    if (!selectedFile) {
        showError('Veuillez sélectionner un fichier');
        return;
    }

    const sortColumn = document.querySelector('input[name="sortColumn"]:checked').value;

    // Récupérer les informations de visite
    const dateDebut = document.getElementById('dateDebut').value;
    const dateFin = document.getElementById('dateFin').value;
    const objetVisite = document.getElementById('objetVisite').value;

    // Validation des champs avec feedback visuel
    if (!validateVisitFields(dateDebut, dateFin, objetVisite)) {
        return;
    }

    // Afficher le loading
    hideAllSections();
    loading.style.display = 'block';

    // Préparer les données pour l'upload
    const formData = new FormData();
    formData.append('file', selectedFile);
    formData.append('sort_column', sortColumn);
    formData.append('date_debut', dateDebut);
    formData.append('date_fin', dateFin);
    formData.append('objet_visite', objetVisite);
    
    // Envoyer la requête
    fetch('/upload/', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        loading.style.display = 'none';

        if (data.success) {
            showStats(data.stats);
            showResults(data.files, data.stats.sort_column, data.files_details);
        } else {
            showError(data.error || 'Une erreur est survenue');
        }
    })
    .catch(err => {
        loading.style.display = 'none';
        showError('Erreur de connexion: ' + err.message);
    });
}

function showStats(stats) {
    statsSection.style.display = 'block';

    // Afficher les statistiques principales
    document.getElementById('totalRows').textContent = stats.total_rows;
    document.getElementById('totalStFo').textContent = stats.total_st_fo;
    document.getElementById('primaryCount').textContent = stats.primary_count;
    document.getElementById('primaryLabel').textContent = `Nombre de ${stats.primary_label}${stats.primary_count > 1 ? 's' : ''}`;
    document.getElementById('fileCount').textContent = stats.file_count;
    document.getElementById('fileSize').textContent = stats.file_size;

    // Gérer l'affichage des détails selon show_details
    const detailsSection = document.querySelector('.details-section');

    if (!stats.show_details) {
        // Pour code site : masquer toute la section des détails
        detailsSection.style.display = 'none';
        return;
    }

    // Afficher la section des détails pour les autres colonnes
    detailsSection.style.display = 'grid';

    // Afficher la répartition principale (adaptée selon la colonne)
    const primaryList = document.getElementById('primaryList');
    const primaryTitle = document.getElementById('primaryTitle');
    const primaryCard = primaryList.closest('.detail-card');

    // Définir l'icône selon le type
    let icon = '🏢';
    if (stats.primary_label === 'ville') icon = '🏙️';
    else if (stats.primary_label === 'ST FO') icon = '📋';
    else if (stats.primary_label === 'DR') icon = '🏢';

    primaryTitle.textContent = `${icon} ${stats.stats_title}`;
    primaryList.innerHTML = '';

    Object.entries(stats.primary_stats)
        .sort((a, b) => {
            // Trier par nombre de lignes décroissant
            const aLines = typeof a[1] === 'object' ? a[1].lines : a[1];
            const bLines = typeof b[1] === 'object' ? b[1].lines : b[1];
            return bLines - aLines;
        })
        .forEach(([name, data]) => {
            const primaryItem = document.createElement('div');
            primaryItem.className = 'primary-item';

            if (typeof data === 'object') {
                // Nouvelles statistiques détaillées
                primaryItem.innerHTML = `
                    <span class="primary-name">${name}</span>
                    <span class="primary-count">
                        ${data.lines} ligne${data.lines > 1 ? 's' : ''} |
                        ${data.st_fo} ST FO |
                        ${data.files} fichier${data.files > 1 ? 's' : ''}
                    </span>
                `;
            } else {
                // Anciennes statistiques simples (pour compatibilité)
                primaryItem.innerHTML = `
                    <span class="primary-name">${name}</span>
                    <span class="primary-count">${data} ligne${data > 1 ? 's' : ''}</span>
                `;
            }

            primaryList.appendChild(primaryItem);
        });

    // Afficher la répartition par colonne de tri (seulement si différente de la principale)
    const sortList = document.getElementById('sortList');
    const sortTitle = document.getElementById('sortTitle');
    const sortCard = sortList.closest('.detail-card');

    if (stats.sort_column === 'DR IAM' || stats.sort_column === 'ville' || stats.sort_column === 'ST FO') {
        // Masquer complètement la section de tri pour éviter la duplication
        sortCard.style.display = 'none';
        sortCard.style.visibility = 'hidden';
        sortCard.style.height = '0';
        sortCard.style.overflow = 'hidden';

    } else {
        // Restaurer l'affichage normal de la section de tri
        sortCard.style.display = 'block';
        sortCard.style.visibility = 'visible';
        sortCard.style.height = 'auto';
        sortCard.style.overflow = 'visible';
        sortTitle.textContent = `📋 Répartition par ${stats.sort_column}`;

        sortList.innerHTML = '';

        Object.entries(stats.sort_stats)
            .sort((a, b) => b[1] - a[1]) // Trier par nombre décroissant
            .forEach(([sortName, count]) => {
                const sortItem = document.createElement('div');
                sortItem.className = 'sort-item';
                sortItem.innerHTML = `
                    <span class="sort-name">${sortName}</span>
                    <span class="sort-count">${count} ligne${count > 1 ? 's' : ''}</span>
                `;
                sortList.appendChild(sortItem);
            });
    }
}

function showResults(files, sortColumn, filesDetails = null) {
    results.style.display = 'block';

    filesList.innerHTML = '';

    if (filesDetails && filesDetails.length > 0) {
        // Utiliser les détails des fichiers avec nombre de lignes
        filesDetails.forEach(fileInfo => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';

            fileItem.innerHTML = `
                <span>📄 ${fileInfo.filename}</span>
                <span class="file-lines">${fileInfo.lines} ligne${fileInfo.lines > 1 ? 's' : ''}</span>
                <a href="/download/${encodeURIComponent(fileInfo.filename)}/" class="btn-download" download>
                    Télécharger
                </a>
            `;

            filesList.appendChild(fileItem);
        });
    } else {
        // Fallback pour compatibilité (ancien format)
        files.forEach(filename => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';

            fileItem.innerHTML = `
                <span>📄 ${filename}</span>
                <a href="/download/${encodeURIComponent(filename)}/" class="btn-download" download>
                    Télécharger
                </a>
            `;

            filesList.appendChild(fileItem);
        });
    }
}

function showError(message) {
    hideAllSections();
    error.style.display = 'block';
    errorMessage.textContent = message;
}

function hideAllSections() {
    loading.style.display = 'none';
    statsSection.style.display = 'none';
    results.style.display = 'none';
    error.style.display = 'none';
}

// Fonction pour recommencer
function resetForm() {
    removeFile();
    hideAllSections();
    resetVisitFields();
}

// Validation des champs de visite avec feedback visuel
function validateVisitFields(dateDebut, dateFin, objetVisite) {
    let isValid = true;

    // Réinitialiser les classes de validation
    document.querySelectorAll('.field-group').forEach(group => {
        group.classList.remove('valid', 'invalid');
    });

    // Validation des dates
    if (dateDebut && dateFin) {
        if (dateDebut > dateFin) {
            showError('La date de début ne peut pas être postérieure à la date de fin');
            // Marquer les champs de date comme invalides
            document.getElementById('dateDebut').parentElement.classList.add('invalid');
            document.getElementById('dateFin').parentElement.classList.add('invalid');
            isValid = false;
        } else {
            // Marquer les champs de date comme valides
            document.getElementById('dateDebut').parentElement.classList.add('valid');
            document.getElementById('dateFin').parentElement.classList.add('valid');
        }
    }

    // Validation de l'objet de visite (optionnel mais feedback visuel)
    if (objetVisite && objetVisite.trim().length > 0) {
        document.getElementById('objetVisite').parentElement.classList.add('valid');
    }

    return isValid;
}

// Réinitialiser les champs de visite
function resetVisitFields() {
    document.getElementById('dateDebut').value = '';
    document.getElementById('dateFin').value = '';
    document.getElementById('objetVisite').value = '';

    document.querySelectorAll('.field-group').forEach(group => {
        group.classList.remove('valid', 'invalid');
    });
}

// Validation en temps réel des champs de visite
function setupVisitFieldsValidation() {
    const dateDebut = document.getElementById('dateDebut');
    const dateFin = document.getElementById('dateFin');
    const objetVisite = document.getElementById('objetVisite');

    // Validation en temps réel des dates
    function validateDates() {
        const debut = dateDebut.value;
        const fin = dateFin.value;

        // Réinitialiser les classes pour les champs de date
        const dateFields = [
            document.getElementById('dateDebut').parentElement,
            document.getElementById('dateFin').parentElement
        ];

        dateFields.forEach(group => {
            group.classList.remove('valid', 'invalid');
        });

        if (debut && fin) {
            if (debut > fin) {
                dateFields.forEach(group => {
                    group.classList.add('invalid');
                });
            } else {
                dateFields.forEach(group => {
                    group.classList.add('valid');
                });
            }
        } else if (debut || fin) {
            dateFields.forEach(group => {
                group.classList.add('valid');
            });
        }
    }

    // Validation en temps réel de l'objet de visite
    function validateObjet() {
        const objetField = document.getElementById('objetVisite').parentElement;
        objetField.classList.remove('valid', 'invalid');

        if (objetVisite.value && objetVisite.value.trim().length > 0) {
            objetField.classList.add('valid');
        }
    }

    // Événements
    dateDebut.addEventListener('change', validateDates);
    dateFin.addEventListener('change', validateDates);
    objetVisite.addEventListener('input', validateObjet);
    objetVisite.addEventListener('blur', validateObjet);
}

// Initialiser la validation au chargement de la page
document.addEventListener('DOMContentLoaded', function() {
    setupVisitFieldsValidation();
});

// Ajouter un bouton pour recommencer dans les résultats
document.addEventListener('DOMContentLoaded', () => {
    // Ajouter un bouton "Nouveau fichier" dans la section résultats
    const resultsSection = document.getElementById('results');
    if (resultsSection) {
        const newFileBtn = document.createElement('button');
        newFileBtn.textContent = 'Traiter un nouveau fichier';
        newFileBtn.className = 'btn-process';
        newFileBtn.style.marginTop = '20px';
        newFileBtn.onclick = resetForm;
        resultsSection.appendChild(newFileBtn);
    }
});
