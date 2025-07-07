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
    
    // Afficher le loading
    hideAllSections();
    loading.style.display = 'block';
    
    // Préparer les données pour l'upload
    const formData = new FormData();
    formData.append('file', selectedFile);
    formData.append('sort_column', sortColumn);
    
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
            showResults(data.files, data.stats.sort_column);
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
    document.getElementById('primaryCount').textContent = stats.primary_count;
    document.getElementById('primaryLabel').textContent = `Nombre de ${stats.primary_label}${stats.primary_count > 1 ? 's' : ''}`;
    document.getElementById('groupCount').textContent = stats.group_count;

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
        .sort((a, b) => b[1] - a[1]) // Trier par nombre décroissant
        .forEach(([name, count]) => {
            const primaryItem = document.createElement('div');
            primaryItem.className = 'primary-item';
            primaryItem.innerHTML = `
                <span class="primary-name">${name}</span>
                <span class="primary-count">${count} ligne${count > 1 ? 's' : ''}</span>
            `;
            primaryList.appendChild(primaryItem);
        });

    // Afficher la répartition par colonne de tri (seulement si différente de la principale)
    const sortList = document.getElementById('sortList');
    const sortTitle = document.getElementById('sortTitle');
    const sortCard = sortList.closest('.detail-card');

    if (stats.sort_column === stats.primary_label ||
        (stats.sort_column === 'DR IAm' && stats.primary_label === 'DR')) {
        // Masquer la section de tri si c'est la même que la principale
        sortCard.style.display = 'none';
    } else {
        sortCard.style.display = 'block';
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

function showResults(files, sortColumn) {
    results.style.display = 'block';

    filesList.innerHTML = '';

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
}

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
