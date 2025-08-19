/**************************** CONFIGURATION GLOBALE ****************************/
// Objet de configuration contenant les noms des feuilles, indices des colonnes,
// statuts valides, icônes associées aux statuts, paramètres de trigger, etc.
const CONFIG = {
  FEUILLES: {
    TACHES: 'Tâches sample',       // Nom de la feuille principale des tâches
    HISTORIQUE: 'Historique',      // Nom de la feuille d'historique des tâches
  },
  COLONNES: {
    PROJET_ID: 1,       // Colonne pour l'ID du projet
    PROJET: 2,          // Colonne pour le nom du projet
    ASSIGNE: 3,         // Colonne pour la personne assignée
    EMAIL: 4,           // Colonne pour l'email de la personne assignée
    DATE_PROJET: 5,     // Colonne pour la date d'échéance du projet
    STATUT: 6,          // Colonne pour le statut de la tâche
    TACHE: 7,           // Colonne pour le nom de la tâche
    TEMPS_ECHEANCE: 8,  // Colonne pour l'heure limite de la tâche
  },
  STATUTS_VALIDES: ['À faire', 'En cours', 'Terminé'],  // Statuts valides
  STATUTS_ICONS: {   // Icônes associées à différents statuts ou alertes
    ATTENTE: '~',
    TERMINE: '✅🔕',
    ECHEANCE_PASSEE: '⌛❌',
    A_RAPPELER: '☑️ à rappeler',
    TEMPS_DEPASSE: ' ⏰ Temps dépassé'
  },
  TRIGGER_HORAIRE: 9,  // Heure du trigger journalier (9h)
  MAX_EMAILS: 50,      // Nombre maximum d'emails envoyés par exécution
  HEADERS_TACHES: [    // Entêtes utilisées dans la feuille tâches
    "ProjetID", "Projet", "Assigné à", "Email", "Date d’échéance (Projet)", 
    "Statut", "Tâche", "Temps d’échéance (Tâche)"
  ],
  HEADERS_HTMLTBL: [   // Entêtes pour le tableau HTML affiché dans le dialogue
    "Projet ID", "Projet", "Assigné à", "Email", "Date d’échéance (Projet)", "Statut", "Ligne",
    "Rappel", "Tâche", "Temps d’échéance (Tâche)"
  ],
  HEADERS_HISTORIQUE: [  // Entêtes pour la feuille historique
    "Projet ID", "Assigné à", "Email", "Projet", "Date d’échéance (Projet)", "Date et Heure de Création (Projet)", "Tâche"   
  ],
  LARGEURS_TACHES: [90, 200, 100, 170, 170, 60, 200, 170],    // Largeurs colonnes feuille tâches
  LARGEURS_HTMLTBL: [90, 200, 100, 170, 170, 60, 50, 60, 200, 170], // Largeurs colonnes tableau HTML
  LARGEURS_HISTORIQUE: [90, 100, 170, 200, 170, 200, 200],    // Largeurs colonnes feuille historique
  UI_MENU_LABELS: {    // Labels du menu UI personnalisé
    MENU: "📋 Menu",
    SYNC_RAPPELS: "⏳ Synchroniser + Rappels",
    ACTIVER_RAPPEL: "📅 Activer rappel automatique",
    MARQUER_TERMINE: "✅ Marquer comme terminé",
    MARQUER_EN_COURS: "🕘 Marquer comme en cours",
    MARQUER_A_Faire: "📝 Marquer comme À faire",
    RESET_TACHES: "🧹 Réinitialiser les tâches",
    RESET_HISTORIQUE: "↺  Réinitialiser Historique"
  }
};

/*************** PROPRIÉTÉS (PropertiesService) ***************/
// Accéder aux propriétés script persistantes
function getProperties() {
  return PropertiesService.getScriptProperties();
}
// Enregistrer une propriété clé-valeur
function setProperty(key, value) {
  getProperties().setProperty(key, value);
}
// Récupérer une propriété par clé
function getProperty(key) {
  return getProperties().getProperty(key);
}

/*************** MENU DÉMARRAGE ***************/
// Fonction appelée à l'ouverture du fichier Google Sheets
// Crée un menu personnalisé avec différentes actions liées aux tâches
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(CONFIG.UI_MENU_LABELS.MENU)
    .addItem(CONFIG.UI_MENU_LABELS.SYNC_RAPPELS, "syncEtRappels")
    .addItem(CONFIG.UI_MENU_LABELS.ACTIVER_RAPPEL, "installerTrigger")
    .addItem(CONFIG.UI_MENU_LABELS.MARQUER_TERMINE, "marquerCommeTermine")
    .addItem(CONFIG.UI_MENU_LABELS.MARQUER_EN_COURS, "marquerCommeEnCours")
    .addItem(CONFIG.UI_MENU_LABELS.MARQUER_A_Faire, "marquerCommeAFaire")
    .addItem(CONFIG.UI_MENU_LABELS.RESET_TACHES, "resetTaches")
    .addItem(CONFIG.UI_MENU_LABELS.RESET_HISTORIQUE, "resetHistorique")
    .addToUi();

  creationEntetesTachesSample(); // Crée les entêtes dans la feuille tâches si nécessaire
  installerTrigger();             // Installe le déclencheur horaire quotidien
  syncEtRappels();               // Synchronise les tâches et envoie les rappels
}

/*************** UTILITAIRES ***************/
// Aligne à droite les colonnes spécifiées dans une feuille donnée, à partir de la 2ème ligne (hors entête)
function alignerColonnesADroiteParFeuille(nomFeuille, colonnes) {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomFeuille);
  if (!feuille) return;
  const lastRow = feuille.getLastRow();
  if (lastRow < 2) return; // Pas de données à aligner

  colonnes.forEach(col => {
    // Aligne horizontalement à droite sur les cellules données
    feuille.getRange(2, col, lastRow - 1).setHorizontalAlignment("right");
  });
}

/*************** RESET HISTORIQUE ***************/
// Vide tout le contenu de la feuille Historique à partir de la 2ème ligne
function resetHistorique() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.FEUILLES.HISTORIQUE);
  if (!feuille) {
    SpreadsheetApp.getUi().alert("Feuille 'Historique' introuvable.");
    return;
  }
  const lastRow = feuille.getLastRow();
  if (lastRow > 1) {
    feuille.getRange(2, 1, lastRow - 1, feuille.getLastColumn()).clearContent();
  }
  SpreadsheetApp.getUi().alert("La feuille 'Historique' a été réinitialisée.");
}

/*************** MARQUAGE DES STATUTS ***************/
// Ces fonctions modifient le statut des lignes sélectionnées dans la feuille active
function marquerCommeTermine() { mettreAJourStatut("Terminé"); }
function marquerCommeEnCours() { mettreAJourStatut("En cours"); }
function marquerCommeAFaire() { mettreAJourStatut("À faire"); }

// Met à jour le statut des lignes sélectionnées avec le statut donné en paramètre
function mettreAJourStatut(nouveauStatut) {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = feuille.getActiveRange();
  if (!range) return;
  const startRow = range.getRow();
  const numRows = range.getNumRows();

  for (let i = 0; i < numRows; i++) {
    feuille.getRange(startRow + i, CONFIG.COLONNES.STATUT).setValue(nouveauStatut);
  }
}

/*************** SYNCHRONISATION + RAPPELS ***************/
// Synchronise les données, prépare un tableau HTML et envoie des emails de rappel
function syncEtRappels() {
  try {
    // Aligne certaines colonnes à droite dans la feuille tâches
    alignerColonnesADroiteParFeuille(CONFIG.FEUILLES.TACHES, [1, 2, 3, 4, 5, 6, 7]);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const src = ss.getSheetByName(CONFIG.FEUILLES.TACHES);
    if (!src) return;

    // Date du jour à minuit (pour comparer uniquement dates sans heures)
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const srcData = src.getDataRange().getValues(); // Récupère toutes les données
    const emails = [];  // Liste des emails à envoyer
    const rows = [];    // Données formatées pour affichage HTML

    for (let i = 1; i < srcData.length; i++) {
      const row = srcData[i];
      const projetIDCell = row[CONFIG.COLONNES.PROJET_ID - 1];
      const projet = row[CONFIG.COLONNES.PROJET - 1];
      const assigne = row[CONFIG.COLONNES.ASSIGNE - 1];
      const email = row[CONFIG.COLONNES.EMAIL - 1];
      const dateProjet = row[CONFIG.COLONNES.DATE_PROJET - 1];
      const statut = row[CONFIG.COLONNES.STATUT - 1];
      const tache = row[CONFIG.COLONNES.TACHE - 1];
      const tempsEcheance = row[CONFIG.COLONNES.TEMPS_ECHEANCE - 1];

      // Si ID projet absent, génère un ID temporaire avec un préfixe "P-"
      const projetID = projetIDCell || "P-" + i.toString().padStart(4, "0");

      // Vérifie que les champs essentiels sont bien remplis et valides
      if (!projet || !assigne || !email || !dateProjet || !statut) continue;
      if (!/@/.test(email.trim())) continue;

      const parsedDate = new Date(dateProjet);
      if (isNaN(parsedDate.getTime())) continue;
      if (!CONFIG.STATUTS_VALIDES.includes(statut)) continue;

      // Calcul différence en jours entre dateProjet et aujourd'hui
      const diff = Math.floor((parsedDate - today) / 86400000);

      // Initialisation de l'icône de rappel et autres variables
      let rappel = CONFIG.STATUTS_ICONS.ATTENTE;
      let tempsDepasse = false;
      let heureFinale = '';

      // Formatage de l'heure limite si spécifiée
      if (tempsEcheance instanceof Date) {
        const maintenant = new Date();
        const heureTotale = new Date(maintenant.getTime());
        heureTotale.setHours(tempsEcheance.getHours(), tempsEcheance.getMinutes());
        heureFinale = Utilities.formatDate(heureTotale, Session.getScriptTimeZone(), "HH:mm");
      }

      if (statut === 'Terminé') {
        rappel = CONFIG.STATUTS_ICONS.TERMINE;
      } else {
        if (diff < 0) {
          // Date dépassée
          rappel = CONFIG.STATUTS_ICONS.ECHEANCE_PASSEE;
        } else if (diff <= 2) {
          // Rappel à envoyer pour échéance proche
          rappel = CONFIG.STATUTS_ICONS.A_RAPPELER;
          emails.push({ email: email.trim(), assigne, tache: projet, date: dateProjet, tempsDepasse: false });
        }
        // Si échéance temps dépassé dans la journée même
        if (tempsEcheance instanceof Date && diff === 0) {
          const maintenant = new Date();
          const heureTache = new Date();
          heureTache.setHours(tempsEcheance.getHours(), tempsEcheance.getMinutes(), 0, 0);
          if (maintenant > heureTache) {
            rappel += CONFIG.STATUTS_ICONS.TEMPS_DEPASSE;
            tempsDepasse = true;
            emails.push({ email: email.trim(), assigne, tache: projet, date: dateProjet, tempsDepasse: true });
          }
        }
      }

      // Prépare les données formatées pour affichage HTML
      rows.push([projetID, projet, assigne, email, dateProjet, statut, i + 1, rappel, tache, heureFinale]);
    }

    // Envoie les emails de rappel, jusqu'au maximum configuré
    emails.slice(0, CONFIG.MAX_EMAILS).forEach(e => {
      try {
        let message = `Bonjour ${e.assigne},\nVotre projet “${e.tache}” est dû le ${new Date(e.date).toLocaleDateString()}.`;
        if (e.tempsDepasse) {
          message += `\n⚠️ Attention : le temps d’échéance de cette tâche est déjà dépassé.`;
        }
        MailApp.sendEmail(e.email, `📌 Rappel - ${e.tache}`, message);
      } catch (err) {
        logErreur(`Erreur lors de l'envoi à ${e.email}`, err);
      }
    });

    afficherTableauHTML(CONFIG.HEADERS_HTMLTBL, rows); // Affiche un tableau HTML dans une fenêtre modale
    enregistrerProjetsEtTaches(); // Synchronise les données dans la feuille historique

  } catch (e) {
    logErreur("Erreur dans syncEtRappels()", e);
  }
}

/*************** INSTALLER TRIGGER ***************/
// Supprime les triggers existants liés à syncEtRappels puis crée un trigger horaire journalier
function installerTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  // Supprime les triggers syncEtRappels existants pour éviter doublons
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'syncEtRappels') ScriptApp.deleteTrigger(t);
  });

  // Crée un nouveau trigger journalier à l'heure définie dans la config
  ScriptApp.newTrigger('syncEtRappels')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.TRIGGER_HORAIRE)
    .create();
}

/*************** RÉINITIALISATION TÂCHES ***************/
// Efface le contenu des tâches à partir de la 2e ligne, colonnes 1 à TEMPS_ECHEANCE
function resetTaches() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.FEUILLES.TACHES);
  if (sheet) sheet.getRange(2, 1, sheet.getLastRow() - 1, CONFIG.COLONNES.TEMPS_ECHEANCE).clearContent();
}

/*************** LOGGING D’ERREURS ***************/
// Affiche dans le log une erreur avec un message personnalisé
function logErreur(msg, e) {
  const message = e?.message || String(e) || 'Erreur inconnue';
  Logger.log(`[ERREUR] ${msg} : ${message}`);
}

/*************** INITIALISATION FEUILLE TÂCHES ***************/
// Crée les entêtes et configure la mise en forme de la feuille "Tâches sample"
function creationEntetesTachesSample() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.FEUILLES.TACHES);
  if (!feuille) {
    SpreadsheetApp.getUi().alert(`Feuille '${CONFIG.FEUILLES.TACHES}' introuvable.`);
    return;
  }

  // Définit les entêtes dans la 1ère ligne
  feuille.getRange(1, 1, 1, CONFIG.HEADERS_TACHES.length).setValues([CONFIG.HEADERS_TACHES]);

  // Définit les largeurs de colonnes
  CONFIG.LARGEURS_TACHES.forEach((width, idx) => {
    feuille.setColumnWidth(idx + 1, width);
  });

  const totalRows = feuille.getMaxRows();

  // Active le retour à la ligne dans toutes les cellules de la table
  feuille.getRange(1, 1, totalRows, CONFIG.HEADERS_TACHES.length).setWrap(true);

  // Mise en forme des entêtes : police, alignement, gras, couleur de fond
  feuille.getRange(1, 1, 1, CONFIG.HEADERS_TACHES.length)
    .setFontFamily("Georgia")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontWeight("bold")
    .setBackground("#d6eaf8");
}

/*************** AFFICHAGE TABLEAU HTML ***************/
// Affiche un tableau HTML interactif dans une fenêtre modale du tableur avec recherche et tri
function afficherTableauHTML(headers, rows) {
  if (!headers || !Array.isArray(headers)) {
    SpreadsheetApp.getUi().alert("Erreur : les en-têtes sont manquants ou invalides.");
    return;
  }
  if (!rows || !Array.isArray(rows)) {
    SpreadsheetApp.getUi().alert("Erreur : les lignes sont manquantes ou invalides.");
    return;
  }

  // Formate la date dans la colonne 4 (index 4 dans rows) au format français
  const timeZone = Session.getScriptTimeZone();
  rows = rows.map(row => {
    const newRow = [...row];
    const dateProjet = row[4];
    if (dateProjet instanceof Date) {
      newRow[4] = Utilities.formatDate(dateProjet, timeZone, "dd/MM/yyyy");
    }
    return newRow;
  });

  // Génération du code HTML complet pour le tableau
  let html = `
    <html>
    <head>
      <style>
        body { font-family: Arial; font-size: 13px; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: center; vertical-align: center; }
        th { background-color: #f0b27a; color: black; cursor: pointer; text-align: center; }
        tr:hover { background-color: #f9f9f9; }
        #searchInput {
          width: 100%;
          padding: 8px;
          border: 1px solid #ccc;
          margin-bottom: 10px;
          font-size: 14px;
        }
      </style>
    </head>
    <body>
      <h2>📋 Tâches enregistrées (HTML)</h2>
      <input type="text" id="searchInput" placeholder="🔍 Rechercher dans le tableau...">

      <table id="tachesTable">
        <thead>
          <tr>${headers.map(h => `<th onclick="sortTable(this)">${h}</th>`).join('')}</tr>
        </thead>
        <tbody>
          ${rows.map(row =>
            `<tr>${row.map(cell => `<td>${cell !== undefined ? cell : ''}</td>`).join('')}</tr>`
          ).join('')}
        </tbody>
      </table>

      <script>
        // Recherche en temps réel dans le tableau
        document.getElementById('searchInput').addEventListener('keyup', function () {
          const filter = this.value.toLowerCase();
          const rows = document.querySelectorAll('#tachesTable tbody tr');
          rows.forEach(row => {
            const text = row.innerText.toLowerCase();
            row.style.display = text.includes(filter) ? '' : 'none';
          });
        });

        // Fonction de tri par colonne (toggle asc/desc)
        function sortTable(th) {
          const table = th.closest('table');
          const tbody = table.querySelector('tbody');
          const index = Array.from(th.parentNode.children).indexOf(th);
          const rows = Array.from(tbody.querySelectorAll('tr'));
          const asc = th.asc = !th.asc;

          rows.sort((a, b) => {
            const cellA = a.children[index].innerText;
            const cellB = b.children[index].innerText;
            return asc
              ? cellA.localeCompare(cellB, undefined, { numeric: true })
              : cellB.localeCompare(cellA, undefined, { numeric: true });
          });

          rows.forEach(row => tbody.appendChild(row));
        }
      </script>
    </body>
    </html>
  `;

  // Affiche la fenêtre modale avec le tableau
  const page = HtmlService.createHtmlOutput(html)
    .setWidth(1200)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(page, 'Tâches générées');
}

/*************** ENREGISTREMENT DES PROJETS ET TÂCHES ***************/
// Synchronise la feuille historique avec les données actuelles de la feuille tâches
function enregistrerProjetsEtTaches() {
  const feuilleSource = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.FEUILLES.TACHES);
  if (!feuilleSource) return;

  const donneesSource = feuilleSource.getDataRange().getValues();
  if (donneesSource.length < 2) return;

  const feuilleHistorique = verifierOuCreerFeuilleHistorique();
  const donneesHistorique = feuilleHistorique.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  const horodatageActuel = Utilities.formatDate(new Date(), timeZone, "dd-MM-yyyy HH:mm");

  // --- Construire dictionnaire des projets dans Tâches ---
  const projetsSource = {};
  for (let i = 1; i < donneesSource.length; i++) {
    const ligne = donneesSource[i];
    const [projetID, projet, assigneA, email, dateProjet, statut, tache] = ligne;
    if (!projetID || !projet) continue;

    const dateProjetFormatee = dateProjet instanceof Date
      ? Utilities.formatDate(dateProjet, timeZone, "yyyy-MM-dd")
      : dateProjet;

    const cleComposite = `${projetID}||${projet}`;
    projetsSource[cleComposite] = { valeurs: [projetID, assigneA, email, projet, dateProjetFormatee, horodatageActuel, tache], statut };
  }

  // --- Construire dictionnaire des projets dans Historique ---
  const projetsHistorique = {};
  for (let i = 1; i < donneesHistorique.length; i++) {
    const ligne = donneesHistorique[i];
    const [projetIDHist, assigneHist, emailHist, projetHist, dateProjetHist, dateCreationHist, tacheHist] = ligne;
    if (!projetIDHist || !projetHist) continue;

    const cleComposite = `${projetIDHist}||${projetHist}`;
    projetsHistorique[cleComposite] = { index: i + 1, dateCreation: dateCreationHist };
  }

  // --- Supprimer projets disparus ---
  const lignesASupprimer = Object.entries(projetsHistorique)
    .filter(([cle]) => !projetsSource.hasOwnProperty(cle))
    .map(([_, info]) => info.index)
    .sort((a, b) => b - a);
  lignesASupprimer.forEach(index => feuilleHistorique.deleteRow(index));

  // --- Mettre à jour ou ajouter ---
  const aFaireLignes = [];
  const enCoursLignes = [];

  Object.entries(projetsSource).forEach(([cle, obj], idx) => {
    const { valeurs, statut } = obj;

    if (projetsHistorique.hasOwnProperty(cle)) {
      const ligneIndex = projetsHistorique[cle].index;
      let ancienneDate = projetsHistorique[cle].dateCreation;

      // Si ancienneDate est "~~~" ou vide, utiliser horodatage actuel
      if (!ancienneDate || ancienneDate === "~~~") {
        ancienneDate = horodatageActuel;
      }

      if (statut === "À faire") {
        valeurs[5] = "~~~";       // Affichage
        aFaireLignes.push(ligneIndex);
      } else if (statut === "En cours") {
        valeurs[5] = ancienneDate; // Restaurer vraie date
        enCoursLignes.push(ligneIndex);
      } else { // Terminé
        valeurs[5] = ancienneDate;
      }

      feuilleHistorique.getRange(ligneIndex, 1, 1, valeurs.length).setValues([valeurs]);

    } else {
      // Nouveau projet
      feuilleHistorique.appendRow(valeurs);
      const lastRow = feuilleHistorique.getLastRow();

      if (statut === "À faire") {
        feuilleHistorique.getRange(lastRow, 6).setValue("~~~");
        aFaireLignes.push(lastRow);
      } else if (statut === "En cours") {
        feuilleHistorique.getRange(lastRow, 6).setValue(horodatageActuel);
        enCoursLignes.push(lastRow);
      }
    }
  });

  // --- Appliquer alignement en bloc ---
  aFaireLignes.forEach(row => feuilleHistorique.getRange(row, 6).setHorizontalAlignment("right"));
  enCoursLignes.forEach(row => feuilleHistorique.getRange(row, 6).setHorizontalAlignment("center"));
}

/*************** VERIFICATION / CREATION FEUILLE HISTORIQUE ***************/
// Vérifie si la feuille historique existe, sinon la crée et configure la mise en forme
function verifierOuCreerFeuilleHistorique() {
  const feuilleNom = CONFIG.FEUILLES.HISTORIQUE;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let feuille = ss.getSheetByName(feuilleNom);

  if (!feuille) {
    feuille = ss.insertSheet(feuilleNom);
  }

  const headers = CONFIG.HEADERS_HISTORIQUE;

  // Définit les entêtes dans la 1ère ligne
  feuille.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Définit les largeurs de colonnes
  const largeurs = CONFIG.LARGEURS_HISTORIQUE;
  for (let i = 0; i < largeurs.length; i++) {
    feuille.setColumnWidth(i + 1, largeurs[i]);
  }

  const totalRows = feuille.getMaxRows();

  // Active le retour à la ligne dans toutes les cellules de la table
  feuille.getRange(1, 1, totalRows, headers.length).setWrap(true);

  // Mise en forme des entêtes : police, alignement, gras, couleur de fond
  feuille.getRange(1, 1, 1, headers.length)
    .setFontFamily("Georgia")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontWeight("bold")
    .setBackground("#F76363");

  // Aligne à droite certaines colonnes
  alignerColonnesADroiteParFeuille(feuilleNom, [1, 2, 3, 4, 5]);

  return feuille;
}