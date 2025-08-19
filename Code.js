/**************************** CONFIGURATION GLOBALE ****************************/
const CONFIG = {
  FEUILLES: {
    TACHES: 'Tâches sample',
    HISTORIQUE: 'Historique',
  },
  COLONNES: {
    PROJET_ID: 1,
    PROJET: 2,
    ASSIGNE: 3,
    EMAIL: 4,
    DATE_PROJET: 5,
    STATUT: 6,
    TACHE: 7,
    TEMPS_ECHEANCE: 8,
  },
  STATUTS_VALIDES: ['À faire', 'En cours', 'Terminé'],
  STATUTS_ICONS: {
    ATTENTE: '~',
    TERMINE: '✅🔕',
    ECHEANCE_PASSEE: '⌛❌',
    A_RAPPELER: '☑️ à rappeler',
    TEMPS_DEPASSE: '⏰ Temps dépassé'
  },
  TRIGGER_HORAIRE: 9,
  MAX_EMAILS: 50,
  HEADERS_TACHES: ["ProjetID", "Projet", "Assigné à", "Email", "Date d’échéance (Projet)", "Statut", "Tâche", "Temps d’échéance (Tâche)"],
  HEADERS_HTMLTBL: ["Projet ID", "Projet", "Assigné à", "Email", "Date d’échéance (Projet)", "Statut", "Ligne", "Rappel", "Tâche", "Temps d’échéance (Tâche)"],
  HEADERS_HISTORIQUE: ["Projet ID", "Assigné à", "Email", "Projet", "Date d’échéance (Projet)", "Date et Heure de Création (Projet)", "Tâche"],
  LARGEURS_TACHES: [90, 200, 100, 170, 170, 60, 200, 170],
  LARGEURS_HTMLTBL: [90, 200, 100, 170, 170, 60, 50, 60, 200, 170],
  LARGEURS_HISTORIQUE: [90, 100, 170, 200, 170, 200, 200],
  UI_MENU_LABELS: {
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

/**************************** PROPRIÉTÉS ****************************/
function getProperties() { return PropertiesService.getScriptProperties(); }
function setProperty(key, value) { getProperties().setProperty(key, value); }
function getProperty(key) { return getProperties().getProperty(key); }

/**************************** MENU DÉMARRAGE ****************************/
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

  creationEntetesTachesSample();
  installerTrigger();
  syncEtRappels();
}

/**************************** UTILITAIRES ****************************/
function alignerColonnesADroiteParFeuille(nomFeuille, colonnes) {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomFeuille);
  if (!feuille) return;
  const lastRow = feuille.getLastRow();
  if (lastRow < 2) return;
  colonnes.forEach(col => feuille.getRange(2, col, lastRow - 1).setHorizontalAlignment("right"));
}

/**************************** RESET HISTORIQUE ****************************/
function resetHistorique() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.FEUILLES.HISTORIQUE);
  if (!feuille) { SpreadsheetApp.getUi().alert("Feuille 'Historique' introuvable."); return; }
  const lastRow = feuille.getLastRow();
  if (lastRow > 1) feuille.getRange(2, 1, lastRow - 1, feuille.getLastColumn()).clearContent();
  SpreadsheetApp.getUi().alert("La feuille 'Historique' a été réinitialisée.");
}

/**************************** MARQUAGE DES STATUTS ****************************/
function marquerCommeTermine() { mettreAJourStatut("Terminé"); }
function marquerCommeEnCours() { mettreAJourStatut("En cours"); }
function marquerCommeAFaire() { mettreAJourStatut("À faire"); }

function mettreAJourStatut(nouveauStatut) {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = feuille.getActiveRange();
  if (!range) return;
  const startRow = range.getRow();
  const numRows = range.getNumRows();
  for (let i = 0; i < numRows; i++) {
    feuille.getRange(startRow + i, CONFIG.COLONNES.STATUT).setValue(nouveauStatut);
  }

  // Mettre à jour l'historique immédiatement
  enregistrerProjetsEtTaches();
}

/**************************** SYNCHRONISATION + RAPPELS ****************************/
function syncEtRappels() {
  try {
    alignerColonnesADroiteParFeuille(CONFIG.FEUILLES.TACHES, [1,2,3,4,5,6,7]);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const src = ss.getSheetByName(CONFIG.FEUILLES.TACHES);
    if (!src) return;
    const today = new Date(); today.setHours(0,0,0,0);
    const srcData = src.getDataRange().getValues();
    const emails = [];
    const rows = [];

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
      const projetID = projetIDCell || "P-"+i.toString().padStart(4,"0");

      if (!projet || !assigne || !email || !dateProjet || !statut) continue;
      if (!/@/.test(email.trim())) continue;
      const parsedDate = new Date(dateProjet);
      if (isNaN(parsedDate.getTime())) continue;
      if (!CONFIG.STATUTS_VALIDES.includes(statut)) continue;

      const diff = Math.floor((parsedDate - today)/86400000);
      let rappel = CONFIG.STATUTS_ICONS.ATTENTE;
      let tempsDepasse = false;
      let heureFinale = '';

      if (tempsEcheance instanceof Date) {
        const maintenant = new Date();
        const heureTotale = new Date(maintenant.getTime());
        heureTotale.setHours(tempsEcheance.getHours(), tempsEcheance.getMinutes());
        heureFinale = Utilities.formatDate(heureTotale, Session.getScriptTimeZone(), "HH:mm");
      }

      if (statut === 'Terminé') {
        rappel = CONFIG.STATUTS_ICONS.TERMINE;
      } else {
        if (diff < 0) rappel = CONFIG.STATUTS_ICONS.ECHEANCE_PASSEE;
        else if (diff <= 2) { rappel = CONFIG.STATUTS_ICONS.A_RAPPELER; emails.push({ email: email.trim(), assigne, tache: projet, date: dateProjet, tempsDepasse: false }); }
        if (tempsEcheance instanceof Date && diff === 0) {
          const maintenant = new Date();
          const heureTache = new Date(); heureTache.setHours(tempsEcheance.getHours(), tempsEcheance.getMinutes(),0,0);
          if (maintenant > heureTache) { rappel += CONFIG.STATUTS_ICONS.TEMPS_DEPASSE; tempsDepasse = true; emails.push({ email: email.trim(), assigne, tache: projet, date: dateProjet, tempsDepasse:true }); }
        }
      }

      rows.push([projetID, projet, assigne, email, dateProjet, statut, i+1, rappel, tache, heureFinale]);
    }

    emails.slice(0, CONFIG.MAX_EMAILS).forEach(e=>{
      try { 
        let message=`Bonjour ${e.assigne},\nVotre projet “${e.tache}” est dû le ${new Date(e.date).toLocaleDateString()}.`;
        if(e.tempsDepasse) message+=`\n⚠️ Attention : le temps d’échéance de cette tâche est déjà dépassé.`;
        MailApp.sendEmail(e.email, `📌 Rappel - ${e.tache}`, message);
      } catch(err){ logErreur(`Erreur lors de l'envoi à ${e.email}`, err); }
    });

    afficherTableauHTML(CONFIG.HEADERS_HTMLTBL, rows);
    enregistrerProjetsEtTaches();
  } catch(e){ logErreur("Erreur dans syncEtRappels()", e); }
}

/**************************** INSTALLER TRIGGER ****************************/
function installerTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t=>{ if(t.getHandlerFunction()==='syncEtRappels') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncEtRappels').timeBased().everyDays(1).atHour(CONFIG.TRIGGER_HORAIRE).create();
}

/**************************** RESET TÂCHES ****************************/
function resetTaches() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.FEUILLES.TACHES);
  if(sheet) sheet.getRange(2,1,sheet.getLastRow()-1,CONFIG.COLONNES.TEMPS_ECHEANCE).clearContent();
}

/**************************** LOGGING D’ERREURS ****************************/
function logErreur(msg,e){ Logger.log(`[ERREUR] ${msg} : ${e?.message || String(e) || 'Erreur inconnue'}`); }

/**************************** INITIALISATION FEUILLE TÂCHES ****************************/
function creationEntetesTachesSample() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.FEUILLES.TACHES);
  if(!feuille){ SpreadsheetApp.getUi().alert(`Feuille '${CONFIG.FEUILLES.TACHES}' introuvable.`); return; }
  feuille.getRange(1,1,1,CONFIG.HEADERS_TACHES.length).setValues([CONFIG.HEADERS_TACHES]);
  CONFIG.LARGEURS_TACHES.forEach((w,i)=>feuille.setColumnWidth(i+1,w));
  feuille.getRange(1,1,feuille.getMaxRows(),CONFIG.HEADERS_TACHES.length).setWrap(true);
  feuille.getRange(1,1,1,CONFIG.HEADERS_TACHES.length).setFontFamily("Georgia").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold").setBackground("#d6eaf8");
}

/**************************** AFFICHAGE TABLEAU HTML ****************************/
function afficherTableauHTML(headers, rows){
  if(!headers||!Array.isArray(headers)){ SpreadsheetApp.getUi().alert("Erreur : en-têtes invalides."); return; }
  if(!rows||!Array.isArray(rows)){ SpreadsheetApp.getUi().alert("Erreur : lignes invalides."); return; }
  const timeZone = Session.getScriptTimeZone();
  rows = rows.map(row=>{ const newRow=[...row]; const d=row[4]; if(d instanceof Date) newRow[4]=Utilities.formatDate(d,timeZone,"dd/MM/yyyy"); return newRow; });

  let html=`<html><head><style>
  body{font-family:Arial;font-size:13px;}
  table{border-collapse:collapse;width:100%;margin-top:10px;}
  th,td{border:1px solid #ddd;padding:8px;text-align:center;vertical-align:center;}
  th{background-color:#f0b27a;color:black;cursor:pointer;}
  tr:hover{background-color:#f9f9f9;}
  #searchInput{width:100%;padding:8px;border:1px solid #ccc;margin-bottom:10px;font-size:14px;}
  </style></head><body>
  <h2>📋 Tâches enregistrées (HTML)</h2>
  <input type="text" id="searchInput" placeholder="🔍 Rechercher dans le tableau...">
  <table id="tachesTable"><thead><tr>${headers.map(h=>`<th onclick="sortTable(this)">${h}</th>`).join('')}</tr></thead>
  <tbody>${rows.map(r=>`<tr>${r.map(c=>`<td>${c!==undefined?c:''}</td>`).join('')}</tr>`).join('')}</tbody>
  </table><script>
  document.getElementById('searchInput').addEventListener('keyup',function(){const f=this.value.toLowerCase();document.querySelectorAll('#tachesTable tbody tr').forEach(r=>{r.style.display=r.innerText.toLowerCase().includes(f)?'':'none';});});
  function sortTable(th){const table=th.closest('table');const tbody=table.querySelector('tbody');const idx=Array.from(th.parentNode.children).indexOf(th);const rows=Array.from(tbody.querySelectorAll('tr'));const asc=th.asc=!th.asc;rows.sort((a,b)=>asc?a.children[idx].innerText.localeCompare(b.children[idx].innerText,undefined,{numeric:true}):b.children[idx].innerText.localeCompare(a.children[idx].innerText,undefined,{numeric:true}));rows.forEach(r=>tbody.appendChild(r));}
  </script></body></html>`;

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(1200).setHeight(600),'Tâches générées');
}

/**************************** ENREGISTREMENT PROJETS & TÂCHES ****************************/
function enregistrerProjetsEtTaches() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Vérifier la feuille Tâches ---
  const feuilleSource = ss.getSheetByName(CONFIG.FEUILLES.TACHES);
  if (!feuilleSource) {
    SpreadsheetApp.getUi().alert(`Feuille '${CONFIG.FEUILLES.TACHES}' introuvable !`);
    return;
  }

  const donneesSource = feuilleSource.getDataRange().getValues();
  if (donneesSource.length < 2) return;

  // --- Vérifier ou créer la feuille Historique ---
  let feuilleHistorique = ss.getSheetByName(CONFIG.FEUILLES.HISTORIQUE);
  if (!feuilleHistorique) {
    feuilleHistorique = ss.insertSheet(CONFIG.FEUILLES.HISTORIQUE);
    feuilleHistorique.getRange(1,1,1,CONFIG.HEADERS_HISTORIQUE.length)
      .setValues([CONFIG.HEADERS_HISTORIQUE]);
  }

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
    projetsSource[cleComposite] = { projetID, assigne: assigneA, email, projet, dateProjet: dateProjetFormatee, statut, tache };
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
  const lignesASupprimer = [];
  Object.entries(projetsHistorique).forEach(([cle, info]) => {
    if (!projetsSource.hasOwnProperty(cle)) lignesASupprimer.push(info.index);
  });

  // --- Mettre à jour ou ajouter ---
  Object.entries(projetsSource).forEach(([cle, valeurs]) => {
    const { projetID, assigne, email, projet, dateProjet, statut, tache } = valeurs;

    if (projetsHistorique.hasOwnProperty(cle)) {
      const ligneIndex = projetsHistorique[cle].index;
      let dateCreation = projetsHistorique[cle].dateCreation;

      // Si statut = "À faire", afficher "~~~"
      const dateAffiche = statut === "À faire" ? "~~~" : (dateCreation && dateCreation !== "~~~" ? dateCreation : horodatageActuel);

      feuilleHistorique.getRange(ligneIndex, 1, 1, CONFIG.HEADERS_HISTORIQUE.length).setValues([
        [projetID, assigne, email, projet, dateProjet, dateAffiche, tache]
      ]);
    } else {
      // Nouveau projet → ajouter à la fin
      const dateAffiche = statut === "À faire" ? "~~~" : horodatageActuel;
      feuilleHistorique.appendRow([projetID, assigne, email, projet, dateProjet, dateAffiche, tache]);
    }
  });

  // --- Supprimer lignes obsolètes ---
  lignesASupprimer.sort((a, b) => b - a).forEach(index => feuilleHistorique.deleteRow(index));
}

/**************************** VERIFICATION / CREATION FEUILLE HISTORIQUE ****************************/
function verifierOuCreerFeuilleHistorique() {
  const feuilleNom=CONFIG.FEUILLES.HISTORIQUE;
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let feuille=ss.getSheetByName(feuilleNom);
  if(!feuille) feuille=ss.insertSheet(feuilleNom);

  const headers=CONFIG.HEADERS_HISTORIQUE;
  feuille.getRange(1,1,1,headers.length).setValues([headers]);
  const largeurs=CONFIG.LARGEURS_HISTORIQUE;
}