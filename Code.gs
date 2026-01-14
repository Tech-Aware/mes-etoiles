// ==================================================
// 🌟 SYSTÈME DE MÉRITE FAMILIAL v4
// Avec section Émotions
// ==================================================

const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const PARIS_TIMEZONE = 'Europe/Paris';
const BASE_TASK_IDS = [
  'rangerChambre',
  'faireLit',
  'rangerJouets',
  'aiderTable',
  'ecouter',
  'gentilSoeur',
  'politesse',
  'pasColere',
  'dentsMatin',
  'dentsSoir',
  'habiller',
  'cartable'
];

const TASK_CATEGORY_KEYS = {
  corvees: 'corvees',
  comportement: 'comportement',
  rituels: 'rituels',
  autres: 'autres'
};


// ==================================================
// CONFIGURATION - COLONNE JOURS (TÂCHES)
// ==================================================
const JOURS_SHEET_NAME = 'Jours';
const JOURS_LISTE = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche'];

function creerFeuilleJoursSemaine() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    let sheet = ss.getSheetByName(JOURS_SHEET_NAME);
    if (!sheet) {
      Logger.log(`[creerFeuilleJoursSemaine] Feuille "${JOURS_SHEET_NAME}" absente, création en cours.`);
      sheet = ss.insertSheet(JOURS_SHEET_NAME);
    }

    const existingValues = sheet.getRange(1, 1, sheet.getMaxRows(), 1).getValues();
    const flattened = existingValues.map(row => String(row[0] || '').trim()).filter(Boolean);
    if (flattened.length === 0) {
      Logger.log('[creerFeuilleJoursSemaine] Initialisation de la liste des jours.');
      sheet.getRange(1, 1, JOURS_LISTE.length, 1).setValues(JOURS_LISTE.map(day => [day]));
    } else {
      Logger.log(`[creerFeuilleJoursSemaine] Liste déjà présente (${flattened.length} valeurs), aucune modification.`);
    }

    return { success: true, message: 'Feuille Jours prête.' };
  } catch (error) {
    Logger.log(`[creerFeuilleJoursSemaine] Erreur lors de la création : ${error}`);
    throw new Error('Impossible de créer ou initialiser la feuille Jours.');
  }
}

function creerColonneJoursTaches() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('Tâches');
    if (!sheet) {
      Logger.log('[creerColonneJoursTaches] Feuille "Tâches" introuvable, création impossible.');
      throw new Error('Feuille "Tâches" introuvable.');
    }

    const headerRow = 1;
    const columnIndex = 9; // Colonne I
    const cell = sheet.getRange(headerRow, columnIndex);
    const currentValue = String(cell.getValue() || '').trim();

    if (currentValue === 'Jours') {
      Logger.log('[creerColonneJoursTaches] Colonne Jours déjà présente en I1.');
    } else if (currentValue && currentValue !== 'Jours') {
      Logger.log(`[creerColonneJoursTaches] Valeur existante en I1 ("${currentValue}"), écrasement avec "Jours".`);
      cell.setValue('Jours');
    } else {
      Logger.log('[creerColonneJoursTaches] Création de la colonne Jours en I1.');
      cell.setValue('Jours');
    }

    return { success: true, message: 'Colonne Jours prête en I1.' };
  } catch (error) {
    Logger.log(`[creerColonneJoursTaches] Erreur lors de la création : ${error}`);
    throw new Error('Impossible de créer la colonne Jours dans la feuille Tâches.');
  }
}

function configurerValidationJoursTaches() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheetTaches = ss.getSheetByName('Tâches');
    if (!sheetTaches) {
      Logger.log('[configurerValidationJoursTaches] Feuille "Tâches" introuvable.');
      throw new Error('Feuille "Tâches" introuvable.');
    }

    const sheetJours = ss.getSheetByName(JOURS_SHEET_NAME);
    if (!sheetJours) {
      Logger.log('[configurerValidationJoursTaches] Feuille "Jours" introuvable.');
      throw new Error('Feuille "Jours" introuvable.');
    }

    const lastRow = Math.max(sheetJours.getLastRow(), 1);
    const rangeJours = sheetJours.getRange(1, 1, lastRow, 1);
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeJours, true)
      .setAllowInvalid(true)
      .build();

    const maxRows = sheetTaches.getMaxRows();
    const columnIndex = 9; // Colonne I
    sheetTaches.getRange(2, columnIndex, maxRows - 1, 1).setDataValidation(validation);

    Logger.log('[configurerValidationJoursTaches] Validation appliquée sur Tâches!I2:I.');
    return { success: true, message: 'Validation Jours appliquée sur la feuille Tâches.' };
  } catch (error) {
    Logger.log(`[configurerValidationJoursTaches] Erreur lors de la configuration : ${error}`);
    throw new Error('Impossible de configurer la validation Jours dans la feuille Tâches.');
  }
}

// ==================================================
// UTILITAIRES DATES (PARIS)
// ==================================================
function getParisDateKey(date) {
  return Utilities.formatDate(date, PARIS_TIMEZONE, 'yyyy-MM-dd');
}

function getParisMidnight(date) {
  const [year, month, day] = getParisDateKey(date).split('-').map(Number);
  return new Date(year, month - 1, day);
}

function parseSheetDate(value, context) {
  if (value instanceof Date) {
    return value;
  }

  if (typeof value === 'number') {
    return new Date(value);
  }

  if (typeof value === 'string') {
    const match = value.trim().match(
      /^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/
    );
    if (match) {
      const [, day, month, year, hour = '00', minute = '00', second = '00'] = match;
      return new Date(
        Number(year),
        Number(month) - 1,
        Number(day),
        Number(hour),
        Number(minute),
        Number(second)
      );
    }
  }

  const fallback = new Date(value);
  if (Number.isNaN(fallback.getTime())) {
    Logger.log(`[parseSheetDate] Date invalide (${context}) : ${value}`);
    return null;
  }
  return fallback;
}

function getParisDateKeyFromValue(value, context) {
  const parsed = parseSheetDate(value, context);
  if (!parsed) {
    return null;
  }
  return getParisDateKey(parsed);
}

function getParisDayIndex(date) {
  const dayIndex = Number(Utilities.formatDate(date, PARIS_TIMEZONE, 'u'));
  if (Number.isNaN(dayIndex)) {
    Logger.log('[getParisDayIndex] Index jour invalide, fallback sur Date.getDay().');
    const fallback = date.getDay();
    return fallback === 0 ? 7 : fallback;
  }
  return dayIndex;
}

function normaliserTexte_(valeur) {
  if (!valeur) {
    return '';
  }
  return String(valeur)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function getMaxPointsParJour_(personne) {
  try {
    const assigned = getTachesAssigneesPourPersonne_(personne);
    const taskCount = assigned.taskIds.length;
    const maxPoints = Math.max(1, taskCount + 1);
    Logger.log(`[getMaxPointsParJour] ${personne} : ${taskCount} tâches, maxPoints=${maxPoints}.`);
    return { maxPoints, taskCount };
  } catch (error) {
    Logger.log(`[getMaxPointsParJour] Erreur pour ${personne} : ${error}`);
    return { maxPoints: 1, taskCount: 0 };
  }
}

function getPointsDisponibles_(personneKey, evalData, claimsData) {
  const evalMeta = getEvaluationsFeuille_();
  const totalJourIndex = getEvaluationColumnIndex_(evalMeta.indexes, 'totalJour', 27, 'TotalJour');
  const personneIndex = getEvaluationColumnIndex_(evalMeta.indexes, 'personne', 3, 'Personne');

  const totalGagnes = evalData
    .filter(row => String(row[personneIndex] || '').trim() === personneKey)
    .reduce((acc, row) => acc + Number(row[totalJourIndex] || 0), 0);

  const totalDepenses = claimsData
    .filter(row => String(row[2] || '').trim() === personneKey)
    .filter(row => {
      const statut = normaliserTexte_(row[5]);
      return statut !== 'annule' && statut !== 'annulé' && statut !== 'refuse' && statut !== 'refusee' && statut !== 'refusée';
    })
    .reduce((acc, row) => acc + Number(row[4] || 0), 0);

  const totalPoints = Math.max(0, totalGagnes - totalDepenses);
  Logger.log(`[getPointsDisponibles] ${personneKey} : gagnés=${totalGagnes}, dépensés=${totalDepenses}, disponibles=${totalPoints}.`);
  return { totalPoints, totalGagnes, totalDepenses };
}

function getEvaluationsFeuille_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Évaluations');
  if (!sheet) {
    Logger.log('[getEvaluationsFeuille] Feuille "Évaluations" introuvable.');
    throw new Error('Feuille "Évaluations" introuvable.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('[getEvaluationsFeuille] Feuille "Évaluations" vide.');
    return { sheet, headers: [], rows: [], indexes: {} };
  }

  const headers = data[0].map(value => String(value || '').trim());
  const indexes = {
    id: headers.indexOf('ID'),
    date: headers.indexOf('Date'),
    heure: headers.indexOf('Heure'),
    personne: headers.indexOf('Personne'),
    emotion1: headers.indexOf('Emotion1'),
    gestionEmotion: headers.indexOf('GestionEmotion'),
    totalJour: headers.indexOf('TotalJour')
  };

  return { sheet, headers, rows: data.slice(1), indexes };
}

function getEvaluationColumnIndex_(indexes, key, fallback, label) {
  const idx = indexes && typeof indexes[key] === 'number' ? indexes[key] : -1;
  if (idx === -1) {
    Logger.log(`[getEvaluationColumnIndex] Colonne "${label}" introuvable, fallback sur index=${fallback}.`);
    return fallback;
  }
  return idx;
}

function buildEvaluationTaskHeader_(task) {
  const taskId = String(task.id || '').trim();
  const taskName = String(task.nom || '').trim();
  if (!taskId || !taskName) {
    return null;
  }
  return `${taskId} - ${taskName}`;
}

function synchroniserColonnesTachesEvaluations_() {
  const evalMeta = getEvaluationsFeuille_();
  const sheet = evalMeta.sheet;
  const headers = evalMeta.headers.slice();
  const headerSet = new Set(headers);

  const taskDefinitions = getTachesDefinitions_();
  if (taskDefinitions.length === 0) {
    Logger.log('[synchroniserColonnesTachesEvaluations] Aucune tâche détectée, aucune colonne ajoutée.');
    return { headers, headerIndex: buildHeaderIndexMap_(headers) };
  }

  let addedCount = 0;
  taskDefinitions.forEach(task => {
    const header = buildEvaluationTaskHeader_(task);
    if (!header) {
      Logger.log(`[synchroniserColonnesTachesEvaluations] En-tête invalide pour la tâche ${JSON.stringify(task)}.`);
      return;
    }
    if (headerSet.has(header)) {
      return;
    }
    headers.push(header);
    headerSet.add(header);
    sheet.getRange(1, headers.length).setValue(header);
    addedCount += 1;
    Logger.log(`[synchroniserColonnesTachesEvaluations] Colonne "${header}" ajoutée en position ${headers.length}.`);
  });

  Logger.log(`[synchroniserColonnesTachesEvaluations] Synchronisation terminée, ${addedCount} colonne(s) ajoutée(s).`);
  return { headers, headerIndex: buildHeaderIndexMap_(headers) };
}

function buildHeaderIndexMap_(headers) {
  return headers.reduce((acc, header, index) => {
    acc[header] = index;
    return acc;
  }, {});
}

function appliquerValeursTachesDynamiques_(sheet, rowIndex, taskIds, getValueFn, definitionsById, headerIndex) {
  taskIds.forEach(taskId => {
    const definition = definitionsById.get(taskId);
    if (!definition) {
      Logger.log(`[appliquerValeursTachesDynamiques] Définition introuvable pour ${taskId}, écriture ignorée.`);
      return;
    }
    const header = buildEvaluationTaskHeader_(definition);
    if (!header) {
      Logger.log(`[appliquerValeursTachesDynamiques] En-tête invalide pour la tâche ${taskId}, écriture ignorée.`);
      return;
    }
    const columnIndex = headerIndex[header];
    if (typeof columnIndex !== 'number') {
      Logger.log(`[appliquerValeursTachesDynamiques] Colonne "${header}" absente, écriture ignorée.`);
      return;
    }
    const value = getValueFn(taskId);
    sheet.getRange(rowIndex, columnIndex + 1).setValue(value);
    Logger.log(`[appliquerValeursTachesDynamiques] Valeur ${value} écrite pour ${header} (ligne ${rowIndex}).`);
  });
}

function mettreAJourCoutsRecompenses() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('Récompenses');
    if (!sheet) {
      Logger.log('[mettreAJourCoutsRecompenses] Feuille "Récompenses" introuvable.');
      throw new Error('Feuille "Récompenses" introuvable.');
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log('[mettreAJourCoutsRecompenses] Aucune récompense à normaliser.');
      return { success: true, message: 'Aucune récompense à mettre à jour.' };
    }

    const headers = data[0].map(value => String(value || '').trim());
    const costIndex = headers.indexOf('Coût');
    if (costIndex === -1) {
      Logger.log('[mettreAJourCoutsRecompenses] Colonne "Coût" introuvable.');
      throw new Error('Colonne "Coût" introuvable.');
    }

    const updated = data.slice(1).map((row, idx) => {
      const rawCost = row[costIndex];
      const cost = Number(rawCost);
      if (Number.isNaN(cost) || cost < 0) {
        Logger.log(`[mettreAJourCoutsRecompenses] Coût invalide ligne ${idx + 2} (${rawCost}), remis à 0.`);
        row[costIndex] = 0;
      } else {
        row[costIndex] = Math.round(cost);
      }
      return row;
    });

    sheet.getRange(2, 1, updated.length, data[0].length).setValues(updated);
    Logger.log('[mettreAJourCoutsRecompenses] Coûts normalisés et mis à jour.');
    return { success: true, message: 'Coûts des récompenses mis à jour.' };
  } catch (error) {
    Logger.log(`[mettreAJourCoutsRecompenses] Erreur : ${error}`);
    throw new Error('Impossible de mettre à jour les coûts des récompenses.');
  }
}

function normaliserCategorieTache_(categorie) {
  const normalized = normaliserTexte_(categorie);
  if (!normalized) {
    return TASK_CATEGORY_KEYS.autres;
  }
  if (normalized.includes('corvee') || normalized.includes('travaux')) {
    return TASK_CATEGORY_KEYS.corvees;
  }
  if (normalized.includes('comportement')) {
    return TASK_CATEGORY_KEYS.comportement;
  }
  if (normalized.includes('rituel')) {
    return TASK_CATEGORY_KEYS.rituels;
  }
  return TASK_CATEGORY_KEYS.autres;
}

function normaliserJoursTache_(rawValue) {
  if (!rawValue) {
    return null;
  }
  const value = String(rawValue).toLowerCase().trim();
  if (!value) {
    return null;
  }

  if (['tous', 'toute', 'toutes', 'toute la semaine', 'toute-semaine', 'toute_semaine', 'touslesjours', 'tous-les-jours', 'tous les jours', '7/7'].includes(value)) {
    return new Set([1, 2, 3, 4, 5, 6, 7]);
  }

  if (['week-end', 'weekend', 'week end', 'weekends'].includes(value)) {
    return new Set([6, 7]);
  }

  if (['lun-ven', 'lundi-vendredi', 'semaine', 'en semaine'].includes(value)) {
    return new Set([1, 2, 3, 4, 5]);
  }

  const separators = /[,;/\n]+/;
  const parts = value.split(separators).map(part => part.trim()).filter(Boolean);
  const dayMap = {
    lun: 1,
    lundi: 1,
    mar: 2,
    mardi: 2,
    mer: 3,
    mercredi: 3,
    jeu: 4,
    jeudi: 4,
    ven: 5,
    vendredi: 5,
    sam: 6,
    samedi: 6,
    dim: 7,
    dimanche: 7,
    '1': 1,
    '2': 2,
    '3': 3,
    '4': 4,
    '5': 5,
    '6': 6,
    '7': 7
  };

  const daySet = new Set();
  parts.forEach(part => {
    const normalized = part.replace(/\s+/g, '');
    if (normalized.includes('-')) {
      const [startRaw, endRaw] = normalized.split('-').map(token => token.trim());
      const start = dayMap[startRaw];
      const end = dayMap[endRaw];
      if (start && end) {
        if (start <= end) {
          for (let day = start; day <= end; day++) {
            daySet.add(day);
          }
        } else {
          for (let day = start; day <= 7; day++) {
            daySet.add(day);
          }
          for (let day = 1; day <= end; day++) {
            daySet.add(day);
          }
        }
      }
      return;
    }

    const mapped = dayMap[normalized];
    if (mapped) {
      daySet.add(mapped);
    }
  });

  return daySet.size > 0 ? daySet : null;
}

function isTacheDisponibleAujourdHui_(rawValue, taskId, rowIndex) {
  if (!rawValue) {
    Logger.log(`[isTacheDisponibleAujourdHui] Tâche ${taskId} sans règle jour (ligne ${rowIndex + 2}), disponible par défaut.`);
    return { available: true, reason: 'aucune_regle' };
  }

  const days = normaliserJoursTache_(rawValue);
  if (!days) {
    Logger.log(`[isTacheDisponibleAujourdHui] Règle jours invalide pour ${taskId} (ligne ${rowIndex + 2}) : "${rawValue}". Tâche conservée par défaut.`);
    return { available: true, reason: 'regle_invalide' };
  }

  const todayIndex = getParisDayIndex(new Date());
  const available = days.has(todayIndex);
  Logger.log(`[isTacheDisponibleAujourdHui] Tâche ${taskId} (ligne ${rowIndex + 2}) ${available ? 'disponible' : 'indisponible'} aujourd'hui (index=${todayIndex}).`);
  return { available, reason: available ? 'jour_ok' : 'jour_ko' };
}

// ==================================================
// DIAGNOSTIC DATES (PARIS)
// ==================================================
function diagnostiquerDatesEvaluation(personne) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('Évaluations');
    const data = sheet.getDataRange().getValues().slice(1);
    const personneKey = String(personne || '').trim();
    const now = new Date();

    const diagnostic = {
      personne: personneKey,
      timezoneScript: Session.getScriptTimeZone(),
      timezoneSpreadsheet: ss.getSpreadsheetTimeZone(),
      nowIso: now.toISOString(),
      nowParis: Utilities.formatDate(now, PARIS_TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
      todayKeyParis: getParisDateKey(now),
      lastEvaluations: []
    };

    const evals = data.filter(row => String(row[3] || '').trim() === personneKey);
    const sorted = evals.sort((a, b) => {
      const dateA = parseSheetDate(a[1], 'Évaluations.Date');
      const dateB = parseSheetDate(b[1], 'Évaluations.Date');
      const timeA = dateA ? dateA.getTime() : 0;
      const timeB = dateB ? dateB.getTime() : 0;
      return timeB - timeA;
    });

    diagnostic.lastEvaluations = sorted.slice(0, 5).map(row => {
      const parsedDate = parseSheetDate(row[1], 'Évaluations.Date');
      return {
        id: row[0],
        rawDate: row[1],
        parsedIso: parsedDate ? parsedDate.toISOString() : null,
        parisKey: parsedDate ? getParisDateKey(parsedDate) : null,
        personne: String(row[3] || '').trim()
      };
    });

    Logger.log(`[diagnostiquerDatesEvaluation] Diagnostic Paris: ${JSON.stringify(diagnostic)}`);
    return diagnostic;
  } catch (error) {
    Logger.log(`[diagnostiquerDatesEvaluation] Erreur diagnostic pour ${personne} : ${error}`);
    throw new Error('Impossible de diagnostiquer les dates d’évaluation (Paris).');
  }
}

// ==================================================
// WEB APP
// ==================================================
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('🌟 Mes Étoiles')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==================================================
// RÉCUPÉRER LES PERSONNES
// ==================================================
function getPersonnes() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Personnes');
  const data = sheet.getDataRange().getValues().slice(1);
  
  return data.map(row => ({
    nom: row[0],
    avatar: row[1],
    couleur: row[2],
    age: row[3]
  }));
}

// ==================================================
// RÉCUPÉRER LES ÉMOTIONS
// ==================================================
function getEmotions() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Émotions');
  const data = sheet.getDataRange().getValues().slice(1);
  
  return data.map(row => ({
    id: row[0],
    nom: row[1],
    emoji: row[2],
    couleur: row[3],
    description: row[4],
    categorie: row[5]
  }));
}

// ==================================================
// RÉCUPÉRER LES SOURCES D'ÉMOTIONS
// ==================================================
function getSources() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Sources_Émotions');
  const data = sheet.getDataRange().getValues().slice(1);
  
  return data.map(row => ({
    id: row[0],
    nom: row[1],
    emoji: row[2],
    description: row[3]
  }));
}

// ==================================================
// RÉCUPÉRER LES TÂCHES ASSIGNÉES
// ==================================================
function getTachesFeuille_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Tâches');
  if (!sheet) {
    Logger.log('[getTachesFeuille] Feuille Tâches introuvable.');
    throw new Error('Feuille Tâches introuvable.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('[getTachesFeuille] Feuille Tâches vide.');
    return { headers: [], rows: [], indexes: {} };
  }

  const headers = data[0].map(value => String(value || '').trim());
  const indexes = {
    id: headers.indexOf('ID'),
    categorie: headers.indexOf('Catégorie'),
    nom: headers.indexOf('Nom'),
    emoji: headers.indexOf('Emoji'),
    description: headers.indexOf('Description'),
    pointsMax: headers.indexOf('PointsMax'),
    ordre: headers.indexOf('Ordre'),
    personne: headers.indexOf('Personne'),
    personnes: headers.indexOf('Personnes'),
    jours: headers.indexOf('Jours')
  };

  return { headers, rows: data.slice(1), indexes };
}

function getTachesDefinitions_() {
  try {
    const { rows, indexes } = getTachesFeuille_();
    if (indexes.id === -1) {
      Logger.log('[getTachesDefinitions] Colonne ID manquante dans la feuille Tâches.');
      return [];
    }

    const tasks = [];
    rows.forEach((row, rowIndex) => {
      const rawId = String(row[indexes.id] || '').trim();
      if (!rawId) {
        Logger.log(`[getTachesDefinitions] ID manquant, ligne ${rowIndex + 2} ignorée.`);
        return;
      }
      const categorie = indexes.categorie !== -1 ? String(row[indexes.categorie] || '').trim() : '';
      const nom = indexes.nom !== -1 ? String(row[indexes.nom] || '').trim() : rawId;
      const emoji = indexes.emoji !== -1 ? String(row[indexes.emoji] || '').trim() : '';
      const description = indexes.description !== -1 ? String(row[indexes.description] || '').trim() : '';
      const pointsMax = indexes.pointsMax !== -1 ? Number(row[indexes.pointsMax]) : 1;
      const ordre = indexes.ordre !== -1 ? Number(row[indexes.ordre]) : null;

      tasks.push({
        id: rawId,
        categorie,
        categorieNormalisee: normaliserCategorieTache_(categorie),
        nom,
        emoji,
        description,
        pointsMax: Number.isNaN(pointsMax) ? 1 : pointsMax,
        ordre: Number.isNaN(ordre) ? null : ordre
      });
    });

    return tasks;
  } catch (error) {
    Logger.log(`[getTachesDefinitions] Erreur lors de la lecture des tâches : ${error}`);
    return [];
  }
}

function getTachesPourPersonne(personne) {
  try {
    const result = getTachesAssigneesPourPersonne_(personne);
    const definitions = getTachesDefinitions_();
    const definitionsById = new Map(definitions.map(task => [task.id, task]));
    const tasks = result.taskIds
      .map(taskId => definitionsById.get(taskId))
      .filter(Boolean);
    const allTaskIds = definitions.length > 0 ? definitions.map(task => task.id) : BASE_TASK_IDS;
    Logger.log(`[getTachesPourPersonne] Tâches filtrées pour ${personne} : ${JSON.stringify(result.taskIds)}`);
    return {
      personne: String(personne || '').trim(),
      taskIds: result.taskIds,
      allowEmpty: result.allowEmpty,
      reason: result.reason,
      allTaskIds,
      tasks
    };
  } catch (error) {
    Logger.log(`[getTachesPourPersonne] Erreur pour ${personne} : ${error}`);
    throw new Error('Impossible de charger les tâches attribuées.');
  }
}

function getTachesAssigneesPourPersonne_(personne) {
  const personneKey = String(personne || '').trim();
  if (!personneKey) {
    Logger.log('[getTachesAssigneesPourPersonne] Personne non renseignée, retour de la liste par défaut.');
    return { taskIds: [...BASE_TASK_IDS], allowEmpty: false, reason: 'personne_absente' };
  }

  let rows = [];
  let indexes = {};
  try {
    const feuille = getTachesFeuille_();
    rows = feuille.rows;
    indexes = feuille.indexes;
  } catch (error) {
    Logger.log(`[getTachesAssigneesPourPersonne] Feuille Tâches introuvable, retour de la liste par défaut. Erreur=${error}`);
    return { taskIds: [...BASE_TASK_IDS], allowEmpty: false, reason: 'feuille_absente' };
  }

  if (rows.length === 0) {
    Logger.log('[getTachesAssigneesPourPersonne] Feuille Tâches vide, retour de la liste par défaut.');
    return { taskIds: [...BASE_TASK_IDS], allowEmpty: false, reason: 'feuille_vide' };
  }

  if (indexes.id === -1) {
    Logger.log('[getTachesAssigneesPourPersonne] Colonne ID manquante, retour de la liste par défaut.');
    return { taskIds: [...BASE_TASK_IDS], allowEmpty: false, reason: 'id_manquant' };
  }

  const assignedTasks = [];
  const addTask = taskId => {
    if (!assignedTasks.includes(taskId)) {
      assignedTasks.push(taskId);
    }
  };

  rows.forEach((row, rowIndex) => {
    const rawId = String(row[indexes.id] || '').trim();
    if (!rawId) {
      Logger.log(`[getTachesAssigneesPourPersonne] ID manquant, ligne ${rowIndex + 2} ignorée.`);
      return;
    }

    const taskId = rawId;
    if (indexes.jours !== -1) {
      const rawJours = row[indexes.jours];
      const availability = isTacheDisponibleAujourdHui_(rawJours, taskId, rowIndex);
      if (!availability.available) {
        Logger.log(`[getTachesAssigneesPourPersonne] Tâche ${taskId} ignorée pour ${personneKey} (ligne ${rowIndex + 2}) - règle jours.`);
        return;
      }
    }
    const targetIndex = indexes.personne !== -1 ? indexes.personne : indexes.personnes;
    if (targetIndex === -1) {
      addTask(taskId);
      return;
    }

    const rawAssignees = String(row[targetIndex] || '').trim();
    if (!rawAssignees) {
      addTask(taskId);
      return;
    }

    const assignees = rawAssignees
      .split(/[,;\n]+/)
      .map(value => value.trim())
      .filter(Boolean);

    if (assignees.length === 0) {
      addTask(taskId);
      return;
    }

    if (assignees.includes(personneKey)) {
      addTask(taskId);
      Logger.log(`[getTachesAssigneesPourPersonne] Tâche ${taskId} assignée (ligne ${rowIndex + 2}).`);
    } else {
      Logger.log(`[getTachesAssigneesPourPersonne] Tâche ${taskId} ignorée pour ${personneKey} (ligne ${rowIndex + 2}).`);
    }
  });

  const reason = assignedTasks.length > 0 ? 'filtrage_ok' : 'aucune_tache';
  Logger.log(`[getTachesAssigneesPourPersonne] Résultat ${reason} pour ${personneKey} : ${JSON.stringify(assignedTasks)}`);
  return { taskIds: assignedTasks, allowEmpty: true, reason };
}

// ==================================================
// VÉRIFIER SI ÉVALUÉ AUJOURD'HUI
// ==================================================
function hasEvaluatedToday(personne) {
  try {
    const evalMeta = getEvaluationsFeuille_();
    const data = evalMeta.rows;
    const dateIndex = getEvaluationColumnIndex_(evalMeta.indexes, 'date', 1, 'Date');
    const personneIndex = getEvaluationColumnIndex_(evalMeta.indexes, 'personne', 3, 'Personne');
    
    const todayKey = getParisDateKey(new Date());
    const personneKey = String(personne || '').trim();
    Logger.log(`[hasEvaluatedToday] Vérification Paris pour ${personneKey} (date=${todayKey}).`);
    
    return data.some(row => {
      const rowKey = getParisDateKeyFromValue(row[dateIndex], 'Évaluations.Date');
      const rowPersonne = String(row[personneIndex] || '').trim();
      if (!rowKey) {
        return false;
      }
      return rowPersonne === personneKey && rowKey === todayKey;
    });
  } catch (error) {
    Logger.log(`[hasEvaluatedToday] Erreur lors de la vérification Paris pour ${personne} : ${error}`);
    throw new Error('Impossible de vérifier l’évaluation du jour (timezone Paris).');
  }
}

// ==================================================
// SOUMETTRE UNE ÉVALUATION
// ==================================================
function submitEvaluation(personne, taches, emotions, humeur, commentaire) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('Évaluations');
    
    if (hasEvaluatedToday(personne)) {
      Logger.log(`[submitEvaluation] Évaluation déjà faite aujourd'hui (Paris) pour ${personne}.`);
      return { success: false, message: 'Tu as déjà fait ton évaluation aujourd\'hui ! 😊' };
    }

    const emotionPairs = [
      { emotion: emotions.emotion1, source: emotions.source1, label: 'emotion1' },
      { emotion: emotions.emotion2, source: emotions.source2, label: 'emotion2' },
      { emotion: emotions.emotion3, source: emotions.source3, label: 'emotion3' }
    ];
    const missingSources = emotionPairs
      .filter(pair => pair.emotion && !pair.source)
      .map(pair => pair.label);
    if (missingSources.length > 0) {
      Logger.log(`[submitEvaluation] Sources manquantes pour ${personne} : ${missingSources.join(', ')}`);
      return { success: false, message: 'Choisis une cause pour chaque émotion, s’il te plaît.' };
    }
    
    const lastRow = sheet.getLastRow();
    const newId = 'E' + String(lastRow).padStart(4, '0');
    const now = new Date();

    const MAX_COMMENTAIRE_LENGTH = 400;
    const commentaireBrut = String(commentaire || '').trim();
    let commentaireSafe = commentaireBrut;
    if (!commentaireSafe) {
      Logger.log(`[submitEvaluation] Commentaire vide pour ${personne}.`);
    } else if (commentaireSafe.length > MAX_COMMENTAIRE_LENGTH) {
      Logger.log(`[submitEvaluation] Commentaire trop long pour ${personne} (longueur=${commentaireSafe.length}, max=${MAX_COMMENTAIRE_LENGTH}). Tronquage appliqué.`);
      commentaireSafe = commentaireSafe.slice(0, MAX_COMMENTAIRE_LENGTH);
    }
    
    const assignedResult = getTachesAssigneesPourPersonne_(personne);
    const assignedTasks = assignedResult.taskIds;
    const assignedSet = new Set(assignedTasks);
    const taskDefinitions = getTachesDefinitions_();
    const definitionsById = new Map(taskDefinitions.map(task => [task.id, task]));

    const safeTaskValue = (taskKey) => {
      const value = Number(taches && taches[taskKey]);
      if (!assignedSet.has(taskKey)) {
        return 0;
      }
      if (Number.isNaN(value)) {
        Logger.log(`[submitEvaluation] Valeur de tâche invalide pour ${taskKey} (${personne}). Valeur remise à 0.`);
        return 0;
      }
      return value;
    };

    const tachesNormalisees = {
      rangerChambre: safeTaskValue('rangerChambre'),
      faireLit: safeTaskValue('faireLit'),
      rangerJouets: safeTaskValue('rangerJouets'),
      aiderTable: safeTaskValue('aiderTable'),
      ecouter: safeTaskValue('ecouter'),
      gentilSoeur: safeTaskValue('gentilSoeur'),
      politesse: safeTaskValue('politesse'),
      pasColere: safeTaskValue('pasColere'),
      dentsMatin: safeTaskValue('dentsMatin'),
      dentsSoir: safeTaskValue('dentsSoir'),
      habiller: safeTaskValue('habiller'),
      cartable: safeTaskValue('cartable')
    };

    const totalsByCategory = {
      corvees: 0,
      comportement: 0,
      rituels: 0,
      autres: 0
    };
    const dynamicTasks = {};

    assignedTasks.forEach(taskId => {
      const value = safeTaskValue(taskId);
      const definition = definitionsById.get(taskId);
      const categorie = definition ? definition.categorieNormalisee : TASK_CATEGORY_KEYS.autres;
      if (!definition) {
        Logger.log(`[submitEvaluation] Définition de tâche introuvable pour ${taskId}, catégorie par défaut appliquée.`);
      }
      totalsByCategory[categorie] = (totalsByCategory[categorie] || 0) + value;
      if (!BASE_TASK_IDS.includes(taskId)) {
        dynamicTasks[taskId] = value;
      }
    });

    // Calculs totaux
    const totalCorvees = totalsByCategory.corvees;
    const totalComportement = totalsByCategory.comportement;
    const totalRituels = totalsByCategory.rituels;
    const totalEmotions = emotions.gestion;
    const totalJour = totalCorvees + totalComportement + totalRituels + totalEmotions;

    Logger.log(`[submitEvaluation] Totaux calculés pour ${personne} : corvées=${totalCorvees}, comportement=${totalComportement}, rituels=${totalRituels}, émotions=${totalEmotions}, totalJour=${totalJour}.`);
    
    Logger.log(`[submitEvaluation] Ajout évaluation ${newId} pour ${personne} (Paris=${getParisDateKey(now)}).`);
    
    const syncResult = synchroniserColonnesTachesEvaluations_();
    const headers = syncResult.headers;
    const dynamicHeader = 'Taches_Dynamiques';
    let dynamicColumnIndex = headers.indexOf(dynamicHeader);
    if (dynamicColumnIndex === -1) {
      dynamicColumnIndex = headers.length;
      sheet.getRange(1, dynamicColumnIndex + 1).setValue(dynamicHeader);
      Logger.log(`[submitEvaluation] Colonne ${dynamicHeader} ajoutée en ${dynamicColumnIndex + 1}.`);
    } else if (dynamicColumnIndex !== headers.length - 1) {
      Logger.log(`[submitEvaluation] Colonne ${dynamicHeader} non en dernière position (index=${dynamicColumnIndex + 1}). Valeur ajoutée en fin de ligne.`);
    }

    const dynamicPayload = Object.keys(dynamicTasks).length > 0 ? JSON.stringify(dynamicTasks) : '';

    // Ajouter la ligne
    sheet.appendRow([
      newId,
      now,
      Utilities.formatDate(now, PARIS_TIMEZONE, 'HH:mm'),
      personne,
      // Tâches
      tachesNormalisees.rangerChambre,
      tachesNormalisees.faireLit,
      tachesNormalisees.rangerJouets,
      tachesNormalisees.aiderTable,
      tachesNormalisees.ecouter,
      tachesNormalisees.gentilSoeur,
      tachesNormalisees.politesse,
      tachesNormalisees.pasColere,
      tachesNormalisees.dentsMatin,
      tachesNormalisees.dentsSoir,
      tachesNormalisees.habiller,
      tachesNormalisees.cartable,
      // Émotions
      emotions.emotion1,
      emotions.emotion2 || '',
      emotions.emotion3 || '',
      emotions.source1 || '',
      emotions.source2 || '',
      emotions.source3 || '',
      emotions.gestion,
      // Totaux
      totalCorvees,
      totalComportement,
      totalRituels,
      totalEmotions,
      totalJour,
      // Meta
      humeur,
      commentaireSafe,
      dynamicPayload
    ]);

    const appendedRowIndex = sheet.getLastRow();
    appliquerValeursTachesDynamiques_(
      sheet,
      appendedRowIndex,
      assignedTasks,
      safeTaskValue,
      definitionsById,
      syncResult.headerIndex
    );
    
    // Enregistrer dans historique émotions
    saveEmotionHistory(personne, emotions);
    
    // Vérifier badges
    const newBadges = checkBadges(personne);
    
    // Message selon score
    const baseTaskCount = assignedTasks.length;
    const maxPoints = baseTaskCount + 1;
    const percent = Math.max(0, Math.round((totalJour / maxPoints) * 100));
    
    let message, stars;
    if (percent >= 90) {
      message = "INCROYABLE ! Tu es une vraie STAR ! 🌟";
      stars = 5;
    } else if (percent >= 75) {
      message = "SUPER journée ! Bravo champion ! 🎉";
      stars = 4;
    } else if (percent >= 60) {
      message = "Bien joué ! Continue comme ça ! 👍";
      stars = 3;
    } else if (percent >= 40) {
      message = "Pas mal ! Tu peux faire encore mieux ! 💪";
      stars = 2;
    } else {
      message = "Demain sera meilleur ! On y croit ! 🌈";
      stars = 1;
    }
    
    return {
      success: true,
      message: message,
      totalJour: totalJour,
      maxJour: maxPoints,
      percent: percent,
      stars: stars,
      newBadges: newBadges,
      details: {
        corvees: totalCorvees,
        comportement: totalComportement,
        rituels: totalRituels,
        emotions: totalEmotions
      }
    };
  } catch (error) {
    Logger.log(`[submitEvaluation] Erreur lors de l'envoi pour ${personne} : ${error}`);
    return { success: false, message: 'Erreur lors de l’enregistrement. Réessaie dans un instant.' };
  }
}

// ==================================================
// SAUVEGARDER HISTORIQUE ÉMOTIONS
// ==================================================
function saveEmotionHistory(personne, emotions) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Historique_Émotions');
  
  sheet.appendRow([
    new Date(),
    personne,
    emotions.emotion1,
    emotions.emotion2 || '',
    emotions.emotion3 || '',
    emotions.source1 || '',
    emotions.source2 || '',
    emotions.source3 || '',
    emotions.gestion,
    ''
  ]);
}

// ==================================================
// DONNÉES PERSONNE
// ==================================================
function getPersonneData(personne) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    
    // Infos personne
    const personnesSheet = ss.getSheetByName('Personnes');
    const personnesData = personnesSheet.getDataRange().getValues().slice(1);
    const personneKey = String(personne || '').trim();
    const personneInfo = personnesData.find(r => String(r[0] || '').trim() === personneKey);
    
    // Évaluations
    const evalMeta = getEvaluationsFeuille_();
    const evalData = evalMeta.rows;
    const evalIndexes = evalMeta.indexes;
    const dateIndex = getEvaluationColumnIndex_(evalIndexes, 'date', 1, 'Date');
    const personneIndex = getEvaluationColumnIndex_(evalIndexes, 'personne', 3, 'Personne');
    const totalJourIndex = getEvaluationColumnIndex_(evalIndexes, 'totalJour', 27, 'TotalJour');

    // Récompenses demandées (dépenses)
    const claimsSheet = ss.getSheetByName('Récompenses_Demandées');
    let claimsData = [];
    if (!claimsSheet) {
      Logger.log('[getPersonneData] Feuille "Récompenses_Demandées" introuvable, dépenses ignorées.');
    } else {
      claimsData = claimsSheet.getDataRange().getValues().slice(1);
    }
    
    // Semaine en cours (Paris)
    const todayParis = getParisMidnight(new Date());
    const weekStart = getMonday(todayParis);
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekEnd.getDate() + 6);
    weekEnd.setHours(23, 59, 59);

    const maxPointsData = getMaxPointsParJour_(personneKey);
    const maxPointsJour = maxPointsData.maxPoints;
    
    let weekPoints = 0;
    let weekDays = 0;
    let dailyScores = [null, null, null, null, null, null, null];
    
    const personneEvals = evalData.filter(r => String(r[personneIndex] || '').trim() === personneKey);
    
    personneEvals.forEach(row => {
      const parsedDate = parseSheetDate(row[dateIndex], 'Évaluations.Date');
      if (!parsedDate) {
        return;
      }
      const date = getParisMidnight(parsedDate);
      if (date >= weekStart && date <= weekEnd) {
        const total = Number(row[totalJourIndex] || 0);
        weekPoints += total;
        weekDays++;
        const dayIndex = date.getDay() === 0 ? 6 : date.getDay() - 1;
        dailyScores[dayIndex] = total;
      }
    });
    
    // Streak (Paris)
    let streak = 0;
    const sortedEvals = personneEvals.sort((a, b) => {
      const dateA = parseSheetDate(a[dateIndex], 'Évaluations.Date');
      const dateB = parseSheetDate(b[dateIndex], 'Évaluations.Date');
      const timeA = dateA ? dateA.getTime() : 0;
      const timeB = dateB ? dateB.getTime() : 0;
      return timeB - timeA;
    });
    
    if (sortedEvals.length > 0) {
      let checkDate = getParisMidnight(new Date());
      
      for (const eval of sortedEvals) {
        const parsedDate = parseSheetDate(eval[dateIndex], 'Évaluations.Date');
        if (!parsedDate) {
          continue;
        }
        const evalDate = getParisMidnight(parsedDate);
        const diffDays = Math.floor((checkDate - evalDate) / (1000 * 60 * 60 * 24));
        
        if (diffDays <= 1) {
          streak++;
          checkDate = new Date(evalDate);
          checkDate.setDate(checkDate.getDate() - 1);
        } else {
          break;
        }
      }
    }
    
    // Émotions récentes
    const emotionSheet = ss.getSheetByName('Historique_Émotions');
    const emotionData = emotionSheet.getDataRange().getValues().slice(1);
    const recentEmotions = emotionData
      .filter(r => String(r[1] || '').trim() === personneKey)
      .sort((a, b) => {
        const dateA = parseSheetDate(a[0], 'Historique_Émotions.Date');
        const dateB = parseSheetDate(b[0], 'Historique_Émotions.Date');
        const timeA = dateA ? dateA.getTime() : 0;
        const timeB = dateB ? dateB.getTime() : 0;
        return timeB - timeA;
      })
      .slice(0, 7)
      .map(r => {
        const parsedDate = parseSheetDate(r[0], 'Historique_Émotions.Date');
        return {
          date: parsedDate ? Utilities.formatDate(parsedDate, PARIS_TIMEZONE, 'dd/MM') : '—',
          emotion1: r[2],
          emotion2: r[3],
          emotion3: r[4],
          source1: r[5],
          source2: r[6],
          source3: r[7],
          gestion: r[8]
        };
      });
    
    // Badges
    const badgesObtSheet = ss.getSheetByName('Badges_Obtenus');
    const badgesObtData = badgesObtSheet.getDataRange().getValues().slice(1);
    const personneBadgesIds = badgesObtData.filter(r => r[0] === personne).map(r => r[1]);
    
    const badgesDefSheet = ss.getSheetByName('Badges');
    const badgesDef = badgesDefSheet.getDataRange().getValues().slice(1);
    const badges = personneBadgesIds.map(bid => {
      const def = badgesDef.find(b => b[0] === bid);
      return def ? { id: def[0], nom: def[1], emoji: def[2] } : null;
    }).filter(b => b);
    
    const pointsDisponibles = getPointsDisponibles_(personneKey, evalData, claimsData);

    // Récompenses
    const rewardsSheet = ss.getSheetByName('Récompenses');
    const rewardsData = rewardsSheet.getDataRange().getValues().slice(1);
    const rewards = rewardsData
      .filter(r => r[5] === 'Oui')
      .map(r => ({
        id: r[0],
        nom: r[1],
        emoji: r[2],
        cout: Number(r[3]) || 0,
        disponible: pointsDisponibles.totalPoints >= (Number(r[3]) || 0)
      }));
    
    return {
      nom: personne,
      avatar: personneInfo ? personneInfo[1] : '👤',
      couleur: personneInfo ? personneInfo[2] : '#6C5CE7',
      age: personneInfo ? personneInfo[3] : 0,
      weekPoints: weekPoints,
      totalPoints: pointsDisponibles.totalPoints,
      totalEarned: pointsDisponibles.totalGagnes,
      totalSpent: pointsDisponibles.totalDepenses,
      weekDays: weekDays,
      dailyScores: dailyScores,
      streak: streak,
      recentEmotions: recentEmotions,
      badges: badges,
      rewards: rewards,
      maxPointsJour: maxPointsJour,
      evaluatedToday: hasEvaluatedToday(personne),
      weekStart: Utilities.formatDate(weekStart, PARIS_TIMEZONE, 'dd/MM'),
      weekEnd: Utilities.formatDate(weekEnd, PARIS_TIMEZONE, 'dd/MM')
    };
  } catch (error) {
    Logger.log(`[getPersonneData] Erreur pour ${personne} : ${error}`);
    throw new Error('Impossible de charger les données (timezone Paris).');
  }
}

// ==================================================
// DONNÉES FAMILLE
// ==================================================
function getFamilyData() {
  const personnes = getPersonnes();
  return personnes.map(p => {
    const data = getPersonneData(p.nom);
    return {
      nom: p.nom,
      avatar: p.avatar,
      couleur: p.couleur,
      totalPoints: data.totalPoints,
      streak: data.streak,
      badgeCount: data.badges.length,
      evaluatedToday: data.evaluatedToday
    };
  }).sort((a, b) => b.totalPoints - a.totalPoints);
}

// ==================================================
// RÉCLAMER RÉCOMPENSE
// ==================================================
function claimReward(personne, rewardId) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const data = getPersonneData(personne);
  
  const rewardsSheet = ss.getSheetByName('Récompenses');
  const rewardsData = rewardsSheet.getDataRange().getValues().slice(1);
  const reward = rewardsData.find(r => r[0] === rewardId);
  
  if (!reward) {
    return { success: false, message: 'Récompense introuvable 😕' };
  }

  const rewardCost = Number(reward[3]) || 0;
  
  if (data.totalPoints < rewardCost) {
    return { success: false, message: `Il te manque ${rewardCost - data.totalPoints} étoiles 😢` };
  }

  Logger.log(`[claimReward] Demande de récompense pour ${personne} (${rewardId}) : coût=${rewardCost}, points dispos=${data.totalPoints}.`);
  
  const claimsSheet = ss.getSheetByName('Récompenses_Demandées');
  const newId = 'C' + String(claimsSheet.getLastRow()).padStart(4, '0');
  
  claimsSheet.appendRow([
    newId,
    new Date(),
    personne,
    reward[1],
    rewardCost,
    'En attente',
    '',
    ''
  ]);
  
  return { 
    success: true, 
    message: `🎉 Super ! "${reward[1]}" demandé !`
  };
}

// ==================================================
// BADGES
// ==================================================
function checkBadges(personne) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const data = getPersonneData(personne);
  const evalMeta = getEvaluationsFeuille_();
  const personneIndex = getEvaluationColumnIndex_(evalMeta.indexes, 'personne', 3, 'Personne');
  const totalJourIndex = getEvaluationColumnIndex_(evalMeta.indexes, 'totalJour', 27, 'TotalJour');
  const gestionEmotionIndex = getEvaluationColumnIndex_(evalMeta.indexes, 'gestionEmotion', 22, 'GestionEmotion');
  const emotion1Index = getEvaluationColumnIndex_(evalMeta.indexes, 'emotion1', 16, 'Emotion1');
  const evals = evalMeta.rows.filter(r => r[personneIndex] === personne);
  
  const newBadges = [];
  
  // B01 - Première étoile
  if (evals.length >= 1 && !data.badges.some(b => b.id === 'B01')) {
    if (awardBadge(personne, 'B01')) {
      newBadges.push({ id: 'B01', nom: 'Première étoile', emoji: '⭐' });
    }
  }
  
  // B05 - Semaine champion (7 jours)
  if (data.weekDays >= 7 && !data.badges.some(b => b.id === 'B05')) {
    if (awardBadge(personne, 'B05')) {
      newBadges.push({ id: 'B05', nom: 'Semaine champion', emoji: '🏆' });
    }
  }
  
  // B06 - Journée parfaite (26/26)
  const hasPerfect = evals.some(r => Number(r[totalJourIndex] || 0) >= 26);
  if (hasPerfect && !data.badges.some(b => b.id === 'B06')) {
    if (awardBadge(personne, 'B06')) {
      newBadges.push({ id: 'B06', nom: 'Journée parfaite', emoji: '🌟' });
    }
  }
  
  // B08 - Zen master (5x gestion émotions = 2)
  const goodGestion = evals.filter(r => Number(r[gestionEmotionIndex] || 0) === 2).length;
  if (goodGestion >= 5 && !data.badges.some(b => b.id === 'B08')) {
    if (awardBadge(personne, 'B08')) {
      newBadges.push({ id: 'B08', nom: 'Zen master', emoji: '🧘' });
    }
  }
  
  // B11 - Explorateur émotions (7 jours avec émotions)
  const daysWithEmotions = evals.filter(r => r[emotion1Index] && r[emotion1Index] !== '').length;
  if (daysWithEmotions >= 7 && !data.badges.some(b => b.id === 'B11')) {
    if (awardBadge(personne, 'B11')) {
      newBadges.push({ id: 'B11', nom: 'Explorateur émotions', emoji: '🎭' });
    }
  }
  
  // B10 - Collectionneur (5 badges)
  if (data.badges.length + newBadges.length >= 5 && !data.badges.some(b => b.id === 'B10')) {
    if (awardBadge(personne, 'B10')) {
      newBadges.push({ id: 'B10', nom: 'Collectionneur', emoji: '👑' });
    }
  }
  
  return newBadges;
}

function awardBadge(personne, badgeId) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Badges_Obtenus');
  const data = sheet.getDataRange().getValues().slice(1);
  
  if (data.some(r => r[0] === personne && r[1] === badgeId)) {
    return false;
  }
  
  sheet.appendRow([personne, badgeId, new Date()]);
  return true;
}

// ==================================================
// UTILITAIRES
// ==================================================
function getMonday(date) {
  const d = getParisMidnight(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  d.setDate(diff);
  d.setHours(0, 0, 0, 0);
  return d;
}
