// ==================================================
// 🌟 SYSTÈME DE MÉRITE FAMILIAL v4
// Avec section Émotions
// ==================================================

const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const PARIS_TIMEZONE = 'Europe/Paris';
const TASK_IDS = [
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

// ==================================================
// CONFIGURATION - COLONNE JOURS (TÂCHES)
// ==================================================
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
      return { success: true, message: 'Colonne Jours déjà présente en I1.' };
    }

    if (currentValue && currentValue !== 'Jours') {
      Logger.log(`[creerColonneJoursTaches] Valeur existante en I1 ("${currentValue}"), écrasement avec "Jours".`);
    } else {
      Logger.log('[creerColonneJoursTaches] Création de la colonne Jours en I1.');
    }

    cell.setValue('Jours');
    return { success: true, message: 'Colonne Jours créée en I1.' };
  } catch (error) {
    Logger.log(`[creerColonneJoursTaches] Erreur lors de la création : ${error}`);
    throw new Error('Impossible de créer la colonne Jours dans la feuille Tâches.');
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

function normaliserJoursTache_(rawValue) {
  if (!rawValue) {
    return null;
  }
  const value = String(rawValue).toLowerCase().trim();
  if (!value) {
    return null;
  }

  if (['tous', 'toute', 'toutes', 'toute la semaine', 'toute-semaine', 'toute_semaine', '7/7'].includes(value)) {
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
    dimanche: 7
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
function getTachesPourPersonne(personne) {
  try {
    const result = getTachesAssigneesPourPersonne_(personne);
    Logger.log(`[getTachesPourPersonne] Tâches filtrées pour ${personne} : ${JSON.stringify(result.taskIds)}`);
    return {
      personne: String(personne || '').trim(),
      taskIds: result.taskIds,
      allowEmpty: result.allowEmpty,
      reason: result.reason,
      allTaskIds: TASK_IDS
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
    return { taskIds: [...TASK_IDS], allowEmpty: false, reason: 'personne_absente' };
  }
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Tâches');
  if (!sheet) {
    Logger.log('[getTachesAssigneesPourPersonne] Feuille Tâches introuvable, retour de la liste par défaut.');
    return { taskIds: [...TASK_IDS], allowEmpty: false, reason: 'feuille_absente' };
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('[getTachesAssigneesPourPersonne] Feuille Tâches vide, retour de la liste par défaut.');
    return { taskIds: [...TASK_IDS], allowEmpty: false, reason: 'feuille_vide' };
  }

  const headers = data[0].map(value => String(value || '').trim());
  const idIndex = headers.indexOf('ID');
  const personneIndex = headers.indexOf('Personne');
  const personnesIndex = headers.indexOf('Personnes');
  const ordreIndex = headers.indexOf('Ordre');
  const joursIndex = headers.indexOf('Jours');

  if (idIndex === -1) {
    Logger.log('[getTachesAssigneesPourPersonne] Colonne ID manquante, retour de la liste par défaut.');
    return { taskIds: [...TASK_IDS], allowEmpty: false, reason: 'id_manquant' };
  }

  const assignedTasks = [];
  const addTask = taskId => {
    if (!TASK_IDS.includes(taskId)) {
      Logger.log(`[getTachesAssigneesPourPersonne] ID de tâche inconnu ignoré : ${taskId}`);
      return;
    }
    if (!assignedTasks.includes(taskId)) {
      assignedTasks.push(taskId);
    }
  };

  const resolveTaskKey = (row) => {
    const rawId = String(row[idIndex] || '').trim();
    if (TASK_IDS.includes(rawId)) {
      return { taskKey: rawId, source: 'id_direct' };
    }

    if (rawId) {
      const match = rawId.match(/\d+/);
      if (match) {
        const index = Number(match[0]) - 1;
        if (index >= 0 && index < TASK_IDS.length) {
          return { taskKey: TASK_IDS[index], source: 'id_numerique' };
        }
      }
    }

    if (ordreIndex !== -1) {
      const ordreValue = Number(row[ordreIndex]);
      const ordreIndexBased = ordreValue - 1;
      if (!Number.isNaN(ordreValue) && ordreIndexBased >= 0 && ordreIndexBased < TASK_IDS.length) {
        return { taskKey: TASK_IDS[ordreIndexBased], source: 'ordre' };
      }
    }

    return { taskKey: '', source: 'inconnu' };
  };

  data.slice(1).forEach((row, rowIndex) => {
    const resolved = resolveTaskKey(row);
    if (!resolved.taskKey) {
      const rawId = String(row[idIndex] || '').trim();
      Logger.log(`[getTachesAssigneesPourPersonne] ID de tâche inconnu ignoré : ${rawId}`);
      return;
    }

    const taskId = resolved.taskKey;
    if (joursIndex !== -1) {
      const rawJours = row[joursIndex];
      const availability = isTacheDisponibleAujourdHui_(rawJours, taskId, rowIndex);
      if (!availability.available) {
        Logger.log(`[getTachesAssigneesPourPersonne] Tâche ${taskId} ignorée pour ${personneKey} (ligne ${rowIndex + 2}) - règle jours.`);
        return;
      }
    }
    const targetIndex = personneIndex !== -1 ? personneIndex : personnesIndex;
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
      Logger.log(`[getTachesAssigneesPourPersonne] Tâche ${taskId} assignée via ${resolved.source} (ligne ${rowIndex + 2}).`);
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
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('Évaluations');
    const data = sheet.getDataRange().getValues().slice(1);
    
    const todayKey = getParisDateKey(new Date());
    const personneKey = String(personne || '').trim();
    Logger.log(`[hasEvaluatedToday] Vérification Paris pour ${personneKey} (date=${todayKey}).`);
    
    return data.some(row => {
      const rowKey = getParisDateKeyFromValue(row[1], 'Évaluations.Date');
      const rowPersonne = String(row[3] || '').trim();
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
    
    const assignedResult = getTachesAssigneesPourPersonne_(personne);
    const assignedTasks = assignedResult.taskIds;
    const assignedSet = new Set(assignedTasks);
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

    // Calculs totaux
    const totalCorvees = tachesNormalisees.rangerChambre + tachesNormalisees.faireLit + tachesNormalisees.rangerJouets + tachesNormalisees.aiderTable;
    const totalComportement = tachesNormalisees.ecouter + tachesNormalisees.gentilSoeur + tachesNormalisees.politesse + tachesNormalisees.pasColere;
    const totalRituels = tachesNormalisees.dentsMatin + tachesNormalisees.dentsSoir + tachesNormalisees.habiller + tachesNormalisees.cartable;
    const totalEmotions = emotions.gestion;
    const totalJour = totalCorvees + totalComportement + totalRituels + totalEmotions;

    Logger.log(`[submitEvaluation] Totaux calculés pour ${personne} : corvées=${totalCorvees}, comportement=${totalComportement}, rituels=${totalRituels}, émotions=${totalEmotions}, totalJour=${totalJour}.`);
    
    Logger.log(`[submitEvaluation] Ajout évaluation ${newId} pour ${personne} (Paris=${getParisDateKey(now)}).`);
    
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
      commentaire || ''
    ]);
    
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
    const evalSheet = ss.getSheetByName('Évaluations');
    const evalData = evalSheet.getDataRange().getValues().slice(1);
    
    // Semaine en cours (Paris)
    const todayParis = getParisMidnight(new Date());
    const weekStart = getMonday(todayParis);
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekEnd.getDate() + 6);
    weekEnd.setHours(23, 59, 59);
    
    let weekPoints = 0;
    let weekDays = 0;
    let dailyScores = [null, null, null, null, null, null, null];
    
    const personneEvals = evalData.filter(r => String(r[3] || '').trim() === personneKey);
    
    personneEvals.forEach(row => {
      const parsedDate = parseSheetDate(row[1], 'Évaluations.Date');
      if (!parsedDate) {
        return;
      }
      const date = getParisMidnight(parsedDate);
      if (date >= weekStart && date <= weekEnd) {
        const total = row[27]; // Colonne TotalJour
        weekPoints += total;
        weekDays++;
        const dayIndex = date.getDay() === 0 ? 6 : date.getDay() - 1;
        dailyScores[dayIndex] = total;
      }
    });
    
    // Streak (Paris)
    let streak = 0;
    const sortedEvals = personneEvals.sort((a, b) => {
      const dateA = parseSheetDate(a[1], 'Évaluations.Date');
      const dateB = parseSheetDate(b[1], 'Évaluations.Date');
      const timeA = dateA ? dateA.getTime() : 0;
      const timeB = dateB ? dateB.getTime() : 0;
      return timeB - timeA;
    });
    
    if (sortedEvals.length > 0) {
      let checkDate = getParisMidnight(new Date());
      
      for (const eval of sortedEvals) {
        const parsedDate = parseSheetDate(eval[1], 'Évaluations.Date');
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
    
    // Récompenses
    const rewardsSheet = ss.getSheetByName('Récompenses');
    const rewardsData = rewardsSheet.getDataRange().getValues().slice(1);
    const rewards = rewardsData
      .filter(r => r[5] === 'Oui')
      .map(r => ({
        id: r[0],
        nom: r[1],
        emoji: r[2],
        cout: r[3],
        disponible: weekPoints >= r[3]
      }));
    
    return {
      nom: personne,
      avatar: personneInfo ? personneInfo[1] : '👤',
      couleur: personneInfo ? personneInfo[2] : '#6C5CE7',
      age: personneInfo ? personneInfo[3] : 0,
      weekPoints: weekPoints,
      weekDays: weekDays,
      dailyScores: dailyScores,
      streak: streak,
      recentEmotions: recentEmotions,
      badges: badges,
      rewards: rewards,
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
      weekPoints: data.weekPoints,
      streak: data.streak,
      badgeCount: data.badges.length,
      evaluatedToday: data.evaluatedToday
    };
  }).sort((a, b) => b.weekPoints - a.weekPoints);
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
  
  if (data.weekPoints < reward[3]) {
    return { success: false, message: `Il te manque ${reward[3] - data.weekPoints} étoiles 😢` };
  }
  
  const claimsSheet = ss.getSheetByName('Récompenses_Demandées');
  const newId = 'C' + String(claimsSheet.getLastRow()).padStart(4, '0');
  
  claimsSheet.appendRow([
    newId,
    new Date(),
    personne,
    reward[1],
    reward[3],
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
  const evalSheet = ss.getSheetByName('Évaluations');
  const evals = evalSheet.getDataRange().getValues().slice(1).filter(r => r[3] === personne);
  
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
  const hasPerfect = evals.some(r => r[27] >= 26);
  if (hasPerfect && !data.badges.some(b => b.id === 'B06')) {
    if (awardBadge(personne, 'B06')) {
      newBadges.push({ id: 'B06', nom: 'Journée parfaite', emoji: '🌟' });
    }
  }
  
  // B08 - Zen master (5x gestion émotions = 2)
  const goodGestion = evals.filter(r => r[22] === 2).length;
  if (goodGestion >= 5 && !data.badges.some(b => b.id === 'B08')) {
    if (awardBadge(personne, 'B08')) {
      newBadges.push({ id: 'B08', nom: 'Zen master', emoji: '🧘' });
    }
  }
  
  // B11 - Explorateur émotions (7 jours avec émotions)
  const daysWithEmotions = evals.filter(r => r[16] && r[16] !== '').length;
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
