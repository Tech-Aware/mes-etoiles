// ==================================================
// 🌟 SYSTÈME DE MÉRITE FAMILIAL v4
// Avec section Émotions
// ==================================================

const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const PARIS_TIMEZONE = 'Europe/Paris';
const TASK_COLUMN_ORDER = [
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

const TASK_SHEET_COLUMNS = {
  id: 0,
  nom: 1,
  description: 2,
  emoji: 3,
  sectionId: 4,
  sectionTitle: 5,
  sectionEmoji: 6,
  sectionType: 7,
  sectionOrder: 8,
  taskOrder: 9,
  active: 10
};

const REWARD_SHEET_COLUMNS = {
  id: 0,
  nom: 1,
  emoji: 2,
  cout: 3,
  active: 5
};

const TASKS_SHEET_PREFIX = 'Tâches ';
const REWARDS_SHEET_PREFIX = 'Récompenses ';

function normalizeHeaderKey(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
}

function getCellValue(row, index) {
  if (index === null || index === undefined || index < 0) {
    return '';
  }
  return row[index];
}

function buildTaskColumnsFromHeader(header, contextLabel) {
  if (!header || header.length === 0) {
    Logger.log(`[buildTaskColumnsFromHeader] En-tête introuvable (${contextLabel}). Utilisation des colonnes par défaut.`);
    return { ...TASK_SHEET_COLUMNS };
  }

  const headerMap = {};
  header.forEach((cell, index) => {
    const key = normalizeHeaderKey(cell);
    if (key) {
      headerMap[key] = index;
    }
  });

  const hasNewFormat = headerMap.idtache !== undefined || headerMap.nomtache !== undefined;
  if (hasNewFormat) {
    return {
      id: headerMap.idtache ?? TASK_SHEET_COLUMNS.id,
      nom: headerMap.nomtache ?? TASK_SHEET_COLUMNS.nom,
      description: headerMap.description ?? TASK_SHEET_COLUMNS.description,
      emoji: headerMap.emojitache ?? TASK_SHEET_COLUMNS.emoji,
      sectionId: headerMap.idsection ?? TASK_SHEET_COLUMNS.sectionId,
      sectionTitle: headerMap.titresection ?? TASK_SHEET_COLUMNS.sectionTitle,
      sectionEmoji: headerMap.emojisection ?? TASK_SHEET_COLUMNS.sectionEmoji,
      sectionType: headerMap.typesection ?? TASK_SHEET_COLUMNS.sectionType,
      sectionOrder: headerMap.ordresection ?? TASK_SHEET_COLUMNS.sectionOrder,
      taskOrder: headerMap.ordretache ?? TASK_SHEET_COLUMNS.taskOrder,
      active: headerMap.actif ?? TASK_SHEET_COLUMNS.active
    };
  }

  const hasLegacyFormat = headerMap.categorie !== undefined || headerMap.pointsmax !== undefined;
  if (hasLegacyFormat) {
    return {
      id: headerMap.id ?? 0,
      nom: headerMap.nom ?? 2,
      description: headerMap.description ?? 4,
      emoji: headerMap.emoji ?? 3,
      sectionId: headerMap.categorie ?? 1,
      sectionTitle: headerMap.categorie ?? 1,
      sectionEmoji: headerMap.emoji ?? 3,
      sectionType: headerMap.categorie ?? 1,
      sectionOrder: headerMap.ordre ?? 6,
      taskOrder: headerMap.ordre ?? 6,
      active: headerMap.actif ?? 7
    };
  }

  Logger.log(`[buildTaskColumnsFromHeader] En-têtes non reconnus (${contextLabel}). Utilisation des colonnes par défaut.`);
  return { ...TASK_SHEET_COLUMNS };
}

function buildRewardColumnsFromHeader(header, contextLabel) {
  if (!header || header.length === 0) {
    Logger.log(`[buildRewardColumnsFromHeader] En-tête introuvable (${contextLabel}). Utilisation des colonnes par défaut.`);
    return { ...REWARD_SHEET_COLUMNS };
  }

  const headerMap = {};
  header.forEach((cell, index) => {
    const key = normalizeHeaderKey(cell);
    if (key) {
      headerMap[key] = index;
    }
  });

  const hasNewFormat = headerMap.idrecompense !== undefined || headerMap.nomrecompense !== undefined;
  if (hasNewFormat) {
    return {
      id: headerMap.idrecompense ?? REWARD_SHEET_COLUMNS.id,
      nom: headerMap.nomrecompense ?? REWARD_SHEET_COLUMNS.nom,
      emoji: headerMap.emoji ?? REWARD_SHEET_COLUMNS.emoji,
      cout: headerMap.cout ?? REWARD_SHEET_COLUMNS.cout,
      active: headerMap.actif ?? REWARD_SHEET_COLUMNS.active
    };
  }

  const hasLegacyFormat = headerMap.categorie !== undefined || headerMap.cout !== undefined;
  if (hasLegacyFormat) {
    return {
      id: headerMap.id ?? 0,
      nom: headerMap.nom ?? 1,
      emoji: headerMap.emoji ?? 2,
      cout: headerMap.cout ?? 3,
      active: headerMap.active ?? headerMap.actif ?? 5
    };
  }

  Logger.log(`[buildRewardColumnsFromHeader] En-têtes non reconnus (${contextLabel}). Utilisation des colonnes par défaut.`);
  return { ...REWARD_SHEET_COLUMNS };
}

function buildTasksConfigFromSheet(sheet, contextLabel) {
  const data = sheet.getDataRange().getValues();
  const header = data.shift() || [];
  const columns = buildTaskColumnsFromHeader(header, contextLabel);
  return buildTasksConfigFromRows(data, columns, contextLabel);
}

function getRewardsRowsWithColumns(sheet, contextLabel) {
  const data = sheet.getDataRange().getValues();
  const header = data.shift() || [];
  const columns = buildRewardColumnsFromHeader(header, contextLabel);
  return {
    rows: data,
    columns
  };
}

function buildDefaultTasksConfig() {
  const sections = [
    {
      id: 'corvees',
      title: 'Mes petits travaux',
      emoji: '🧹',
      type: 'corvees',
      order: 1,
      tasks: [
        { id: 'rangerChambre', nom: 'Ranger ma chambre', description: 'Mes affaires sont bien rangées', emoji: '🛏️', order: 1 },
        { id: 'faireLit', nom: 'Faire mon lit', description: 'Ma couette est bien mise', emoji: '🛌', order: 2 },
        { id: 'rangerJouets', nom: 'Ranger mes jouets', description: 'Mes jouets sont à leur place', emoji: '🧸', order: 3 },
        { id: 'aiderTable', nom: 'Aider à table', description: 'Mettre ou débarrasser', emoji: '🍽️', order: 4 }
      ]
    },
    {
      id: 'comportement',
      title: 'Mon comportement',
      emoji: '💛',
      type: 'comportement',
      order: 2,
      tasks: [
        { id: 'ecouter', nom: 'Écouter papa et maman', description: 'J\'écoute quand on me parle', emoji: '👂', order: 1 },
        { id: 'gentilSoeur', nom: 'Gentil avec ma sœur', description: 'On joue bien ensemble', emoji: '👭', order: 2 },
        { id: 'politesse', nom: 'Les mots magiques', description: 'S\'il te plaît, merci, pardon', emoji: '🙏', order: 3 },
        { id: 'pasColere', nom: 'Calme et zen', description: 'Pas de grosse colère', emoji: '😌', order: 4 }
      ]
    },
    {
      id: 'rituels',
      title: 'Mes rituels',
      emoji: '🌅',
      type: 'rituels',
      order: 3,
      tasks: [
        { id: 'dentsMatin', nom: 'Brosser mes dents', description: 'Après le petit-déjeuner', emoji: '🦷', order: 1 },
        { id: 'dentsSoir', nom: 'Brosser mes dents', description: 'Avant le dodo', emoji: '🦷', order: 2 },
        { id: 'habiller', nom: 'M\'habiller tout seul', description: 'Comme un grand !', emoji: '👕', order: 3 },
        { id: 'cartable', nom: 'Préparer mes affaires', description: 'Mon sac est prêt', emoji: '🎒', order: 4 }
      ]
    }
  ];

  const taskIds = [];
  const tasksById = {};
  sections.forEach(section => {
    section.tasks.forEach(task => {
      taskIds.push(task.id);
      tasksById[task.id] = {
        sectionId: section.id,
        sectionType: section.type,
        nom: task.nom
      };
    });
  });

  return {
    sections,
    taskIds,
    tasksById,
    maxPoints: taskIds.length + 1,
    minPoints: -taskIds.length - 1
  };
}

function getTasksConfig(personne) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const personneKey = String(personne || '').trim();
    const personSheetName = `${TASKS_SHEET_PREFIX}${personneKey}`;
    const personSheet = personneKey ? ss.getSheetByName(personSheetName) : null;
    if (personSheet) {
      const configFromPerson = buildTasksConfigFromSheet(personSheet, personSheetName);
      if (configFromPerson) {
        return configFromPerson;
      }
      Logger.log(`[getTasksConfig] Aucune tâche active dans "${personSheetName}".`);
      return buildDefaultTasksConfig();
    }

    const sheet = ss.getSheetByName('Tâches');
    if (!sheet) {
      Logger.log('[getTasksConfig] Feuille "Tâches" introuvable. Utilisation des valeurs par défaut.');
      return buildDefaultTasksConfig();
    }

    const configFromSheet = buildTasksConfigFromSheet(sheet, 'Tâches');
    if (configFromSheet) {
      return configFromSheet;
    }

    Logger.log('[getTasksConfig] Aucune tâche active détectée. Utilisation des valeurs par défaut.');
    return buildDefaultTasksConfig();
  } catch (error) {
    Logger.log(`[getTasksConfig] Erreur lors du chargement : ${error}`);
    return buildDefaultTasksConfig();
  }
}

function buildTasksConfigFromRows(rows, columns, contextLabel) {
  const sectionsMap = {};
  const taskIds = [];
  const tasksById = {};
  let hasRows = false;

  rows.forEach((row, index) => {
    const activeValue = String(getCellValue(row, columns.active) || '').trim();
    if (activeValue && activeValue.toLowerCase() !== 'oui') {
      return;
    }

    const id = String(getCellValue(row, columns.id) || '').trim();
    const nom = String(getCellValue(row, columns.nom) || '').trim();
    if (!id || !nom) {
      Logger.log(`[buildTasksConfigFromRows] Ligne ${index + 2} ignorée (${contextLabel}) : id/nom manquants.`);
      return;
    }

    const rawSectionId = String(getCellValue(row, columns.sectionId) || 'section').trim();
    const sectionId = rawSectionId || 'section';
    const sectionTitle = String(getCellValue(row, columns.sectionTitle) || rawSectionId || 'Mes tâches').trim();
    const sectionEmoji = String(getCellValue(row, columns.sectionEmoji) || getCellValue(row, columns.emoji) || '⭐').trim();
    const sectionType = String(getCellValue(row, columns.sectionType) || sectionId).trim();
    const sectionOrder = Number(getCellValue(row, columns.sectionOrder) || 0);
    const taskOrder = Number(getCellValue(row, columns.taskOrder) || 0);
    const description = String(getCellValue(row, columns.description) || '').trim();
    const emoji = String(getCellValue(row, columns.emoji) || '⭐').trim();

    if (!sectionsMap[sectionId]) {
      sectionsMap[sectionId] = {
        id: sectionId,
        title: sectionTitle,
        emoji: sectionEmoji,
        type: sectionType,
        order: sectionOrder,
        tasks: []
      };
    }

    sectionsMap[sectionId].tasks.push({
      id,
      nom,
      description,
      emoji,
      order: taskOrder
    });

    taskIds.push(id);
    tasksById[id] = {
      sectionId,
      sectionType,
      nom
    };
    hasRows = true;
  });

  if (!hasRows || taskIds.length === 0) {
    return null;
  }

  const sections = Object.values(sectionsMap)
    .sort((a, b) => a.order - b.order)
    .map(section => ({
      ...section,
      tasks: section.tasks.sort((a, b) => a.order - b.order)
    }));

  return {
    sections,
    taskIds,
    tasksById,
    maxPoints: taskIds.length + 1,
    minPoints: -taskIds.length - 1
  };
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

function normalizeSectionType(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
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
    const tasksConfig = getTasksConfig(personne);
    const activeTaskIds = tasksConfig.taskIds || [];
    const tasksById = tasksConfig.tasksById || {};
    
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

    // Calculs totaux dynamiques
    const totalsByType = { corvees: 0, comportement: 0, rituels: 0 };
    let totalTasks = 0;
    const invalidTasks = [];
    const taskScores = {};

    activeTaskIds.forEach(taskId => {
      const value = Number(taches[taskId]);
      if (![-1, 0, 1].includes(value)) {
        invalidTasks.push(taskId);
        return;
      }

      taskScores[taskId] = value;
      totalTasks += value;

      const sectionType = normalizeSectionType(tasksById[taskId]?.sectionType || tasksById[taskId]?.sectionId);
      if (sectionType === 'corvees') totalsByType.corvees += value;
      if (sectionType === 'comportement') totalsByType.comportement += value;
      if (sectionType === 'rituels') totalsByType.rituels += value;
    });

    if (invalidTasks.length > 0) {
      Logger.log(`[submitEvaluation] Notes invalides pour ${personne} : ${invalidTasks.join(', ')}`);
      return { success: false, message: 'Merci de remplir toutes les tâches avant de valider.' };
    }

    const totalEmotions = Number(emotions.gestion);
    if (Number.isNaN(totalEmotions)) {
      Logger.log(`[submitEvaluation] Score émotion invalide pour ${personne} : ${emotions.gestion}`);
      return { success: false, message: 'Merci de sélectionner la gestion des émotions.' };
    }
    const totalJour = totalTasks + totalEmotions;
    const totalCorvees = totalsByType.corvees;
    const totalComportement = totalsByType.comportement;
    const totalRituels = totalsByType.rituels;

    Logger.log(`[submitEvaluation] Totaux calculés pour ${personne} : corvées=${totalCorvees}, comportement=${totalComportement}, rituels=${totalRituels}, émotions=${totalEmotions}, totalJour=${totalJour}.`);
    
    Logger.log(`[submitEvaluation] Ajout évaluation ${newId} pour ${personne} (Paris=${getParisDateKey(now)}).`);
    
    // Ajouter la ligne
    sheet.appendRow([
      newId,
      now,
      Utilities.formatDate(now, PARIS_TIMEZONE, 'HH:mm'),
      personne,
      // Tâches
      ...TASK_COLUMN_ORDER.map(taskId => taskScores[taskId] ?? ''),
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

    const extraTaskIds = activeTaskIds.filter(taskId => !TASK_COLUMN_ORDER.includes(taskId));
    if (extraTaskIds.length > 0) {
      saveEvaluationTasks(newId, personne, now, taskScores, extraTaskIds);
    }

    // Enregistrer dans historique émotions
    saveEmotionHistory(personne, emotions);
    
    // Vérifier badges
    const newBadges = checkBadges(personne);
    
    // Message selon score
    const maxPoints = tasksConfig.maxPoints || 1;
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

function saveEvaluationTasks(evaluationId, personne, date, taskScores, taskIds) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    let sheet = ss.getSheetByName('Évaluations_Tâches');
    if (!sheet) {
      sheet = ss.insertSheet('Évaluations_Tâches');
      sheet.appendRow(['EvaluationID', 'Date', 'Personne', 'TaskID', 'Score']);
      Logger.log('[saveEvaluationTasks] Feuille "Évaluations_Tâches" créée.');
    }

    taskIds.forEach(taskId => {
      sheet.appendRow([
        evaluationId,
        date,
        personne,
        taskId,
        taskScores[taskId]
      ]);
    });
  } catch (error) {
    Logger.log(`[saveEvaluationTasks] Erreur lors de l’enregistrement des tâches extra : ${error}`);
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
    const rewards = getRewardsForPerson(personneKey, weekPoints);
    
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

function getRewardsForPerson(personneKey, weekPoints) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const personSheetName = `${REWARDS_SHEET_PREFIX}${personneKey}`;
  const rewardsPersonSheet = personneKey ? ss.getSheetByName(personSheetName) : null;
  if (rewardsPersonSheet) {
    const rewardsData = getRewardsRowsWithColumns(rewardsPersonSheet, personSheetName);
    const rows = rewardsData.rows;
    const columns = rewardsData.columns;
    const rewards = rows
      .filter(row => String(getCellValue(row, columns.active) || '').trim().toLowerCase() === 'oui')
      .map(row => ({
        id: getCellValue(row, columns.id),
        nom: getCellValue(row, columns.nom),
        emoji: getCellValue(row, columns.emoji),
        cout: getCellValue(row, columns.cout),
        disponible: weekPoints >= getCellValue(row, columns.cout)
      }));

    Logger.log(`[getRewardsForPerson] Récompenses chargées depuis "${personSheetName}" : ${rewards.length}.`);
    return rewards;
  }

  const rewardsSheet = ss.getSheetByName('Récompenses');
  if (!rewardsSheet) {
    Logger.log('[getRewardsForPerson] Feuille "Récompenses" introuvable.');
    return [];
  }

  const rewardsData = getRewardsRowsWithColumns(rewardsSheet, 'Récompenses');
  const rows = rewardsData.rows;
  const columns = rewardsData.columns;
  return rows
    .filter(row => String(getCellValue(row, columns.active) || '').trim().toLowerCase() === 'oui')
    .map(row => ({
      id: getCellValue(row, columns.id),
      nom: getCellValue(row, columns.nom),
      emoji: getCellValue(row, columns.emoji),
      cout: getCellValue(row, columns.cout),
      disponible: weekPoints >= getCellValue(row, columns.cout)
    }));
}

function findRewardForPerson(personneKey, rewardId) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const personSheetName = `${REWARDS_SHEET_PREFIX}${personneKey}`;
  const rewardsPersonSheet = personneKey ? ss.getSheetByName(personSheetName) : null;
  if (rewardsPersonSheet) {
    const rewardsData = getRewardsRowsWithColumns(rewardsPersonSheet, personSheetName);
    const rows = rewardsData.rows;
    const columns = rewardsData.columns;
    const reward = rows.find(row => String(getCellValue(row, columns.id) || '').trim() === rewardId);
    return reward
      ? {
          id: getCellValue(reward, columns.id),
          nom: getCellValue(reward, columns.nom),
          emoji: getCellValue(reward, columns.emoji),
          cout: getCellValue(reward, columns.cout),
          active: String(getCellValue(reward, columns.active) || '').trim().toLowerCase() === 'oui'
        }
      : null;
  }

  const rewardsSheet = ss.getSheetByName('Récompenses');
  if (!rewardsSheet) {
    return null;
  }

  const rewardsData = getRewardsRowsWithColumns(rewardsSheet, 'Récompenses');
  const rows = rewardsData.rows;
  const columns = rewardsData.columns;
  const reward = rows.find(row => String(getCellValue(row, columns.id) || '').trim() === rewardId);
  return reward
    ? {
        id: getCellValue(reward, columns.id),
        nom: getCellValue(reward, columns.nom),
        emoji: getCellValue(reward, columns.emoji),
        cout: getCellValue(reward, columns.cout),
        active: String(getCellValue(reward, columns.active) || '').trim().toLowerCase() === 'oui'
      }
    : null;
}

function ensureRewardsClaimsSheet() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName('Récompenses_Demandées');
  if (!sheet) {
    sheet = ss.insertSheet('Récompenses_Demandées');
    sheet.appendRow(['ID', 'Date', 'Personne', 'Récompense', 'Coût', 'Statut', 'Note', 'Validation']);
    Logger.log('[ensureRewardsClaimsSheet] Feuille "Récompenses_Demandées" créée avec colonne Personne.');
    return sheet;
  }

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
  if (header.length < 3 || String(header[2] || '').trim().toLowerCase() !== 'personne') {
    Logger.log('[ensureRewardsClaimsSheet] Colonne "Personne" absente ou mal positionnée (colonne C).');
  }
  return sheet;
}

// ==================================================
// RÉCLAMER RÉCOMPENSE
// ==================================================
function claimReward(personne, rewardId) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const data = getPersonneData(personne);

  const reward = findRewardForPerson(String(personne || '').trim(), String(rewardId || '').trim());
  if (!reward) {
    return { success: false, message: 'Récompense introuvable 😕' };
  }

  if (!reward.active) {
    Logger.log(`[claimReward] Récompense inactive : ${rewardId} pour ${personne}.`);
    return { success: false, message: 'Récompense indisponible pour le moment.' };
  }
  
  if (data.weekPoints < reward.cout) {
    return { success: false, message: `Il te manque ${reward.cout - data.weekPoints} étoiles 😢` };
  }

  const claimsSheet = ensureRewardsClaimsSheet();
  const newId = 'C' + String(claimsSheet.getLastRow()).padStart(4, '0');
  
  claimsSheet.appendRow([
    newId,
    new Date(),
    personne,
    reward.nom,
    reward.cout,
    'En attente',
    '',
    ''
  ]);
  
  return { 
    success: true, 
    message: `🎉 Super ! "${reward.nom}" demandé !`
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
