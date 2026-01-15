// ==================================================
// MES ETOILES - SYSTEME DE MERITE FAMILIAL v5
// Architecture Code-First avec auto-synchronisation
// ==================================================

const PARIS_TIMEZONE = 'Europe/Paris';

// ==================================================
// CONFIGURATION DES SCORES
// ==================================================
const SCORES = {
  // Valeurs stockees pour les taches
  TASK_MIN: -1,
  TASK_MAX: 3,
  TASK_ALLOWED: [-1, 0, 1, 2, 3],
  // Conversion etoiles affichees → points reels
  // 0 etoiles (penalite) → -1 point
  // 1 etoile → 0 point
  // 2 etoiles → 1 point
  // 3 etoiles → 2 points
  // Formule: points = etoiles - 1 (minimum -1)
  TASK_TO_POINTS: { '-1': -1, '0': -1, '1': 0, '2': 1, '3': 2 },
  // Points max par tache = 2 (quand 3 etoiles)
  TASK_POINTS_MAX: 2,
  // Gestion des emotions
  GESTION_MIN: -1,
  GESTION_MAX: 1,
  GESTION_ALLOWED: [-1, 0, 1]
};

// ==================================================
// SCHEMAS DES FEUILLES (SOURCE DE VERITE)
// ==================================================
const SCHEMAS = {
  Personnes: {
    headers: ['Nom', 'Avatar', 'Couleur', 'Age'],
    required: true
  },
  Taches: {
    headers: ['ID', 'Categorie', 'Nom', 'Emoji', 'Description', 'PointsMax', 'Ordre', 'Personnes', 'Jours'],
    required: true
  },
  Evaluations: {
    headers: ['ID', 'Date', 'Heure', 'Personne', 'Emotion1', 'Emotion2', 'Emotion3', 'Source1', 'Source2', 'Source3', 'GestionEmotion', 'TotalJour', 'Humeur', 'Commentaire'],
    required: true,
    dynamic: true // Colonnes de taches ajoutees dynamiquement
  },
  Recompenses: {
    headers: ['ID', 'Nom', 'Emoji', 'Cout', 'Description', 'Actif'],
    required: true
  },
  Recompenses_Demandees: {
    headers: ['ID', 'Date', 'Personne', 'Recompense', 'Cout', 'Statut', 'ValidePar', 'Commentaire'],
    required: true
  },
  Badges: {
    headers: ['ID', 'Nom', 'Emoji', 'Description', 'Condition'],
    required: true
  },
  Badges_Obtenus: {
    headers: ['Personne', 'BadgeID', 'Date'],
    required: true
  }
};

// ==================================================
// JOURS DE LA SEMAINE
// ==================================================
const JOURS_MAP = {
  lun: 1, lundi: 1,
  mar: 2, mardi: 2,
  mer: 3, mercredi: 3,
  jeu: 4, jeudi: 4,
  ven: 5, vendredi: 5,
  sam: 6, samedi: 6,
  dim: 7, dimanche: 7,
  '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6, '7': 7
};

// ==================================================
// UTILITAIRES
// ==================================================

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function normaliserTexte_(valeur) {
  if (!valeur) return '';
  return String(valeur)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function getParisDateKey_(date) {
  return Utilities.formatDate(date, PARIS_TIMEZONE, 'yyyy-MM-dd');
}

function getParisMidnight_(date) {
  const [year, month, day] = getParisDateKey_(date).split('-').map(Number);
  return new Date(year, month - 1, day);
}

function getParisDayIndex_(date) {
  const dayIndex = Number(Utilities.formatDate(date, PARIS_TIMEZONE, 'u'));
  if (Number.isNaN(dayIndex)) {
    const fallback = date.getDay();
    return fallback === 0 ? 7 : fallback;
  }
  return dayIndex;
}

function parseDate_(value) {
  if (value instanceof Date) return value;
  if (typeof value === 'number') return new Date(value);
  if (typeof value === 'string') {
    const match = value.trim().match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/);
    if (match) {
      const [, day, month, year, hour = '00', minute = '00', second = '00'] = match;
      return new Date(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second));
    }
  }
  const fallback = new Date(value);
  return Number.isNaN(fallback.getTime()) ? null : fallback;
}

function getMonday_(date) {
  const d = getParisMidnight_(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  d.setDate(diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

// ==================================================
// GESTION DES POINTS (SCRIPT PROPERTIES)
// ==================================================
// Les points sont stockes dans les Script Properties
// Cle: "points_<NomPersonne>" (normalise)
// Valeur: nombre entier (total des points disponibles)

function getPointsPropertyKey_(personne) {
  return 'points_' + normaliserTexte_(personne);
}

function getPointsProperty_(personne) {
  const props = PropertiesService.getScriptProperties();
  const key = getPointsPropertyKey_(personne);
  const value = props.getProperty(key);
  if (value === null) {
    // Premiere utilisation: recalculer depuis les feuilles
    const calculated = recalculerPointsDepuisFeuilles_(personne);
    props.setProperty(key, String(calculated));
    Logger.log(`[getPointsProperty] Initialisation ${personne}: ${calculated} pts`);
    return calculated;
  }
  return Number(value) || 0;
}

function setPointsProperty_(personne, points) {
  const props = PropertiesService.getScriptProperties();
  const key = getPointsPropertyKey_(personne);
  const safePoints = Math.max(0, Math.round(points));
  props.setProperty(key, String(safePoints));
  Logger.log(`[setPointsProperty] ${personne} = ${safePoints} pts`);
  return safePoints;
}

function addPointsProperty_(personne, delta) {
  const current = getPointsProperty_(personne);
  const newValue = Math.max(0, current + delta);
  setPointsProperty_(personne, newValue);
  Logger.log(`[addPointsProperty] ${personne}: ${current} + (${delta}) = ${newValue} pts`);
  return newValue;
}

function recalculerPointsDepuisFeuilles_(personne) {
  const personneKey = String(personne || '').trim();

  // Total gagne = somme des TotalJour
  let totalGagnes = 0;
  try {
    const { rows, headerIndex } = getFeuilleAvecHeaders_('Evaluations');
    const persIdx = headerIndex['Personne'] ?? 3;
    const totalIdx = headerIndex['TotalJour'] ?? 11;

    rows.forEach(row => {
      if (String(row[persIdx] || '').trim() === personneKey) {
        totalGagnes += Number(row[totalIdx]) || 0;
      }
    });
  } catch (error) {
    Logger.log(`[recalculerPoints] Erreur evaluations: ${error}`);
  }

  // Total depense = somme des couts (hors annule/refuse)
  let totalDepenses = 0;
  try {
    const { rows, headerIndex } = getFeuilleAvecHeaders_('Recompenses_Demandees');
    const persIdx = headerIndex['Personne'] ?? 2;
    const coutIdx = headerIndex['Cout'] ?? 4;
    const statutIdx = headerIndex['Statut'] ?? 5;

    rows.forEach(row => {
      if (String(row[persIdx] || '').trim() !== personneKey) return;
      const statut = normaliserTexte_(row[statutIdx]);
      if (statut === 'annule' || statut === 'refuse' || statut === 'refusee') return;
      totalDepenses += Number(row[coutIdx]) || 0;
    });
  } catch (error) {
    Logger.log(`[recalculerPoints] Erreur recompenses: ${error}`);
  }

  return Math.max(0, totalGagnes - totalDepenses);
}

// Fonction utilitaire pour reinitialiser tous les points (admin)
function reinitialiserTousLesPoints() {
  const props = PropertiesService.getScriptProperties();
  const personnes = getPersonnes();
  const resultats = [];

  personnes.forEach(p => {
    const key = getPointsPropertyKey_(p.nom);
    const calculated = recalculerPointsDepuisFeuilles_(p.nom);
    props.setProperty(key, String(calculated));
    resultats.push({ nom: p.nom, points: calculated });
    Logger.log(`[reinitialiserPoints] ${p.nom}: ${calculated} pts`);
  });

  return resultats;
}

// ==================================================
// SYNCHRONISATION DES FEUILLES
// ==================================================

function synchroniserFeuilles_() {
  const ss = getSpreadsheet_();
  const resultats = [];

  Object.entries(SCHEMAS).forEach(([nomFeuille, schema]) => {
    try {
      let sheet = ss.getSheetByName(nomFeuille);

      // Creer la feuille si elle n'existe pas
      if (!sheet) {
        sheet = ss.insertSheet(nomFeuille);
        Logger.log(`[sync] Feuille "${nomFeuille}" creee.`);
        resultats.push({ feuille: nomFeuille, action: 'creee' });
      }

      // Verifier et corriger les en-tetes
      const lastCol = Math.max(sheet.getLastColumn(), 1);
      const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
        .map(h => String(h || '').trim())
        .filter(Boolean);

      if (existingHeaders.length === 0) {
        // Feuille vide, ecrire tous les en-tetes
        sheet.getRange(1, 1, 1, schema.headers.length).setValues([schema.headers]);
        Logger.log(`[sync] En-tetes initialises pour "${nomFeuille}".`);
        resultats.push({ feuille: nomFeuille, action: 'headers_init' });
      } else {
        // Verifier les en-tetes manquants
        const existingNormalized = new Set(existingHeaders.map(h => normaliserTexte_(h)));
        let colIndex = existingHeaders.length + 1;

        schema.headers.forEach(header => {
          if (!existingNormalized.has(normaliserTexte_(header))) {
            sheet.getRange(1, colIndex).setValue(header);
            Logger.log(`[sync] En-tete "${header}" ajoute a "${nomFeuille}" (col ${colIndex}).`);
            colIndex++;
          }
        });
      }

      // Synchroniser les colonnes de taches pour Evaluations
      if (nomFeuille === 'Evaluations' && schema.dynamic) {
        synchroniserColonnesTaches_(sheet);
      }

    } catch (error) {
      Logger.log(`[sync] Erreur pour "${nomFeuille}": ${error}`);
      resultats.push({ feuille: nomFeuille, action: 'erreur', message: error.toString() });
    }
  });

  return resultats;
}

function synchroniserColonnesTaches_(sheet) {
  const taches = getTachesDefinitions_();
  if (taches.length === 0) return;

  const lastCol = sheet.getLastColumn();
  const headers = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
    : [];
  const headersNormalized = new Set(headers.map(h => normaliserTexte_(h)));

  let colIndex = headers.length + 1;
  taches.forEach(tache => {
    const header = tache.id;
    if (!headersNormalized.has(normaliserTexte_(header))) {
      sheet.getRange(1, colIndex).setValue(header);
      Logger.log(`[sync] Colonne tache "${header}" ajoutee a Evaluations (col ${colIndex}).`);
      colIndex++;
    }
  });
}

function getFeuilleAvecHeaders_(nomFeuille) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(nomFeuille);
  if (!sheet) {
    throw new Error(`Feuille "${nomFeuille}" introuvable.`);
  }

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    return { sheet, headers: [], rows: [], headerIndex: {} };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h || '').trim());
  const headerIndex = headers.reduce((acc, h, i) => { acc[h] = i; return acc; }, {});
  const rows = data.slice(1);

  return { sheet, headers, rows, headerIndex };
}

// ==================================================
// WEB APP
// ==================================================

function doGet(e) {
  // Synchroniser les feuilles a chaque ouverture
  try {
    synchroniserFeuilles_();
  } catch (error) {
    Logger.log(`[doGet] Erreur sync: ${error}`);
  }

  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Mes Etoiles')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==================================================
// PERSONNES
// ==================================================

function getPersonnes() {
  const { rows } = getFeuilleAvecHeaders_('Personnes');
  return rows
    .filter(row => row[0])
    .map(row => ({
      nom: String(row[0] || '').trim(),
      avatar: String(row[1] || '').trim() || '👤',
      couleur: String(row[2] || '').trim() || '#6C5CE7',
      age: Number(row[3]) || 0
    }));
}

// ==================================================
// TACHES
// ==================================================

function getTachesDefinitions_() {
  try {
    const { rows, headerIndex } = getFeuilleAvecHeaders_('Taches');
    const idIdx = headerIndex['ID'] ?? 0;
    const catIdx = headerIndex['Categorie'] ?? 1;
    const nomIdx = headerIndex['Nom'] ?? 2;
    const emojiIdx = headerIndex['Emoji'] ?? 3;
    const descIdx = headerIndex['Description'] ?? 4;
    const maxIdx = headerIndex['PointsMax'] ?? 5;
    const ordreIdx = headerIndex['Ordre'] ?? 6;
    const persIdx = headerIndex['Personnes'] ?? 7;
    const joursIdx = headerIndex['Jours'] ?? 8;

    return rows
      .filter(row => row[idIdx])
      .map(row => ({
        id: String(row[idIdx] || '').trim(),
        categorie: normaliserCategorie_(String(row[catIdx] || '').trim()),
        nom: String(row[nomIdx] || '').trim() || String(row[idIdx] || '').trim(),
        emoji: String(row[emojiIdx] || '').trim() || '✨',
        description: String(row[descIdx] || '').trim(),
        pointsMax: Number(row[maxIdx]) || SCORES.TASK_MAX,
        ordre: Number(row[ordreIdx]) || 999,
        personnes: String(row[persIdx] || '').trim(),
        jours: String(row[joursIdx] || '').trim()
      }))
      .sort((a, b) => a.ordre - b.ordre);
  } catch (error) {
    Logger.log(`[getTachesDefinitions] Erreur: ${error}`);
    return [];
  }
}

function normaliserCategorie_(categorie) {
  const n = normaliserTexte_(categorie);
  if (n.includes('corvee') || n.includes('travaux')) return 'corvees';
  if (n.includes('comportement')) return 'comportement';
  if (n.includes('rituel')) return 'rituels';
  return 'autres';
}

function parseJours_(rawValue) {
  if (!rawValue) return null;
  const value = String(rawValue).toLowerCase().trim();
  if (!value) return null;

  // Valeurs speciales
  if (['tous', 'toute', 'toutes', 'toute la semaine', '7/7', 'tous les jours'].includes(value)) {
    return new Set([1, 2, 3, 4, 5, 6, 7]);
  }
  if (['week-end', 'weekend'].includes(value)) {
    return new Set([6, 7]);
  }
  if (['semaine', 'lun-ven', 'en semaine'].includes(value)) {
    return new Set([1, 2, 3, 4, 5]);
  }

  // Parser les jours individuels
  const daySet = new Set();
  value.split(/[,;/\n]+/).forEach(part => {
    const normalized = part.trim().replace(/\s+/g, '');
    if (normalized.includes('-')) {
      const [startRaw, endRaw] = normalized.split('-');
      const start = JOURS_MAP[startRaw];
      const end = JOURS_MAP[endRaw];
      if (start && end) {
        for (let d = start; d <= (end >= start ? end : 7); d++) daySet.add(d);
        if (end < start) for (let d = 1; d <= end; d++) daySet.add(d);
      }
    } else {
      const mapped = JOURS_MAP[normalized];
      if (mapped) daySet.add(mapped);
    }
  });

  return daySet.size > 0 ? daySet : null;
}

function estTacheDisponibleAujourdhui_(tache) {
  if (!tache.jours) return true;
  const jours = parseJours_(tache.jours);
  if (!jours) return true;
  return jours.has(getParisDayIndex_(new Date()));
}

function estTacheAssigneePourPersonne_(tache, personne) {
  if (!tache.personnes) return true;
  const assignees = tache.personnes.split(/[,;\n]+/).map(s => s.trim()).filter(Boolean);
  if (assignees.length === 0) return true;
  return assignees.includes(personne);
}

function getTachesPourPersonne(personne) {
  const personneKey = String(personne || '').trim();
  const taches = getTachesDefinitions_();

  const tachesFiltrees = taches.filter(t =>
    estTacheAssigneePourPersonne_(t, personneKey) &&
    estTacheDisponibleAujourdhui_(t)
  );

  return {
    personne: personneKey,
    tasks: tachesFiltrees,
    taskIds: tachesFiltrees.map(t => t.id)
  };
}

// ==================================================
// EVALUATIONS
// ==================================================

function hasEvaluatedToday(personne) {
  try {
    const { rows, headerIndex } = getFeuilleAvecHeaders_('Evaluations');
    const dateIdx = headerIndex['Date'] ?? 1;
    const persIdx = headerIndex['Personne'] ?? 3;
    const todayKey = getParisDateKey_(new Date());
    const personneKey = String(personne || '').trim();

    return rows.some(row => {
      const rowDate = parseDate_(row[dateIdx]);
      if (!rowDate) return false;
      const rowKey = getParisDateKey_(rowDate);
      return String(row[persIdx] || '').trim() === personneKey && rowKey === todayKey;
    });
  } catch (error) {
    Logger.log(`[hasEvaluatedToday] Erreur: ${error}`);
    return false;
  }
}

function submitEvaluation(personne, taches, emotions, humeur, commentaire) {
  try {
    // Verifier si deja evalue
    if (hasEvaluatedToday(personne)) {
      return { success: false, message: 'Tu as deja fait ton evaluation aujourd\'hui !' };
    }

    // Verifier les sources pour chaque emotion
    const emotionPairs = [
      { emotion: emotions.emotion1, source: emotions.source1 },
      { emotion: emotions.emotion2, source: emotions.source2 },
      { emotion: emotions.emotion3, source: emotions.source3 }
    ];
    const missingSources = emotionPairs.filter(p => p.emotion && !p.source);
    if (missingSources.length > 0) {
      return { success: false, message: 'Choisis une cause pour chaque emotion.' };
    }

    // Recuperer les taches assignees
    const { tasks: assignedTasks, taskIds } = getTachesPourPersonne(personne);
    const assignedSet = new Set(taskIds);

    // Valider et calculer les scores
    const safeTaskValue = (taskId) => {
      if (!assignedSet.has(taskId)) return 0;
      const value = Number(taches && taches[taskId]);
      if (Number.isNaN(value) || !SCORES.TASK_ALLOWED.includes(value)) return 0;
      return value;
    };

    const gestionValue = SCORES.GESTION_ALLOWED.includes(emotions.gestion) ? emotions.gestion : 0;

    // Calculer le total en convertissant etoiles → points
    // Formule: 0 etoiles=-1pt, 1 etoile=0pt, 2 etoiles=1pt, 3 etoiles=2pts
    let totalJour = gestionValue;
    const taskScores = {};
    taskIds.forEach(taskId => {
      const stars = safeTaskValue(taskId);
      taskScores[taskId] = stars; // On stocke les etoiles
      const points = SCORES.TASK_TO_POINTS[String(stars)] ?? (stars - 1);
      totalJour += points;
    });

    // Preparer la ligne
    const { sheet, headers, headerIndex } = getFeuilleAvecHeaders_('Evaluations');
    const now = new Date();
    const newId = 'E' + String(sheet.getLastRow()).padStart(4, '0');

    const commentaireSafe = String(commentaire || '').trim().slice(0, 400);

    // Construire la ligne
    const rowValues = new Array(headers.length).fill('');
    const setValue = (header, value) => {
      const idx = headerIndex[header];
      if (typeof idx === 'number') rowValues[idx] = value;
    };

    setValue('ID', newId);
    setValue('Date', now);
    setValue('Heure', Utilities.formatDate(now, PARIS_TIMEZONE, 'HH:mm'));
    setValue('Personne', personne);
    setValue('Emotion1', emotions.emotion1 || '');
    setValue('Emotion2', emotions.emotion2 || '');
    setValue('Emotion3', emotions.emotion3 || '');
    setValue('Source1', emotions.source1 || '');
    setValue('Source2', emotions.source2 || '');
    setValue('Source3', emotions.source3 || '');
    setValue('GestionEmotion', gestionValue);
    setValue('TotalJour', totalJour);
    setValue('Humeur', humeur || '');
    setValue('Commentaire', commentaireSafe);

    // Ajouter les scores de chaque tache
    Object.entries(taskScores).forEach(([taskId, score]) => {
      setValue(taskId, score);
    });

    // Ecrire la ligne
    sheet.appendRow(rowValues);
    Logger.log(`[submitEvaluation] Evaluation ${newId} ajoutee pour ${personne}, total=${totalJour}.`);

    // Mettre a jour les points dans Script Properties
    const newTotalPoints = addPointsProperty_(personne, totalJour);
    Logger.log(`[submitEvaluation] Points mis a jour: ${newTotalPoints} pts`);

    // Verifier les badges
    const newBadges = checkBadges_(personne);

    // Calculer le pourcentage et le message
    // Max = (nb taches * 2 points) + 1 point gestion
    const maxPoints = taskIds.length * SCORES.TASK_POINTS_MAX + SCORES.GESTION_MAX;
    const percent = maxPoints > 0 ? Math.max(0, Math.round((totalJour / maxPoints) * 100)) : 0;

    let message, stars;
    if (percent >= 90) { message = 'INCROYABLE ! Tu es une vraie STAR !'; stars = 5; }
    else if (percent >= 75) { message = 'SUPER journee ! Bravo champion !'; stars = 4; }
    else if (percent >= 60) { message = 'Bien joue ! Continue comme ca !'; stars = 3; }
    else if (percent >= 40) { message = 'Pas mal ! Tu peux faire encore mieux !'; stars = 2; }
    else { message = 'Demain sera meilleur ! On y croit !'; stars = 1; }

    return {
      success: true,
      message,
      totalJour,
      maxJour: maxPoints,
      percent,
      stars,
      newBadges
    };

  } catch (error) {
    Logger.log(`[submitEvaluation] Erreur: ${error}`);
    return { success: false, message: 'Erreur lors de l\'enregistrement.' };
  }
}

// ==================================================
// CALCUL DES POINTS
// ==================================================

function calculerPoints_(personne) {
  // Utilise les Script Properties pour performance
  // Les points sont mis a jour incrementalement lors des evaluations/recompenses
  const totalPoints = getPointsProperty_(personne);
  return { totalPoints };
}

// ==================================================
// DONNEES PERSONNE
// ==================================================

function getPersonneData(personne) {
  try {
    const personneKey = String(personne || '').trim();
    const personnes = getPersonnes();
    const personneInfo = personnes.find(p => p.nom === personneKey);

    // Points
    const points = calculerPoints_(personneKey);

    // Taches assignees
    const { tasks: assignedTasks } = getTachesPourPersonne(personneKey);
    const maxPointsJour = assignedTasks.length * SCORES.TASK_MAX + SCORES.GESTION_MAX;

    // Evaluations de la semaine
    const { rows, headerIndex } = getFeuilleAvecHeaders_('Evaluations');
    const dateIdx = headerIndex['Date'] ?? 1;
    const persIdx = headerIndex['Personne'] ?? 3;
    const totalIdx = headerIndex['TotalJour'] ?? 11;
    const emo1Idx = headerIndex['Emotion1'] ?? 4;
    const emo2Idx = headerIndex['Emotion2'] ?? 5;
    const emo3Idx = headerIndex['Emotion3'] ?? 6;
    const gestionIdx = headerIndex['GestionEmotion'] ?? 10;

    const todayParis = getParisMidnight_(new Date());
    const weekStart = getMonday_(todayParis);
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekEnd.getDate() + 6);
    weekEnd.setHours(23, 59, 59);

    let weekPoints = 0;
    let weekDays = 0;
    const dailyScores = [null, null, null, null, null, null, null];

    const personneEvals = rows.filter(r => String(r[persIdx] || '').trim() === personneKey);

    personneEvals.forEach(row => {
      const parsedDate = parseDate_(row[dateIdx]);
      if (!parsedDate) return;
      const date = getParisMidnight_(parsedDate);
      if (date >= weekStart && date <= weekEnd) {
        const total = Number(row[totalIdx]) || 0;
        weekPoints += total;
        weekDays++;
        const dayIndex = date.getDay() === 0 ? 6 : date.getDay() - 1;
        dailyScores[dayIndex] = total;
      }
    });

    // Streak
    let streak = 0;
    const sortedEvals = personneEvals.sort((a, b) => {
      const dateA = parseDate_(a[dateIdx]);
      const dateB = parseDate_(b[dateIdx]);
      return (dateB?.getTime() || 0) - (dateA?.getTime() || 0);
    });

    if (sortedEvals.length > 0) {
      let checkDate = getParisMidnight_(new Date());
      for (const evalRow of sortedEvals) {
        const parsedDate = parseDate_(evalRow[dateIdx]);
        if (!parsedDate) continue;
        const evalDate = getParisMidnight_(parsedDate);
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

    // Emotions recentes (depuis Evaluations)
    const recentEmotions = sortedEvals.slice(0, 7).map(row => {
      const parsedDate = parseDate_(row[dateIdx]);
      return {
        date: parsedDate ? Utilities.formatDate(parsedDate, PARIS_TIMEZONE, 'dd/MM') : '-',
        emotion1: row[emo1Idx] || '',
        emotion2: row[emo2Idx] || '',
        emotion3: row[emo3Idx] || '',
        gestion: Number(row[gestionIdx]) || 0
      };
    });

    // Badges
    const badges = getBadgesObtenus_(personneKey);

    // Recompenses
    const rewards = getRecompensesDisponibles_(points.totalPoints);

    return {
      nom: personneKey,
      avatar: personneInfo?.avatar || '👤',
      couleur: personneInfo?.couleur || '#6C5CE7',
      age: personneInfo?.age || 0,
      weekPoints,
      totalPoints: points.totalPoints,
      totalEarned: points.totalGagnes,
      totalSpent: points.totalDepenses,
      weekDays,
      dailyScores,
      streak,
      recentEmotions,
      badges,
      rewards,
      maxPointsJour,
      evaluatedToday: hasEvaluatedToday(personneKey),
      weekStart: Utilities.formatDate(weekStart, PARIS_TIMEZONE, 'dd/MM'),
      weekEnd: Utilities.formatDate(weekEnd, PARIS_TIMEZONE, 'dd/MM')
    };

  } catch (error) {
    Logger.log(`[getPersonneData] Erreur: ${error}`);
    throw new Error('Impossible de charger les donnees.');
  }
}

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
// RECOMPENSES
// ==================================================

function getRecompensesDisponibles_(pointsDisponibles) {
  try {
    const { rows, headerIndex } = getFeuilleAvecHeaders_('Recompenses');
    const idIdx = headerIndex['ID'] ?? 0;
    const nomIdx = headerIndex['Nom'] ?? 1;
    const emojiIdx = headerIndex['Emoji'] ?? 2;
    const coutIdx = headerIndex['Cout'] ?? 3;
    const actifIdx = headerIndex['Actif'] ?? 5;

    return rows
      .filter(row => {
        const actif = String(row[actifIdx] || '').trim().toLowerCase();
        return actif === 'oui' || actif === 'true' || actif === '1';
      })
      .map(row => {
        const cout = Number(row[coutIdx]) || 0;
        return {
          id: String(row[idIdx] || '').trim(),
          nom: String(row[nomIdx] || '').trim(),
          emoji: String(row[emojiIdx] || '').trim() || '🎁',
          cout,
          disponible: pointsDisponibles >= cout
        };
      });
  } catch (error) {
    Logger.log(`[getRecompensesDisponibles] Erreur: ${error}`);
    return [];
  }
}

function claimReward(personne, rewardId) {
  try {
    const personneKey = String(personne || '').trim();
    const points = calculerPoints_(personneKey);

    // Trouver la recompense
    const { rows: rewardsRows, headerIndex: rewardsIndex } = getFeuilleAvecHeaders_('Recompenses');
    const reward = rewardsRows.find(r => String(r[rewardsIndex['ID'] ?? 0] || '').trim() === rewardId);

    if (!reward) {
      return { success: false, message: 'Recompense introuvable.' };
    }

    const rewardNom = String(reward[rewardsIndex['Nom'] ?? 1] || '').trim();
    const rewardCout = Number(reward[rewardsIndex['Cout'] ?? 3]) || 0;

    if (points.totalPoints < rewardCout) {
      return { success: false, message: `Il te manque ${rewardCout - points.totalPoints} etoiles.` };
    }

    // Enregistrer la demande
    const { sheet, headers, headerIndex } = getFeuilleAvecHeaders_('Recompenses_Demandees');
    const newId = 'C' + String(sheet.getLastRow()).padStart(4, '0');

    const rowValues = new Array(headers.length).fill('');
    const setValue = (header, value) => {
      const idx = headerIndex[header];
      if (typeof idx === 'number') rowValues[idx] = value;
    };

    setValue('ID', newId);
    setValue('Date', new Date());
    setValue('Personne', personneKey);
    setValue('Recompense', rewardNom);
    setValue('Cout', rewardCout);
    setValue('Statut', 'En attente');
    setValue('ValidePar', '');
    setValue('Commentaire', '');

    sheet.appendRow(rowValues);
    Logger.log(`[claimReward] Demande ${newId} pour ${personneKey}: ${rewardNom} (${rewardCout} pts).`);

    // Soustraire les points du total stocke
    const newTotalPoints = addPointsProperty_(personneKey, -rewardCout);
    Logger.log(`[claimReward] Points apres deduction: ${newTotalPoints} pts`);

    return { success: true, message: `"${rewardNom}" demande !`, newTotalPoints };

  } catch (error) {
    Logger.log(`[claimReward] Erreur: ${error}`);
    return { success: false, message: 'Erreur lors de la demande.' };
  }
}

// ==================================================
// BADGES
// ==================================================

function getBadgesObtenus_(personne) {
  try {
    const { rows: obtRows, headerIndex: obtIndex } = getFeuilleAvecHeaders_('Badges_Obtenus');
    const { rows: defRows, headerIndex: defIndex } = getFeuilleAvecHeaders_('Badges');

    const persIdx = obtIndex['Personne'] ?? 0;
    const badgeIdIdx = obtIndex['BadgeID'] ?? 1;
    const defIdIdx = defIndex['ID'] ?? 0;
    const defNomIdx = defIndex['Nom'] ?? 1;
    const defEmojiIdx = defIndex['Emoji'] ?? 2;

    const badgeIds = obtRows
      .filter(r => String(r[persIdx] || '').trim() === personne)
      .map(r => String(r[badgeIdIdx] || '').trim());

    return badgeIds.map(id => {
      const def = defRows.find(r => String(r[defIdIdx] || '').trim() === id);
      if (!def) return null;
      return {
        id,
        nom: String(def[defNomIdx] || '').trim(),
        emoji: String(def[defEmojiIdx] || '').trim() || '🏆'
      };
    }).filter(Boolean);

  } catch (error) {
    Logger.log(`[getBadgesObtenus] Erreur: ${error}`);
    return [];
  }
}

function checkBadges_(personne) {
  const personneKey = String(personne || '').trim();
  const newBadges = [];
  const existingBadges = getBadgesObtenus_(personneKey).map(b => b.id);

  try {
    const { rows, headerIndex } = getFeuilleAvecHeaders_('Evaluations');
    const persIdx = headerIndex['Personne'] ?? 3;
    const totalIdx = headerIndex['TotalJour'] ?? 11;
    const gestionIdx = headerIndex['GestionEmotion'] ?? 10;
    const emo1Idx = headerIndex['Emotion1'] ?? 4;
    const dateIdx = headerIndex['Date'] ?? 1;

    const evals = rows.filter(r => String(r[persIdx] || '').trim() === personneKey);

    // Calculer les donnees pour les badges
    const todayParis = getParisMidnight_(new Date());
    const weekStart = getMonday_(todayParis);
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekEnd.getDate() + 6);

    let weekDays = 0;
    evals.forEach(row => {
      const parsedDate = parseDate_(row[dateIdx]);
      if (!parsedDate) return;
      const date = getParisMidnight_(parsedDate);
      if (date >= weekStart && date <= weekEnd) weekDays++;
    });

    const { tasks } = getTachesPourPersonne(personneKey);
    const maxPoints = tasks.length * SCORES.TASK_MAX + SCORES.GESTION_MAX;

    // B01 - Premiere etoile (1 evaluation)
    if (evals.length >= 1 && !existingBadges.includes('B01')) {
      if (awardBadge_(personneKey, 'B01')) {
        newBadges.push({ id: 'B01', nom: 'Premiere etoile', emoji: '⭐' });
      }
    }

    // B05 - Semaine champion (7 jours dans la semaine)
    if (weekDays >= 7 && !existingBadges.includes('B05')) {
      if (awardBadge_(personneKey, 'B05')) {
        newBadges.push({ id: 'B05', nom: 'Semaine champion', emoji: '🏆' });
      }
    }

    // B06 - Journee parfaite (max points)
    const hasPerfect = evals.some(r => Number(r[totalIdx] || 0) >= maxPoints);
    if (hasPerfect && !existingBadges.includes('B06')) {
      if (awardBadge_(personneKey, 'B06')) {
        newBadges.push({ id: 'B06', nom: 'Journee parfaite', emoji: '🌟' });
      }
    }

    // B08 - Zen master (5x gestion = max)
    const goodGestion = evals.filter(r => Number(r[gestionIdx] || 0) === SCORES.GESTION_MAX).length;
    if (goodGestion >= 5 && !existingBadges.includes('B08')) {
      if (awardBadge_(personneKey, 'B08')) {
        newBadges.push({ id: 'B08', nom: 'Zen master', emoji: '🧘' });
      }
    }

    // B11 - Explorateur emotions (7 jours avec emotions)
    const daysWithEmotions = evals.filter(r => r[emo1Idx] && String(r[emo1Idx]).trim() !== '').length;
    if (daysWithEmotions >= 7 && !existingBadges.includes('B11')) {
      if (awardBadge_(personneKey, 'B11')) {
        newBadges.push({ id: 'B11', nom: 'Explorateur emotions', emoji: '🎭' });
      }
    }

    // B10 - Collectionneur (5 badges)
    if (existingBadges.length + newBadges.length >= 5 && !existingBadges.includes('B10')) {
      if (awardBadge_(personneKey, 'B10')) {
        newBadges.push({ id: 'B10', nom: 'Collectionneur', emoji: '👑' });
      }
    }

  } catch (error) {
    Logger.log(`[checkBadges] Erreur: ${error}`);
  }

  return newBadges;
}

function awardBadge_(personne, badgeId) {
  try {
    const { sheet, rows, headerIndex } = getFeuilleAvecHeaders_('Badges_Obtenus');
    const persIdx = headerIndex['Personne'] ?? 0;
    const badgeIdIdx = headerIndex['BadgeID'] ?? 1;

    // Verifier si deja obtenu
    const alreadyHas = rows.some(r =>
      String(r[persIdx] || '').trim() === personne &&
      String(r[badgeIdIdx] || '').trim() === badgeId
    );
    if (alreadyHas) return false;

    sheet.appendRow([personne, badgeId, new Date()]);
    Logger.log(`[awardBadge] Badge ${badgeId} attribue a ${personne}.`);
    return true;

  } catch (error) {
    Logger.log(`[awardBadge] Erreur: ${error}`);
    return false;
  }
}

// ==================================================
// FONCTIONS UTILITAIRES EXPOSEES
// ==================================================

function forcerSynchronisation() {
  const resultats = synchroniserFeuilles_();
  Logger.log(`[forcerSynchronisation] Resultats: ${JSON.stringify(resultats)}`);
  return resultats;
}
