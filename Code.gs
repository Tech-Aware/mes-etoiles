// ==================================================
// 🌟 SYSTÈME DE MÉRITE FAMILIAL v4
// Avec section Émotions
// ==================================================

const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

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
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Évaluations');
  const data = sheet.getDataRange().getValues().slice(1);
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  return data.some(row => {
    const rowDate = new Date(row[1]);
    rowDate.setHours(0, 0, 0, 0);
    return row[3] === personne && rowDate.getTime() === today.getTime();
  });
}

// ==================================================
// SOUMETTRE UNE ÉVALUATION
// ==================================================
function submitEvaluation(personne, taches, emotions, humeur, commentaire) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Évaluations');
  
  if (hasEvaluatedToday(personne)) {
    return { success: false, message: 'Tu as déjà fait ton évaluation aujourd\'hui ! 😊' };
  }
  
  const lastRow = sheet.getLastRow();
  const newId = 'E' + String(lastRow).padStart(4, '0');
  const now = new Date();
  
  // Calculs totaux
  const totalCorvees = taches.rangerChambre + taches.faireLit + taches.rangerJouets + taches.aiderTable;
  const totalComportement = taches.ecouter + taches.gentilSoeur + taches.politesse + taches.pasColere;
  const totalRituels = taches.dentsMatin + taches.dentsSoir + taches.habiller + taches.cartable;
  const totalEmotions = emotions.gestion;
  const totalJour = totalCorvees + totalComportement + totalRituels + totalEmotions;
  
  // Ajouter la ligne
  sheet.appendRow([
    newId,
    now,
    Utilities.formatDate(now, 'Europe/Paris', 'HH:mm'),
    personne,
    // Tâches
    taches.rangerChambre,
    taches.faireLit,
    taches.rangerJouets,
    taches.aiderTable,
    taches.ecouter,
    taches.gentilSoeur,
    taches.politesse,
    taches.pasColere,
    taches.dentsMatin,
    taches.dentsSoir,
    taches.habiller,
    taches.cartable,
    // Émotions
    emotions.emotion1,
    emotions.emotion2 || '',
    emotions.source,
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
  const maxPoints = 52;
  const percent = Math.round((totalJour / maxPoints) * 100);
  
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
    emotions.source,
    emotions.gestion,
    ''
  ]);
}

// ==================================================
// DONNÉES PERSONNE
// ==================================================
function getPersonneData(personne) {
  const ss = SpreadsheetApp.openById(SS_ID);
  
  // Infos personne
  const personnesSheet = ss.getSheetByName('Personnes');
  const personnesData = personnesSheet.getDataRange().getValues().slice(1);
  const personneInfo = personnesData.find(r => r[0] === personne);
  
  // Évaluations
  const evalSheet = ss.getSheetByName('Évaluations');
  const evalData = evalSheet.getDataRange().getValues().slice(1);
  
  // Semaine en cours
  const today = new Date();
  const weekStart = getMonday(today);
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekEnd.getDate() + 6);
  weekEnd.setHours(23, 59, 59);
  
  let weekPoints = 0;
  let weekDays = 0;
  let dailyScores = [null, null, null, null, null, null, null];
  
  const personneEvals = evalData.filter(r => r[3] === personne);
  
  personneEvals.forEach(row => {
    const date = new Date(row[1]);
    if (date >= weekStart && date <= weekEnd) {
      const total = row[24]; // Colonne TotalJour
      weekPoints += total;
      weekDays++;
      const dayIndex = date.getDay() === 0 ? 6 : date.getDay() - 1;
      dailyScores[dayIndex] = total;
    }
  });
  
  // Streak
  let streak = 0;
  const sortedEvals = personneEvals.sort((a, b) => new Date(b[1]) - new Date(a[1]));
  
  if (sortedEvals.length > 0) {
    let checkDate = new Date();
    checkDate.setHours(0, 0, 0, 0);
    
    for (const eval of sortedEvals) {
      const evalDate = new Date(eval[1]);
      evalDate.setHours(0, 0, 0, 0);
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
    .filter(r => r[1] === personne)
    .sort((a, b) => new Date(b[0]) - new Date(a[0]))
    .slice(0, 7)
    .map(r => ({
      date: Utilities.formatDate(new Date(r[0]), 'Europe/Paris', 'dd/MM'),
      emotion1: r[2],
      emotion2: r[3],
      source: r[4],
      gestion: r[5]
    }));
  
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
    weekStart: Utilities.formatDate(weekStart, 'Europe/Paris', 'dd/MM'),
    weekEnd: Utilities.formatDate(weekEnd, 'Europe/Paris', 'dd/MM')
  };
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
  
  // B06 - Journée parfaite (52/52)
  const hasPerfect = evals.some(r => r[24] >= 50);
  if (hasPerfect && !data.badges.some(b => b.id === 'B06')) {
    if (awardBadge(personne, 'B06')) {
      newBadges.push({ id: 'B06', nom: 'Journée parfaite', emoji: '🌟' });
    }
  }
  
  // B08 - Zen master (5x gestion émotions = 4)
  const goodGestion = evals.filter(r => r[19] === 4).length;
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
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  d.setDate(diff);
  d.setHours(0, 0, 0, 0);
  return d;
}