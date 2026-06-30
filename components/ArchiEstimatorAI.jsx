"use client";
import { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── RATIOS BASE ──────────────────────────────────────────────────────────────
const RATIOS = {
  logement: {
    label: "Logement / Maison individuelle",
    // ⚠️ Ratios issus de la DPGF réelle Eiffage — 39 maisons M3 / Val-de-Reuil (Jan. 2026)
    // Base : 168 064 € travaux/maison · ~85 m² · Zone PPRI (coeff surcoût intégré)
    // Source : Marché Global de Performance 3F Normanvie / STUDIOS Architecture / AML
    neuf: {
      // LLS/LLI logement social — base DPGF Eiffage M3 (charpente+couverture+façades inclus dans grosOeuvre)
      "Bas de gamme":   { demolition:0, grosOeuvre:580, cloisons:140, menuiseries:140, electricite:75,  plomberie:90,  cvc:120, sols:70,  peinture:60,  mobilier:0, cuisine:0 },
      // Moyen de gamme — DPGF Eiffage M3 Val-de-Reuil (valeurs réelles 2026, zone PPRI, poteaux-poutres, bac acier, bardage bois)
      "Moyen de gamme": { demolition:0, grosOeuvre:685, cloisons:175, menuiseries:170, electricite:88,  plomberie:103, cvc:143, sols:82,  peinture:75,  mobilier:0, cuisine:0 },
      // Haut de gamme — surcoût matériaux premium estimé +35% vs base Eiffage
      "Haut de gamme":  { demolition:0, grosOeuvre:900, cloisons:230, menuiseries:240, electricite:130, plomberie:150, cvc:200, sols:130, peinture:110, mobilier:0, cuisine:0 },
    },
    renovation: {
      "Bas de gamme":   { demolition:30, grosOeuvre:90,  cloisons:120, menuiseries:130, electricite:90,  plomberie:110, cvc:100, sols:65,  peinture:55,  mobilier:0, cuisine:0 },
      "Moyen de gamme": { demolition:40, grosOeuvre:120, cloisons:160, menuiseries:170, electricite:110, plomberie:140, cvc:130, sols:85,  peinture:70,  mobilier:0, cuisine:0 },
      "Haut de gamme":  { demolition:60, grosOeuvre:180, cloisons:220, menuiseries:240, electricite:160, plomberie:200, cvc:185, sols:130, peinture:100, mobilier:0, cuisine:0 },
    },
  },
  restaurant: {
    label: "Restaurant",
    neuf: {
      "Bas de gamme":   { demolition:0, grosOeuvre:280, cloisons:95, menuiseries:70, electricite:120, plomberie:90, cvc:160, sols:75, peinture:40, mobilier:80, cuisine:200 },
      "Moyen de gamme": { demolition:0, grosOeuvre:350, cloisons:120, menuiseries:95, electricite:150, plomberie:115, cvc:200, sols:100, peinture:55, mobilier:130, cuisine:280 },
      "Haut de gamme":  { demolition:0, grosOeuvre:480, cloisons:160, menuiseries:140, electricite:210, plomberie:160, cvc:280, sols:150, peinture:80, mobilier:220, cuisine:400 },
    },
    renovation: {
      "Bas de gamme":   { demolition:25, grosOeuvre:80, cloisons:85, menuiseries:60, electricite:110, plomberie:80, cvc:140, sols:65, peinture:35, mobilier:70, cuisine:180 },
      "Moyen de gamme": { demolition:35, grosOeuvre:100, cloisons:110, menuiseries:85, electricite:140, plomberie:100, cvc:180, sols:90, peinture:50, mobilier:120, cuisine:260 },
      "Haut de gamme":  { demolition:50, grosOeuvre:150, cloisons:150, menuiseries:130, electricite:200, plomberie:145, cvc:260, sols:140, peinture:75, mobilier:200, cuisine:380 },
    },
  },
  bureaux: {
    label: "Bureaux",
    neuf: {
      "Bas de gamme":   { demolition:0, grosOeuvre:320, cloisons:100, menuiseries:80, electricite:130, plomberie:40, cvc:120, sols:60, peinture:35, mobilier:60, cuisine:0 },
      "Moyen de gamme": { demolition:0, grosOeuvre:400, cloisons:140, menuiseries:110, electricite:170, plomberie:55, cvc:160, sols:85, peinture:50, mobilier:100, cuisine:0 },
      "Haut de gamme":  { demolition:0, grosOeuvre:560, cloisons:200, menuiseries:160, electricite:240, plomberie:80, cvc:220, sols:130, peinture:70, mobilier:180, cuisine:0 },
    },
    renovation: {
      "Bas de gamme":   { demolition:20, grosOeuvre:60, cloisons:90, menuiseries:70, electricite:120, plomberie:35, cvc:105, sols:50, peinture:30, mobilier:50, cuisine:0 },
      "Moyen de gamme": { demolition:30, grosOeuvre:80, cloisons:130, menuiseries:100, electricite:160, plomberie:50, cvc:145, sols:75, peinture:45, mobilier:90, cuisine:0 },
      "Haut de gamme":  { demolition:45, grosOeuvre:120, cloisons:190, menuiseries:150, electricite:230, plomberie:75, cvc:205, sols:120, peinture:65, mobilier:165, cuisine:0 },
    },
  },
  commerce: {
    label: "Commerce / Retail",
    neuf: {
      "Bas de gamme":   { demolition:0, grosOeuvre:260, cloisons:80, menuiseries:100, electricite:115, plomberie:30, cvc:100, sols:70, peinture:30, mobilier:90, cuisine:0 },
      "Moyen de gamme": { demolition:0, grosOeuvre:330, cloisons:110, menuiseries:140, electricite:150, plomberie:45, cvc:140, sols:95, peinture:45, mobilier:150, cuisine:0 },
      "Haut de gamme":  { demolition:0, grosOeuvre:460, cloisons:160, menuiseries:210, electricite:210, plomberie:65, cvc:190, sols:145, peinture:65, mobilier:260, cuisine:0 },
    },
    renovation: {
      "Bas de gamme":   { demolition:20, grosOeuvre:50, cloisons:70, menuiseries:90, electricite:105, plomberie:25, cvc:85, sols:60, peinture:25, mobilier:75, cuisine:0 },
      "Moyen de gamme": { demolition:30, grosOeuvre:70, cloisons:100, menuiseries:130, electricite:140, plomberie:40, cvc:125, sols:85, peinture:40, mobilier:135, cuisine:0 },
      "Haut de gamme":  { demolition:45, grosOeuvre:110, cloisons:150, menuiseries:200, electricite:200, plomberie:60, cvc:175, sols:135, peinture:60, mobilier:245, cuisine:0 },
    },
  },
  hotel: {
    label: "Hôtel",
    neuf: {
      "Bas de gamme":   { demolition:0, grosOeuvre:420, cloisons:130, menuiseries:110, electricite:160, plomberie:180, cvc:200, sols:90, peinture:50, mobilier:150, cuisine:80 },
      "Moyen de gamme": { demolition:0, grosOeuvre:560, cloisons:170, menuiseries:160, electricite:210, plomberie:240, cvc:260, sols:130, peinture:70, mobilier:250, cuisine:120 },
      "Haut de gamme":  { demolition:0, grosOeuvre:780, cloisons:240, menuiseries:240, electricite:300, plomberie:340, cvc:360, sols:200, peinture:100, mobilier:420, cuisine:180 },
    },
    renovation: {
      "Bas de gamme":   { demolition:40, grosOeuvre:100, cloisons:120, menuiseries:100, electricite:150, plomberie:165, cvc:180, sols:80, peinture:45, mobilier:135, cuisine:70 },
      "Moyen de gamme": { demolition:55, grosOeuvre:140, cloisons:160, menuiseries:150, electricite:200, plomberie:225, cvc:240, sols:120, peinture:65, mobilier:230, cuisine:110 },
      "Haut de gamme":  { demolition:75, grosOeuvre:200, cloisons:230, menuiseries:230, electricite:285, plomberie:325, cvc:340, sols:190, peinture:95, mobilier:400, cuisine:165 },
    },
  },
  erp: {
    label: "ERP (Équipement Recevant du Public)",
    neuf: {
      "Bas de gamme":   { demolition:0, grosOeuvre:400, cloisons:110, menuiseries:90, electricite:170, plomberie:70, cvc:180, sols:70, peinture:40, mobilier:40, cuisine:0 },
      "Moyen de gamme": { demolition:0, grosOeuvre:500, cloisons:150, menuiseries:130, electricite:220, plomberie:95, cvc:230, sols:100, peinture:55, mobilier:70, cuisine:0 },
      "Haut de gamme":  { demolition:0, grosOeuvre:700, cloisons:210, menuiseries:190, electricite:310, plomberie:135, cvc:320, sols:150, peinture:80, mobilier:120, cuisine:0 },
    },
    renovation: {
      "Bas de gamme":   { demolition:30, grosOeuvre:90, cloisons:100, menuiseries:80, electricite:160, plomberie:60, cvc:165, sols:60, peinture:35, mobilier:35, cuisine:0 },
      "Moyen de gamme": { demolition:40, grosOeuvre:120, cloisons:140, menuiseries:120, electricite:210, plomberie:85, cvc:215, sols:90, peinture:50, mobilier:60, cuisine:0 },
      "Haut de gamme":  { demolition:60, grosOeuvre:175, cloisons:200, menuiseries:180, electricite:295, plomberie:125, cvc:305, sols:140, peinture:75, mobilier:110, cuisine:0 },
    },
  },
  // ── DONNÉES RÉELLES DPGF EIFFAGE ─────────────────────────────────────────
  // Source : Marché Global de Performance 3F Normanvie — 39 maisons M3
  // Val-de-Reuil (27100) — Zone PPRI — Janvier 2026
  // Architectes : STUDIOS Architecture + Atelier Marie Leguillon
  // Total exécution : 6 554 518 € HT / 39 maisons / ~85 m² moyen
  logement_social: {
    label: "Logement social / MGP (DPGF Eiffage 2026)",
    neuf: {
      "Bas de gamme":   { demolition:0, grosOeuvre:620, cloisons:150, menuiseries:145, electricite:80,  plomberie:92,  cvc:128, sols:74, peinture:68, mobilier:0, cuisine:0 },
      "Moyen de gamme": { demolition:0, grosOeuvre:685, cloisons:175, menuiseries:170, electricite:88,  plomberie:103, cvc:143, sols:82, peinture:75, mobilier:0, cuisine:0 },
      "Haut de gamme":  { demolition:0, grosOeuvre:790, cloisons:200, menuiseries:195, electricite:100, plomberie:120, cvc:165, sols:95, peinture:85, mobilier:0, cuisine:0 },
    },
    renovation: {
      "Bas de gamme":   { demolition:35, grosOeuvre:100, cloisons:130, menuiseries:140, electricite:85,  plomberie:95,  cvc:110, sols:70, peinture:60, mobilier:0, cuisine:0 },
      "Moyen de gamme": { demolition:45, grosOeuvre:130, cloisons:165, menuiseries:170, electricite:105, plomberie:118, cvc:138, sols:90, peinture:75, mobilier:0, cuisine:0 },
      "Haut de gamme":  { demolition:60, grosOeuvre:180, cloisons:200, menuiseries:200, electricite:130, plomberie:150, cvc:170, sols:115,peinture:95, mobilier:0, cuisine:0 },
    },
  },
};

const LOT_LABELS = {
  demolition:  "Démolition / Curage",
  grosOeuvre:  "Gros Œuvre / Structure",
  cloisons:    "Cloisons / Doublages",
  menuiseries: "Menuiseries int. & ext.",
  electricite: "Électricité / CFO-CFA",
  plomberie:   "Plomberie / Sanitaires",
  cvc:         "CVC — Chauffage / Ventilation / Clim.",
  sols:        "Revêtements de sols",
  peinture:    "Peinture / Décoration",
  mobilier:    "Mobilier / Agencement",
  cuisine:     "Équipements de cuisine",
};

const COEFF_LOC = {
  paris:    { label: "Paris / Île-de-France", coeff: 1.25 },
  lyon:     { label: "Lyon / Bordeaux / Marseille", coeff: 1.10 },
  autres:   { label: "Autres grandes villes", coeff: 1.00 },
  province: { label: "Province / Rural", coeff: 0.90 },
};

// ─── DETECTION MATERIAUX IFC → profil gamme ──────────────────────────────────
function detecterGammeDepuisMateriaux(materiaux) {
  const hauts = ["ardoise","zinc","bardage bois","chêne","pierre","marbre","cuivre"];
  const bas   = ["parpaing","agglo","pvc","osb"];
  const m = (materiaux || []).map(x => x.toLowerCase()).join(" ");
  const scoreHaut = hauts.filter(k => m.includes(k)).length;
  const scoreBas  = bas.filter(k => m.includes(k)).length;
  if (scoreHaut >= 2) return "Haut de gamme";
  if (scoreBas >= 2)  return "Bas de gamme";
  return "Moyen de gamme";
}

function detecterLocalisationDepuisNom(nom) {
  const n = (nom || "").toLowerCase();
  if (n.includes("paris") || n.includes("75") || n.includes("idf")) return "paris";
  if (n.includes("lyon") || n.includes("bordeaux") || n.includes("marseille")) return "lyon";
  return "province";
}

// ─── CALCUL HONORAIRES ────────────────────────────────────────────────────────
function calcHonoraires(totalTravaux, missionType) {
  let taux;
  if (totalTravaux < 100000) taux = 0.14;
  else if (totalTravaux < 300000) taux = 0.12;
  else if (totalTravaux < 700000) taux = 0.10;
  else if (totalTravaux < 1500000) taux = 0.09;
  else taux = 0.08;
  const missionFull = missionType === "complete" ? 1.0 : missionType === "conception" ? 0.52 : 0.35;
  const base = totalTravaux * taux * missionFull;
  const repartition = { DIAG:0.10, DEPOT_DP:0.08, PRO_DCE:0.42, ACT:0.10, DET:0.22, AOR:0.08 };
  const phases = {};
  for (const [k, p] of Object.entries(repartition)) phases[k] = Math.round(base * p / 100) * 100;
  const honorairesArchitecte = Object.values(phases).reduce((s, v) => s + v, 0);
  return {
    base, phases, taux, missionFull, honorairesArchitecte,
    honorairesBET:  Math.round(totalTravaux * 0.025 / 100) * 100,
    honorairesOPC:  Math.round(totalTravaux * 0.015 / 100) * 100,
    honorairesSPS:  Math.round(totalTravaux * 0.008 / 100) * 100,
  };
}

const fmt = n => new Intl.NumberFormat("fr-FR", { style:"currency", currency:"EUR", maximumFractionDigits:0 }).format(n);
const round1k = n => Math.round(n / 1000) * 1000;

// ─── EXPORT EXCEL ─────────────────────────────────────────────────────────────
function exportToExcel(form, result, ifcData) {
  const { lots, totalTravaux, hon, totalOperation, surface } = result;
  const today = new Date().toLocaleDateString("fr-FR");
  const nomProjet = form.nomProjet || "MON PROJET";
  const wb = XLSX.utils.book_new();

  // ── FEUILLE 1 : Proposition d'honoraires (modèle exact de l'architecte) ──
  const ws1 = {};
  const sc = (a, v) => { ws1[a] = { v, t: typeof v === "number" ? "n" : "s" }; };
  const sf = (a, f)  => { ws1[a] = { f, t: "n" }; };

  sc("C3", `PROPOSITION D'HONORAIRES  /  ${nomProjet.toUpperCase()}  /  ${(form.adresse||"").toUpperCase()}`);
  sc("H3", today);
  sc("C4", `MISSION DE MAÎTRISE D'ŒUVRE — ${form.missionType === "complete" ? "COMPLÈTE" : form.missionType === "conception" ? "CONCEPTION SEULE" : "PARTIELLE"}`);

  sc("B7",  "MONTANT TRAVAUX ESTIMATIF HT");
  sc("B8",  "TRAVAUX DE BASE HT");         sc("H8",  totalTravaux);
  sc("B9",  "TRAVAUX DEMOLITION");         sc("H9",  lots.demolition?.montant || 0);
  sc("B10", "TRAVAUX VRD");                sc("H10", 0);
  sf("G7",  "H8+H9+H10");

  const honConception = hon.phases.DIAG + hon.phases.DEPOT_DP + hon.phases.PRO_DCE + hon.phases.ACT;
  const honChantier   = hon.phases.DET  + hon.phases.AOR;
  sc("B11", "TAUX MOE CONCEPTION"); sc("G11", honConception); sf("E11", "G11/H8");
  sc("B12", "TAUX EXE");           sc("G12", honChantier);   sf("E12", "G12/H8");
  sc("B13", "TAUX OPC");           sc("G13", hon.honorairesOPC); sf("E13", "G13/G7");
  sc("B15", "RÉMUNÉRATION MOE + OPC"); sf("G15", "G11+G12+G13"); sf("E15", "G15/H8");
  sc("B16", "RÉMUNÉRATION MOE + OPC TTC (TVA 20%)"); sf("G16", "1.2*G15");

  sc("B18", "MISSIONS"); sc("C18", "TAUX PAR PHASE"); sc("D18", "MONTANT HT");
  sc("E18", "MOE Architecte"); sc("G18", "MOE EXE/ OPC");
  sc("I18", "BUREAUX DE CONTRÔLE"); sc("K18", "BET PLURI\nELEC/ CVC + THERMIQUE");
  sc("E19", form.nomAgence || "ATELIER ARCHITECTURE");
  sc("E20","Forfait en %"); sc("F20","Montant HT");
  sc("G20","Forfait en %"); sc("H20","Montant HT");
  sc("I20","Forfait en %"); sc("J20","Montant HT");
  sc("K20","Forfait en %"); sc("L20","Montant HT");

  sc("B21", "MISSION MOE CONCEPTION");
  sc("A22", "ETUDES");

  const phasesConc = [
    { r:22, l:"DIAG",                              v:hon.phases.DIAG },
    { r:23, l:"DEPOT DP / AT / DOSSIER SURETE",    v:hon.phases.DEPOT_DP },
    { r:24, l:"PRO / DCE / DQE",                   v:hon.phases.PRO_DCE },
    { r:25, l:"ACT - DOSSIER MARCHE",              v:hon.phases.ACT },
  ];
  for (const { r, l, v } of phasesConc) {
    sc(`B${r}`, l); sf(`C${r}`, `F${r}/F26`); sf(`D${r}`, `SUM(F${r})`);
    sf(`E${r}`, `F${r}/F26`); sc(`F${r}`, v);
    sc(`G${r}`,0); sc(`H${r}`,0); sc(`I${r}`,0); sc(`J${r}`,0); sc(`K${r}`,0); sc(`L${r}`,0);
  }
  sc("B26","TOTAL CONCEPTION");
  sf("D26","SUM(D22:D25)"); sf("F26","SUM(F22:F25)");
  sf("H26","SUM(H22:H25)"); sf("J26","SUM(J22:J25)"); sf("L26","SUM(L22:L25)");

  sc("B27","MISSION EXE"); sc("A28","CHANTIER");
  const phasesChant = [
    { r:28, l:"DET/VISA TRAVAUX", v:hon.phases.DET },
    { r:29, l:"AOR",              v:hon.phases.AOR },
    { r:30, l:"DOE/ Mis à jour conception des plans EXE", v:0 },
  ];
  for (const { r, l, v } of phasesChant) {
    sc(`B${r}`, l); sf(`C${r}`, `F${r}/F31`); sc(`D${r}`, v); sc(`F${r}`, v);
    sc(`G${r}`,0); sc(`H${r}`,0); sc(`I${r}`,0); sc(`J${r}`,0); sc(`K${r}`,0); sc(`L${r}`,0);
  }
  sc("B31","TOTAL CHANTIER");
  sf("D31","H31+L31+F31"); sf("F31","SUM(F28:F30)");
  sf("H31","SUM(H28:H30)"); sf("J31","SUM(J28:J30)"); sf("L31","SUM(L28:L30)");

  sc("B33","TOTAL OPC"); sc("C33",1); sc("D33", hon.honorairesOPC);

  sc("B34","TOTAUX HT");
  sf("C34","D34/D34"); sf("D34","D26+D31+D33");
  sf("E34","F34/D34*C34"); sf("F34","F26+F31");
  sf("G34","H34/D34*C34"); sf("H34","H31+H26");
  sf("I34","J34/D34*C34"); sf("J34","J31+J26");
  sf("K34","L34/D34");     sf("L34","L31+L26");

  sc("B35","TOTAUX TTC\nTAUX  TVA à revoir en fonction au montant travaux par poste");
  sf("C35","D35/D35"); sf("D35","F35+H35+J35+L35");
  sf("E35","F35/D35"); sf("F35","1.2*F34");
  sf("G35","H35/D35"); sf("H35","1.2*H34");
  sf("I35","J35/D35"); sf("J35","1.2*J34");
  sf("K35","L35/D35"); sf("L35","1.2*L34");

  ws1["!cols"] = [
    {wch:9},{wch:39},{wch:14},{wch:14},{wch:12},{wch:15},
    {wch:21},{wch:18},{wch:12},{wch:13},{wch:14},{wch:13},
  ];
  ws1["!ref"] = "A1:L38";
  ws1["!merges"] = [
    {s:{r:2,c:2},e:{r:2,c:6}},{s:{r:2,c:7},e:{r:2,c:11}},
    {s:{r:3,c:2},e:{r:3,c:6}},
    {s:{r:10,c:4},e:{r:10,c:5}},{s:{r:10,c:6},e:{r:10,c:7}},
    {s:{r:11,c:4},e:{r:11,c:5}},{s:{r:11,c:6},e:{r:11,c:7}},
    {s:{r:12,c:4},e:{r:12,c:5}},{s:{r:12,c:6},e:{r:12,c:7}},
    {s:{r:14,c:4},e:{r:14,c:5}},{s:{r:14,c:6},e:{r:14,c:7}},
    {s:{r:15,c:4},e:{r:15,c:5}},{s:{r:15,c:6},e:{r:15,c:7}},
    {s:{r:17,c:4},e:{r:17,c:5}},{s:{r:17,c:6},e:{r:17,c:7}},
    {s:{r:17,c:8},e:{r:17,c:9}},{s:{r:17,c:10},e:{r:17,c:11}},
    {s:{r:21,c:0},e:{r:23,c:0}},{s:{r:27,c:0},e:{r:29,c:0}},
  ];
  XLSX.utils.book_append_sheet(wb, ws1, "Honoraires");

  // ── FEUILLE 2 : Estimatif + données IFC ──────────────────────────────────
  const ws2 = {};
  const sc2 = (a, v) => { ws2[a] = { v, t: typeof v === "number" ? "n" : "s" }; };
  const sf2 = (a, f)  => { ws2[a] = { f, t: "n" }; };

  sc2("A1", `ESTIMATIF PRÉVISIONNEL DES TRAVAUX — ${nomProjet.toUpperCase()}`);
  sc2("A2", `Généré le ${today} | Surface: ${surface} m² | ${COEFF_LOC[form.localisation]?.label} | ${form.gamme} | ${form.typeOperation === "neuf" ? "Construction neuve" : "Rénovation"}`);
  if (ifcData) sc2("A3", `Source IFC: ${ifcData.fichier} | Phase: ${ifcData.phase_detectee} | SHAB extraite: ${ifcData.shab_total} m²`);

  // Tableau estimatif par lot
  sc2("A5","N°"); sc2("B5","CORPS D'ÉTAT"); sc2("C5","SURFACE (m²)"); sc2("D5","RATIO €/m²"); sc2("E5","MONTANT HT");
  let r = 6; let lotNum = 1;
  const lotRefs = [];
  for (const [key, label] of Object.entries(LOT_LABELS)) {
    if (!lots[key] || lots[key].montant === 0) continue;
    const { montant } = lots[key];
    const ratio = Math.round(montant / surface);
    sc2(`A${r}`, lotNum); sc2(`B${r}`, label.toUpperCase());
    sc2(`C${r}`, surface); sc2(`D${r}`, ratio); sc2(`E${r}`, montant);
    lotRefs.push(`E${r}`); r++; lotNum++;
  }
  r++;
  sc2(`B${r}`, "TOTAL TRAVAUX HT");
  sf2(`E${r}`, lotRefs.join("+")); const totalRow = r; r++;
  sc2(`B${r}`, "TOTAL TRAVAUX TTC (TVA 20%)");
  sf2(`E${r}`, `1.2*E${totalRow}`); r += 2;

  sc2(`B${r}`, "RÉCAPITULATIF BUDGET TOTAL OPÉRATION"); r++;
  const recap = [
    ["Travaux HT", totalTravaux],
    ["Honoraires architecte HT", hon.honorairesArchitecte],
    ["BET (Bureaux d'Études) HT", hon.honorairesBET],
    ["OPC HT", hon.honorairesOPC],
    ["Coordination SPS HT", hon.honorairesSPS],
    ["", ""],
    ["TOTAL OPÉRATION HT", totalOperation],
    ["TOTAL OPÉRATION TTC (TVA 20%)", totalOperation * 1.2],
  ];
  for (const [l, v] of recap) {
    sc2(`B${r}`, l); if (v !== "") sc2(`E${r}`, v); r++;
  }

  // Données IFC si disponibles
  if (ifcData && ifcData.pieces?.length) {
    r += 2;
    sc2(`A${r}`, "DONNÉES EXTRAITES DU FICHIER IFC"); r++;
    sc2(`A${r}`,"PROGRAMME — SURFACES PAR PIÈCE"); r++;
    sc2(`A${r}`,"Pièce"); sc2(`B${r}`,"Surface (m²)"); r++;
    for (const piece of ifcData.pieces) {
      sc2(`A${r}`, piece.nom); sc2(`B${r}`, piece.surface); r++;
    }
    sc2(`A${r}`, "SHAB TOTALE"); sc2(`B${r}`, ifcData.shab_total); r += 2;

    sc2(`A${r}`,"ÉLÉMENTS BIM COMPTABILISÉS"); r++;
    const elems = [
      ["Murs", ifcData.elements.murs],
      ["Dalles", ifcData.elements.dalles],
      ["Toitures", ifcData.elements.toitures],
      ["Portes", ifcData.elements.portes],
      ["Fenêtres", ifcData.elements.fenetres],
      ["Colonnes", ifcData.elements.colonnes],
      ["Poutres", ifcData.elements.poutres],
      ["Escaliers", ifcData.elements.escaliers],
    ];
    for (const [l, v] of elems) { sc2(`A${r}`, l); sc2(`B${r}`, v); r++; }

    if (ifcData.surfaces?.murs_m2 > 0) {
      r++;
      sc2(`A${r}`, "SURFACES BIM"); r++;
      sc2(`A${r}`,"Surface murs (m²)"); sc2(`B${r}`, ifcData.surfaces.murs_m2); r++;
      sc2(`A${r}`,"Surface dalles (m²)"); sc2(`B${r}`, ifcData.surfaces.dalles_m2); r++;
    }

    if (ifcData.materiaux?.length) {
      r++;
      sc2(`A${r}`, "MATÉRIAUX DÉTECTÉS"); r++;
      for (const m of ifcData.materiaux.slice(0,20)) { sc2(`A${r}`, m); r++; }
    }

    if (ifcData.niveaux?.length) {
      r++;
      sc2(`A${r}`, "NIVEAUX / ÉTAGES"); r++;
      for (const n of ifcData.niveaux) { sc2(`A${r}`, n.nom); sc2(`B${r}`, `${n.elevation} m`); r++; }
    }
  }

  ws2["!cols"] = [{wch:32},{wch:40},{wch:16},{wch:14},{wch:18}];
  ws2["!ref"] = `A1:E${r+5}`;
  ws2["!merges"] = [
    {s:{r:0,c:0},e:{r:0,c:4}},
    {s:{r:1,c:0},e:{r:1,c:4}},
  ];
  XLSX.utils.book_append_sheet(wb, ws2, "Estimatif & IFC");

  const nom = `${nomProjet.replace(/\s+/g,"_")}_Estimation_Honoraires.xlsx`;
  XLSX.writeFile(wb, nom);
}

// ─── ICÔNES SVG ──────────────────────────────────────────────────────────────
const IconUpload = () => (
  <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/>
  </svg>
);
const IconCheck = () => (
  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7"/>
  </svg>
);
const AMLLogo = ({ className = "" }) => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 349.51 115.9" className={className}>
    <defs><style>{`.cls-1{fill:#4a4a49;}.cls-2{fill:#caba9f;}`}</style></defs>
    <path className="cls-1" d="M16.57,78.86l5.72,13.57H10.83ZM6.13,99.66H7.9l2.36-5.85H22.82l2.46,5.85h1.8L16.57,74.87Z" transform="translate(-6.13 -4.99)"/>
    <polygon className="cls-1" points="35.73 71.4 26.56 71.4 26.56 94.82 25.18 94.82 25.18 71.4 16.01 71.4 16.01 69.99 35.73 69.99 35.73 71.4"/>
    <polygon className="cls-1" points="39.47 71.44 39.47 82.02 49.09 82.02 49.09 83.39 39.47 83.39 39.47 93.3 51.77 93.3 51.77 94.71 38.05 94.71 38.05 70.09 52.09 70.09 52.09 71.44 39.47 71.44"/>
    <polygon className="cls-1" points="55.94 69.99 55.94 93.3 67.29 93.3 67.29 94.71 54.38 94.71 54.38 69.99 55.94 69.99"/>
    <rect className="cls-1" x="69.76" y="69.81" width="1.34" height="24.79"/>
    <polygon className="cls-1" points="74.88 71.44 74.88 82.02 84.5 82.02 84.5 83.39 74.88 83.39 74.88 93.3 87.18 93.3 87.18 94.71 73.47 94.71 73.47 70.09 87.5 70.09 87.5 71.44 74.88 71.44"/>
    <path className="cls-1" d="M102.28,89.26a6.53,6.53,0,0,0,0-13.05h-5V89.26Zm.14,1.31-5.15,0v9.13H95.93V74.91h6.49a7.87,7.87,0,0,1,7.69,7.83,7.79,7.79,0,0,1-3.53,6.55l-.78.5,4.48,9.94H108.8l-4.27-9.45a7,7,0,0,1-2.11.29" transform="translate(-6.13 -4.99)"/>
    <polygon className="cls-1" points="129.93 69.99 129.93 94.64 126.54 94.64 126.54 77.68 118.89 94.67 111.48 77.68 111.48 94.64 108.1 94.64 108.1 69.99 112.15 69.99 118.89 86 125.91 69.99 129.93 69.99"/>
    <path className="cls-1" d="M148.8,84l-3.74,8.6h7.47Zm-5.22,11.81-1.73,3.88h-3.74L148.8,74.87l10.72,24.79h-3.85L154,95.78Z" transform="translate(-6.13 -4.99)"/>
    <path className="cls-1" d="M175.18,83.26a5,5,0,0,0-5-4.93h-3.56v9.91h3.56a5,5,0,0,0,5-5m4.23,16.33-3.81,0,0-.07-4.16-8.08a7.53,7.53,0,0,1-1,.07l-.08,0h-3.7v8.07h-3.31l0-24.68h7.06a8.3,8.3,0,0,1,4.37,15.2Z" transform="translate(-6.13 -4.99)"/>
    <rect className="cls-1" x="176.03" y="69.99" width="3.42" height="24.72"/>
    <polygon className="cls-1" points="196.48 91.39 196.48 94.67 181.78 94.67 181.78 69.99 196.31 69.99 196.31 73.3 185.16 73.3 185.16 80.89 194.01 80.89 194.01 84.17 185.16 84.17 185.16 91.39 196.48 91.39"/>
    <polygon className="cls-1" points="201.88 69.99 201.88 93.3 213.23 93.3 213.23 94.71 200.33 94.71 200.33 69.99 201.88 69.99"/>
    <polygon className="cls-1" points="216.94 71.44 216.94 82.02 226.56 82.02 226.56 83.39 216.94 83.39 216.94 93.3 229.25 93.3 229.25 94.71 215.53 94.71 215.53 70.09 229.56 70.09 229.56 71.44 216.94 71.44"/>
    <path className="cls-1" d="M257.46,93.46a11.7,11.7,0,0,1-10.19,6.17,12,12,0,0,1-11.78-12.27c0-6.81,5.29-12.35,11.78-12.35a11.44,11.44,0,0,1,6.38,2l-.74,1.3a10.38,10.38,0,0,0-5.64-1.8c-5.68,0-10.37,4.94-10.37,10.87a10.69,10.69,0,0,0,10.37,10.79c5.21,0,9.55-4.23,10.19-9.49H250v-1.3h8.85v12.3h-1.37Z" transform="translate(-6.13 -4.99)"/>
    <path className="cls-1" d="M270.05,99.63c-4.58,0-8.85-3.88-8.64-8.64V74.91h1.52V91a7.12,7.12,0,0,0,14.24,0V74.91h1.48v16a8.71,8.71,0,0,1-8.6,8.71" transform="translate(-6.13 -4.99)"/>
    <rect className="cls-1" x="275.03" y="69.81" width="1.34" height="24.79"/>
    <polygon className="cls-1" points="280.32 69.99 280.32 93.3 291.67 93.3 291.67 94.71 278.76 94.71 278.76 69.99 280.32 69.99"/>
    <polygon className="cls-1" points="295.55 69.99 295.55 93.3 306.91 93.3 306.91 94.71 294 94.71 294 69.99 295.55 69.99"/>
    <path className="cls-1" d="M333.08,87.25a10.52,10.52,0,1,0-10.51,10.9,10.82,10.82,0,0,0,10.51-10.9m1.3,0c0,6.81-5.29,12.38-11.81,12.38s-11.82-5.57-11.82-12.38a11.83,11.83,0,1,1,23.63,0" transform="translate(-6.13 -4.99)"/>
    <polygon className="cls-1" points="347.82 70.03 347.82 93.05 335.02 70.03 332.69 70.03 332.69 94.64 334.17 94.64 334.1 71.68 347.15 94.67 349.3 94.67 349.3 70.03 347.82 70.03"/>
    <path className="cls-2" d="M325.85,38.58a53.25,53.25,0,0,0-11.8-17.64A53.86,53.86,0,0,0,296.61,9.25a54.11,54.11,0,0,0-41.61,0,54,54,0,0,0-26.33,23.44A54,54,0,0,0,202.28,9.15,53.29,53.29,0,0,0,181.47,5a61,61,0,0,0-10.7,1A48.34,48.34,0,0,0,160,9.25a46.87,46.87,0,0,0-10,5.95,43.29,43.29,0,0,0-8.72,9.21V9.55H127.37V59.09c0,.25,0,.5,0,.76h13.89c0-.26,0-.5,0-.76a39,39,0,0,1,3.17-15.65A40.34,40.34,0,0,1,165.82,22a40.24,40.24,0,0,1,31.31,0,40.37,40.37,0,0,1,21.4,21.41,39,39,0,0,1,3.17,15.65c0,.26,0,.5,0,.76H235.6c0-.16,0-.31,0-.47a39.61,39.61,0,0,1,3.17-15.75,40.13,40.13,0,0,1,21.4-21.5,40.24,40.24,0,0,1,31.31,0,40.13,40.13,0,0,1,21.4,21.5A39.62,39.62,0,0,1,316,59.38c0,.16,0,.31,0,.47h13.9v-.47a54.44,54.44,0,0,0-4.06-20.8" transform="translate(-6.13 -4.99)"/>
    <rect className="cls-1" x="215.47" y="61.77" width="14.02" height="1.34"/>
    <path className="cls-2" d="M114.25,38.58a53.22,53.22,0,0,0-11.79-17.64A53.86,53.86,0,0,0,85,9.25a54.11,54.11,0,0,0-41.61,0A54.37,54.37,0,0,0,14.17,38.58a54.44,54.44,0,0,0-4.06,20.8c0,.16,0,.31,0,.47H24c0-.16,0-.31,0-.47a39.61,39.61,0,0,1,3.17-15.75,40.13,40.13,0,0,1,21.4-21.5,40.24,40.24,0,0,1,31.31,0,40.13,40.13,0,0,1,21.4,21.5,39.61,39.61,0,0,1,3.17,15.75c0,.16,0,.31,0,.47h13.9v-.47a54.26,54.26,0,0,0-4.07-20.8" transform="translate(-6.13 -4.99)"/>
    <rect className="cls-2" x="335.57" y="3.85" width="13.87" height="51.01"/>
  </svg>
);
const IconDownload = () => (
  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
  </svg>
);

// ─── COMPOSANTS UI ────────────────────────────────────────────────────────────
const Card = ({ children, className="" }) => (
  <div className={`bg-white rounded-2xl border border-slate-200 shadow-sm ${className}`}>{children}</div>
);
const Badge = ({ children, color="slate" }) => {
  const c = { slate:"bg-slate-100 text-slate-700", emerald:"bg-emerald-100 text-emerald-700", amber:"bg-amber-100 text-amber-700", violet:"bg-violet-100 text-violet-700", blue:"bg-blue-100 text-blue-700", rose:"bg-rose-100 text-rose-700" };
  return <span className={`inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium ${c[color]}`}>{children}</span>;
};
const ProgressBar = ({ value, max }) => (
  <div className="w-full bg-slate-100 rounded-full h-1.5">
    <div className="h-1.5 rounded-full bg-slate-700 transition-all duration-700" style={{ width:`${Math.min((value/max)*100,100)}%` }}/>
  </div>
);

// ─── MODULE IFC ───────────────────────────────────────────────────────────────
function IFCUploader({ onDataExtracted }) {
  const [state, setState]     = useState("idle"); // idle | parsing | done | error
  const [ifcData, setIfcData] = useState(null);
  const [fileName, setFileName] = useState("");
  const inputRef = useRef(null);

  const parseIFC = useCallback(async (file) => {
    setState("parsing");
    setFileName(file.name);
    try {
      // Lecture du fichier texte IFC
      const text = await file.text();
      const lines = text.split("\n");

      // Parser IFC simplifié — extraction par regex des entités clés
      const data = {
        fichier: file.name,
        phase_detectee: "inconnu",
        batiment: null,
        niveaux: [],
        pieces: [],
        shab_total: 0,
        elements: { murs:0, dalles:0, toitures:0, portes:0, fenetres:0, colonnes:0, poutres:0, escaliers:0 },
        surfaces: { murs_m2: 0, dalles_m2: 0 },
        materiaux: [],
        menuiseries: [],
      };

      // Détection phase depuis nom fichier
      const fn = file.name.toUpperCase();
      if (fn.includes("PRO")) data.phase_detectee = "PRO";
      else if (fn.includes("APD")) data.phase_detectee = "APD";
      else if (fn.includes("APS")) data.phase_detectee = "APS";
      else if (fn.includes("ESQ") || fn.includes("CONCOURS") || fn.includes("PILOTIS")) data.phase_detectee = "ESQ/Concours";

      // Comptage entités
      const counts = {};
      const spaces = [];
      const materiaux = new Set();

      for (const line of lines) {
        const l = line.trim();

        // Compter types
        const typeMatch = l.match(/=\s*IFC(\w+)\s*\(/);
        if (typeMatch) {
          const t = "IFC" + typeMatch[1].toUpperCase();
          counts[t] = (counts[t] || 0) + 1;
        }

        // Extraire nom bâtiment
        if (l.includes("IFCBUILDING(") && !data.batiment) {
          const m = l.match(/IFCBUILDING\([^,]*,[^,]*,'([^']+)'/);
          if (m) data.batiment = m[1];
        }

        // Niveaux
        if (l.includes("IFCBUILDINGSTOREY(")) {
          const m = l.match(/IFCBUILDINGSTOREY\([^,]*,[^,]*,'([^']+)'[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,([^,)]+)/);
          if (m) {
            const elev = parseFloat(m[2]);
            data.niveaux.push({ nom: m[1], elevation: isNaN(elev) ? 0 : Math.round(elev*100)/100 });
          }
        }

        // Espaces / pièces
        if (l.includes("IFCSPACE(")) {
          const m = l.match(/IFCSPACE\([^,]*,[^,]*,'([^']*)'[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,'([^']*)'/);
          if (m) spaces.push({ nom: m[2] || m[1] });
        }

        // Surfaces quantités
        if (l.includes("IFCQUANTITYAREA(")) {
          const m = l.match(/IFCQUANTITYAREA\('[^']*Area[^']*'[^,]*,[^,]*,([^,)]+)/i);
          if (m) {
            const v = parseFloat(m[1]);
            if (!isNaN(v) && v > 0.5 && v < 2000) {
              // Approximation : si petite surface → pièce, sinon → mur/dalle
              if (v < 100) data.shab_total += v;
            }
          }
        }

        // Matériaux
        if (l.includes("IFCMATERIAL(")) {
          const m = l.match(/IFCMATERIAL\('([^']+)'/);
          if (m && m[1].length > 2) materiaux.add(m[1]);
        }

        // Menuiseries
        if (l.includes("IFCWINDOW(")) {
          const m = l.match(/IFCWINDOW\([^,]*,[^,]*,'([^']*)'[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,([^,]*),([^,)]*)/);
          if (m) {
            data.menuiseries.push({ type:"fenetre", nom:m[1], hauteur:parseFloat(m[2])||0, largeur:parseFloat(m[3])||0 });
          }
        }
        if (l.includes("IFCDOOR(")) {
          const m = l.match(/IFCDOOR\([^,]*,[^,]*,'([^']*)'[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,([^,]*),([^,)]*)/);
          if (m) {
            data.menuiseries.push({ type:"porte", nom:m[1], hauteur:parseFloat(m[2])||0, largeur:parseFloat(m[3])||0 });
          }
        }
      }

      // Remplir depuis counts
      data.elements.murs      = (counts["IFCWALL"] || 0) + (counts["IFCWALLSTANDARDCASE"] || 0);
      data.elements.dalles    = counts["IFCSLAB"] || 0;
      data.elements.toitures  = counts["IFCROOF"] || 0;
      data.elements.portes    = counts["IFCDOOR"] || 0;
      data.elements.fenetres  = counts["IFCWINDOW"] || 0;
      data.elements.colonnes  = counts["IFCCOLUMN"] || 0;
      data.elements.poutres   = counts["IFCBEAM"] || 0;
      data.elements.escaliers = counts["IFCSTAIR"] || 0;

      // Pièces depuis spaces (dédupliquer)
      const seen = new Set();
      for (const s of spaces) {
        if (s.nom && !seen.has(s.nom)) {
          seen.add(s.nom);
          data.pieces.push({ nom: s.nom, surface: null });
        }
      }

      // Matériaux
      data.materiaux = [...materiaux].sort();

      // Estimation SHAB depuis surfaces quantités (approximation)
      data.shab_total = Math.round(data.shab_total * 10) / 10;

      setIfcData(data);
      setState("done");
      onDataExtracted(data);
    } catch(e) {
      console.error(e);
      setState("error");
    }
  }, [onDataExtracted]);

  const onDrop = useCallback(e => {
    e.preventDefault();
    const file = e.dataTransfer?.files?.[0] || e.target.files?.[0];
    if (file && (file.name.endsWith(".ifc") || file.name.endsWith(".IFC"))) parseIFC(file);
  }, [parseIFC]);

  return (
    <div>
      {/* Zone de drop */}
      {state === "idle" && (
        <div
          onDrop={onDrop}
          onDragOver={e => e.preventDefault()}
          onClick={() => inputRef.current?.click()}
          className="border-2 border-dashed border-slate-300 rounded-xl p-6 text-center cursor-pointer hover:border-slate-500 hover:bg-slate-50 transition-all group"
        >
          <input ref={inputRef} type="file" accept=".ifc" className="hidden" onChange={e => { const f = e.target.files?.[0]; if(f) parseIFC(f); }}/>
          <div className="w-10 h-10 bg-slate-100 rounded-lg flex items-center justify-center mx-auto mb-3 group-hover:bg-slate-200 transition-colors">
            <IconUpload />
          </div>
          <p className="text-sm font-semibold text-slate-700">Déposer un fichier IFC</p>
          <p className="text-xs text-slate-400 mt-1">ou cliquer pour parcourir</p>
          <p className="text-xs text-slate-400 mt-2">Formats supportés : IFC2x3, IFC4 (Revit, ArchiCAD, Archicad, etc.)</p>
        </div>
      )}

      {/* Parsing */}
      {state === "parsing" && (
        <div className="border border-slate-200 rounded-xl p-5 text-center">
          <div className="flex items-center justify-center gap-2 mb-2">
            <svg className="w-4 h-4 animate-spin text-slate-600" fill="none" viewBox="0 0 24 24"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/></svg>
            <span className="text-sm font-medium text-slate-700">Analyse BIM en cours…</span>
          </div>
          <p className="text-xs text-slate-400">{fileName}</p>
          <div className="mt-3 space-y-1.5">
            {["Lecture entités IFC","Extraction espaces & surfaces","Comptage éléments","Détection matériaux"].map((s,i) => (
              <div key={i} className="flex items-center gap-2 text-xs text-slate-500">
                <div className="w-3 h-3 rounded-full bg-slate-200 animate-pulse"/>
                {s}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Done */}
      {state === "done" && ifcData && (
        <div className="border border-emerald-200 bg-emerald-50 rounded-xl p-4">
          <div className="flex items-start justify-between mb-3">
            <div className="flex items-center gap-2">
              <div className="w-6 h-6 bg-emerald-600 rounded-full flex items-center justify-center flex-shrink-0">
                <IconCheck />
              </div>
              <div>
                <p className="text-sm font-semibold text-emerald-800">Fichier IFC analysé</p>
                <p className="text-xs text-emerald-600">{ifcData.fichier}</p>
              </div>
            </div>
            <Badge color="emerald">{ifcData.phase_detectee}</Badge>
          </div>

          {/* Stats extraites */}
          <div className="grid grid-cols-2 gap-2 mb-3">
            {[
              { label:"SHAB extraite", value: ifcData.shab_total > 0 ? `${ifcData.shab_total} m²` : `~${ifcData.elements.murs > 50 ? "120+" : "80"} m²` },
              { label:"Pièces", value: ifcData.pieces.length > 0 ? `${ifcData.pieces.length} pièces` : `${ifcData.elements.dalles} niveaux` },
              { label:"Murs / Dalles", value: `${ifcData.elements.murs} / ${ifcData.elements.dalles}` },
              { label:"Portes / Fenêtres", value: `${ifcData.elements.portes} / ${ifcData.elements.fenetres}` },
            ].map(({ label, value }) => (
              <div key={label} className="bg-white rounded-lg px-3 py-2 border border-emerald-100">
                <p className="text-xs text-emerald-600">{label}</p>
                <p className="text-sm font-bold text-emerald-900">{value}</p>
              </div>
            ))}
          </div>

          {/* Matériaux détectés */}
          {ifcData.materiaux.length > 0 && (
            <div className="mb-3">
              <p className="text-xs font-semibold text-emerald-700 mb-1.5">Matériaux détectés</p>
              <div className="flex flex-wrap gap-1">
                {ifcData.materiaux.filter(m => m.length < 30).slice(0, 8).map(m => (
                  <span key={m} className="text-xs bg-white border border-emerald-200 text-emerald-700 px-2 py-0.5 rounded">{m}</span>
                ))}
              </div>
            </div>
          )}

          {/* Pièces */}
          {ifcData.pieces.length > 0 && (
            <div>
              <p className="text-xs font-semibold text-emerald-700 mb-1.5">Programme détecté</p>
              <div className="grid grid-cols-2 gap-1">
                {ifcData.pieces.slice(0,8).map(p => (
                  <div key={p.nom} className="flex justify-between text-xs bg-white border border-emerald-100 rounded px-2 py-1">
                    <span className="text-emerald-700 truncate">{p.nom}</span>
                    {p.surface && <span className="font-mono font-medium text-emerald-900 ml-1 flex-shrink-0">{p.surface} m²</span>}
                  </div>
                ))}
              </div>
            </div>
          )}

          <button onClick={() => { setState("idle"); setIfcData(null); onDataExtracted(null); }}
            className="mt-3 text-xs text-emerald-600 hover:text-emerald-800 underline">
            Changer de fichier
          </button>
        </div>
      )}

      {state === "error" && (
        <div className="border border-rose-200 bg-rose-50 rounded-xl p-4 text-center">
          <p className="text-sm font-semibold text-rose-700">Fichier non reconnu</p>
          <p className="text-xs text-rose-500 mt-1">Vérifiez que le fichier est bien au format IFC</p>
          <button onClick={() => setState("idle")} className="mt-2 text-xs text-rose-600 underline">Réessayer</button>
        </div>
      )}
    </div>
  );
}

// ─── APP PRINCIPALE ───────────────────────────────────────────────────────────
export default function ArchiEstimatorAI() {
  const [step, setStep]       = useState("form");
  const [ifcData, setIfcData] = useState(null);
  const [form, setForm]       = useState({
    nomProjet:"", adresse:"", nomAgence:"",
    typeProjet:"logement", surface:"", gamme:"Moyen de gamme",
    localisation:"province", typeOperation:"neuf", missionType:"complete",
    budgetConnu:"", commentaire:"",
  });
  const [result, setResult]       = useState(null);
  const [aiAnalysis, setAiAnalysis] = useState("");
  const [loadingAI, setLoadingAI]   = useState(false);
  const [activeTab, setActiveTab]   = useState("esq");
  const [downloading, setDownloading] = useState(false);
  const [ifcApplied, setIfcApplied]   = useState(false);
  const resultRef = useRef(null);

  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  // ── Pré-remplissage depuis IFC ─────────────────────────────────────────────
  const onIFCExtracted = useCallback((data) => {
    if (!data) { setIfcData(null); setIfcApplied(false); return; }
    setIfcData(data);

    // Surface : SHAB extraite ou estimation depuis nb murs
    let surfaceEstimee = data.shab_total > 10 ? Math.round(data.shab_total) : null;
    if (!surfaceEstimee && data.elements.murs > 0) {
      // Estimation grossière : ~1.5 m² de mur par m² de SHAB en maison individuelle
      surfaceEstimee = Math.round(data.surfaces.murs_m2 / 4) || null;
    }

    // Gamme depuis matériaux
    const gammeDetectee = detecterGammeDepuisMateriaux(data.materiaux);

    // Localisation depuis nom fichier
    const locDetectee = detecterLocalisationDepuisNom(data.fichier + " " + (data.batiment || ""));

    // Phase → type opération
    const typeOp = data.phase_detectee?.includes("ESQ") ? "neuf" : "neuf";

    setForm(f => ({
      ...f,
      typeProjet: "logement",
      surface: surfaceEstimee ? String(surfaceEstimee) : f.surface,
      gamme: gammeDetectee,
      localisation: locDetectee,
      typeOperation: typeOp,
      commentaire: [
        data.elements.escaliers > 0 ? "Avec escalier" : "",
        data.elements.colonnes > 10 ? `${data.elements.colonnes} colonnes/pilotis` : "",
        data.materiaux.some(m => m.toLowerCase().includes("ardoise")) ? "Toiture ardoise" : "",
        data.materiaux.some(m => m.toLowerCase().includes("bardage")) ? "Bardage bois" : "",
        data.materiaux.some(m => m.toLowerCase().includes("zinc")) ? "Zinc en couverture" : "",
      ].filter(Boolean).join(" — ") || f.commentaire,
    }));
    setIfcApplied(true);
  }, []);

  // ── Calcul ────────────────────────────────────────────────────────────────
  function calculate() {
    const surface = parseFloat(form.surface);
    if (!surface || surface <= 0) return;
    const ratios   = RATIOS[form.typeProjet][form.typeOperation][form.gamme];
    const locCoeff = COEFF_LOC[form.localisation].coeff;
    const lots = {};
    let totalTravaux = 0;
    for (const [key, ratio] of Object.entries(ratios)) {
      if (ratio === 0) continue;
      const montant = round1k(ratio * surface * locCoeff);
      lots[key] = { ratio, montant };
      totalTravaux += montant;
    }
    const esq = {
      "Gros Œuvre & Structure": Object.entries(lots).filter(([k])=>["demolition","grosOeuvre"].includes(k)).reduce((s,[,v])=>s+v.montant,0),
      "Second Œuvre":            Object.entries(lots).filter(([k])=>["cloisons","menuiseries","sols","peinture","mobilier","cuisine"].includes(k)).reduce((s,[,v])=>s+v.montant,0),
      "Corps d'État Techniques": Object.entries(lots).filter(([k])=>["electricite","plomberie","cvc"].includes(k)).reduce((s,[,v])=>s+v.montant,0),
    };
    const apd = Object.fromEntries(Object.entries(lots).map(([k,v])=>[k,{...v,min:round1k(v.montant*.85),max:round1k(v.montant*1.15)}]));
    const hon = calcHonoraires(totalTravaux, form.missionType);
    const totalOperation = totalTravaux + hon.honorairesArchitecte + hon.honorairesBET + hon.honorairesOPC + hon.honorairesSPS;
    const alertes = [];
    if (form.budgetConnu && parseFloat(form.budgetConnu) > 0) {
      const ecart = ((totalTravaux - parseFloat(form.budgetConnu)) / parseFloat(form.budgetConnu)) * 100;
      if (Math.abs(ecart) > 20) alertes.push({ type:"danger", msg:`Écart de ${ecart>0?"+":""}${ecart.toFixed(0)}% vs budget cible (${fmt(parseFloat(form.budgetConnu))})` });
      else if (Math.abs(ecart) > 10) alertes.push({ type:"warning", msg:`Écart de ${ecart>0?"+":""}${ecart.toFixed(0)}% vs budget cible` });
      else alertes.push({ type:"ok", msg:`Budget cohérent — écart ${ecart>0?"+":""}${ecart.toFixed(0)}%` });
    }
    setResult({ lots, esq, apd, totalTravaux, hon, totalOperation, alertes, surface, locCoeff });
    setStep("result");
    setTimeout(()=>resultRef.current?.scrollIntoView({behavior:"smooth"}),100);
  }

  async function generateAIAnalysis() {
    if (!result) return;
    setLoadingAI(true); setAiAnalysis("");
    const { totalTravaux, hon, totalOperation, surface } = result;
    const ifcContext = ifcData ? `\nDonnées BIM IFC (${ifcData.phase_detectee}) :\n- SHAB: ${ifcData.shab_total} m²\n- ${ifcData.pieces.length} pièces détectées\n- ${ifcData.elements.murs} murs, ${ifcData.elements.dalles} dalles, ${ifcData.elements.colonnes} colonnes\n- Matériaux: ${ifcData.materiaux.slice(0,6).join(", ")}` : "";
    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method:"POST", headers:{"Content-Type":"application/json"},
        body: JSON.stringify({
          model:"claude-sonnet-4-6", max_tokens:1000,
          system:"Tu es un économiste de la construction expert en maîtrise d'œuvre française. Synthèse professionnelle concise, 3 paragraphes max, ton expert direct.",
          messages:[{ role:"user", content:`Projet : ${RATIOS[form.typeProjet]?.label} — ${form.gamme} — ${form.typeOperation}\nSurface : ${surface} m² | ${COEFF_LOC[form.localisation]?.label} (coeff. ${result.locCoeff})\nTravaux : ${fmt(totalTravaux)} (${fmt(totalTravaux/surface)}/m²)\nHonoraires archi : ${fmt(hon.honorairesArchitecte)} (taux ${(hon.taux*100).toFixed(1)}%)\nTotal opération : ${fmt(totalOperation)}${form.budgetConnu?`\nBudget cible : ${fmt(parseFloat(form.budgetConnu))}`:""}${form.commentaire?`\nSpécificités : ${form.commentaire}`:""}${ifcContext}\nAnalyse professionnelle : cohérence ratios, points d'attention, recommandations.` }],
        }),
      });
      const data = await res.json();
      setAiAnalysis(data.content?.[0]?.text || "");
    } catch { setAiAnalysis("Analyse IA indisponible."); }
    setLoadingAI(false);
  }

  function handleExport() {
    setDownloading(true);
    try { exportToExcel(form, result, ifcData); }
    finally { setTimeout(()=>setDownloading(false),1500); }
  }

  function reset() { setStep("form"); setResult(null); setAiAnalysis(""); window.scrollTo({top:0,behavior:"smooth"}); }

  // ─── FORMULAIRE ─────────────────────────────────────────────────────────────
  if (step === "form") return (
    <div className="min-h-screen bg-slate-50">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-5xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <AMLLogo className="h-9 w-auto" />
            <div className="border-l border-slate-200 pl-3">
              <p className="text-xs font-bold text-slate-900 tracking-widest uppercase">Atelier Marie Leguillon</p>
              <p className="text-xs text-slate-400">Estimation · BIM · IA</p>
            </div>
          </div>
          <Badge color="violet">BIM Ready</Badge>
        </div>
      </header>

      <div className="max-w-5xl mx-auto px-6 py-10">
        <div className="mb-8">
          <h1 className="text-3xl font-bold text-slate-900 mb-2">Estimation budgétaire</h1>
          <p className="text-slate-500">Importez un fichier IFC pour pré-remplir automatiquement le formulaire, ou saisissez manuellement.</p>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          <div className="lg:col-span-2 space-y-5">

            {/* ── BLOC IFC ── */}
            <Card className="p-6">
              <div className="flex items-center gap-2 mb-4">
                <div className="w-6 h-6 bg-violet-100 rounded flex items-center justify-center">
                  <svg className="w-3.5 h-3.5 text-violet-700" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 21a4 4 0 01-4-4V5a2 2 0 012-2h4l2 2h4a2 2 0 012 2v3"/><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M14 21v-4a2 2 0 012-2h4"/></svg>
                </div>
                <h2 className="text-sm font-semibold text-slate-900">Import BIM — Fichier IFC</h2>
                <Badge color="violet">Optionnel</Badge>
              </div>
              <IFCUploader onDataExtracted={onIFCExtracted} />
              {ifcApplied && (
                <div className="mt-3 flex items-center gap-2 text-xs text-violet-700 bg-violet-50 px-3 py-2 rounded-lg border border-violet-100">
                  <IconCheck />
                  <span>Formulaire pré-rempli depuis le fichier IFC — vérifiez et ajustez si nécessaire</span>
                </div>
              )}
            </Card>

            {/* Identité */}
            <Card className="p-6">
              <h2 className="text-sm font-semibold text-slate-700 uppercase tracking-wider mb-4">Identité du projet</h2>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">Nom du projet</label>
                  <input type="text" value={form.nomProjet} onChange={e=>set("nomProjet",e.target.value)} placeholder="ex. MAISON TÉMOIN" className="w-full border border-slate-300 rounded-lg px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-slate-900"/>
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">Adresse / Ville</label>
                  <input type="text" value={form.adresse} onChange={e=>set("adresse",e.target.value)} placeholder="ex. LES ANDELYS (27)" className="w-full border border-slate-300 rounded-lg px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-slate-900"/>
                </div>
              </div>
              <div className="mt-3">
                <label className="block text-xs font-medium text-slate-600 mb-1.5">Nom de l'agence</label>
                <input type="text" value={form.nomAgence} onChange={e=>set("nomAgence",e.target.value)} placeholder="ex. ATELIER MARIE LEGUILLON" className="w-full border border-slate-300 rounded-lg px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-slate-900"/>
              </div>
            </Card>

            {/* Projet */}
            <Card className="p-6">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-sm font-semibold text-slate-700 uppercase tracking-wider">Paramètres du projet</h2>
                {ifcApplied && <Badge color="violet">Pré-rempli IFC</Badge>}
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">Type de projet</label>
                  <select value={form.typeProjet} onChange={e=>set("typeProjet",e.target.value)} className="w-full border border-slate-300 rounded-lg px-3 py-2.5 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-slate-900">
                    {Object.entries(RATIOS).map(([k,v])=><option key={k} value={k}>{v.label}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">Type d'opération</label>
                  <div className="flex rounded-lg border border-slate-300 overflow-hidden">
                    {["renovation","neuf"].map(t=>(
                      <button key={t} onClick={()=>set("typeOperation",t)} className={`flex-1 py-2.5 text-sm font-medium transition-colors ${form.typeOperation===t?"bg-slate-900 text-white":"bg-white text-slate-600 hover:bg-slate-50"}`}>
                        {t==="renovation"?"Rénovation":"Neuf"}
                      </button>
                    ))}
                  </div>
                </div>
              </div>
              <div className="grid grid-cols-2 gap-4 mt-4">
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Surface (m²) <span className="text-rose-500">*</span>
                    {ifcApplied && ifcData?.shab_total > 0 && <span className="ml-1 text-violet-600 font-normal">← SHAB IFC</span>}
                  </label>
                  <input type="number" value={form.surface} onChange={e=>set("surface",e.target.value)} placeholder="ex. 124" className={`w-full border rounded-lg px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-slate-900 ${ifcApplied && ifcData?.shab_total > 0 ? "border-violet-300 bg-violet-50" : "border-slate-300"}`}/>
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">Localisation</label>
                  <select value={form.localisation} onChange={e=>set("localisation",e.target.value)} className="w-full border border-slate-300 rounded-lg px-3 py-2.5 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-slate-900">
                    {Object.entries(COEFF_LOC).map(([k,v])=><option key={k} value={k}>{v.label}</option>)}
                  </select>
                </div>
              </div>
            </Card>

            {/* Gamme */}
            <Card className="p-6">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-sm font-semibold text-slate-700 uppercase tracking-wider">Niveau de prestation</h2>
                {ifcApplied && <span className="text-xs text-violet-600">Détecté : {detecterGammeDepuisMateriaux(ifcData?.materiaux)}</span>}
              </div>
              <div className="grid grid-cols-3 gap-3">
                {["Bas de gamme","Moyen de gamme","Haut de gamme"].map(g=>(
                  <button key={g} onClick={()=>set("gamme",g)} className={`p-3 rounded-xl border-2 text-left transition-all ${form.gamme===g?"border-slate-900 bg-slate-900 text-white":"border-slate-200 bg-white text-slate-700 hover:border-slate-400"}`}>
                    <div className="text-lg mb-1">{g==="Bas de gamme"?"○":g==="Moyen de gamme"?"◑":"●"}</div>
                    <div className="text-xs font-medium">{g}</div>
                  </button>
                ))}
              </div>
            </Card>

            {/* Mission */}
            <Card className="p-6">
              <h2 className="text-sm font-semibold text-slate-700 uppercase tracking-wider mb-4">Mission architecte</h2>
              <div className="grid grid-cols-3 gap-3">
                {[{k:"complete",l:"Mission complète",s:"ESQ → AOR"},{k:"conception",l:"Conception seule",s:"ESQ → DCE"},{k:"partielle",l:"Mission partielle",s:"Sur devis"}].map(({k,l,s})=>(
                  <button key={k} onClick={()=>set("missionType",k)} className={`p-3 rounded-xl border-2 text-left transition-all ${form.missionType===k?"border-slate-900 bg-slate-900 text-white":"border-slate-200 bg-white text-slate-700 hover:border-slate-400"}`}>
                    <div className="text-xs font-semibold mb-0.5">{l}</div>
                    <div className={`text-xs ${form.missionType===k?"text-slate-300":"text-slate-400"}`}>{s}</div>
                  </button>
                ))}
              </div>
            </Card>

            {/* Contexte */}
            <Card className="p-6">
              <h2 className="text-sm font-semibold text-slate-700 uppercase tracking-wider mb-4">Contexte (optionnel)</h2>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">Budget cible client (€)</label>
                  <input type="number" value={form.budgetConnu} onChange={e=>set("budgetConnu",e.target.value)} placeholder="Laissez vide si inconnu" className="w-full border border-slate-300 rounded-lg px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-slate-900"/>
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">Spécificités</label>
                  <input type="text" value={form.commentaire} onChange={e=>set("commentaire",e.target.value)} placeholder="Ex. : toiture ardoise, bardage bois…" className={`w-full border rounded-lg px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-slate-900 ${ifcApplied && form.commentaire ? "border-violet-300 bg-violet-50" : "border-slate-300"}`}/>
                </div>
              </div>
            </Card>

            <button onClick={calculate} disabled={!form.surface} className="w-full bg-slate-900 text-white py-4 rounded-xl font-semibold text-base hover:bg-slate-800 disabled:opacity-40 disabled:cursor-not-allowed transition-all shadow-lg">
              Générer l'estimation budgétaire →
            </button>
          </div>

          {/* Sidebar */}
          <div className="space-y-4">
            {ifcData && (
              <Card className="p-5 border-violet-200 bg-violet-50">
                <p className="text-xs font-semibold text-violet-600 uppercase tracking-wider mb-3">Données BIM extraites</p>
                <div className="space-y-1.5">
                  {[
                    { l:"Phase IFC", v: ifcData.phase_detectee },
                    { l:"SHAB", v: ifcData.shab_total > 0 ? `${ifcData.shab_total} m²` : "Non disponible" },
                    { l:"Pièces", v: `${ifcData.pieces.length}` },
                    { l:"Murs", v: `${ifcData.elements.murs}` },
                    { l:"Dalles", v: `${ifcData.elements.dalles}` },
                    { l:"Portes", v: `${ifcData.elements.portes}` },
                    { l:"Fenêtres", v: `${ifcData.elements.fenetres}` },
                    { l:"Colonnes", v: `${ifcData.elements.colonnes}` },
                    { l:"Escaliers", v: `${ifcData.elements.escaliers}` },
                  ].map(({ l, v }) => (
                    <div key={l} className="flex justify-between text-xs">
                      <span className="text-violet-500">{l}</span>
                      <span className="font-semibold text-violet-800">{v}</span>
                    </div>
                  ))}
                </div>
                {ifcData.niveaux.length > 0 && (
                  <div className="mt-3 pt-3 border-t border-violet-200">
                    <p className="text-xs font-semibold text-violet-600 mb-1.5">Niveaux</p>
                    {ifcData.niveaux.map(n => (
                      <div key={n.nom} className="flex justify-between text-xs py-0.5">
                        <span className="text-violet-500">{n.nom}</span>
                        <span className="text-violet-700">{n.elevation} m</span>
                      </div>
                    ))}
                  </div>
                )}
              </Card>
            )}

            <Card className="p-5">
              <p className="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-3">Ratios de référence</p>
              <p className="text-xs text-slate-500 mb-2">Coeff. {COEFF_LOC[form.localisation]?.coeff}×</p>
              {form.surface && Object.entries(RATIOS[form.typeProjet]?.[form.typeOperation]?.[form.gamme] || {}).map(([k,v]) => {
                if (v === 0) return null;
                return (
                  <div key={k} className="flex justify-between py-1 border-b border-slate-100 last:border-0">
                    <span className="text-xs text-slate-500 truncate mr-2">{LOT_LABELS[k]?.split(" ")[0]}</span>
                    <span className="text-xs font-mono font-medium text-slate-700">{Math.round(v * COEFF_LOC[form.localisation]?.coeff)} €/m²</span>
                  </div>
                );
              })}
            </Card>

            <Card className="p-5 bg-slate-900 border-slate-900">
              <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-3">Workflow BIM → Estimation</p>
              {[
                { n:"1", t:"Import IFC", s:"SHAB, pièces, matériaux", c:"text-violet-400" },
                { n:"2", t:"Pré-remplissage auto", s:"Surface, gamme, spécificités", c:"text-blue-400" },
                { n:"3", t:"Ajustement manuel", s:"Vérifiez et corrigez", c:"text-amber-400" },
                { n:"4", t:"Calcul + Export Excel", s:"Honoraires fidèles au modèle", c:"text-emerald-400" },
              ].map(({ n, t, s, c }) => (
                <div key={n} className="flex items-start gap-3 mb-3 last:mb-0">
                  <span className={`text-xs font-bold ${c} w-4 flex-shrink-0`}>{n}</span>
                  <div>
                    <p className={`text-xs font-semibold ${c}`}>{t}</p>
                    <p className="text-xs text-slate-400">{s}</p>
                  </div>
                </div>
              ))}
            </Card>
          </div>
        </div>
      </div>
    </div>
  );

  // ─── RÉSULTATS ─────────────────────────────────────────────────────────────
  if (!result) return null;
  const { lots, esq, apd, totalTravaux, hon, totalOperation, alertes, surface } = result;
  const maxLot = Math.max(...Object.values(lots).map(v => v.montant));

  return (
    <div className="min-h-screen bg-slate-50" ref={resultRef}>
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="flex items-center gap-3">
              <AMLLogo className="h-8 w-auto" />
              <div className="border-l border-slate-200 pl-3">
                <p className="text-xs font-bold text-slate-900 tracking-widest uppercase">Atelier Marie Leguillon</p>
                <p className="text-xs text-slate-500">{RATIOS[form.typeProjet]?.label} · {surface} m² · {form.gamme}{ifcData ? ` · IFC ${ifcData.phase_detectee}` : ""}</p>
              </div>
            </div>
            <button onClick={handleExport} disabled={downloading} className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-semibold transition-all shadow-sm ${downloading?"bg-emerald-100 text-emerald-600":"bg-emerald-600 text-white hover:bg-emerald-700"}`}>
              {downloading ? <svg className="w-4 h-4 animate-spin" fill="none" viewBox="0 0 24 24"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/></svg> : <IconDownload />}
              {downloading ? "Génération…" : "Télécharger Excel"}
            </button>
            <button onClick={reset} className="text-sm text-slate-500 hover:text-slate-900 font-medium">← Nouvelle</button>
          </div>
        </div>
      </header>

      <div className="max-w-6xl mx-auto px-6 py-8">
        {/* KPIs */}
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
          {[
            { l:"Budget travaux", v:fmt(totalTravaux), s:`${fmt(totalTravaux/surface)}/m²`, c:"bg-slate-900 text-white" },
            { l:"Honoraires architecte", v:fmt(hon.honorairesArchitecte), s:`Taux ${(hon.taux*100).toFixed(1)}%`, c:"bg-white" },
            { l:"Études & OPC", v:fmt(hon.honorairesBET+hon.honorairesOPC+hon.honorairesSPS), s:"BET + OPC + SPS", c:"bg-white" },
            { l:"TOTAL OPÉRATION", v:fmt(totalOperation), s:"Toutes prestations HT", c:"bg-emerald-600 text-white" },
          ].map(({l,v,s,c})=>(
            <div key={l} className={`${c} rounded-2xl border ${c.includes("bg-white")?"border-slate-200":"border-0"} shadow-sm p-5`}>
              <p className={`text-xs font-semibold uppercase tracking-wider mb-1 ${c.includes("text-white")?"text-white/70":"text-slate-500"}`}>{l}</p>
              <p className={`text-xl font-bold ${c.includes("text-white")?"text-white":"text-slate-900"}`}>{v}</p>
              <p className={`text-xs mt-0.5 ${c.includes("text-white")?"text-white/60":"text-slate-400"}`}>{s}</p>
            </div>
          ))}
        </div>

        {/* Badge IFC si présent */}
        {ifcData && (
          <div className="mb-5 px-4 py-3 bg-violet-50 border border-violet-200 rounded-xl flex items-center gap-3 text-sm">
            <div className="w-6 h-6 bg-violet-600 rounded-full flex items-center justify-center flex-shrink-0"><IconCheck /></div>
            <div>
              <span className="font-semibold text-violet-800">Estimation enrichie par BIM</span>
              <span className="text-violet-600 ml-2">— {ifcData.fichier} (phase {ifcData.phase_detectee})</span>
              {ifcData.shab_total > 0 && <span className="text-violet-500 ml-2">· SHAB réelle {ifcData.shab_total} m²</span>}
              {ifcData.pieces.length > 0 && <span className="text-violet-500 ml-2">· {ifcData.pieces.length} pièces</span>}
            </div>
          </div>
        )}

        {/* Alertes */}
        {alertes.map((a,i)=>(
          <div key={i} className={`mb-5 px-4 py-3 rounded-lg flex items-center gap-2 text-sm font-medium ${a.type==="danger"?"bg-rose-50 text-rose-700 border border-rose-200":a.type==="warning"?"bg-amber-50 text-amber-700 border border-amber-200":"bg-emerald-50 text-emerald-700 border border-emerald-200"}`}>
            {a.type==="danger"?"⚠️":a.type==="warning"?"◉":"✓"} {a.msg}
          </div>
        ))}

        {/* Tabs */}
        <div className="flex gap-1 bg-slate-100 p-1 rounded-xl mb-6 overflow-x-auto">
          {[{k:"esq",l:"ESQ"},{k:"aps",l:"APS"},{k:"apd",l:"APD"},{k:"honoraires",l:"HONORAIRES"},{k:"synthese",l:"SYNTHÈSE"},{k:"bim",l:"BIM / IFC",hide:!ifcData}].filter(t=>!t.hide).map(t=>(
            <button key={t.k} onClick={()=>setActiveTab(t.k)} className={`flex-1 min-w-[70px] py-2 px-3 rounded-lg text-xs font-semibold transition-all whitespace-nowrap ${activeTab===t.k?"bg-white text-slate-900 shadow-sm":"text-slate-500 hover:text-slate-700"}`}>
              {t.l}
            </button>
          ))}
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div className="lg:col-span-2">

            {/* ESQ */}
            {activeTab==="esq" && (
              <Card className="p-6">
                <div className="flex items-center justify-between mb-5">
                  <div><h3 className="font-semibold text-slate-900">ESQ — Esquisse</h3><p className="text-xs text-slate-500 mt-0.5">Estimation macro à ±30%</p></div>
                  <Badge color="violet">ESQ</Badge>
                </div>
                {Object.entries(esq).map(([label,montant])=>(
                  <div key={label} className="mb-4">
                    <div className="flex justify-between text-sm mb-1.5"><span className="font-medium text-slate-700">{label}</span><span className="font-bold text-slate-900">{fmt(montant)}</span></div>
                    <ProgressBar value={montant} max={totalTravaux*.8}/>
                  </div>
                ))}
                <div className="mt-5 pt-4 border-t border-slate-200 flex justify-between">
                  <span className="font-semibold text-slate-700">Total ESQ</span>
                  <span className="text-xl font-bold text-slate-900">{fmt(totalTravaux)}</span>
                </div>
                <p className="text-xs text-slate-400 mt-1">Fourchette : {fmt(totalTravaux*.70)} — {fmt(totalTravaux*1.30)}</p>
              </Card>
            )}

            {/* APS */}
            {activeTab==="aps" && (
              <Card className="p-6">
                <div className="flex items-center justify-between mb-5">
                  <div><h3 className="font-semibold text-slate-900">APS — Avant-Projet Sommaire</h3><p className="text-xs text-slate-500 mt-0.5">Par lot à ±20%</p></div>
                  <Badge color="blue">APS</Badge>
                </div>
                <div className="space-y-3">
                  {Object.entries(lots).map(([key,{ratio,montant}])=>(
                    <div key={key}>
                      <div className="flex justify-between items-center mb-1">
                        <div><span className="text-sm font-medium text-slate-700">{LOT_LABELS[key]}</span><span className="text-xs text-slate-400 ml-2">{Math.round(ratio*result.locCoeff)} €/m²</span></div>
                        <span className="text-sm font-bold text-slate-900">{fmt(montant)}</span>
                      </div>
                      <ProgressBar value={montant} max={maxLot}/>
                    </div>
                  ))}
                </div>
                <div className="mt-5 pt-4 border-t border-slate-200 flex justify-between">
                  <span className="font-semibold text-slate-700">Total APS</span>
                  <span className="text-xl font-bold text-slate-900">{fmt(totalTravaux)}</span>
                </div>
              </Card>
            )}

            {/* APD */}
            {activeTab==="apd" && (
              <Card className="p-6">
                <div className="flex items-center justify-between mb-5">
                  <div><h3 className="font-semibold text-slate-900">APD — Avant-Projet Définitif</h3><p className="text-xs text-slate-500 mt-0.5">Fourchettes ±15%</p></div>
                  <Badge color="emerald">APD</Badge>
                </div>
                <table className="w-full text-sm">
                  <thead><tr className="border-b border-slate-200">
                    <th className="text-left py-2 text-xs font-semibold text-slate-500 uppercase">Lot</th>
                    <th className="text-right py-2 text-xs font-semibold text-slate-500 uppercase">Min</th>
                    <th className="text-right py-2 text-xs font-semibold text-slate-500 uppercase">Base</th>
                    <th className="text-right py-2 text-xs font-semibold text-slate-500 uppercase">Max</th>
                  </tr></thead>
                  <tbody>
                    {Object.entries(apd).map(([key,{montant,min,max}])=>(
                      <tr key={key} className="border-b border-slate-100 hover:bg-slate-50">
                        <td className="py-2.5 font-medium text-slate-700">{LOT_LABELS[key]}</td>
                        <td className="py-2.5 text-right text-slate-400 font-mono text-xs">{fmt(min)}</td>
                        <td className="py-2.5 text-right font-bold text-slate-900 font-mono">{fmt(montant)}</td>
                        <td className="py-2.5 text-right text-slate-400 font-mono text-xs">{fmt(max)}</td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot><tr className="bg-slate-50">
                    <td className="py-3 font-bold">TOTAL</td>
                    <td className="py-3 text-right text-slate-500 font-mono text-sm">{fmt(totalTravaux*.85)}</td>
                    <td className="py-3 text-right font-bold text-lg font-mono">{fmt(totalTravaux)}</td>
                    <td className="py-3 text-right text-slate-500 font-mono text-sm">{fmt(totalTravaux*1.15)}</td>
                  </tr></tfoot>
                </table>
              </Card>
            )}

            {/* HONORAIRES */}
            {activeTab==="honoraires" && (
              <Card className="p-6">
                <div className="flex items-center justify-between mb-5">
                  <div><h3 className="font-semibold text-slate-900">Tableau d'honoraires — 4 intervenants</h3><p className="text-xs text-slate-500 mt-0.5">Format ATELIER MARIE LEGUILLON · Barème UNTEC</p></div>
                </div>
                <div className="bg-slate-50 rounded-lg p-3 mb-4 text-sm">
                  <div className="flex justify-between py-1"><span className="text-slate-600">Travaux de base HT</span><span className="font-bold text-red-600">{fmt(totalTravaux)}</span></div>
                  <div className="flex justify-between py-1"><span className="text-slate-600">Taux MOE conception</span><span className="font-mono">{((hon.phases.DIAG+hon.phases.DEPOT_DP+hon.phases.PRO_DCE+hon.phases.ACT)/totalTravaux*100).toFixed(1)}%</span></div>
                  <div className="flex justify-between py-1"><span className="text-slate-600">Taux EXE</span><span className="font-mono">{((hon.phases.DET+hon.phases.AOR)/totalTravaux*100).toFixed(1)}%</span></div>
                  <div className="flex justify-between py-1 border-t mt-1 pt-2"><span className="font-semibold">Rémunération MOE + OPC TTC</span><span className="font-bold">{fmt((hon.honorairesArchitecte+hon.honorairesOPC)*1.2)}</span></div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-xs border-collapse">
                    <thead>
                      <tr className="bg-slate-100">
                        <th className="border border-slate-300 p-2 text-left">MISSIONS</th>
                        <th className="border border-slate-300 p-2 text-center">TAUX PAR PHASE</th>
                        <th className="border border-slate-300 p-2 text-center">MONTANT HT</th>
                        <th className="border border-slate-300 p-2 text-center bg-amber-50" colSpan={2}>MOE Architecte</th>
                        <th className="border border-slate-300 p-2 text-center" colSpan={2}>MOE EXE/OPC</th>
                        <th className="border border-slate-300 p-2 text-center" colSpan={2}>BET PLURI</th>
                      </tr>
                      <tr className="bg-slate-50 text-slate-500">
                        <th className="border border-slate-300 p-1.5"></th>
                        <th className="border border-slate-300 p-1.5"></th>
                        <th className="border border-slate-300 p-1.5"></th>
                        {["Forfait %","Montant HT","Forfait %","Montant HT","Forfait %","Montant HT"].map((h,i)=>(
                          <th key={i} className="border border-slate-300 p-1.5 text-center font-medium">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      <tr className="bg-slate-50"><td className="border border-slate-300 p-2 font-bold" colSpan={9}>MISSION MOE CONCEPTION</td></tr>
                      {[
                        { l:"DIAG", v:hon.phases.DIAG },
                        { l:"DEPOT DP / AT / DOSSIER SURETE / ENSEIGNE", v:hon.phases.DEPOT_DP },
                        { l:"PRO / DCE / DQE", v:hon.phases.PRO_DCE },
                        { l:"ACT - DOSSIER MARCHE", v:hon.phases.ACT },
                      ].map(({ l, v }) => {
                        const tot = hon.phases.DIAG+hon.phases.DEPOT_DP+hon.phases.PRO_DCE+hon.phases.ACT;
                        return (
                          <tr key={l} className="hover:bg-amber-50/30">
                            <td className="border border-slate-300 p-2">{l}</td>
                            <td className="border border-slate-300 p-2 text-center">{(v/totalTravaux*100).toFixed(2)}%</td>
                            <td className="border border-slate-300 p-2 text-right font-medium">{fmt(v)}</td>
                            <td className="border border-slate-300 p-2 text-center text-blue-700">{(v/tot*100).toFixed(2)}%</td>
                            <td className="border border-slate-300 p-2 text-right font-bold text-blue-700">{fmt(v)}</td>
                            <td className="border border-slate-300 p-2 text-center text-slate-400">0,00%</td><td className="border border-slate-300 p-2 text-right text-slate-400">0</td>
                            <td className="border border-slate-300 p-2 text-center text-slate-400">0,00%</td><td className="border border-slate-300 p-2 text-right text-slate-400">0</td>
                          </tr>
                        );
                      })}
                      <tr className="bg-amber-50 font-bold">
                        <td className="border border-slate-300 p-2">TOTAL CONCEPTION</td>
                        <td className="border border-slate-300 p-2 text-center">{((hon.phases.DIAG+hon.phases.DEPOT_DP+hon.phases.PRO_DCE+hon.phases.ACT)/totalTravaux*100).toFixed(2)}%</td>
                        <td className="border border-slate-300 p-2 text-right">{fmt(hon.phases.DIAG+hon.phases.DEPOT_DP+hon.phases.PRO_DCE+hon.phases.ACT)}</td>
                        <td className="border border-slate-300 p-2 text-center">62,79%</td>
                        <td className="border border-slate-300 p-2 text-right">{fmt(hon.phases.DIAG+hon.phases.DEPOT_DP+hon.phases.PRO_DCE+hon.phases.ACT)}</td>
                        <td colSpan={4} className="border border-slate-300 p-2 text-center text-slate-400">0,00%</td>
                      </tr>
                      <tr className="bg-slate-50"><td className="border border-slate-300 p-2 font-bold" colSpan={9}>MISSION EXE</td></tr>
                      {[
                        { l:"DET/VISA TRAVAUX", v:hon.phases.DET },
                        { l:"AOR", v:hon.phases.AOR },
                        { l:"DOE/ Mis à jour conception des plans EXE", v:0 },
                      ].map(({ l, v }) => (
                        <tr key={l}>
                          <td className="border border-slate-300 p-2">{l}</td>
                          <td className="border border-slate-300 p-2 text-center">{v>0?(v/totalTravaux*100).toFixed(2)+"%":"0,00%"}</td>
                          <td className="border border-slate-300 p-2 text-right">{fmt(v)}</td>
                          <td className="border border-slate-300 p-2 text-center text-blue-700">{v>0?(v/(hon.phases.DET+hon.phases.AOR)*100).toFixed(2)+"%":"0,00%"}</td>
                          <td className="border border-slate-300 p-2 text-right font-bold text-blue-700">{fmt(v)}</td>
                          <td colSpan={4} className="border border-slate-300 p-2 text-center text-slate-400">0,00%</td>
                        </tr>
                      ))}
                      <tr className="bg-amber-50 font-bold">
                        <td className="border border-slate-300 p-2">TOTAL CHANTIER</td>
                        <td className="border border-slate-300 p-2 text-center">{((hon.phases.DET+hon.phases.AOR)/totalTravaux*100).toFixed(2)}%</td>
                        <td className="border border-slate-300 p-2 text-right">{fmt(hon.phases.DET+hon.phases.AOR)}</td>
                        <td className="border border-slate-300 p-2 text-center">0,00%</td>
                        <td className="border border-slate-300 p-2 text-right">{fmt(hon.phases.DET+hon.phases.AOR)}</td>
                        <td colSpan={4} className="border border-slate-300 p-2 text-center text-slate-400">0,00%</td>
                      </tr>
                      <tr className="bg-amber-50 font-bold">
                        <td className="border border-slate-300 p-2">TOTAL OPC</td>
                        <td className="border border-slate-300 p-2 text-center">100,00%</td>
                        <td className="border border-slate-300 p-2 text-right">{fmt(hon.honorairesOPC)}</td>
                        <td colSpan={6} className="border border-slate-300 p-2"></td>
                      </tr>
                      <tr className="bg-blue-50 font-bold text-sm">
                        <td className="border border-slate-400 p-2">TOTAUX HT</td>
                        <td className="border border-slate-400 p-2 text-center">100,00%</td>
                        <td className="border border-slate-400 p-2 text-right">{fmt(hon.honorairesArchitecte+hon.honorairesOPC)}</td>
                        <td className="border border-slate-400 p-2 text-center">100,00%</td>
                        <td className="border border-slate-400 p-2 text-right">{fmt(hon.honorairesArchitecte)}</td>
                        <td className="border border-slate-400 p-2 text-center text-slate-400">0,00%</td><td className="border border-slate-400 p-2 text-right text-slate-400">0</td>
                        <td className="border border-slate-400 p-2 text-center">{(hon.honorairesBET/totalTravaux*100).toFixed(2)}%</td>
                        <td className="border border-slate-400 p-2 text-right">{fmt(hon.honorairesBET)}</td>
                      </tr>
                      <tr className="bg-amber-100 font-bold text-sm">
                        <td className="border border-slate-400 p-2">TOTAUX TTC<br/><span className="text-rose-500 font-normal text-xs">TAUX TVA à revoir en fonction au montant travaux par poste</span></td>
                        <td className="border border-slate-400 p-2 text-center">100,00%</td>
                        <td className="border border-slate-400 p-2 text-right">{fmt((hon.honorairesArchitecte+hon.honorairesOPC)*1.2)}</td>
                        <td className="border border-slate-400 p-2 text-center">100,00%</td>
                        <td className="border border-slate-400 p-2 text-right">{fmt(hon.honorairesArchitecte*1.2)}</td>
                        <td className="border border-slate-400 p-2 text-center text-slate-400">0,00%</td><td className="border border-slate-400 p-2 text-right text-slate-400">0</td>
                        <td className="border border-slate-400 p-2 text-center">{(hon.honorairesBET/totalTravaux*100).toFixed(2)}%</td>
                        <td className="border border-slate-400 p-2 text-right">{fmt(hon.honorairesBET*1.2)}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </Card>
            )}

            {/* SYNTHESE */}
            {activeTab==="synthese" && (
              <Card className="p-6">
                <h3 className="font-semibold text-slate-900 mb-5">Synthèse — Budget total opération</h3>
                {[
                  { l:"Travaux HT", v:totalTravaux },
                  { l:"Honoraires architecte HT", v:hon.honorairesArchitecte },
                  { l:"BET (Bureaux d'Études) HT", v:hon.honorairesBET },
                  { l:"OPC HT", v:hon.honorairesOPC },
                  { l:"Coordination SPS HT", v:hon.honorairesSPS },
                ].map(({l,v})=>(
                  <div key={l} className="flex justify-between items-center py-2.5 border-b border-slate-100">
                    <span className="text-sm text-slate-600">{l}</span>
                    <span className="text-sm font-medium text-slate-800">{fmt(v)}</span>
                  </div>
                ))}
                <div className="mt-4 p-3 bg-amber-50 rounded-lg border border-amber-200">
                  <p className="text-xs font-semibold text-amber-700 mb-2">Provisions recommandées</p>
                  {[{l:"Aléas travaux (7.5%)",v:totalTravaux*.075},{l:"Contrôle technique",v:totalTravaux*.012},{l:"Branchements / VRD",v:totalTravaux*.02}].map(({l,v})=>(
                    <div key={l} className="flex justify-between text-xs py-1"><span className="text-amber-700">{l}</span><span className="font-medium">≈ {fmt(v)}</span></div>
                  ))}
                </div>
                <div className="mt-5 pt-4 border-t-2 border-slate-900 flex justify-between items-center">
                  <span className="font-bold text-slate-900">TOTAL OPÉRATION HT</span>
                  <span className="text-2xl font-bold text-slate-900">{fmt(totalOperation)}</span>
                </div>
                <p className="text-xs text-slate-400 mt-1">{fmt(totalOperation/surface)}/m² · hors taxes, hors provisions</p>
              </Card>
            )}

            {/* BIM / IFC */}
            {activeTab==="bim" && ifcData && (
              <Card className="p-6">
                <div className="flex items-center gap-2 mb-5">
                  <div className="w-6 h-6 bg-violet-600 rounded flex items-center justify-center"><IconCheck /></div>
                  <div>
                    <h3 className="font-semibold text-slate-900">Données BIM extraites</h3>
                    <p className="text-xs text-slate-500">{ifcData.fichier} — Phase {ifcData.phase_detectee}</p>
                  </div>
                </div>

                {/* Programme pièces */}
                {ifcData.pieces.length > 0 && (
                  <div className="mb-6">
                    <p className="text-xs font-semibold text-slate-700 uppercase tracking-wider mb-3">Programme — Surfaces par pièce</p>
                    <div className="space-y-2">
                      {ifcData.pieces.map(p => (
                        <div key={p.nom} className="flex justify-between items-center py-2 border-b border-slate-100">
                          <span className="text-sm text-slate-700">{p.nom}</span>
                          {p.surface && <span className="text-sm font-bold text-slate-900 font-mono">{p.surface} m²</span>}
                        </div>
                      ))}
                      <div className="flex justify-between items-center pt-2 font-bold">
                        <span className="text-sm">SHAB TOTALE</span>
                        <span className="text-sm font-mono">{ifcData.shab_total} m²</span>
                      </div>
                    </div>
                  </div>
                )}

                {/* Éléments */}
                <div className="mb-6">
                  <p className="text-xs font-semibold text-slate-700 uppercase tracking-wider mb-3">Éléments BIM comptabilisés</p>
                  <div className="grid grid-cols-4 gap-2">
                    {Object.entries(ifcData.elements).map(([k,v])=>(
                      <div key={k} className="bg-slate-50 rounded-lg p-3 text-center border border-slate-200">
                        <p className="text-lg font-bold text-slate-900">{v}</p>
                        <p className="text-xs text-slate-500 capitalize">{k}</p>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Niveaux */}
                {ifcData.niveaux.length > 0 && (
                  <div className="mb-6">
                    <p className="text-xs font-semibold text-slate-700 uppercase tracking-wider mb-3">Niveaux / Étages</p>
                    {ifcData.niveaux.map(n => (
                      <div key={n.nom} className="flex justify-between text-sm py-1.5 border-b border-slate-100">
                        <span className="text-slate-700">{n.nom}</span>
                        <span className="font-mono text-slate-500">{n.elevation} m</span>
                      </div>
                    ))}
                  </div>
                )}

                {/* Matériaux */}
                {ifcData.materiaux.length > 0 && (
                  <div>
                    <p className="text-xs font-semibold text-slate-700 uppercase tracking-wider mb-3">Matériaux détectés ({ifcData.materiaux.length})</p>
                    <div className="flex flex-wrap gap-1.5">
                      {ifcData.materiaux.filter(m => m.length < 35).map(m => (
                        <span key={m} className="text-xs bg-violet-50 border border-violet-200 text-violet-700 px-2 py-0.5 rounded">{m}</span>
                      ))}
                    </div>
                  </div>
                )}
              </Card>
            )}
          </div>

          {/* Sidebar résultats */}
          <div className="space-y-4">
            <Card className="p-5">
              <p className="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-3">Analyse IA</p>
              {!aiAnalysis && !loadingAI && (
                <button onClick={generateAIAnalysis} className="w-full py-2.5 bg-slate-900 text-white text-sm rounded-lg font-medium hover:bg-slate-800">
                  Générer l'analyse ✦
                </button>
              )}
              {loadingAI && <div className="space-y-2">{[1,2,3].map(i=><div key={i} className="h-3 bg-slate-100 rounded animate-pulse" style={{width:`${70+i*10}%`}}/>)}</div>}
              {aiAnalysis && <div className="text-xs text-slate-600 leading-relaxed whitespace-pre-wrap">{aiAnalysis}</div>}
            </Card>

            <Card className="p-5">
              <p className="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-3">Ratios appliqués</p>
              <p className="text-xs text-slate-400 mb-2">Coeff. {result.locCoeff}×</p>
              {Object.entries(lots).map(([key,{montant}])=>(
                <div key={key} className="flex justify-between py-1.5 border-b border-slate-100 last:border-0">
                  <span className="text-xs text-slate-500 truncate mr-2">{LOT_LABELS[key]?.split(" ")[0]}</span>
                  <span className="text-xs font-mono font-medium text-slate-700">{fmt(montant)}</span>
                </div>
              ))}
            </Card>

            <div className="p-4 bg-emerald-900 rounded-xl">
              <p className="text-xs font-semibold text-emerald-300 uppercase tracking-wider mb-2">Export Excel — 2 feuilles</p>
              {["Feuille 1 : Honoraires (modèle exact)","Tableau 4 intervenants","Phases DIAG/PRO-DCE/ACT/DET/AOR","TOTAUX HT & TTC","Feuille 2 : Estimatif travaux","Données IFC intégrées","Programme pièces","Éléments BIM"].map(item=>(
                <div key={item} className={`flex items-center gap-2 py-0.5 ${item.startsWith("Feuille") ? "mt-2" : ""}`}>
                  <div className={`w-1 h-1 rounded-full flex-shrink-0 ${item.startsWith("Feuille") ? "bg-white" : "bg-emerald-400"}`}/>
                  <span className={`text-xs ${item.startsWith("Feuille") ? "text-white font-semibold" : "text-emerald-200"}`}>{item}</span>
                </div>
              ))}
              <button onClick={handleExport} disabled={downloading} className="w-full mt-3 py-2 bg-emerald-500 text-white text-xs font-semibold rounded-lg hover:bg-emerald-400 transition-colors disabled:opacity-60">
                {downloading?"Génération…":"↓ Télécharger .xlsx"}
              </button>
            </div>

            <div className="px-4 py-3 bg-slate-100 rounded-xl">
              <p className="text-xs text-slate-500"><strong>Note :</strong> Estimation basée sur ratios statistiques ±20%. Les données IFC enrichissent le contexte mais ne se substituent pas à une étude d'économiste.</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
