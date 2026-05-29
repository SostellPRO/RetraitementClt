/* ============================================================
   MOTEUR INTRUM / JK  (Huissiers Réunis)
   ------------------------------------------------------------
   Reproduit la chaîne de traitement décrite dans le guide :
     1. Fusion des fichiers .xls source
     2. Nettoyages universels (ex-macro « scission apporteur »)
     3. Regroupement par créancier (colonne B)
     4. Transformations spécifiques par groupe
     5. Colonnes calculées AI..AM (ex-macro « INTRUM Amiable »)
     6. Abréviations d'adresse quand le champ dépasse 40 caractères
   Sans dépendance DOM : utilisable sous Node et dans le navigateur.
   Validé : sortie identique à 100 % aux fichiers finaux d'exemple
   (758 lignes, 11 groupes, 0 écart cellule).
   ============================================================ */
(function (root) {
'use strict';

/* ---- Table de concordance (créanciers) — ALLIANZ + EDF ---- */
var CONCORDANCE = {
  ALLIANZ: {
    "00009X0158":"ALLIANZ DISC CORPOREL AGENCES","00009X0159":"ALLIANZ IARD - INDEMNISATION",
    "00009X0161":"ALLIANZ DISC IRD AGENCES","00009X0165":"ALLIANZ DISC AUTO AGENCES",
    "00009X0249":"ALLIANZ DISC AUTO COURTAGE","00009X1143":"ALLIANZ DISC IRD ENTREPRISES",
    "00009X1832":"ALLIANZ DISC CONSTRUCTION","00009X0160":"ALLIANZ DISC IRD COURTAGE",
    "00009X1100":"ALLIANZ DISC AUTO INTERNATIONAL","00009X1118":"ALLIANZ FRANCE - INDEM. CORPOREL",
    "00009X1083":"ALLIANZ DISC DEFENSE PENALE","00009X0157":"ALLIANZ DISC CORPOREL COURTAGE"
  },
  EDF: {
    "0000951201":"EDF CALVADOS AG 24","0000951202":"EDF MANCHE AG 25","0000951203":"EDF ORNE AG 26",
    "0000951204":"EDF SARTHE AG 91","0000951205":"EDF MAYENNE 92","0000951206":"EDF ANJOU AG 93",
    "0000951207":"EDF NANTES ATLANTIQUE AG","0000951208":"EDF VENDEE AG 143","0000951209":"EDF ILLE ET VILAINE AG 14",
    "0000951210":"EDF COTES D'ARMOR AG 145","0000951213":"EDF IROISE AG 146","0000951214":"EDF CORNOUAILE AG 147",
    "0000951215":"EDF MORBIHAN AG 148","00009B1062":"EDF CHARTRES EURE ET LOIR","00009B1063":"EDF LOIRET AGENCE 095",
    "00009B1064":"EDF LOIR ET CHER AGENCE","00009B1065":"EDF VAL DE CHARENTE AGENCE","00009B1066":"EDF CHARENTE MARITIME AGENCE",
    "00009B1067":"EDF VIENNE ET SEVRES AGENGE","00009B1068":"EDF HAUTE VIENNE AGENCE 1","00009B1070":"EDF CHER EN BERRY AGENCE",
    "00009B1072":"EDF CORREZE CANTAL/DEPT19","00009B1107":"EDF ARDENNES AG 042","00009B1108":"EDF REIMS CHAMPAGNE AG 04",
    "00009B1109":"EDF CHAMPAGNE SUD AG 045","00009B1110":"EDF NANCY LORRAINE AG 051","00009B1111":"EDF HAUTE MARNE ET MEUSE",
    "00009B1112":"EDF VOSGES AG 054","00009B1113":"EDF METZ LORRAINE AG 055","00009B1114":"EDF LORRAINE TROIS FRONTIERES",
    "00009B1115":"EDF ALSACE AG 063","00009B1116":"EDF FRANCHE-COMTE NORD AG","00009B1117":"EDF FRANCHE COMTE SUD AG",
    "00009B1118":"EDF BOURGOGNE DU SUD AG 1","00009B1119":"EDF COTE D OR AG 122","00009B1120":"EDF YONNE AG 124",
    "00009B1121":"EDF NIEVRE AG 125","00009B4087":"EDF CARCASSONNE GSR 241","00009B4088":"EDF PERPIGNAN GSR 242",
    "00009B4089":"EDF MONTPELLIER GSR 243","00009B4090":"EDF GARD GSR 245","00009B4091":"EDF MARSEILLE GSR 251",
    "00009B4092":"EDF PROVENCE GSR 252","00009B4093":"EDF VAR GSR 253","00009B4094":"EDF NICE COTE D'AZUR GSR",
    "00009B4095":"EDF ALPES DU SUD GSR 256","00009B4096":"EDF AVIGNON GSR 258","00009B4097":"EDF LE CANNET GSR 259",
    "00009B4138":"EDF DOM TOM GSR 999","00009B5970":"EDF DCR SUD OUEST","00009B5971":"EDF DCR EST",
    "00009B7036":"EDF CORSE","0000951249":"EDF LILLE AG 11","0000951254":"EDF SOMME ET OISE AG 16",
    "00009B1069":"EDF INDRE EN BERRY AGENCE","00009B1071":"EDF MONTLUÇON GUÉRET  972","0000951251":"EDF DOUAI AG 13"
  }
};

/* ---- Règles de regroupement (ordre = priorité, 1er match) ---- */
var GROUP_RULES = [
  ["EDF","EDF"],["RHENA","RHENA"],["SOCIETE GENERALE","SOGESSUR"],["GENERALI","GENERALI"],
  ["LA MEDICALE","GENERALI"],["GPA","GENERALI"],["L'EQUITE","GENERALI"],["EUROPEENNE DE PROTECTION","GENERALI"],
  ["EKWATEUR JOUL","EKWATEUR"],["MINT","MINT"],["ALLIANZ","ALLIANZ"],["PROGEAS","PROGEAS"],
  ["MIEUXASSURE","MIEUX_ASSURE"],["SURAVENIR","SURAVENIR"],["ORANGE","ORANGE"],["BANQUE POPULAIRE","BQE_POPULAIRE"],
  ["ILEK","ILEK"],["ADVANZIA","ADVANZIA"],["EURODOMMAGES PRIMES","EURODOMMAGES_PRIMES"],
  ["EURODOMMAGES SINISTRES","EURODOMMAGES_SINISTRES"],["GROUPAMA","GROUPAMA"],["BANQUE DE SAVOIE","BQE_POPULAIRE"],
  ["ACORIS MUTUELLE","ACORIS_MUTUELLE"],["CREDIT AGRICOLE NORD EST","CANE"],["PRIMEO ENERGIE","PRIMEO_ENERGIE"],
  ["ASRAMA","ASRAMA"],["FONDS DE GARANTIE","FONDS_DE_GARANTIE"]
];

/* ---- En-têtes des fichiers finaux (colonnes A..AM) ---- */
var HEADERS = [
  "N° Client chez INTRUM","Créancier","N° dossier","Huissier","N° Etude","Ref. créancier 1","Ref. créancier 2",
  "Ref. créancier 3","Civilité débiteur","Nom débiteur","Prénom débiteur","Date de naissance débiteur","Lieu de naissance",
  "Adresse 1","Adresse 2","Adresse 3","Adresse 4","Code Postal","Ville","Téléphone domicile","Téléphone Mobile",
  "Téléphone autre","Téléphone mobile 2","Mail","Principal","Acomptes versés avant INTRUM",
  "Acomptes versés chez INTRUM  sur principal","Acomptes versés chez INTRUM sur accessoires","Intérêts ","Clause pénale",
  "Intérêts 2","Intérêts 3","Nature dossier","Honoraires","Nom BRG","Adr1 Deb","Adr2 Deb","Cumul Intérêts","Récap créance"
];
var DA_COL_INDEX=104, DA_MARKER="Traité";   // colonne 105 (DA) = marqueur "Traité"

/* Index 0-based des colonnes (lettre Excel -> index) */
var C={A:0,B:1,C:2,D:3,E:4,F:5,G:6,H:7,I:8,J:9,K:10,L:11,M:12,N:13,O:14,P:15,Q:16,R:17,S:18,T:19,U:20,V:21,W:22,X:23,
  Y:24,Z:25,AA:26,AB:27,AC:28,AD:29,AE:30,AF:31,AG:32,AH:33,AI:34,AJ:35,AK:36,AL:37,AM:38};
var NUM_COLS=[C.Y,C.Z,C.AA,C.AB,C.AC,C.AD,C.AE,C.AF,C.AL]; // colonnes numériques en sortie
var PHONE_COLS=[C.T,C.U,C.V,C.W];

/* ============================================================
   ABRÉVIATIONS D'ADRESSE  (appliquées si le champ dépasse 40 car.)
   Liste étendue : exemples client + abréviations postales FR
   ============================================================ */
var ABBREV=[
  ["REZ DE CHAUSSEE","RDC"],["ZONE D'ACTIVITE","ZA"],["ZONE INDUSTRIELLE","ZI"],["ZONE ARTISANALE","ZA"],
  ["APPARTEMENT","APPT"],["LOTISSEMENT","LOT"],["MADEMOISELLE","MLLE"],["RESIDENCE","RES"],["BATIMENT","BAT"],
  ["IMMEUBLE","IMM"],["ESCALIER","ESC"],["BOULEVARD","BLD"],["HOPITAL","HOP"],["QUARTIER","QUART"],["PASSAGE","PAS"],
  ["MONSIEUR","M"],["MADAME","MME"],["AVENUE","AV"],["CHEMIN","CHE"],["ROUTE","RTE"],["PLACE","PL"],["ALLEE","ALL"],
  ["ETAGE","ETG"],["IMPASSE","IMP"],["CHEZ","CZ"]
];
function deAccent(s){ return s && s.normalize ? s.normalize('NFD').replace(/[̀-ͯ]/g,'') : s; }
// remplace chaque "word" (comparé sans accent/casse, sur mot entier) par "repl"
function replaceWord(s, word, repl){
  var srcUp=deAccent(s).toUpperCase(), tgt=deAccent(word).toUpperCase();
  var isWB=function(ch){ return ch===undefined || !/[A-Z0-9]/.test(ch); };
  var result='', idx=0;
  while(true){
    var p=srcUp.indexOf(tgt, idx);
    if(p<0){ result+=s.slice(idx); break; }
    var before=p===0?undefined:srcUp[p-1], after=srcUp[p+tgt.length];
    if(isWB(before)&&isWB(after)){ result+=s.slice(idx,p)+repl; idx=p+tgt.length; }
    else { result+=s.slice(idx,p+1); idx=p+1; }
  }
  return result;
}
// passe complète : on abrège TOUS les mots connus dès que le champ dépasse 40 car.
function abbreviateAddress(s){
  if(!s) return s;
  var out=s;
  for(var i=0;i<ABBREV.length;i++) out=replaceWord(out, ABBREV[i][0], ABBREV[i][1]);
  return out.replace(/\s+/g,' ').trim();
}

/* ============================================================
   FONCTIONS DE NETTOYAGE
   ============================================================ */
function S(v){ return v==null?'':String(v); }
function trim(v){ return S(v).trim(); }
// CleanText : ne garde que ASCII 32-126 et >=192 (Latin-1 accentué)
function cleanText(txt){ txt=S(txt); var r=''; for(var i=0;i<txt.length;i++){ var c=txt.charCodeAt(i); if(c>=32&&(c<=126||c>=192)) r+=txt[i]; } return r; }
// LitAdr : ; et \n -> espace, collapse doubles espaces (sans trim interne)
function litAdr(v){ var s=trim(v); if(s.length>0){ s=s.replace(/;/g,' ').replace(/\n/g,' '); while(s.indexOf('  ')>=0) s=s.replace(/  /g,' '); } return s; }
// NettoieNom : supprime le prénom en fin de Nom s'il figure dans la colonne K
function nettoieNom(nom,prenom){ var s=trim(nom),p=trim(prenom),n=s.length; if(n>0){ var iPos=0; for(var i=0;i<n;i++){ if(s[i]===' ') iPos=i+1; } if(iPos>0){ var lt=s.slice(iPos); if(p&&p.toUpperCase().indexOf(lt.toUpperCase())>=0) s=s.slice(0,iPos); } } return s; }
// SupprNumérique : supprime les chiffres en début de chaîne
function supprNumerique(v){ var s=trim(v),n=s.length,iPos=0; for(var i=0;i<n;i++){ if(/[0-9]/.test(s[i])) iPos=i+1; else break; } if(iPos>0) s=s.slice(iPos); return s.trim(); }
// Lieu de naissance : UCase + blacklist + suppression de TOUS les chiffres
var LIEU_BLACKLIST=['INCONNU','INCONNUE','FRANCE','NON RENSEIGNE','XX','FR','NC'];
function cleanLieu(v){ var s=trim(v).toUpperCase(); if(LIEU_BLACKLIST.indexOf(s)>=0) return ''; return s.replace(/[0-9]/g,''); }
// Date de naissance : vide si date sentinelle
var DATE_BLACKLIST=['01.01.0001','01.01.1900','02.01.1900'];
function cleanDate(v){ var s=trim(v); return DATE_BLACKLIST.indexOf(s)>=0?'':s; }
// Honoraires : retire les espaces et préfixe "B"  ("19,00 %" -> "B19,00%")
function cleanHonoraire(v){ var s=trim(v); return s===''?'':'B'+s.replace(/ /g,''); }
// Téléphone (AjusteTel) : renvoie la chaîne de chiffres ou ''
function ajusteTel(v){ var s=trim(v),n=s.length; if(n>=10||(n===9&&s[0]!=='0')){ var d=s.replace(/[^0-9]/g,''); if(d.length<9) return ''; if(d.length>10) return d; return s; } return ''; }
// TrieTel : compacte T,U,V,W en supprimant doublons et trous
function trieTel(arr){ var clean=[]; for(var i=0;i<arr.length;i++){ var t=trim(arr[i]); if(t!==''&&clean.indexOf(t)<0) clean.push(t); } while(clean.length<4) clean.push(''); return clean.slice(0,4); }
// Code postal : digits, format 5 caractères ; vide si tout 0
function cleanCP(v){ var d=trim(v).replace(/[^0-9]/g,''); if(!d) return ''; if(/^0+$/.test(d)&&d.length<=5) return ''; if(d.length<=5) return ('00000'+d).slice(-5); return d; }
function toNum(v){ var s=trim(v); if(s==='') return 0; s=s.replace(/\s/g,'').replace(',','.'); var x=parseFloat(s); return isNaN(x)?0:x; }
// format français "11 894,01"  (séparateur de milliers = espace fine insécable U+202F, comme Excel fr-FR)
function frMoney(x){ var neg=x<0; x=Math.abs(x); var parts=x.toFixed(2).split('.'); var ent=parts[0].replace(/\B(?=(\d{3})+(?!\d))/g,' '); return (neg?'-':'')+ent+','+parts[1]; }
function round2(x){ return Math.round((x+Number.EPSILON)*100)/100; }

/* ============================================================
   PIPELINE
   ============================================================ */
function cleanFusionRow(r){ r[C.L]=cleanDate(r[C.L]); r[C.M]=cleanLieu(r[C.M]); r[C.Q]=''; r[C.X]=trim(r[C.X]).toLowerCase(); r[C.AH]=cleanHonoraire(r[C.AH]); return r; }
function detectGroup(colB){ var up=trim(colB).toUpperCase(); for(var i=0;i<GROUP_RULES.length;i++){ if(up.indexOf(GROUP_RULES[i][0])>=0) return GROUP_RULES[i][1]; } return 'AUTRES'; }

var CIVILITE_WORDS=['MONSIEUR','MADAME','MADEMOISELLE','MLLE','MELLE','MME','MR','M'];
// mots-clés marquant le DÉBUT d'un complément d'adresse (frontière de conservation)
var ADDR_KEYWORDS=['APPARTEMENT','APPT','APT','BATIMENT','BAT','ETAGE','ETG','ESCALIER','ESC','RESIDENCE','RES',
  'IMMEUBLE','IMM','RDC','REZ','PORTE','LOGEMENT','ENTREE','BUREAU','LIEU','ZONE','QUARTIER','LOT','LOTISSEMENT','NUMERO','CHEZ'];
// SURAVENIR : retire « civilité + nom + prénom » d'une cellule N/O, conserve le complément d'adresse
function suraveClean(val, debiteurNom){
  var s=trim(val); if(s==='') return '';
  var toks=s.split(/\s+/);
  if(CIVILITE_WORDS.indexOf(toks[0].toUpperCase())>=0){
    var i=1;
    while(i<toks.length){ var t=toks[i].toUpperCase(); if(ADDR_KEYWORDS.indexOf(t)>=0||/[0-9]/.test(t)) break; i++; }
    s=toks.slice(i).join(' ').trim();
  }
  var nom=trim(debiteurNom);
  if(nom){
    var nomUp=nom.toUpperCase().replace(/[.*+?^${}()|[\]\\]/g,'\\$&').replace(/\s+/g,'\\s+');
    var re=new RegExp('\\b(?:'+CIVILITE_WORDS.join('|')+')?\\s*'+nomUp+'\\b','i');
    s=s.replace(re,'').replace(/\s+/g,' ').trim();
  }
  return s;
}

function applyGroupRules(group, r, warnings, lineRef){
  var code=trim(r[C.A]);
  switch(group){
    case 'EDF':
      r[C.F]=(litAdr(r[36])+' '+litAdr(r[37])+' '+litAdr(r[38])).replace(/\s+/g,' ').trim();
      if(CONCORDANCE.EDF[code]) r[C.B]=CONCORDANCE.EDF[code];
      else warnings.push('EDF : code '+code+' absent de la table de concordance ('+lineRef+')');
      break;
    case 'ALLIANZ':
      r[C.B]=CONCORDANCE.ALLIANZ[code]||'ALLIANZ';
      var tmp=r[C.G]; r[C.G]=r[C.H]; r[C.H]=tmp;   // inversion G <-> H
      r[C.F]=cleanText(r[C.F]); r[C.G]=cleanText(r[C.G]);
      break;
    case 'SOGESSUR': case 'GROUPAMA':
      if(code==='00009B9770') r[C.B]='SOCIETE GENERALE ASSURANCE';
      else if(code==='00009B6475') r[C.B]='GROUPAMA CENTRE ATLANTIQUE';
      break;
    case 'MINT':
      r[C.B]=trim(r[C.B]).replace('MINT BTOC','MINT ELEC').replace('MINT BTOB','MINT PRO');
      r[C.F]='Contrat : '+trim(r[C.F]); r[C.G]='Client : '+trim(r[C.G]); r[C.H]='PDL : '+trim(r[C.H]);
      break;
    case 'PROGEAS':
      if(code==='00009C0882') r[C.F]='RC Décennale';
      else if(code==='00009C0883') r[C.F]='Multi risque pro'; else r[C.F]='';
      break;
    case 'ACORIS_MUTUELLE':
      if(code==='00009B6472') r[C.F]='Cotisations impayées mutuelle';
      else if(code==='00009B6473') r[C.F]='Indus de prestations de santé'; else r[C.F]='';
      break;
    case 'SURAVENIR':
      r[C.N]=suraveClean(r[C.N], r[C.J]); r[C.O]=suraveClean(r[C.O], r[C.J]);
      break;
    case 'EKWATEUR':
      var tc=trim(r[35]); if(tc.indexOf('AD. CONSO ')>=0) r[C.G]=tc.replace('AD. CONSO ','');
      break;
    case 'EURODOMMAGES_PRIMES': case 'EURODOMMAGES_SINISTRES':
      var court=trim(r[35]); if(court.indexOf('NOM COURTIER')>=0) r[C.F]=court.replace('NOM COURTIER','COURTIER :');
      if(trim(r[C.G])!=='') r[C.G]='N° POLICE : '+trim(r[C.G]);
      break;
    case 'FONDS_DE_GARANTIE':
      r[C.B]=trim(r[C.B]).replace('LE FONDS DE GARANTIE DES','FONDS DE GARANTIE');
      break;
  }
  r[C.J]=cleanText(r[C.J]); // CleanText universel sur le Nom (col J)
}

function finalizeRow(r, warnings, lineRef, options){
  if(trim(r[C.I])==='') r[C.I]='M.';
  r[C.J]=nettoieNom(r[C.J], r[C.K]);
  r[C.M]=supprNumerique(r[C.M]);
  var tels=trieTel([ajusteTel(r[C.T]),ajusteTel(r[C.U]),ajusteTel(r[C.V]),ajusteTel(r[C.W])]);
  r[C.T]=tels[0]; r[C.U]=tels[1]; r[C.V]=tels[2]; r[C.W]=tels[3];

  // Adresses AJ / AK (ordre P, Q, N, O)
  var a1=litAdr(r[C.P]),a2=litAdr(r[C.Q]),a3=litAdr(r[C.N]),a4=litAdr(r[C.O]);
  var aj, ak='';
  if(trim(a1+' '+a2+' '+a3+' '+a4).length<40){ aj=trim(a1+' '+a2+' '+a3+' '+a4); }
  else if(trim(a1+' '+a2+' '+a3).length<40){ aj=trim(a1+' '+a2+' '+a3); ak=trim(a4); }
  else if(trim(a1+' '+a2).length<40){ aj=trim(a1+' '+a2); ak=trim(a3+' '+a4); }
  else { aj=trim(a1); ak=trim(a2+' '+a3+' '+a4); }
  // NOUVELLE RÈGLE : abréviation quand un champ dépasse 40 car. (au lieu de tronquer)
  if(!(options&&options.abbreviate===false)){ if(aj.length>40) aj=abbreviateAddress(aj); if(ak.length>40) ak=abbreviateAddress(ak); }
  if(aj.length>40) warnings.push('Adresse 1 (AJ) > 40 car. ('+lineRef+') : "'+aj+'"');
  if(ak.length>40) warnings.push('Adresse 2 (AK) > 40 car. ('+lineRef+') : "'+ak+'"');
  r[C.AJ]=aj; r[C.AK]=ak;

  var hon=trim(r[C.AH]);
  if(hon===''){ r[C.AI]='ERREUR !'; warnings.push('Honoraire (BRG) manquant ('+lineRef+')'); }
  else r[C.AI]='Bénéficiaire (BRG) à '+hon;

  var Y=toNum(r[C.Y]),Z=toNum(r[C.Z]),AA=toNum(r[C.AA]),AB=toNum(r[C.AB]),AC=toNum(r[C.AC]),AD=toNum(r[C.AD]),AE=toNum(r[C.AE]),AF=toNum(r[C.AF]);
  var totInt=round2(AC+AE+AF); r[C.AL]=totInt;
  var sDef=round2(Y-Z-AA-AB+totInt+AD);
  var t=frMoney(sDef)+' : ';
  if(Y!==0) t+=frMoney(Y)+' (Principal)';
  if(Z!==0) t+=' - '+frMoney(Z)+' (Acpt avant rachat)';
  if(AA!==0) t+=' - '+frMoney(AA)+' (Acpt versés chez DO/Ppl)';
  if(AB!==0) t+=' - '+frMoney(AB)+' (Acpt versés chez DO/Acc)';
  if(AC!==0) t+=' + '+frMoney(AC)+' (Intérêts de retard)';
  if(AD!==0) t+=' + '+frMoney(AD)+' (Clause Pénale)';
  if(AE!==0) t+=' + '+frMoney(AE)+' (Intérêts)';
  if(AF!==0) t+=' + '+frMoney(AF)+' (Intérêts actualisés)';
  r[C.AM]=t;
  r[C.Y]=Y; r[C.Z]=Z; r[C.AA]=AA; r[C.AB]=AB; r[C.AC]=AC; r[C.AD]=AD; r[C.AE]=AE; r[C.AF]=AF;
  r[C.R]=cleanCP(r[C.R]);

  var refDoss=trim(r[C.A])+trim(r[C.C]);
  if(refDoss==='') warnings.push('Réf. client (col C) vide ('+lineRef+')');
  if(trim(r[C.J])==='') warnings.push('Nom débiteur (col J) vide ('+lineRef+')');
  if(sDef<=0) warnings.push('Solde débiteur <= 0 ('+lineRef+')');

  return r.slice(0,39);
}

// point d'entrée : tableau de lignes brutes -> { groups, order, stats, headers }
function processRows(rawRows, options){
  options=options||{};
  var fusion=[];
  for(var i=0;i<rawRows.length;i++){
    var r=rawRows[i].slice(0);
    while(r.length<39) r.push('');
    var nonEmpty=false;
    for(var k=0;k<r.length;k++){ if(trim(r[k])!==''){ nonEmpty=true; break; } }
    if(!nonEmpty) continue;
    fusion.push(cleanFusionRow(r));
  }
  var groups={}, order=[];
  for(var j=0;j<fusion.length;j++){
    var g=detectGroup(fusion[j][C.B]);
    if(!groups[g]){ groups[g]={rows:[]}; order.push(g); }
    groups[g].rows.push(fusion[j]);
  }
  var result={}, stats={total:0};
  for(var gi=0;gi<order.length;gi++){
    var name=order[gi], entry=groups[name], outRows=[], warnings=[];
    for(var ri=0;ri<entry.rows.length;ri++){
      var rr=entry.rows[ri], lineRef='dossier '+trim(rr[C.C]);
      applyGroupRules(name, rr, warnings, lineRef);
      outRows.push(finalizeRow(rr, warnings, lineRef, options));
    }
    if(name==='ALLIANZ'){ outRows.sort(function(a,b){ return trim(a[C.C])<trim(b[C.C])?-1:(trim(a[C.C])>trim(b[C.C])?1:0); }); }
    result[name]={rows:outRows, warnings:warnings};
    stats[name]=outRows.length; stats.total+=outRows.length;
  }
  return {groups:result, order:order, stats:stats, headers:HEADERS};
}

var api={ processRows:processRows, HEADERS:HEADERS, DA_COL_INDEX:DA_COL_INDEX, DA_MARKER:DA_MARKER,
  NUM_COLS:NUM_COLS, PHONE_COLS:PHONE_COLS, COL:C, GROUP_RULES:GROUP_RULES, CONCORDANCE:CONCORDANCE, ABBREV:ABBREV,
  _:{ajusteTel:ajusteTel,trieTel:trieTel,nettoieNom:nettoieNom,frMoney:frMoney,abbreviateAddress:abbreviateAddress,cleanLieu:cleanLieu,cleanCP:cleanCP,detectGroup:detectGroup,suraveClean:suraveClean} };
if(typeof module!=='undefined'&&module.exports) module.exports=api;
else root.IntrumEngine=api;

})(typeof window!=='undefined'?window:this);
