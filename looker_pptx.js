const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Formation Looker Studio – RESOBIL";

const C = {
  dark:"0D1B2A", teal:"028090", tealL:"E0F5F4", orange:"F97316",
  orL:"FFF3E8", blue:"1E6F9F", blueL:"EFF6FF", green:"166534",
  greenL:"F0FDF4", gray:"64748B", grayL:"F8F9FA", white:"FFFFFF",
};
const noline = { type:"none" };

function stepCircle(s, x, y, num, c) {
  s.addShape(pres.shapes.OVAL,{x,y,w:.55,h:.55,fill:{color:c},line:noline});
  s.addText(String(num),{x,y,w:.55,h:.55,fontSize:16,color:C.white,bold:true,align:"center",valign:"middle"});
}
function card(s,x,y,w,h,bg,border){
  s.addShape(pres.shapes.RECTANGLE,{x,y,w,h,fill:{color:bg},line:{color:border||"E2E8F0",width:1},shadow:{type:"outer",color:"000000",blur:6,offset:2,angle:135,opacity:.07}});
}
function topBar(s,x,y,w,c){
  s.addShape(pres.shapes.RECTANGLE,{x,y,w,h:.06,fill:{color:c},line:noline});
}
function header(s,t,sub){
  s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:.72,fill:{color:C.dark},line:noline});
  s.addText(t,{x:.3,y:0,w:9.4,h:.72,fontSize:14,color:C.orange,bold:true,align:"left",valign:"middle",charSpacing:2});
  if(sub) s.addText(sub,{x:.3,y:.72,w:9.4,h:.38,fontSize:11,color:C.gray,italic:true,align:"left",valign:"middle"});
}

// ── SLIDE 1 TITRE ─────────────────────────────────────────────────────────
{
  const s=pres.addSlide(); s.background={color:C.dark};
  s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:.2,h:5.625,fill:{color:C.teal},line:noline});
  s.addShape(pres.shapes.OVAL,{x:6.8,y:2.0,w:4.2,h:4.2,fill:{color:C.teal,transparency:88},line:noline});
  s.addShape(pres.shapes.OVAL,{x:7.8,y:.5,w:2.4,h:2.4,fill:{color:C.orange,transparency:88},line:noline});
  s.addText("SYNTEYA CONSULTING LAB",{x:.4,y:.4,w:9,h:.38,fontSize:11,color:C.orange,bold:true,charSpacing:4});
  s.addText("Formation",{x:.4,y:.95,w:9,h:.7,fontSize:46,color:C.white,bold:true});
  s.addText("Google Looker Studio",{x:.4,y:1.68,w:8,h:.65,fontSize:38,color:C.teal});
  s.addText("De zéro au dashboard RESOBIL",{x:.4,y:2.55,w:7.5,h:.45,fontSize:20,color:C.gray});
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:.4,y:3.3,w:3.6,h:.5,fill:{color:C.teal},line:noline,rectRadius:.08});
  s.addText("Session 2h+  •  Débutants complets",{x:.4,y:3.3,w:3.6,h:.5,fontSize:13,color:C.white,align:"center",valign:"middle"});
  s.addText("lookerstudio.google.com  •  Gratuit  •  Sans installation",{x:.4,y:5.1,w:9,h:.35,fontSize:11,color:C.gray});
}

// ── SLIDE 2 AGENDA ─────────────────────────────────────────────────────────
{
  const s=pres.addSlide(); s.background={color:C.grayL};
  header(s,"PROGRAMME DE LA SÉANCE","2h+ pour maîtriser l'essentiel");
  const items=[
    {n:"1",t:"L'interface Looker Studio",d:"Navigation, zones clés, panneau des propriétés",c:C.teal,time:"20 min"},
    {n:"2",t:"Connecter vos CSV RESOBIL",d:"Via Google Sheets — méthode pas à pas",c:C.blue,time:"20 min"},
    {n:"3",t:"Dimensions vs Métriques",d:"Le concept fondamental — exercice pratique",c:C.orange,time:"15 min"},
    {n:"4",t:"Les 6 graphiques essentiels",d:"Scorecard, barres, courbes, camembert, tableau, carte",c:C.green,time:"30 min"},
    {n:"5",t:"Filtres et contrôles",d:"Rendre le dashboard interactif",c:C.teal,time:"15 min"},
    {n:"6",t:"Design & bonnes pratiques",d:"Thème, grille, hiérarchie visuelle",c:C.blue,time:"10 min"},
    {n:"7",t:"Pratique guidée RESOBIL",d:"Construire votre 1er graphique sur vos données",c:C.orange,time:"30 min"},
  ];
  items.forEach((it,i)=>{
    const y=1.22+i*.59;
    card(s,.3,y,9.4,.52,C.white,"E2E8F0");
    topBar(s,.3,y,9.4,it.c);
    stepCircle(s,.42,y+.09,it.n,it.c);
    s.addText(it.t,{x:1.1,y:y+.04,w:5.8,h:.28,fontSize:13,color:C.dark,bold:true});
    s.addText(it.d,{x:1.1,y:y+.25,w:5.8,h:.22,fontSize:10,color:C.gray});
    s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:7.3,y:y+.1,w:.9,h:.3,fill:{color:it.c},line:noline,rectRadius:.08});
    s.addText(it.time,{x:7.3,y:y+.1,w:.9,h:.3,fontSize:10,color:C.white,bold:true,align:"center",valign:"middle"});
  });
}

// ── SLIDE 3 QU'EST-CE QUE LOOKER STUDIO ───────────────────────────────────
{
  const s=pres.addSlide(); s.background={color:C.white};
  header(s,"QU'EST-CE QUE LOOKER STUDIO ?","L'outil BI gratuit de Google");
  card(s,.3,1.18,5.5,3.9,C.blueL,C.blue);
  s.addText("En une phrase :",{x:.5,y:1.32,w:5.1,h:.35,fontSize:13,color:C.blue,bold:true});
  s.addText("Looker Studio connecte vos données à des graphiques interactifs, sans code, dans le navigateur.",{x:.5,y:1.68,w:5.1,h:.85,fontSize:14,color:C.dark});
  s.addShape(pres.shapes.RECTANGLE,{x:.5,y:2.65,w:5.1,h:.03,fill:{color:C.blue},line:noline});
  ["✅  100% gratuit","✅  Aucune installation","✅  Partage par lien URL","✅  800+ sources de données"].forEach((f,i)=>
    s.addText(f,{x:.5,y:2.78+i*.48,w:5.1,h:.38,fontSize:13,color:C.dark})
  );
  [{icon:"📊",t:"Tableaux de bord",d:"Graphiques interactifs mis à jour automatiquement",c:C.teal},
   {icon:"🔗",t:"Sources variées",d:"Google Sheets, CSV, MySQL, BigQuery...",c:C.orange},
   {icon:"👁️",t:"Partage facile",d:"Lien URL, PDF, email programmé, intégration web",c:C.blue}
  ].forEach((r,i)=>{
    const y=1.18+i*1.3;
    card(s,6.1,y,3.6,1.15,C.grayL,"E2E8F0");
    s.addText(r.icon,{x:6.2,y:y+.12,w:.7,h:.7,fontSize:24,align:"center"});
    s.addText(r.t,{x:6.95,y:y+.1,w:2.6,h:.35,fontSize:13,color:r.c,bold:true});
    s.addText(r.d,{x:6.95,y:y+.43,w:2.6,h:.55,fontSize:10.5,color:C.gray});
  });
  s.addText("Pour RESOBIL : vos CSV de production → dashboard interactif → décisions terrain",{x:.3,y:5.1,w:9.4,h:.38,fontSize:12,color:C.blue,italic:true,align:"center"});
}

// ── SLIDE 4 INTERFACE ─────────────────────────────────────────────────────
{
  const s=pres.addSlide(); s.background={color:C.white};
  header(s,"L'INTERFACE EN UN COUP D'ŒIL","Mode Édition — les 4 zones à connaître");
  // Zone 1 menus
  s.addShape(pres.shapes.RECTANGLE,{x:.3,y:1.1,w:9.4,h:.48,fill:{color:"E2E8F0"},line:{color:"94A3B8",width:1}});
  s.addShape(pres.shapes.RECTANGLE,{x:.3,y:1.1,w:.05,h:.48,fill:{color:C.teal},line:noline});
  s.addText("① Barre de menus  |  Fichier  ·  Éditer  ·  Affichage  ·  Insertion  ·  Page  ·  Organiser  ·  Ressource",{x:.45,y:1.1,w:9.1,h:.48,fontSize:10,color:C.dark,valign:"middle"});
  // Zone 2 outils
  s.addShape(pres.shapes.RECTANGLE,{x:.3,y:1.62,w:9.4,h:.42,fill:{color:"F1F5F9"},line:{color:"94A3B8",width:1}});
  s.addShape(pres.shapes.RECTANGLE,{x:.3,y:1.62,w:.05,h:.42,fill:{color:C.orange},line:noline});
  s.addText("② Barre d'outils  |  ↩↪  │  ✚page  Ajouter données  Ajouter graphique ▾  Contrôle ▾  Texte  Image  Forme  🎨Thème",{x:.45,y:1.62,w:9.1,h:.42,fontSize:9.5,color:C.dark,valign:"middle"});
  // Zone 3 canevas
  s.addShape(pres.shapes.RECTANGLE,{x:.3,y:2.08,w:6.6,h:3.0,fill:{color:"FAFAFA"},line:{color:"CBD5E1",width:1}});
  s.addShape(pres.shapes.RECTANGLE,{x:.3,y:2.08,w:.05,h:3.0,fill:{color:C.green},line:noline});
  s.addText("③ CANEVAS\n(zone de travail)\nDrag & drop  •  Libre positionnement  •  Multi-pages",{x:1.5,y:3.0,w:4.0,h:1.5,fontSize:11,color:"94A3B8",align:"center",valign:"middle"});
  // mini graphique mock
  s.addShape(pres.shapes.RECTANGLE,{x:.6,y:2.25,w:1.5,h:.65,fill:{color:"E2E8F0"},line:{color:"94A3B8",width:1}});
  s.addText("KPI",{x:.6,y:2.25,w:1.5,h:.65,fontSize:9,color:"94A3B8",align:"center",valign:"middle"});
  s.addShape(pres.shapes.RECTANGLE,{x:2.3,y:2.25,w:2.5,h:1.1,fill:{color:"E2E8F0"},line:{color:"94A3B8",width:1}});
  s.addText("Graphique",{x:2.3,y:2.25,w:2.5,h:1.1,fontSize:9,color:"94A3B8",align:"center",valign:"middle"});
  // Zone 4 propriétés
  s.addShape(pres.shapes.RECTANGLE,{x:7.05,y:2.08,w:2.65,h:3.0,fill:{color:"F8FAFC"},line:{color:"CBD5E1",width:1}});
  s.addShape(pres.shapes.RECTANGLE,{x:7.05,y:2.08,w:.05,h:3.0,fill:{color:C.blue},line:noline});
  s.addText("④ Panneau des propriétés",{x:7.1,y:2.2,w:2.55,h:.55,fontSize:12,color:C.blue,bold:true,align:"center"});
  s.addShape(pres.shapes.RECTANGLE,{x:7.15,y:2.9,w:1.05,h:.32,fill:{color:C.teal},line:noline});
  s.addText("Configuration",{x:7.15,y:2.9,w:1.05,h:.32,fontSize:8,color:C.white,bold:true,align:"center",valign:"middle"});
  s.addShape(pres.shapes.RECTANGLE,{x:8.25,y:2.9,w:1.1,h:.32,fill:{color:"E2E8F0"},line:noline});
  s.addText("Style",{x:8.25,y:2.9,w:1.1,h:.32,fontSize:8,color:C.dark,align:"center",valign:"middle"});
  s.addText("Dimension ●",{x:7.15,y:3.32,w:2.55,h:.28,fontSize:10,color:"166534"});
  s.addText("Métrique ●",{x:7.15,y:3.6,w:2.55,h:.28,fontSize:10,color:"1E6F9F"});
  s.addText("Tri / Filtre / Dates",{x:7.15,y:3.88,w:2.55,h:.28,fontSize:10,color:C.gray});
  // Légende
  [{c:C.teal,l:"① Menus"},{c:C.orange,l:"② Outils"},{c:C.green,l:"③ Canevas"},{c:C.blue,l:"④ Propriétés"}
  ].forEach((lg,i)=>{
    s.addShape(pres.shapes.OVAL,{x:.4+i*2.3,y:5.18,w:.2,h:.2,fill:{color:lg.c},line:noline});
    s.addText(lg.l,{x:.68+i*2.3,y:5.15,w:1.8,h:.28,fontSize:11,color:lg.c,bold:true});
  });
}

// ── SLIDE 5 CONNECTER CSV ─────────────────────────────────────────────────
{
  const s=pres.addSlide(); s.background={color:C.white};
  header(s,"CONNECTER VOS CSV RESOBIL","Méthode Google Sheets — la plus fiable pour débutants");
  const steps=[
    {n:"1",icon:"📁",t:"Exporter le CSV depuis MySQL Workbench",d:"Server → Data Export → Format CSV → Démarrer l'export",c:C.teal},
    {n:"2",icon:"☁️",t:"Uploader dans Google Drive",d:"drive.google.com → + Nouveau → Téléverser un fichier → Sélectionner le CSV",c:C.blue},
    {n:"3",icon:"📊",t:"Ouvrir avec Google Sheets",d:"Double-clic sur le CSV → Ouvrir avec → Google Sheets → Vérifier 1ère ligne = en-têtes",c:C.orange},
    {n:"4",icon:"🔗",t:"Connecter dans Looker Studio",d:"Barre d'outils → Ajouter des données → Google Sheets → Sélectionner classeur & onglet → AJOUTER",c:C.green},
    {n:"5",icon:"✅",t:"Vérifier et configurer les champs",d:"Renommer les champs, vérifier les types (Date, Nombre, Géo), ajuster les agrégations par défaut",c:C.teal},
  ];
  steps.forEach((st,i)=>{
    const y=1.18+i*.86;
    card(s,.3,y,9.4,.78,i%2===0?C.white:C.grayL,"E2E8F0");
    stepCircle(s,.42,y+.17,st.n,st.c);
    s.addText(st.icon+"  "+st.t,{x:.78+.18,y:y+.06,w:8.4,h:.34,fontSize:13,color:st.c,bold:true});
    s.addText(st.d,{x:.78+.18,y:y+.4,w:8.4,h:.3,fontSize:11,color:C.gray});
  });
  s.addShape(pres.shapes.RECTANGLE,{x:0,y:5.18,w:10,h:.44,fill:{color:C.orL},line:noline});
  s.addText("⚠️  Prérequis CSV : encodage UTF-8  •  1ère ligne = en-têtes  •  Pas de cellules fusionnées  •  Pas de ligne de totaux",{x:.3,y:5.18,w:9.4,h:.44,fontSize:11,color:C.orange,bold:true,align:"center",valign:"middle"});
}

// ── SLIDE 6 DIMENSIONS VS METRIQUES ───────────────────────────────────────
{
  const s=pres.addSlide(); s.background={color:C.white};
  header(s,"DIMENSIONS VS MÉTRIQUES","Le concept fondamental — à maîtriser absolument");
  // Dimension
  card(s,.3,1.1,4.4,3.95,C.greenL,C.green);
  s.addShape(pres.shapes.RECTANGLE,{x:.3,y:1.1,w:4.4,h:.55,fill:{color:C.green},line:noline});
  s.addText("DIMENSION  ●  (vert)",{x:.4,y:1.1,w:4.2,h:.55,fontSize:14,color:C.white,bold:true,valign:"middle"});
  s.addText("Répond à : QUI ? QUOI ? OÙ ? QUAND ?",{x:.5,y:1.75,w:4.0,h:.35,fontSize:12,color:C.green,bold:true});
  s.addText("Catégoriel, non agrégé, sert à regrouper",{x:.5,y:2.1,w:4.0,h:.3,fontSize:11,color:C.gray});
  ["Nom_Antenne","Type_Culture","Date","Région","Nom_Agriculteur"].forEach((d,i)=>{
    s.addShape(pres.shapes.OVAL,{x:.52,y:2.52+i*.42,w:.18,h:.18,fill:{color:C.green},line:noline});
    s.addText(d,{x:.8,y:2.47+i*.42,w:3.7,h:.3,fontSize:12,color:C.dark});
  });
  // Métrique
  card(s,5.3,1.1,4.4,3.95,C.blueL,C.blue);
  s.addShape(pres.shapes.RECTANGLE,{x:5.3,y:1.1,w:4.4,h:.55,fill:{color:C.blue},line:noline});
  s.addText("MÉTRIQUE  ●  (bleu)",{x:5.4,y:1.1,w:4.2,h:.55,fontSize:14,color:C.white,bold:true,valign:"middle"});
  s.addText("Répond à : COMBIEN ?",{x:5.5,y:1.75,w:4.0,h:.35,fontSize:12,color:C.blue,bold:true});
  s.addText("Numérique, agrégé (SUM, AVG, COUNT...)",{x:5.5,y:2.1,w:4.0,h:.3,fontSize:11,color:C.gray});
  ["SUM(Production_kg)","AVG(Rendement_ha)","COUNT_DISTINCT(Agriculteurs)","SUM(Revenus_FCFA)","MAX(Surface_ha)"].forEach((m,i)=>{
    s.addShape(pres.shapes.OVAL,{x:5.5,y:2.52+i*.42,w:.18,h:.18,fill:{color:C.blue},line:noline});
    s.addText(m,{x:5.78,y:2.47+i*.42,w:3.8,h:.3,fontSize:12,color:C.dark});
  });
  s.addShape(pres.shapes.RECTANGLE,{x:0,y:5.12,w:10,h:.5,fill:{color:"1E2D3D"},line:noline});
  s.addText("Analogie : Dimension = axe X (par antenne, par culture)  •  Métrique = hauteur des barres (combien produit)",{x:.3,y:5.12,w:9.4,h:.5,fontSize:12,color:C.white,align:"center",valign:"middle"});
}

// ── SLIDE 7 LES 6 GRAPHIQUES ──────────────────────────────────────────────
{
  const s=pres.addSlide(); s.background={color:C.white};
  header(s,"LES 6 GRAPHIQUES ESSENTIELS","Quel graphique pour quel besoin ?");
  const charts=[
    {icon:"🔢",t:"Scorecard",q:"Un KPI en grand",dim:"—",met:"SUM(Production_kg)",c:C.teal},
    {icon:"📊",t:"Barres",q:"Comparer des antennes",dim:"Nom_Antenne",met:"SUM(Production)",c:C.blue},
    {icon:"📈",t:"Courbes",q:"Évolution dans le temps",dim:"Date",met:"SUM(Production)",c:C.orange},
    {icon:"🥧",t:"Camembert",q:"Répartition des cultures",dim:"Type_Culture",met:"SUM(Surface_ha)",c:C.green},
    {icon:"📋",t:"Tableau",q:"Données détaillées",dim:"Antenne + Culture",met:"SUM, AVG, COUNT",c:C.blue},
    {icon:"🗺️",t:"Carte Google Maps",q:"Répartition géographique",dim:"Ville (type Géo)",met:"SUM(Production)",c:C.teal},
  ];
  charts.forEach((ch,i)=>{
    const col=i%3, row=Math.floor(i/3);
    const x=.3+col*3.22, y=1.12+row*2.2;
    card(s,x,y,3.0,2.0,C.white,"E2E8F0");
    topBar(s,x,y,3.0,ch.c);
    s.addText(ch.icon,{x,y:y+.1,w:3.0,h:.55,fontSize:24,align:"center"});
    s.addText(ch.t,{x:x+.12,y:y+.65,w:2.78,h:.32,fontSize:14,color:ch.c,bold:true});
    s.addText(ch.q,{x:x+.12,y:y+.95,w:2.78,h:.28,fontSize:10.5,color:C.gray});
    s.addShape(pres.shapes.RECTANGLE,{x:x+.12,y:y+1.28,w:2.78,h:.02,fill:{color:"E2E8F0"},line:noline});
    s.addText("Dim : "+ch.dim,{x:x+.12,y:y+1.35,w:2.78,h:.25,fontSize:9.5,color:C.gray});
    s.addText("Met : "+ch.met,{x:x+.12,y:y+1.6,w:2.78,h:.25,fontSize:9.5,color:C.blue});
  });
}

// ── SLIDE 8 SCORECARD & BARRES ────────────────────────────────────────────
{
  const s=pres.addSlide(); s.background={color:C.white};
  header(s,"SCORECARD ET GRAPHIQUE EN BARRES","Configuration détaillée dans l'onglet Configuration");
  // Scorecard
  card(s,.3,1.12,4.55,4.1,C.tealL,C.teal);
  s.addShape(pres.shapes.RECTANGLE,{x:.3,y:1.12,w:4.55,h:.45,fill:{color:C.teal},line:noline});
  s.addText("🔢  SCORECARD — KPI en grand",{x:.4,y:1.12,w:4.4,h:.45,fontSize:13,color:C.white,bold:true,valign:"middle"});
  card(s,.5,1.7,4.1,.9,C.white,"E2E8F0");
  s.addText("4 250 sacs",{x:.5,y:1.7,w:3.0,h:.9,fontSize:22,color:C.teal,bold:true,align:"center",valign:"middle"});
  s.addText("▲ +8%",{x:3.4,y:1.9,w:1.1,h:.5,fontSize:13,color:"166534",bold:true,align:"center",valign:"middle"});
  s.addText("Configuration :",{x:.5,y:2.72,w:4.1,h:.3,fontSize:12,color:C.teal,bold:true});
  ["Métrique → SUM(quantite_recoltee_sac)","Dimension date → periode_de_production","Activer : Comparaison de période","Style : Nombres compacts ON","Variation en % (vert positif / rouge négatif)"].forEach((c,i)=>{
    s.addShape(pres.shapes.OVAL,{x:.52,y:3.1+i*.38,w:.14,h:.14,fill:{color:C.teal},line:noline});
    s.addText(c,{x:.74,y:3.05+i*.38,w:3.9,h:.3,fontSize:10.5,color:C.dark});
  });
  // Barres
  card(s,5.15,1.12,4.55,4.1,C.blueL,C.blue);
  s.addShape(pres.shapes.RECTANGLE,{x:5.15,y:1.12,w:4.55,h:.45,fill:{color:C.blue},line:noline});
  s.addText("📊  BARRES — Comparer les antennes",{x:5.25,y:1.12,w:4.4,h:.45,fontSize:13,color:C.white,bold:true,valign:"middle"});
  [{l:"NKOLO",v:.85},{l:"ONDONDO",v:.65},{l:"EFOK",v:.5},{l:"LELA",v:.3}].forEach((b,i)=>{
    s.addText(b.l,{x:5.3,y:1.73+i*.35,w:1.2,h:.3,fontSize:9,color:C.dark,align:"right"});
    s.addShape(pres.shapes.RECTANGLE,{x:6.58,y:1.78+i*.35,w:b.v*2.7,h:.2,fill:{color:C.blue},line:noline});
  });
  s.addText("Configuration :",{x:5.3,y:2.72,w:4.25,h:.3,fontSize:12,color:C.blue,bold:true});
  ["Dimension → nom (antenne)","Métrique → SUM(quantite_recoltee_sac)","Tri → SUM décroissant","Style : Barres horizontales","Activer : Étiquettes de données"].forEach((c,i)=>{
    s.addShape(pres.shapes.OVAL,{x:5.32,y:3.1+i*.38,w:.14,h:.14,fill:{color:C.blue},line:noline});
    s.addText(c,{x:5.54,y:3.05+i*.38,w:4.0,h:.3,fontSize:10.5,color:C.dark});
  });
}

// ── SLIDE 9 COURBES, CAMEMBERT, CARTE ─────────────────────────────────────
{
  const s=pres.addSlide(); s.background={color:C.white};
  header(s,"COURBES, CAMEMBERT ET CARTE","Configuration des 3 graphiques restants");
  const panels=[
    {icon:"📈",t:"Courbes (Série temporelle)",c:C.orange,cfg:["Dimension date → periode_de_production","Métrique → SUM(quantite_recoltee_sac)","Drill-down Année → Mois → Semaine","Comparaison : période précédente","Style : Courbe de tendance ON"]},
    {icon:"🥧",t:"Camembert / Anneau",c:C.green,cfg:["Dimension → type_de_cultures","Métrique → SUM(superficie_exploitee_ha)","Max 5-6 tranches (regrouper les autres)","Style : % affiché sur chaque part","⚠️ Jamais plus de 6 catégories"]},
    {icon:"🗺️",t:"Carte Google Maps",c:C.teal,cfg:["Emplacement → localisation (type Géo > Ville)","Taille bulle → SUM(quantite_recoltee_sac)","Couleur → AVG(rendement)","⚠️ Changer le type en Géo > Ville !","Vue par défaut → naviguer Cameroun"]},
  ];
  panels.forEach((p,i)=>{
    const x=.3+i*3.23;
    card(s,x,1.1,3.1,4.15,C.white,"E2E8F0");
    topBar(s,x,1.1,3.1,p.c);
    s.addText(p.icon+"  "+p.t,{x:x+.12,y:1.18,w:2.9,h:.55,fontSize:12,color:p.c,bold:true});
    s.addShape(pres.shapes.RECTANGLE,{x:x+.12,y:1.75,w:2.88,h:.02,fill:{color:"E2E8F0"},line:noline});
    p.cfg.forEach((c,j)=>{
      const isWarn=c.startsWith("⚠️");
      s.addShape(pres.shapes.OVAL,{x:x+.15,y:1.86+j*.46,w:.14,h:.14,fill:{color:isWarn?C.orange:p.c},line:noline});
      s.addText(c.replace("⚠️ ",""),{x:x+.38,y:1.81+j*.46,w:2.7,h:.38,fontSize:10,color:isWarn?C.orange:C.dark,bold:isWarn});
    });
  });
  s.addShape(pres.shapes.RECTANGLE,{x:0,y:5.25,w:10,h:.38,fill:{color:C.orL},line:noline});
  s.addText("💡  Changer le type d'un graphique : sélectionner → menu déroulant dans le panneau Propriétés → nouveau type (les données sont conservées !)",{x:.3,y:5.25,w:9.4,h:.38,fontSize:11,color:C.orange,bold:true,align:"center",valign:"middle"});
}

// ── SLIDE 10 FILTRES ET CONTROLES ─────────────────────────────────────────
{
  const s=pres.addSlide(); s.background={color:C.white};
  header(s,"FILTRES ET CONTRÔLES INTERACTIFS","Rendre le dashboard vivant");
  const ctrls=[
    {t:"Liste déroulante",icon:"☰",c:C.teal,steps:["Insérer → Contrôle → Liste déroulante","Champ de contrôle → nom (antenne)","Toutes les valeurs distinctes affichées auto","Affecte tous les graphiques de la page"]},
    {t:"Plage de dates",icon:"📅",c:C.blue,steps:["Insérer → Contrôle → Plage de dates","Placer en haut à droite (convention)","Plages prédéfinies : 30j, ce mois, ce trimestre","Chaque graphique doit être sur 'Auto'"]},
    {t:"Filtrage croisé",icon:"🔗",c:C.orange,steps:["Sélectionner un graphique","Onglet Config → Interactions","Activer le filtrage croisé","Cliquer sur une barre filtre tous les autres"]},
  ];
  ctrls.forEach((ct,i)=>{
    const x=.3+i*3.22;
    card(s,x,1.08,3.05,3.2,i===0?C.tealL:i===1?C.blueL:C.orL,i===0?C.teal:i===1?C.blue:C.orange);
    s.addShape(pres.shapes.RECTANGLE,{x,y:1.08,w:3.05,h:.5,fill:{color:i===0?C.teal:i===1?C.blue:C.orange},line:noline});
    s.addText(ct.icon+"  "+ct.t,{x:x+.1,y:1.08,w:2.9,h:.5,fontSize:13,color:C.white,bold:true,valign:"middle"});
    ct.steps.forEach((st,j)=>{
      s.addShape(pres.shapes.OVAL,{x:x+.14,y:1.7+j*.5,w:.16,h:.16,fill:{color:i===0?C.teal:i===1?C.blue:C.orange},line:noline});
      s.addText(st,{x:x+.38,y:1.65+j*.5,w:2.62,h:.42,fontSize:10.5,color:C.dark});
    });
  });
  s.addShape(pres.shapes.RECTANGLE,{x:.3,y:4.42,w:9.4,h:1.0,fill:{color:"F1F5F9"},line:{color:"CBD5E1",width:1}});
  s.addText("Hiérarchie des filtres :",{x:.5,y:4.5,w:2.5,h:.3,fontSize:12,color:C.dark,bold:true});
  ["Rapport (toutes pages)","Page","Groupe","Graphique"].forEach((h,i)=>{
    const hx=2.9+i*1.7;
    s.addShape(pres.shapes.RECTANGLE,{x:hx,y:4.52,w:1.55,h:.3,fill:{color:[C.teal,C.blue,C.orange,C.green][i]},line:noline});
    s.addText(h,{x:hx,y:4.52,w:1.55,h:.3,fontSize:9,color:C.white,align:"center",valign:"middle"});
    if(i<3) s.addText("→",{x:hx+1.58,y:4.52,w:.1,h:.3,fontSize:11,color:C.gray,valign:"middle"});
  });
  s.addText("Plus restrictif →",{x:2.9,y:4.85,w:9.4,h:.25,fontSize:9,color:C.gray});
}

// ── SLIDE 11 DESIGN ───────────────────────────────────────────────────────
{
  const s=pres.addSlide(); s.background={color:C.white};
  header(s,"DESIGN & BONNES PRATIQUES","Un dashboard qui sera réellement utilisé");
  // Mise en page mock
  card(s,.3,1.1,5.5,4.15,"F8FAFC","CBD5E1");
  s.addText("Mise en page recommandée",{x:.5,y:1.18,w:5.1,h:.35,fontSize:13,color:C.dark,bold:true});
  s.addShape(pres.shapes.RECTANGLE,{x:.5,y:1.6,w:4.9,h:.42,fill:{color:C.dark},line:noline});
  s.addText("En-tête : Logo + Titre + Contrôle dates + Filtres  →  Rendre au niveau du rapport",{x:.55,y:1.6,w:4.85,h:.42,fontSize:9,color:C.white,valign:"middle"});
  [C.teal,C.blue,C.orange,C.green].forEach((c,i)=>{
    s.addShape(pres.shapes.RECTANGLE,{x:.5+i*1.21,y:2.08,w:1.13,h:.6,fill:{color:c,transparency:85},line:{color:c,width:1}});
    s.addText("KPI "+(i+1),{x:.5+i*1.21,y:2.08,w:1.13,h:.6,fontSize:9,color:c,bold:true,align:"center",valign:"middle"});
  });
  s.addShape(pres.shapes.RECTANGLE,{x:.5,y:2.76,w:2.3,h:.9,fill:{color:C.blueL},line:{color:C.blue,width:1}});
  s.addText("📈 Courbes\ntendances",{x:.5,y:2.76,w:2.3,h:.9,fontSize:9,color:C.blue,align:"center",valign:"middle"});
  s.addShape(pres.shapes.RECTANGLE,{x:2.88,y:2.76,w:2.3,h:.9,fill:{color:C.orL},line:{color:C.orange,width:1}});
  s.addText("📊 Barres\ncomparaison",{x:2.88,y:2.76,w:2.3,h:.9,fontSize:9,color:C.orange,align:"center",valign:"middle"});
  s.addShape(pres.shapes.RECTANGLE,{x:.5,y:3.74,w:4.7,h:1.1,fill:{color:C.grayL},line:{color:"CBD5E1",width:1}});
  s.addText("📋 Tableau détaillé en bas (données exactes, tri, pagination)",{x:.5,y:3.74,w:4.7,h:1.1,fontSize:9,color:C.gray,align:"center",valign:"middle"});
  // Règles
  card(s,6.1,1.1,3.6,4.15,C.white,"E2E8F0");
  s.addText("Les règles clés",{x:6.25,y:1.18,w:3.3,h:.35,fontSize:13,color:C.dark,bold:true});
  [{icon:"🎨",t:"Thème AVANT les graphiques",d:"Couleurs et polices en premier"},{icon:"📐",t:"Grille & alignement",d:"Affichage → Grille + Organiser → Aligner"},{icon:"✂️",t:"Max 8-10 graphiques/page",d:"Au-delà : créer une nouvelle page"},{icon:"🔤",t:"2 polices max",d:"1 titre, 1 corps — sans-serif uniquement"},{icon:"⬜",t:"Espace blanc = ami",d:"Ne pas tout remplir — lisibilité prime"},{icon:"🏷️",t:"Titres explicites",d:'« Production par antenne (sacs) » pas « Graphique 1 »'}
  ].forEach((r,i)=>{
    s.addText(r.icon+" "+r.t,{x:6.25,y:1.62+i*.57,w:3.3,h:.3,fontSize:11,color:C.dark,bold:true});
    s.addText(r.d,{x:6.25,y:1.9+i*.57,w:3.3,h:.26,fontSize:9.5,color:C.gray});
  });
}

// ── SLIDE 12 WORKFLOW ─────────────────────────────────────────────────────
{
  const s=pres.addSlide(); s.background={color:C.grayL};
  header(s,"WORKFLOW COMPLET — CRÉER UN DASHBOARD DE ZÉRO","6 étapes pour aller du CSV au dashboard professionnel RESOBIL");
  const wf=[
    {n:"1",t:"Définir les KPIs",d:"Qui consulte ? Quelles décisions ? 5-10 KPIs max",c:C.teal},
    {n:"2",t:"Préparer les données",d:"Nettoyer CSV, uploader Sheets, vérifier les types",c:C.blue},
    {n:"3",t:"Configurer le thème",d:"Couleurs, polices, grille — AVANT les graphiques",c:C.orange},
    {n:"4",t:"Construire les graphiques",d:"En-tête → KPIs → Graphiques → Tableau",c:C.green},
    {n:"5",t:"Ajouter l'interactivité",d:"Filtres, dates, filtrage croisé — tester tout",c:C.teal},
    {n:"6",t:"Partager",d:"Lien URL, PDF, email programmé, iframe",c:C.blue},
  ];
  wf.forEach((w,i)=>{
    const col=i%3, row=Math.floor(i/3);
    const x=.3+col*3.22, y=1.12+row*2.1;
    card(s,x,y,3.05,1.9,C.white,"E2E8F0");
    topBar(s,x,y,3.05,w.c);
    stepCircle(s,x+.1,y+.15,w.n,w.c);
    s.addText(w.t,{x:x+.78,y:y+.1,w:2.15,h:.42,fontSize:13,color:w.c,bold:true});
    s.addText(w.d,{x:x+.12,y:y+.65,w:2.85,h:.95,fontSize:11,color:C.dark});
  });
}

// ── SLIDE 13 EXERCICE PRATIQUE ────────────────────────────────────────────
{
  const s=pres.addSlide(); s.background={color:C.white};
  header(s,"EXERCICE PRATIQUE — VOS DONNÉES RESOBIL","30 minutes pour construire votre premier graphique");
  s.addShape(pres.shapes.RECTANGLE,{x:.3,y:1.08,w:9.4,h:.48,fill:{color:C.tealL},line:{color:C.teal,width:1}});
  s.addText("🎯  Objectif : Créer un scorecard + un graphique en barres à partir de votre CSV de production",{x:.4,y:1.08,w:9.2,h:.48,fontSize:12,color:C.teal,bold:true,valign:"middle"});
  const tasks=[
    {n:"A",t:"Préparer la source",items:["Ouvrir votre CSV nettoyé dans Google Sheets","Vérifier : 1ère ligne = en-têtes","Connecter à Looker Studio → Ajouter des données"],c:C.teal},
    {n:"B",t:"Créer le scorecard",items:["Insérer → Graphique → Scorecard","Métrique → SUM(quantite_recoltee_sac)","Activer la comparaison de période","Style : Nombres compacts ON"],c:C.blue},
    {n:"C",t:"Créer le graphique en barres",items:["Insérer → Graphique → Barres horizontales","Dimension → nom (antenne)","Métrique → SUM(quantite_recoltee_sac)","Tri décroissant + étiquettes ON"],c:C.orange},
  ];
  tasks.forEach((t,i)=>{
    const x=.3+i*3.22;
    card(s,x,1.72,3.05,3.62,C.white,"E2E8F0");
    s.addShape(pres.shapes.RECTANGLE,{x,y:1.72,w:3.05,h:.48,fill:{color:t.c},line:noline});
    s.addText(t.n+"  "+t.t,{x:x+.12,y:1.72,w:2.9,h:.48,fontSize:14,color:C.white,bold:true,valign:"middle"});
    t.items.forEach((item,j)=>{
      s.addShape(pres.shapes.OVAL,{x:x+.14,y:2.32+j*.72,w:.2,h:.2,fill:{color:t.c},line:noline});
      s.addText(item,{x:x+.44,y:2.27+j*.72,w:2.62,h:.6,fontSize:11,color:C.dark});
    });
  });
  s.addShape(pres.shapes.RECTANGLE,{x:0,y:5.22,w:10,h:.4,fill:{color:C.dark},line:noline});
  s.addText("N'hésitez pas à appeler à l'aide — déblocage immédiat !",{x:.3,y:5.22,w:9.4,h:.4,fontSize:13,color:C.orange,bold:true,align:"center",valign:"middle"});
}

// ── SLIDE 14 CHAMPS CALCULES ──────────────────────────────────────────────
{
  const s=pres.addSlide(); s.background={color:C.white};
  header(s,"CHAMPS CALCULÉS — ALLER PLUS LOIN","Créer de nouvelles métriques sans toucher à la source");
  card(s,.3,1.08,9.4,.88,C.blueL,C.blue);
  s.addText("Un champ calculé = une formule créant un nouveau champ à partir des données existantes. Marqué fx. Sans modifier la source.",{x:.5,y:1.15,w:9.0,h:.72,fontSize:12,color:C.dark,valign:"middle"});
  [{t:"Rendement par hectare",f:"SUM(quantite_recoltee_sac) / SUM(superficie_exploitee_ha)",c:C.teal,d:"Divise deux métriques"},
   {t:"Catégorie de production",f:'CASE WHEN quantite_recoltee_sac > 200 THEN "Élevée" WHEN quantite_recoltee_sac > 100 THEN "Moyenne" ELSE "Faible" END',c:C.orange,d:"Crée une nouvelle dimension"},
   {t:"Taux de perte (%)",f:"SUM(pertes_eventuelles_sac) / NULLIF(SUM(quantite_recoltee_sac),0) * 100",c:C.blue,d:"Ratio avec protection division par zéro"},
  ].forEach((ex,i)=>{
    const y=2.12+i*1.08;
    card(s,.3,y,9.4,.98,i%2===0?C.white:C.grayL,"E2E8F0");
    s.addText("fx  "+ex.t,{x:.5,y:y+.08,w:3.5,h:.3,fontSize:12,color:ex.c,bold:true});
    s.addText(ex.d,{x:.5,y:y+.42,w:3.5,h:.3,fontSize:10,color:C.gray});
    s.addShape(pres.shapes.RECTANGLE,{x:4.2,y:y+.12,w:5.3,h:.65,fill:{color:"1E293B"},line:noline});
    s.addText(ex.f,{x:4.3,y:y+.12,w:5.1,h:.65,fontSize:9.5,color:"7DD3FC",valign:"middle",fontFace:"Courier New"});
  });
  s.addShape(pres.shapes.RECTANGLE,{x:0,y:5.23,w:10,h:.4,fill:{color:C.orL},line:noline});
  s.addText("Ressource → Gérer les sources → Modifier → + Ajouter un champ → Écrire la formule → Enregistrer",{x:.3,y:5.23,w:9.4,h:.4,fontSize:11,color:C.orange,bold:true,align:"center",valign:"middle"});
}

// ── SLIDE 15 CLÔTURE ──────────────────────────────────────────────────────
{
  const s=pres.addSlide(); s.background={color:C.dark};
  s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:.2,h:5.625,fill:{color:C.orange},line:noline});
  s.addShape(pres.shapes.OVAL,{x:6.5,y:1.5,w:5.0,h:5.0,fill:{color:C.teal,transparency:90},line:noline});
  s.addText("Ce soir, vous avez\ncréé votre premier dashboard.",{x:.4,y:.5,w:7.5,h:1.8,fontSize:34,color:C.white,bold:true});
  s.addText("La semaine prochaine, vous livrez à RESOBIL.",{x:.4,y:2.4,w:8,h:.6,fontSize:20,color:C.orange,italic:true});
  ["📌  CSV nettoyé → Google Sheets → Looker Studio","📌  Dimension (vert) = par quoi  •  Métrique (bleu) = combien","📌  Tableau de bord sur votre compte Google personnel","📌  Partager avec Célestine avant la prochaine revue"].forEach((r,i)=>
    s.addText(r,{x:.4,y:3.2+i*.44,w:8.2,h:.38,fontSize:12,color:C.white})
  );
  s.addShape(pres.shapes.RECTANGLE,{x:.4,y:5.1,w:5.5,h:.38,fill:{color:C.teal},line:noline});
  s.addText("Synteya Partners  ×  RESOBIL  |  lookerstudio.google.com",{x:.4,y:5.1,w:5.5,h:.38,fontSize:12,color:C.white,bold:true,align:"center",valign:"middle"});
}

pres.writeFile({fileName:"Formation_LookerStudio_RESOBIL.pptx"}).then(()=>console.log("✅ Done !"));
