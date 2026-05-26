document.getElementById("generate").addEventListener("click", async () => {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  // tutto il tuo codice dentro qui

// ── Colores ──────────────────────────────────────────────────────────────────
const D  = "0B3D6B";  // azul oscuro — headers
const M  = "1A73A7";  // azul medio
const AC = "00B89F";  // verde menta
const BG = "F5F8FB";  // fondo claro
const WH = "FFFFFF";
const TX = "1C2B3A";  // texto oscuro
const MU = "607D8B";  // texto secundario
const RV = "1565C0";  // azul RV
const RF = "00897B";  // verde RF
const PO = "2E7D32";  // positivo
const NE = "C62828";  // negativo
const F  = "Calibri";

// ── Helpers ──────────────────────────────────────────────────────────────────
const sl = () => { const s = pres.addSlide(); s.background = { color: BG }; return s; };
const dkSl = () => { const s = pres.addSlide(); s.background = { color: D }; return s; };

const bar = (s) => s.addShape(pres.shapes.RECTANGLE,
  { x:0, y:0, w:10, h:0.07, fill:{color:AC}, line:{color:AC} });

const ttl = (s, t, y=0.18, c=D, sz=28) =>
  s.addText(t, {x:0.45,y,w:9.1,h:0.55,fontSize:sz,fontFace:F,bold:true,color:c,valign:"middle"});

const sub = (s, t, y=0.72, c=MU, sz=13) =>
  s.addText(t, {x:0.45,y,w:9.1,h:0.28,fontSize:sz,fontFace:F,color:c,valign:"middle"});

const bx = (s,x,y,w,h,c=WH) => s.addShape(pres.shapes.RECTANGLE,
  {x,y,w,h,fill:{color:c},line:{color:"D0DDE8",width:1},
   shadow:{type:"outer",blur:3,offset:1,angle:135,color:"000000",opacity:0.06}});

const pill = (s,t,x,y,w=1.7,h=0.3,bg=M,fg=WH) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y,w,h,fill:{color:bg},line:{color:bg},rectRadius:0.06});
  s.addText(t,{x,y,w,h,fontSize:11,fontFace:F,bold:true,color:fg,align:"center",valign:"middle"});
};

const vc = v => String(v).startsWith("-") ? NE : (parseFloat(v) > 0 ? PO : TX);

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 1 — PORTADA
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = dkSl();
  s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.22,h:5.625,fill:{color:AC},line:{color:AC}});
  s.addText("Cartera de Inversión ESG",
    {x:0.55,y:0.85,w:9,h:1.2,fontSize:44,fontFace:F,bold:true,color:WH,valign:"middle",lineSpacingMultiple:1.1});
  s.addText("Construyendo el Plan de Jubilación — Del Plan a la Acción",
    {x:0.55,y:2.12,w:9,h:0.38,fontSize:15,fontFace:F,color:"B0CCE0",valign:"middle"});
  s.addShape(pres.shapes.LINE,{x:0.55,y:2.58,w:3.8,h:0,line:{color:AC,width:2}});
  s.addText("Jessica Elizabeth Rojas Rodriguez",
    {x:0.55,y:2.72,w:9,h:0.38,fontSize:14,fontFace:F,bold:true,color:WH});
  s.addText("MBA UVa  ·  Activos e Inversiones Financieras  ·  Curso 2025-2026",
    {x:0.55,y:3.12,w:9,h:0.32,fontSize:12,fontFace:F,color:"B0CCE0"});
  ["Horizonte: 45 años","Perfil Moderado","6 Fondos iShares","100% ESG","Gestión Pasiva"]
    .forEach((c,i)=>pill(s,c,0.55+i*1.9,4.72,1.78,0.33,M,WH));
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 2 — PLANIFICACIÓN
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = sl(); bar(s);
  ttl(s,"Planificación: Objetivo y Aportaciones");
  sub(s,"Objetivo: Jubilación a los 69 años  ·  Horizonte: 45 años  ·  Perfil: Moderado (100 − edad)  ·  Estrategia DCA creciente");

  // Tabla aportaciones 3 fases
  const fases = [
    {f:"Fase 1",a:"25 años",sal:"25.000 €",ano:"5.000 €/año",mes:"417 €/mes",rv:"75%",rf:"25%"},
    {f:"Fase 2",a:"40 años",sal:"33.647 €",ano:"6.729 €/año",mes:"561 €/mes",rv:"60%",rf:"40%"},
    {f:"Fase 3",a:"60 años",sal:"49.997 €",ano:"9.999 €/año",mes:"833 €/mes",rv:"40%",rf:"60%"},
  ];
  fases.forEach((f,i)=>{
    const x=0.4+i*3.12;
    bx(s,x,1.05,2.9,3.4);
    s.addShape(pres.shapes.RECTANGLE,{x,y:1.05,w:2.9,h:0.42,fill:{color:D},line:{color:D}});
    s.addText(`${f.f}  ·  ${f.a}`,{x,y:1.05,w:2.9,h:0.42,fontSize:13,fontFace:F,bold:true,color:WH,align:"center",valign:"middle"});
    const rows=[[f.sal,"Salario neto"],[f.ano,"Ahorro anual"],[f.mes,"Mensual"]];
    rows.forEach(([v,l],j)=>{
      s.addText(v,{x:x+0.1,y:1.56+j*0.5,w:2.7,h:0.32,fontSize:j===0?14:16,fontFace:F,bold:true,color:M,align:"center",valign:"middle"});
      s.addText(l,{x:x+0.1,y:1.88+j*0.5,w:2.7,h:0.22,fontSize:10,fontFace:F,color:MU,align:"center",valign:"middle"});
    });
    s.addShape(pres.shapes.LINE,{x:x+0.2,y:3.18,w:2.5,h:0,line:{color:"D0DDE8",width:0.5}});
    s.addText(`RV ${f.rv}  /  RF ${f.rf}`,{x:x+0.1,y:3.25,w:2.7,h:0.38,fontSize:13,fontFace:F,bold:true,color:i<2?RV:RF,align:"center",valign:"middle"});
  });

  // DCA + Total
  bx(s,0.4,4.55,9.2,0.75,D);
  s.addText("Total aportado en 45 años:",{x:0.6,y:4.55,w:3.5,h:0.75,fontSize:13,fontFace:F,bold:true,color:WH,valign:"middle"});
  s.addText("359.464 €",{x:4.1,y:4.55,w:3.0,h:0.75,fontSize:26,fontFace:F,bold:true,color:AC,align:"center",valign:"middle"});
  s.addText("sin rentabilidad — cualquier retorno positivo multiplica este resultado",
    {x:7.1,y:4.55,w:2.4,h:0.75,fontSize:10,fontFace:F,color:"B0CCE0",valign:"middle",wrap:true});
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 3 — SESGOS COGNITIVOS + GESTIÓN PASIVA
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = sl(); bar(s);
  ttl(s,"Sesgos Mitigados y Por Qué Gestión Pasiva");

  const sesgos=[
    ["Exceso de confianza","Fondos indexados. <20% de gestores activos baten al índice a 20 años."],
    ["Aversión a pérdidas","DCA automático. Impide pausar la inversión en mercados bajistas."],
    ["Efecto disposición","Rebalanceos a los 40 y 60 fuerzan a vender lo que sube y comprar lo que baja."],
    ["Seguimiento del rebaño","IPS con criterios ESG escritos actúa como ancla racional."],
  ];
  sesgos.forEach(([nm,desc],i)=>{
    const x=0.4+i*2.4;
    bx(s,x,1.02,2.18,1.62);
    s.addShape(pres.shapes.RECTANGLE,{x,y:1.02,w:2.18,h:0.38,fill:{color:M},line:{color:M}});
    s.addText(nm,{x,y:1.02,w:2.18,h:0.38,fontSize:11,fontFace:F,bold:true,color:WH,align:"center",valign:"middle"});
    s.addText(desc,{x:x+0.1,y:1.44,w:1.98,h:1.14,fontSize:10.5,fontFace:F,color:TX,valign:"top",wrap:true});
  });

  // Gestión pasiva
  s.addText("Justificación de la gestión 100% pasiva",
    {x:0.4,y:2.82,w:9.2,h:0.3,fontSize:13,fontFace:F,bold:true,color:D});
  const razones=[
    ["11.1% fondos activos USA Large Cap superan su índice a 20 años (Morningstar)","Coste cartera: 0.17% vs 1.5% media activa → diferencia de cientos de miles € en 45 años"],
    ["Mercados desarrollados eficientes: precios ya reflejan toda la información disponible","Sin decisiones discrecionales = sin sesgos. Replica mecánica del mercado."],
  ];
  razones.forEach((par,i)=>par.forEach((t,j)=>{
    const x=0.4+(i*2+j)*2.35;
    bx(s,x,3.18,2.22,1.95,"F0F7FF");
    s.addShape(pres.shapes.RECTANGLE,{x,y:3.18,w:0.08,h:1.95,fill:{color:AC},line:{color:AC}});
    s.addText(t,{x:x+0.16,y:3.22,w:1.98,h:1.87,fontSize:11,fontFace:F,color:TX,valign:"middle",wrap:true});
  }));
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 4 — RESUMEN CARTERA (tabla 6 fondos)
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = sl(); bar(s);
  ttl(s,"Resumen de la Cartera — 6 Fondos iShares ESG");
  sub(s,"Todos los fondos son de BlackRock/iShares · Misma metodología ESG MSCI Extended Focus · 100% gestión pasiva indexada");

  const cols=["Ticker","Nombre","ISIN","Tipo","F1","F2","F3","Rent. 5a","OCF"];
  const cw=[0.72,3.0,1.2,0.68,0.72,0.72,0.72,0.88,0.64];
  const rows=[
    ["ESGU","iShares ESG Aware MSCI USA ETF","US46435G4257","RV","35%","30%","20%","13.07%","0.15%"],
    ["ESGD","iShares ESG Aware MSCI EAFE ETF","US46435G5163","RV","20%","15%","10%","8.62%","0.20%"],
    ["ESGE","iShares ESG MSCI Emerging Mkts ETF","US46434G8630","RV","10%","10%","5%","3.24%","0.25%"],
    ["ESML","iShares ESG Aware MSCI USA Sm-Cap","US46435U6635","RV","10%","5%","5%","7.50%","0.17%"],
    ["EAGG","iShares ESG Aware US Agg Bond ETF","US46435U5496","RF","13%","20%","30%","-0.43%","0.10%"],
    ["QLTA","iShares Aaa-A Rated Corp Bond ETF","US46429B2916","RF","10%","15%","25%","-0.62%","0.15%"],
  ];

  s.addTable(
    [cols.map(c=>({text:c,options:{bold:true,fill:{color:D},color:WH,fontSize:11,fontFace:F,align:"center"}})),
     ...rows.map((r,ri)=>r.map((v,ci)=>{
       const isRF=r[3]==="RF";
       const fc=ri%2===0?"FFFFFF":"EBF4FA";
       let vc2=TX;
       if(ci===3) vc2=isRF?RF:RV;
       else if(ci===7) vc2=v.startsWith("-")?NE:PO;
       return {text:v,options:{fill:{color:fc},color:vc2,fontSize:ci===1?10:11,fontFace:F,bold:ci<2||ci===3,align:"center"}};
     }))],
    {x:0.35,y:1.02,w:9.3,h:4.3,rowH:0.6,
     colW:cw,border:{pt:0.5,color:"D0DDE8"},align:"center"}
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 5 — FICHA FONDOS RV (1-2)
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = sl(); bar(s);
  ttl(s,"Fondos de Renta Variable — ESGU y ESGD");
  sub(s,"Núcleo americano + Diversificación internacional desarrollada");

  const fondos2=[
    {tk:"ESGU",nm:"iShares ESG Aware MSCI USA ETF",isin:"US46435G4257",
     geo:"EE.UU. 100%",
     sect:"Tech 37.2% · Financiero 12.0% · Comunicación 10.0% · Salud 8.6%",
     mets:[["Rent. 5a","13.07%"],["Volatilidad","15.66%"],["Sharpe","0.51"],["Beta","1.02"],["Max DD","−25.56%"],["OCF","0.15%"]],
     why:"Núcleo (35%). EE.UU.=65% cap. mundial. Mayor ETF ESG USA (15.500M$). Sharpe 0.51 el más alto en RV. OCF 0.15% el más bajo.",
     f1:"35%",f2:"30%",f3:"20%"},
    {tk:"ESGD",nm:"iShares ESG Aware MSCI EAFE ETF",isin:"US46435G5163",
     geo:"Japón 22.7% · UK 13.4% · Suiza 10.2% · Francia 8.7% · Alemania 8.5%",
     sect:"Financiero 26.1% · Industrials 19.1% · Salud 10.4% · Tech 11.7%",
     mets:[["Rent. 5a","8.62%"],["Volatilidad","15.84%"],["Sharpe","0.34"],["Beta","1.03"],["Max DD","−28.05%"],["OCF","0.20%"]],
     why:"Diversificación geográfica (20%). Rating MSCI ESG AAA (top 1% mundial). Alpha +0.90% positivo. Financials/Industrials complementan el sesgo tech de ESGU.",
     f1:"20%",f2:"15%",f3:"10%"},
  ];

  fondos2.forEach((f,i)=>{
    const x=0.35+i*4.8;
    bx(s,x,1.0,4.6,4.35);
    s.addShape(pres.shapes.RECTANGLE,{x,y:1.0,w:4.6,h:0.38,fill:{color:RV},line:{color:RV}});
    s.addText(`${f.tk}  ·  ${f.nm}`,{x:x+0.1,y:1.0,w:4.4,h:0.38,fontSize:12,fontFace:F,bold:true,color:WH,valign:"middle"});
    s.addText(`ISIN: ${f.isin}`,{x:x+0.1,y:1.42,w:4.4,h:0.22,fontSize:10,fontFace:F,color:MU,valign:"middle"});
    s.addText(f.geo,{x:x+0.1,y:1.66,w:4.4,h:0.22,fontSize:10,fontFace:F,color:TX,bold:true,valign:"middle",wrap:true});
    s.addText(f.sect,{x:x+0.1,y:1.9,w:4.4,h:0.3,fontSize:9.5,fontFace:F,color:MU,valign:"middle",wrap:true});

    // métricas en grid 3x2
    f.mets.forEach(([lbl,v],mi)=>{
      const mx=x+0.12+(mi%3)*1.5;
      const my=2.28+Math.floor(mi/3)*0.62;
      s.addShape(pres.shapes.RECTANGLE,{x:mx,y:my,w:1.38,h:0.56,
        fill:{color:mi%2===0?"F0F7FF":"FFFFFF"},line:{color:"D0DDE8",width:0.5}});
      s.addText(v,{x:mx,y:my,w:1.38,h:0.34,fontSize:14,fontFace:F,bold:true,
        color:vc(v),align:"center",valign:"middle"});
      s.addText(lbl,{x:mx,y:my+0.34,w:1.38,h:0.22,fontSize:9,fontFace:F,
        color:MU,align:"center",valign:"middle"});
    });

    s.addShape(pres.shapes.LINE,{x:x+0.1,y:3.56,w:4.4,h:0,line:{color:"D0DDE8",width:0.5}});
    s.addText(`F1 ${f.f1} · F2 ${f.f2} · F3 ${f.f3}`,
      {x:x+0.1,y:3.6,w:4.4,h:0.26,fontSize:11,fontFace:F,bold:true,color:RV,align:"center"});
    s.addText(f.why,{x:x+0.1,y:3.9,w:4.4,h:0.55,fontSize:9.5,fontFace:F,color:TX,valign:"top",wrap:true});
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 6 — FICHA FONDOS RV (3-4)
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = sl(); bar(s);
  ttl(s,"Fondos de Renta Variable — ESGE y ESML");
  sub(s,"Mercados emergentes + Prima de tamaño small-cap (Fama-French)");

  const fondos2=[
    {tk:"ESGE",nm:"iShares ESG MSCI Emerging Mkts ETF",isin:"US46434G8630",
     geo:"Taiwán 26.2% · China 21.5% · Corea del Sur 19.1% · India 10.9%",
     sect:"Tech 37.5% · Financiero 24.0% · Comunicación 8.4% · Cíclico 9.7%",
     mets:[["Rent. 5a","3.24%"],["Volatilidad","16.98%"],["Sharpe","0.09"],["Beta","1.00"],["Max DD","−37.80%"],["OCF","0.25%"]],
     why:"Prima de crecimiento emergente (>40% PIB mundial). Taiwán sobreexpuesto por ESG (26% vs 19% cat). Tech 37.5% vs 30.5% cat. — filtro ESG mejora calidad. Peso reducido por volatilidad y max DD.",
     f1:"10%",f2:"10%",f3:"5%"},
    {tk:"ESML",nm:"iShares ESG Aware MSCI USA Small-Cap",isin:"US46435U6635",
     geo:"EE.UU. ~100%",
     sect:"Industrials 19.2% · Tech 17.4% · Salud 12.9% · Financiero 14.4%",
     mets:[["Rent. 5a","7.50%"],["Volatilidad","19.00%"],["Sharpe","0.19"],["Beta","1.07"],["Max DD","−23.87%"],["OCF","0.17%"]],
     why:"Prima Fama-French a 45 años. Alpha −6.14% es artefacto metodológico (Morningstar compara small-cap vs S&P 500). Max DD contenido gracias al filtro ESG vs índices sin filtros.",
     f1:"10%",f2:"5%",f3:"5%"},
  ];

  fondos2.forEach((f,i)=>{
    const x=0.35+i*4.8;
    bx(s,x,1.0,4.6,4.35);
    s.addShape(pres.shapes.RECTANGLE,{x,y:1.0,w:4.6,h:0.38,fill:{color:RV},line:{color:RV}});
    s.addText(`${f.tk}  ·  ${f.nm}`,{x:x+0.1,y:1.0,w:4.4,h:0.38,fontSize:11.5,fontFace:F,bold:true,color:WH,valign:"middle"});
    s.addText(`ISIN: ${f.isin}`,{x:x+0.1,y:1.42,w:4.4,h:0.22,fontSize:10,fontFace:F,color:MU,valign:"middle"});
    s.addText(f.geo,{x:x+0.1,y:1.66,w:4.4,h:0.22,fontSize:10,fontFace:F,color:TX,bold:true,valign:"middle",wrap:true});
    s.addText(f.sect,{x:x+0.1,y:1.9,w:4.4,h:0.3,fontSize:9.5,fontFace:F,color:MU,valign:"middle",wrap:true});
    f.mets.forEach(([lbl,v],mi)=>{
      const mx=x+0.12+(mi%3)*1.5;
      const my=2.28+Math.floor(mi/3)*0.62;
      s.addShape(pres.shapes.RECTANGLE,{x:mx,y:my,w:1.38,h:0.56,
        fill:{color:mi%2===0?"F0F7FF":"FFFFFF"},line:{color:"D0DDE8",width:0.5}});
      s.addText(v,{x:mx,y:my,w:1.38,h:0.34,fontSize:14,fontFace:F,bold:true,color:vc(v),align:"center",valign:"middle"});
      s.addText(lbl,{x:mx,y:my+0.34,w:1.38,h:0.22,fontSize:9,fontFace:F,color:MU,align:"center",valign:"middle"});
    });
    s.addShape(pres.shapes.LINE,{x:x+0.1,y:3.56,w:4.4,h:0,line:{color:"D0DDE8",width:0.5}});
    s.addText(`F1 ${f.f1} · F2 ${f.f2} · F3 ${f.f3}`,
      {x:x+0.1,y:3.6,w:4.4,h:0.26,fontSize:11,fontFace:F,bold:true,color:RV,align:"center"});
    s.addText(f.why,{x:x+0.1,y:3.9,w:4.4,h:0.55,fontSize:9.5,fontFace:F,color:TX,valign:"top",wrap:true});
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 7 — FICHA FONDOS RF (5-6)
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = sl(); bar(s);
  ttl(s,"Fondos de Renta Fija — EAGG y QLTA");
  sub(s,"Estabilizadores de la cartera · Alta calidad crediticia · Correlación baja con RV");

  const fondos2=[
    {tk:"EAGG",nm:"iShares ESG Aware US Agg Bond ETF",isin:"US46435U5496",
     geo:"EE.UU. ~97%  ·  Composición: Gobierno 47% | Corp IG 24% | Titulizado 24%",
     sect:"Govt (Tesoro+agencias) 47.03% · Securitized 24.25% · Corporate 24.11%",
     cred:"AAA impl. (Tesoro 47%) · Investment Grade 100% · Sin high-yield",
     mets:[["Rent. 5a","−0.43%"],["Volatilidad","6.39%"],["Sharpe","−0.50"],["Beta","1.00"],["Max DD","−6.16%"],["OCF","0.10%"]],
     why:"Primer estabilizador (13%/20%/30%). OCF 0.10% = el más bajo de toda la cartera. 47% en Treasuries. Rent. negativa por ciclo de tipos 2022: a 3 años recupera +3.81%. Alpha −0.10% confirma réplica perfecta.",
     f1:"13%",f2:"20%",f3:"30%"},
    {tk:"QLTA",nm:"iShares Aaa-A Rated Corp Bond ETF",isin:"US46429B2916",
     geo:"EE.UU. ~95%  ·  100% bonos corporativos — Sin BBB ni inferior por definición",
     sect:"Corporate (Aaa-A): 99.80%  ·  Cash/Deriv.: 0.20%",
     cred:"AAA 1.9% · AA 14.1% · A 83.2% · BBB 0% · Sin high-yield",
     mets:[["Rent. 5a","−0.62%"],["Volatilidad","6.43%"],["Sharpe","0.00"],["Beta","1.22"],["Max DD","−20.59%"],["OCF","0.15%"]],
     why:"Segundo estabilizador RF (10%/15%/25%). 98% en AAA/AA/A — máxima calidad crediticia. Alpha +0.63% POSITIVO: único en RF. Medalla Bronze Morningstar (mar 2026). Clave en Fase 3 cuando RF=60%.",
     f1:"10%",f2:"15%",f3:"25%"},
  ];

  fondos2.forEach((f,i)=>{
    const x=0.35+i*4.8;
    bx(s,x,1.0,4.6,4.35);
    s.addShape(pres.shapes.RECTANGLE,{x,y:1.0,w:4.6,h:0.38,fill:{color:RF},line:{color:RF}});
    s.addText(`${f.tk}  ·  ${f.nm}`,{x:x+0.1,y:1.0,w:4.4,h:0.38,fontSize:11.5,fontFace:F,bold:true,color:WH,valign:"middle"});
    s.addText(`ISIN: ${f.isin}`,{x:x+0.1,y:1.42,w:4.4,h:0.22,fontSize:10,fontFace:F,color:MU,valign:"middle"});
    s.addText(f.cred,{x:x+0.1,y:1.66,w:4.4,h:0.22,fontSize:10,fontFace:F,color:RF,bold:true,valign:"middle",wrap:true});
    s.addText(f.sect,{x:x+0.1,y:1.9,w:4.4,h:0.3,fontSize:9.5,fontFace:F,color:MU,valign:"middle",wrap:true});
    f.mets.forEach(([lbl,v],mi)=>{
      const mx=x+0.12+(mi%3)*1.5;
      const my=2.28+Math.floor(mi/3)*0.62;
      s.addShape(pres.shapes.RECTANGLE,{x:mx,y:my,w:1.38,h:0.56,
        fill:{color:mi%2===0?"F0FBF9":"FFFFFF"},line:{color:"D0DDE8",width:0.5}});
      s.addText(v,{x:mx,y:my,w:1.38,h:0.34,fontSize:14,fontFace:F,bold:true,color:vc(v),align:"center",valign:"middle"});
      s.addText(lbl,{x:mx,y:my+0.34,w:1.38,h:0.22,fontSize:9,fontFace:F,color:MU,align:"center",valign:"middle"});
    });
    s.addShape(pres.shapes.LINE,{x:x+0.1,y:3.56,w:4.4,h:0,line:{color:"D0DDE8",width:0.5}});
    s.addText(`F1 ${f.f1} · F2 ${f.f2} · F3 ${f.f3}`,
      {x:x+0.1,y:3.6,w:4.4,h:0.26,fontSize:11,fontFace:F,bold:true,color:RF,align:"center"});
    s.addText(f.why,{x:x+0.1,y:3.9,w:4.4,h:0.55,fontSize:9.5,fontFace:F,color:TX,valign:"top",wrap:true});
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 8 — DISTRIBUCIÓN POR FASES
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = sl(); bar(s);
  ttl(s,"Distribución de la Cartera por Fases");
  sub(s,"Los mismos 6 fondos en las tres fases — solo cambian los pesos · Rebalanceo redirigiendo aportaciones, no vendiendo");

  const faseInfo=[
    {lbl:"Fase 1 · 25-39 años",rv:"75%",rf:"25%",aport:"5.000 €/año · 417 €/mes",
     rows:[["ESGU","RV","35%","1.750 €"],["ESGD","RV","20%","1.000 €"],["ESGE","RV","10%","500 €"],
           ["ESML","RV","10%","500 €"],["EAGG","RF","13%","625 €"],["QLTA","RF","10%","500 €"]],bg:"EBF5FF"},
    {lbl:"Fase 2 · 40-59 años",rv:"60%",rf:"40%",aport:"6.729 €/año · 561 €/mes",
     rows:[["ESGU","RV","30%","2.019 €"],["ESGD","RV","15%","1.009 €"],["ESGE","RV","10%","673 €"],
           ["ESML","RV","5%","336 €"],["EAGG","RF","20%","1.346 €"],["QLTA","RF","15%","1.009 €"]],bg:"EBF9F6"},
    {lbl:"Fase 3 · 60-69 años",rv:"40%",rf:"60%",aport:"9.999 €/año · 833 €/mes",
     rows:[["ESGU","RV","20%","2.000 €"],["ESGD","RV","10%","1.000 €"],["ESGE","RV","5%","500 €"],
           ["ESML","RV","5%","500 €"],["EAGG","RF","30%","3.000 €"],["QLTA","RF","25%","2.500 €"]],bg:"FFF8EE"},
  ];

  faseInfo.forEach((f,col)=>{
    const x=0.32+col*3.24;
    const w=3.05;
    bx(s,x,0.98,w,4.35,f.bg);
    s.addShape(pres.shapes.RECTANGLE,{x,y:0.98,w,h:0.38,fill:{color:D},line:{color:D}});
    s.addText(f.lbl,{x,y:0.98,w,h:0.38,fontSize:12,fontFace:F,bold:true,color:WH,align:"center",valign:"middle"});
    s.addText(`RV ${f.rv}  /  RF ${f.rf}  ·  ${f.aport}`,
      {x:x+0.06,y:1.4,w:w-0.12,h:0.32,fontSize:10.5,fontFace:F,bold:true,color:D,align:"center",valign:"middle"});

    s.addTable(
      [[{text:"Fondo",options:{bold:true,fill:{color:M},color:WH,fontSize:10,fontFace:F}},
        {text:"T",options:{bold:true,fill:{color:M},color:WH,fontSize:10,fontFace:F}},
        {text:"Peso",options:{bold:true,fill:{color:M},color:WH,fontSize:10,fontFace:F}},
        {text:"€/año",options:{bold:true,fill:{color:M},color:WH,fontSize:10,fontFace:F}}],
       ...f.rows.map((r,ri)=>r.map((v,ci)=>({text:v,options:{
         fill:{color:ri%2===0?"FFFFFF":f.bg},
         color:ci===1?(v==="RV"?RV:RF):ci===3?M:TX,
         fontSize:10,fontFace:F,bold:ci===3,align:"center"
       }})))],
      {x,y:1.76,w,h:3.14,rowH:0.44,
       colW:[0.82,0.38,0.55,0.9],
       border:{pt:0.5,color:"D0DDE8"},align:"center"}
    );
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 9 — GRÁFICO EVOLUCIÓN RV/RF + MÉTRICAS COMPARATIVAS
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = sl(); bar(s);
  ttl(s,"Evolución RV/RF y Comparativa de Métricas");

  // Gráfico barras apiladas
  s.addChart(pres.charts.BAR,[
    {name:"Renta Variable",labels:["Fase 1\n25-39a","Fase 2\n40-59a","Fase 3\n60-69a"],values:[75,60,40]},
    {name:"Renta Fija",   labels:["Fase 1\n25-39a","Fase 2\n40-59a","Fase 3\n60-69a"],values:[25,40,60]},
  ],{
    x:0.35,y:0.9,w:4.8,h:4.4,
    barDir:"col",barGrouping:"stacked",
    chartColors:[RV,RF],
    chartArea:{fill:{color:WH},roundedCorners:false},
    catAxisLabelColor:TX,valAxisLabelColor:MU,
    valGridLine:{color:"E2E8F0",size:0.5},catGridLine:{style:"none"},
    showValue:true,dataLabelColor:WH,dataLabelFontSize:13,dataLabelFontBold:true,
    showLegend:true,legendPos:"b",legendFontSize:12,legendColor:TX,
    valAxisMaxVal:100,
  });

  // Tabla métricas comparativas
  const mrows=[
    ["","ESGU","ESGD","ESGE","ESML","EAGG","QLTA"],
    ["Rent. 5a","13.07%","8.62%","3.24%","7.50%","−0.43%","−0.62%"],
    ["Volatilidad","15.66%","15.84%","16.98%","19.00%","6.39%","6.43%"],
    ["Sharpe","0.51","0.34","0.09","0.19","−0.50","0.00"],
    ["Beta","1.02","1.03","1.00","1.07","1.00","1.22"],
    ["Alpha","−1.40%","+0.90%","−2.95%","−6.14%","−0.10%","+0.63%"],
    ["Max DD","−25.56%","−28.05%","−37.80%","−23.87%","−6.16%","−20.59%"],
    ["OCF","0.15%","0.20%","0.25%","0.17%","0.10%","0.15%"],
  ];

  s.addTable(
    mrows.map((r,ri)=>r.map((v,ci)=>{
      const isHdr=ri===0||ci===0;
      const fc=isHdr?(ri===0?D:M):ri%2===0?"FFFFFF":"EBF4FA";
      const tc=isHdr?WH:(v.startsWith("−")||v.startsWith("-")?NE:(parseFloat(v)>0&&(ri===1||ri===5)?PO:TX));
      return {text:v,options:{fill:{color:fc},color:tc,fontSize:10,fontFace:F,bold:isHdr,align:"center"}};
    })),
    {x:5.35,y:0.9,w:4.3,h:4.4,rowH:0.49,
     colW:[0.85,0.57,0.57,0.57,0.57,0.6,0.57],
     border:{pt:0.5,color:"D0DDE8"},align:"center"}
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 10 — CONCLUSIONES
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = sl(); bar(s);
  ttl(s,"Conclusiones");

  const concs=[
    ["La asignación estratégica lo es todo",
     "RV/RF por fase (75/25→60/40→40/60) explica la mayor parte del rendimiento. La selección de fondos importa menos que la asignación de activos."],
    ["ESG no sacrifica rentabilidad",
     "ESGU +13.07%, ESGD +8.62%, ESML +7.50% a 5 años. Los criterios sostenibles son compatibles con rentabilidades competitivas. ESGD incluso genera alpha positivo (+0.90%)."],
    ["Los costes son decisivos a 45 años",
     "Coste medio cartera: 0.17% vs 1.5% activos. Con 359.464 € aportados, la diferencia capitalizada puede suponer más de 400.000 € adicionales al jubilarse."],
    ["Rentabilidad negativa en RF: efecto transitorio",
     "El ciclo de tipos 2022-2023 fue el peor en 40 años para la RF. EAGG a 3 años: +3.81%. QLTA a 3 años: +5.39%. La recuperación ya está en curso."],
    ["El proceso estructurado evita errores costosos",
     "Un plan escrito con IPS, criterios ESG y rebalanceos definidos protege contra los sesgos cognitivos identificados. El proceso vale más que cualquier predicción de mercado."],
  ];

  concs.forEach(([t,d],i)=>{
    const y=1.02+i*0.87;
    const bg=i%2===0?WH:"F0F7FF";
    bx(s,0.38,y,9.24,0.78,bg);
    s.addShape(pres.shapes.RECTANGLE,{x:0.38,y,w:0.08,h:0.78,fill:{color:AC},line:{color:AC}});
    s.addText(t,{x:0.54,y:y+0.04,w:3.0,h:0.7,fontSize:12.5,fontFace:F,bold:true,color:D,valign:"middle",wrap:true});
    s.addText(d,{x:3.6,y:y+0.04,w:5.9,h:0.7,fontSize:11,fontFace:F,color:TX,valign:"middle",wrap:true});
  });

  // Pie de cierre
  bx(s,0.38,5.32,9.24,0.52,D);
  s.addText("Jessica Elizabeth Rojas Rodriguez  ·  MBA UVa  ·  Activos e Inversiones Financieras  ·  2025-2026",
    {x:0.5,y:5.32,w:9,h:0.52,fontSize:12,fontFace:F,color:WH,align:"center",valign:"middle"});
}

// ═══════════════════════════════════════════════════════════════════════════

  try {
    await pres.writeFile({ fileName: "Cartera_ESG_Final.pptx" });
    console.log("OK");
  } catch (e) {
    console.error(e);
  }
});
