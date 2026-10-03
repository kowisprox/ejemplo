const pptxgen = require("pptxgenjs");
const fs = require("fs");
const S = fs.readFileSync("./S.txt", "utf8").trim();
const { applyTheme } = require(S + "/scripts/apply_theme.js");
const T = JSON.parse(fs.readFileSync("../trace.json", "utf8"));

// ---------- tokens ----------
const INK = "1F1B16", CREAM = "F4F1EA", RED = "8C2F23", GREEN = "2F5D50";
const MUTE = "6E6657", RULE = "B5AD9B", TINT = "EAE5D8";
const H = "EB Garamond", B = "IBM Plex Sans", M = "JetBrains Mono";
const THEME = {
  name: "Editorial ADA", headFontFace: H, bodyFontFace: B,
  colors: { dk1: INK, lt1: CREAM, dk2: MUTE, lt2: TINT, accent1: RED, accent2: GREEN,
    accent3: "B5AD9B", accent4: "3D372E", accent5: "9A9184", accent6: "CFC7B6", hlink: RED, folHlink: MUTE },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5, 16:9
pres.theme = { headFontFace: H, bodyFontFace: B };
pres.title = "Counting sort, Radix sort y Bucket sort";
pres.subject = "Análisis y Diseño de Algoritmos";

const FOOT = { x: 0.6, y: 7.08, w: 9, h: 0.25, fontFace: B, fontSize: 10, color: MUTE, margin: 0 };
pres.defineSlideMaster({
  title: "CONTENIDO", background: { color: CREAM },
  objects: [
    { line: { x: 0.6, y: 0.72, w: 12.13, h: 0, line: { color: INK, width: 0.75 } } },
    { line: { x: 0.6, y: 7.0, w: 12.13, h: 0, line: { color: RULE, width: 0.5 } } },
    { text: { text: "Análisis y Diseño de Algoritmos · Counting, Radix y Bucket sort", options: FOOT } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.88, w: 12.13, h: 0.85, fontFace: H, fontSize: 36, color: INK, bold: false, margin: 0, valign: "top", align: "left" }, text: "" } },
  ],
  slideNumber: { x: 11.73, y: 7.08, w: 1.0, h: 0.25, fontFace: M, fontSize: 10, color: MUTE, align: "right" },
});
pres.defineSlideMaster({ title: "PORTADA", background: { color: CREAM }, objects: [] });

for (const title of ['Apertura','El problema','Counting sort','Radix sort','Bucket sort','Balance','Cierre']) pres.addSection({ title });
// ---------- helpers ----------
const hair = (s, x, y, w, c = RULE, wd = 0.5) => s.addShape(pres.ShapeType.line, { x, y, w, h: 0, line: { color: c, width: wd } });
const vhair = (s, x, y, h, c = RULE, wd = 0.5) => s.addShape(pres.ShapeType.line, { x, y, w: 0, h, line: { color: c, width: wd } });
const txt = (s, t, o) => s.addText(t, Object.assign({ isTextBox: true, margin: 0, valign: "top", fontFace: B, fontSize: 14, color: INK }, o));
const kicker = (s, t) => txt(s, t, { x: 0.6, y: 0.36, w: 9, h: 0.25, fontFace: M, fontSize: 11, color: RED, charSpacing: 2 });
function slide(section, kick, title, notes) {
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: section });
  kicker(s, kick);
  s.addText(title, { placeholder: "title" });
  s.addNotes(notes);
  return s;
}
function cell(s, v, x, y, w, h, o = {}) {
  s.addText(v, Object.assign({ shape: pres.ShapeType.rect, isTextBox: true, x, y, w, h, fill: { color: CREAM }, line: { color: INK, width: 0.75 },
    fontFace: M, fontSize: 14, color: INK, align: "center", valign: "middle", margin: 0 }, o));
}
const strip = (s, arr, x, y, cw, ch, o) => arr.forEach((v, i) => cell(s, String(v), x + i * cw, y, cw, ch, typeof o === "function" ? o(v, i) : o));
function paras(list, size = 14) {
  const runs = [];
  list.forEach(([lead, body], i) => {
    const last = i === list.length - 1;
    runs.push({ text: lead + " ", options: { bold: true, color: RED, fontFace: B, fontSize: size } });
    runs.push({ text: body, options: { breakLine: !last, fontFace: B, fontSize: size, color: INK, paraSpaceAfter: 9 } });
  });
  return runs;
}
function twoCol(s, intro, L, R) {
  txt(s, intro, { x: 0.6, y: 1.8, w: 12.13, h: 0.85, fontFace: H, fontSize: 22, italic: true, color: INK });
  hair(s, 0.6, 2.72, 12.13, INK, 0.5);
  txt(s, paras(L), { x: 0.6, y: 2.92, w: 5.85, h: 3.95, lineSpacingMultiple: 1.05 });
  vhair(s, 6.665, 2.92, 3.85);
  txt(s, paras(R), { x: 6.88, y: 2.92, w: 5.85, h: 3.95, lineSpacingMultiple: 1.05 });
}
const KW = /(\bpara\b|\bhasta\b|\bdescendiendo\b|\bdevolver\b|\bsea\b|\bañadir\b|\bal final de\b)/;
function codeRuns(line) {
  const first = line.trim().match(/^[A-Za-z\-]+\(/) && !line.startsWith(" ") ? "head" : "";
  if (first) { const i = line.indexOf("("); return [{ text: line.slice(0, i), options: { bold: true, color: RED } }, { text: line.slice(i), options: { color: INK } }]; }
  return line.split(KW).filter(Boolean).map((t) => KW.test(t) ? { text: t, options: { bold: true, color: GREEN } } : { text: t, options: { color: INK } });
}
// blocks: [{ from, lines:[...], lead, note }]  -> numbered code on the left, reason on the right
function pseudo(s, blocks, y0) {
  const LH = 20, codeX = 1.05, codeW = 4.25, comX = 5.65, comW = 7.08;
  let y = y0, n = 1;
  blocks.forEach((b) => {
    const codeH = b.lines.length * (LH / 72);
    const cpl = Math.floor((comW * 72) / (13 * 0.53));
    const comH = Math.ceil((b.lead.length + b.note.length + 1) / cpl) * 0.235;
    const h = Math.max(codeH, comH) + 0.14;
    hair(s, 0.6, y, 12.13, RULE, 0.5);
    const nums = b.lines.map((_, i) => ({ text: String(n + i), options: { breakLine: i < b.lines.length - 1 } }));
    txt(s, nums, { x: 0.6, y: y + 0.07, w: 0.35, h: codeH, fontFace: M, fontSize: 12, color: RED, align: "right", lineSpacing: LH });
    const runs = [];
    b.lines.forEach((l, i) => {
      const r = codeRuns(l); r.forEach((x) => (x.options = Object.assign({ fontFace: M, fontSize: 13 }, x.options)));
      r[r.length - 1].options.breakLine = i < b.lines.length - 1; runs.push(...r);
    });
    txt(s, runs, { x: codeX, y: y + 0.07, w: codeW, h: codeH, lineSpacing: LH });
    txt(s, [{ text: b.lead + " ", options: { bold: true, color: INK } }, { text: b.note, options: { color: INK } }],
      { x: comX, y: y + 0.07, w: comW, h: comH, fontSize: 13 });
    n += b.lines.length; y += h;
  });
  hair(s, 0.6, y, 12.13, INK, 0.75);
  console.log("pseudo bottom y =", y.toFixed(2));
  return y;
}
function table(s, rows, o) {
  const none = { type: "none" }, line = (c, pt) => ({ type: "solid", color: c, pt });
  const data = rows.map((r, ri) => r.map((c, ci) => {
    const head = ri === 0;
    const base = { fontFace: head || ci === 0 ? H : (o.mono && o.mono.includes(ci) ? M : B), fontSize: head ? 16 : (o.mono && o.mono.includes(ci) ? 13 : 14),
      bold: head, color: head ? INK : (ci === 0 ? INK : INK), valign: "middle", margin: [0.04, 0.08, 0.04, 0],
      border: [none, none, head ? line(INK, 1) : line(RULE, 0.5), none] };
    if (ri === rows.length - 1) base.border[2] = line(INK, 0.75);
    if (typeof c === "object") return { text: c.t, options: Object.assign(base, c.o) };
    return { text: c, options: base };
  }));
  s.addTable(data, { x: o.x, y: o.y, w: o.w, colW: o.colW, rowH: o.rowH, autoPage: false, fill: { color: CREAM } });
}

// =============== 1. PORTADA ===============
{
  const s = pres.addSlide({ masterName: "PORTADA", sectionTitle: "Apertura" });
  txt(s, "ANÁLISIS Y DISEÑO DE ALGORITMOS", { x: 0.6, y: 0.5, w: 8, h: 0.3, fontFace: M, fontSize: 12, color: RED, charSpacing: 3 });
  hair(s, 0.6, 0.9, 12.13, INK, 0.75);
  txt(s, "Ordenar sin comparar", { x: 0.6, y: 1.6, w: 12.13, h: 1.5, fontFace: H, fontSize: 68, color: INK });
  txt(s, "Counting sort, Radix sort y Bucket sort: los tres algoritmos que rompen el techo de n log n.", { x: 0.6, y: 3.3, w: 8.4, h: 1.0, fontFace: H, fontSize: 28, italic: true, color: MUTE });
  txt(s, "ENTRADA DE TRABAJO", { x: 0.6, y: 5.0, w: 4, h: 0.25, fontFace: M, fontSize: 11, color: MUTE, charSpacing: 2 });
  strip(s, T.A, 0.6, 5.35, 1.0, 0.7, (v) => ({ fontSize: 18 }));
  txt(s, "n = 7 · k = 100", { x: 8.0, y: 5.5, w: 4.7, h: 0.4, fontFace: M, fontSize: 16, color: GREEN, align: "right" });
  hair(s, 0.6, 6.7, 12.13, RULE, 0.5);
  txt(s, "Counting sort · Radix sort · Bucket sort — descripción, pseudocódigo, traza, costos y comparación", { x: 0.6, y: 6.85, w: 12.13, h: 0.3, fontSize: 12, color: MUTE });
  s.addNotes("Buenas. Hoy vamos a hablar de tres algoritmos de ordenamiento que no se parecen a los que todos conocemos: Counting sort, Radix sort y Bucket sort. Lo que tienen en común es que no comparan elementos entre sí, y por eso pueden esquivar el límite de n log n que vale para burbuja, mergesort o quicksort. Todos los ejemplos de la exposición usan la misma entrada, estos siete números: 50, 1, 100, 20, 19, 40 y 99. Los valores de las trazas no son estimados: los obtuve ejecutando los tres algoritmos en Python antes de armar las láminas. Y les adelanto un punto que normalmente se omite: que un algoritmo sea lineal no significa que sea rápido. Lo vamos a ver con números en la lámina 15.");
}

// =============== 2. EL PROBLEMA ===============
{
  const s = slide("El problema", "§ 01 · EL PROBLEMA", "La barrera de las comparaciones",
    "Empecemos por entender por qué hay un techo. Burbuja, inserción, selección, mergesort, quicksort y heapsort solo hacen una pregunta sobre los datos: ¿a es menor que b? Cualquier algoritmo así se puede dibujar como un árbol de decisión binario. Cada nodo es una comparación y cada hoja es un orden posible de salida. Como hay n factorial permutaciones, el árbol necesita al menos n factorial hojas, y un árbol de altura h tiene como máximo 2 a la h. De ahí sale h mayor o igual a log de n factorial, que es theta de n log n. Para nuestros siete números: 7 factorial es 5040 hojas, y la altura mínima es 13 comparaciones en el peor caso. Cuando ejecuté mergesort sobre la entrada hizo 13 comparaciones, o sea que toca el piso. Mergesort y heapsort son óptimos dentro del modelo. Los tres algoritmos de hoy salen del modelo: no comparan, leen la clave y calculan la posición. El teorema no les aplica, pero cada uno exige supuestos sobre los datos.");
  txt(s, paras([
    ["Una sola pregunta.", "Burbuja, inserción, selección, mergesort, quicksort y heapsort solo saben preguntar si a es menor que b. Cualquier algoritmo de ese tipo se puede dibujar como un árbol de decisión binario: cada nodo es una comparación y cada hoja, una permutación posible de la entrada."],
    ["De ahí sale el techo.", "Para ordenar n elementos el árbol necesita al menos una hoja por cada permutación, y un árbol binario de altura h no tiene más de 2^h hojas. Es un resultado demostrado, no una limitación de ingenio."],
    ["Lo que sí se puede hacer.", "Mergesort y heapsort llegan a ese límite: son óptimos dentro del modelo y no hay nada mejor que buscar. Los tres algoritmos de hoy salen del modelo: no comparan elementos, leen la estructura interna de la clave y calculan la posición. El teorema no les aplica, pero cada uno compra velocidad con supuestos sobre los datos."],
  ]), { x: 0.6, y: 1.9, w: 6.7, h: 4.9, lineSpacingMultiple: 1.05 });
  vhair(s, 7.65, 1.95, 4.8);
  txt(s, "EL LÍMITE EN TRES PASOS", { x: 8.0, y: 1.9, w: 4.7, h: 0.25, fontFace: M, fontSize: 11, color: RED, charSpacing: 2 });
  const f = [["2^h ≥ n!", "h: altura del árbol = comparaciones en el peor caso."], ["h ≥ log₂(n!)", "Se toma logaritmo a ambos lados."], ["log₂(n!) = Θ(n log n)", "Aproximación de Stirling."]];
  f.forEach(([a, b], i) => {
    const y = 2.3 + i * 0.95; hair(s, 8.0, y, 4.73);
    txt(s, a, { x: 8.0, y: y + 0.08, w: 4.7, h: 0.5, fontFace: H, fontSize: 28, color: INK });
    txt(s, b, { x: 8.0, y: y + 0.58, w: 4.7, h: 0.28, fontSize: 12, color: MUTE });
  });
  hair(s, 8.0, 5.2, 4.73, INK, 0.75);
  txt(s, [{ text: "Con n = 7: ", options: { bold: true, color: GREEN } }, { text: "7! = 5 040 hojas, así que h ≥ ⌈log₂ 5 040⌉ = 13. Mergesort sobre nuestra entrada hizo 13 comparaciones (medido).", options: {} }],
    { x: 8.0, y: 5.32, w: 4.7, h: 1.2, fontSize: 14 });
}

// =============== 3. POSICIÓN ===============
{
  const s = slide("El problema", "§ 01 · EL PROBLEMA", "Calcular la posición en vez de buscarla",
    "Los tres algoritmos reemplazan la pregunta «¿a es menor que b?» por una cuenta aritmética sobre la clave. Counting sort lee el valor entero completo y lo usa como índice. Radix sort lee un dígito por pasada. Bucket sort usa la posición del valor dentro del rango para elegir una cubeta. El precio es que dejan de ser de propósito general: cada uno asume algo sobre los datos y, si el supuesto falla, el algoritmo falla con él. Miren la fila «Es lineal cuando»: los tres dicen tiempo lineal, pero cada uno con su condición. Los tres son estables y ninguno es in-place: la velocidad se paga con memoria.");
  txt(s, "Los tres reemplazan la pregunta «¿a < b?» por una cuenta aritmética sobre la clave. El precio: dejan de ser de propósito general.", { x: 0.6, y: 1.8, w: 12.13, h: 0.8, fontFace: H, fontSize: 22, italic: true });
  table(s, [
    ["", "Counting sort", "Radix sort", "Bucket sort"],
    ["Qué lee de la clave", "El valor entero completo", "Un dígito por pasada", "Su posición dentro del rango"],
    ["Qué asume", "Enteros en [0, k], con k conocido", "Claves de d dígitos en base b", "Distribución aproximadamente uniforme"],
    ["Memoria extra", "C[0..k] y el arreglo de salida B", "b contadores, reutilizados d veces", "m listas encadenadas"],
    ["Es lineal cuando", { t: "k = O(n)", o: { color: RED, bold: true } }, { t: "d es constante", o: { color: RED, bold: true } }, { t: "el reparto es parejo", o: { color: RED, bold: true } }],
    ["Estable", "Sí", "Sí, y es imprescindible", "Sí, si la cubeta usa inserción"],
  ], { x: 0.6, y: 2.75, w: 12.13, colW: [2.5, 3.2, 3.2, 3.23], rowH: [0.5, 0.55, 0.55, 0.55, 0.55, 0.55] });
  txt(s, "Ninguno es in-place: la velocidad se paga con memoria.", { x: 0.6, y: 6.2, w: 12.13, h: 0.4, fontSize: 14, color: MUTE });
}

// =============== 4. COUNTING: CÓMO FUNCIONA ===============
{
  const s = slide("Counting sort", "§ 02 · COUNTING SORT", "Counting sort: cómo funciona",
    "Counting sort no mueve ningún elemento hasta saber exactamente dónde va. Toda la inteligencia está en el arreglo de conteo C. Primero contamos: por cada valor v que aparece, sumamos uno a C[v]. Eso obliga a que las claves sean enteros en un rango conocido de 0 a k: en nuestro ejemplo k es 100, así que C tiene 101 casillas y solo siete se ocupan. Segundo, acumulamos: cada casilla se suma con la anterior, y entonces C[v] pasa a decir cuántos elementos son menores o iguales que v. Para el 50 el acumulado es 5, así que el 50 va en la posición 4 contando desde cero. Tercero, colocamos recorriendo A de derecha a izquierda: restamos uno a C[x] y escribimos x en B[C[x]]. Recorrer al revés es lo que hace al algoritmo estable, y la estabilidad es lo que después necesita Radix sort. En ningún momento comparamos un elemento con otro.");
  twoCol(s, "No mueve ningún elemento hasta saber exactamente dónde va: primero calcula la posición final de cada valor y después lo escribe allí, una sola vez.", [
    ["Contar.", "Se recorre A una vez y, por cada valor v que aparece, se incrementa C[v]. Al terminar, C es un histograma de la entrada. Esto exige claves enteras en un rango conocido [0, k]: C necesita una casilla por cada valor posible, aparezca o no. Con nuestra entrada k = 100, así que C tiene 101 casillas y solo siete se ocupan."],
    ["Acumular.", "Se recorre C sumando cada casilla a la anterior. Ahora C[v] ya no dice cuántos v hay, sino cuántos elementos son menores o iguales que v, y ese número es la posición final del último v más uno. Para el 50 el acumulado es 5: hay cinco elementos ≤ 50, así que el 50 irá en B[4]."],
  ], [
    ["Colocar.", "Se recorre A de derecha a izquierda. Para cada x se resta uno a C[x] y se escribe x en B[C[x]]. El decremento deja lista la posición de un siguiente x igual, una casilla más a la izquierda. En la traza: A[6] = 99 cae en B[5], A[5] = 40 en B[3], y así hasta A[0] = 50 en B[4]."],
    ["Por qué al revés.", "Así, de dos elementos con la misma clave, el que estaba más a la derecha se coloca primero y queda más a la derecha: el orden relativo se conserva y el algoritmo es estable. No es un detalle estético, es lo que hace posible Radix sort. Nunca se compara un elemento con otro: cada posición sale de una suma."],
  ]);
}

// =============== 5. COUNTING: PSEUDOCÓDIGO ===============
{
  const s = slide("Counting sort", "§ 02 · COUNTING SORT", "Counting sort: pseudocódigo",
    "Leamos el pseudocódigo por bloques. Las líneas 2 y 3 crean dos arreglos: C del tamaño del rango y B del tamaño de la entrada. Ahí se ve por qué el algoritmo no puede ser in-place. Las líneas 4 a 7 hacen el histograma: poner C en ceros cuesta theta de k, y contar apariciones cuesta theta de n. Las líneas 8 y 9 hacen la suma acumulada, otro theta de k. Al salir, C[i] es la cantidad de elementos menores o iguales que i. Las líneas 10 a 12 colocan cada elemento. Se resta antes de escribir porque B empieza en cero: si hay cinco elementos menores o iguales que 20, el último 20 va en B[4]. El costo total es theta de n más k, y no hay mejor ni peor caso: el mismo trabajo con datos ordenados, invertidos o al azar.");
  pseudo(s, [
    { lines: ["COUNTING-SORT(A, n, k)"], lead: "Firma.", note: "k, el valor máximo posible, viaja en la firma: ya delata el supuesto de que el rango se conoce de antemano." },
    { lines: ["  sea C[0..k] arreglo nuevo", "  sea B[0..n-1] arreglo nuevo"], lead: "Líneas 2–3.", note: "C del tamaño del rango, B del tamaño de la entrada. Aquí se ve por qué no puede ser in-place." },
    { lines: ["  para i = 0 hasta k", "    C[i] = 0"], lead: "Líneas 4–5.", note: "Poner C en ceros recorre k + 1 casillas: Θ(k), aunque casi todas queden vacías." },
    { lines: ["  para j = 0 hasta n-1", "    C[A[j]] = C[A[j]] + 1"], lead: "Líneas 6–7.", note: "El histograma: una lectura por elemento, Θ(n). Aquí C[v] cuenta cuántos v hay." },
    { lines: ["  para i = 1 hasta k", "    C[i] = C[i] + C[i-1]"], lead: "Líneas 8–9.", note: "Suma acumulada. Al salir, C[i] es la cantidad de elementos ≤ i: la posición final del último i, más uno." },
    { lines: ["  para j = n-1 descendiendo hasta 0", "    C[A[j]] = C[A[j]] - 1", "    B[C[A[j]]] = A[j]"], lead: "Líneas 10–12.", note: "Se resta antes de escribir porque B empieza en 0: si hay cinco elementos ≤ 20, el último 20 va en B[4], no en B[5]. Ir de derecha a izquierda es lo que da estabilidad." },
    { lines: ["  devolver B"], lead: "Línea 13.", note: "Costo Θ(n + k), sin mejor ni peor caso: mismo trabajo con datos ordenados, invertidos o al azar." },
  ], 1.78);
}

// =============== 6. COUNTING: TRAZA ===============
{
  const s = slide("Counting sort", "§ 02 · COUNTING SORT", "Counting sort sobre la entrada",
    "Esta es la traza real, ejecutada en Python. Arriba, la entrada A. En la tabla solo mostramos los siete índices de C que se ocupan: 1, 19, 20, 40, 50, 99 y 100. Primera fila de datos: cada uno cuenta uno. Segunda: el acumulado va de 1 a 7, porque no hay repetidos. Tercera: la posición final en B, que sale de restar uno al acumulado. A la derecha está el orden de colocación, de derecha a izquierda: primero el 99 va a B[5], luego el 40 a B[3], el 19 a B[1], el 20 a B[2], el 100 a B[6], el 1 a B[0] y por último el 50 a B[4]. Abajo, la salida ordenada. Ahora fíjense en la nota: las otras 94 casillas de C valen cero y se recorren igual, dos veces: al inicializar y al acumular. Ese es el costo theta de k, y es el motivo por el que esta entrada, con k igual a 100, no es un buen caso para Counting sort.");
  const occ = Object.keys(T.occ).map(Number), acc = occ.map((i) => T.acc[i]);
  const X0 = 2.45, CW = 0.95;
  txt(s, "ENTRADA A", { x: 0.6, y: 2.0, w: 1.8, h: 0.4, fontFace: M, fontSize: 11, color: MUTE, valign: "middle" });
  strip(s, T.A, X0, 2.0, CW, 0.48, { fontSize: 15 });
  const rows = [["índice i", occ, { fill: { color: TINT }, bold: true }], ["C[i] contado", occ.map((i) => T.occ[i]), {}], ["C[i] acumulado", acc, { color: GREEN, bold: true }], ["va en B[…]", acc.map((a) => a - 1), { color: RED, bold: true }]];
  rows.forEach(([lab, vals, o], r) => {
    const y = 2.85 + r * 0.5;
    txt(s, lab, { x: 0.6, y, w: 1.8, h: 0.5, fontFace: M, fontSize: 11, color: MUTE, valign: "middle" });
    strip(s, vals, X0, y, CW, 0.5, Object.assign({ fontSize: 15 }, o));
  });
  const outY = 5.2;
  txt(s, "SALIDA B", { x: 0.6, y: outY, w: 1.8, h: 0.48, fontFace: M, fontSize: 11, color: MUTE, valign: "middle" });
  strip(s, T.bk.length ? [1, 19, 20, 40, 50, 99, 100] : [], X0, outY, CW, 0.48, { fontSize: 15, color: GREEN, bold: true });
  [0, 1, 2, 3, 4, 5, 6].forEach((i) => txt(s, "B[" + i + "]", { x: X0 + i * CW, y: outY + 0.52, w: CW, h: 0.22, fontFace: M, fontSize: 10, color: MUTE, align: "center" }));
  txt(s, "Solo se muestran los siete índices ocupados: las otras 94 casillas de C[0..100] valen cero y se recorren igual, dos veces. Ese es el costo Θ(k).", { x: 0.6, y: 6.2, w: 8.5, h: 0.65, fontSize: 13, color: INK });
  vhair(s, 9.45, 1.95, 4.85);
  txt(s, "COLOCAR, DE DERECHA A IZQUIERDA", { x: 9.75, y: 1.95, w: 3.0, h: 0.4, fontFace: M, fontSize: 11, color: RED, charSpacing: 1 });
  T.place.forEach(([j, x, pos], i) => {
    const y = 2.45 + i * 0.6; hair(s, 9.75, y, 2.98);
    txt(s, [{ text: `A[${j}]=${x}`, options: { bold: true, color: INK, breakLine: true } }, { text: `C[${x}]=${pos} → B[${pos}]`, options: { color: GREEN } }], { x: 9.75, y: y + 0.06, w: 2.98, h: 0.5, fontFace: M, fontSize: 12 });
  });
}

// =============== 7. RADIX: CÓMO FUNCIONA ===============
{
  const s = slide("Radix sort", "§ 03 · RADIX SORT", "Radix sort: cómo funciona",
    "Radix sort no mira el valor completo de la clave sino un dígito a la vez, y delega el trabajo pesado a Counting sort. La versión LSD empieza por el dígito menos significativo. Si el mayor número tiene d dígitos hay exactamente d pasadas: en nuestro caso el máximo es 100, así que d es 3. Cada pasada es un Counting sort, pero sobre un alfabeto diminuto: en base 10 solo hay diez dígitos, así que el arreglo de conteo tiene diez casillas aunque los números tengan nueve cifras. Aquí aparece la ventaja sobre Counting sort puro. La estabilidad es la pieza clave: al ordenar por un dígito, los empates conservan el orden anterior, que ya era correcto para los dígitos menos significativos. En la pasada 2 de nuestro ejemplo, 100 y 1 empatan y salen en el orden que traían; la pasada 3 los separa. Si una pasada fuera inestable, se pierde el trabajo de las anteriores.");
  twoCol(s, "No mira la clave completa sino un dígito a la vez, y delega el trabajo pesado en Counting sort, aplicado una vez por cada dígito.", [
    ["Una pasada por dígito.", "La versión LSD empieza por el dígito menos significativo, las unidades, y avanza hacia el más significativo. Si el mayor valor tiene d dígitos hay exactamente d pasadas, ni una más. Con máximo 100 en base 10, d = 3: unidades, decenas y centenas."],
    ["Cada pasada es un Counting sort.", "Pero sobre un alfabeto diminuto: en base b solo existen b dígitos, así que k = b − 1. En base 10 el arreglo de conteo tiene diez casillas aunque los números tengan nueve cifras. Por eso Radix sí ordena enteros enormes, donde Counting sort puro pediría miles de millones de casillas."],
  ], [
    ["La estabilidad es la pieza clave.", "Al ordenar por el dígito i, los elementos que empatan conservan el orden que traían, y ese orden ya era correcto respecto de los dígitos menos significativos. Cada pasada refina la anterior sin destruirla. En la pasada 2, 100 y 1 caen en la cubeta 0 y salen en ese orden; la pasada 3 los separa."],
    ["Extraer el dígito es gratis.", "El dígito i de x es ⌊x / b^(i−1)⌋ mod b: una división y un módulo, O(1). Si una sola pasada fuera inestable, el trabajo de las anteriores se perdería y el resultado final sería incorrecto."],
  ]);
}

// =============== 8. RADIX: PSEUDOCÓDIGO ===============
{
  const s = slide("Radix sort", "§ 03 · RADIX SORT", "Radix sort: pseudocódigo",
    "El pseudocódigo de Radix sort es cortísimo porque toda la maquinaria está en el Counting sort estable. La línea 2 recorre los d dígitos, no los n elementos: con máximo 100 en base 10, d es 3. Las líneas 3 y 4 llaman a Counting sort con una sola diferencia respecto de la lámina anterior: cuenta el dígito de A[j], no el valor de A[j]. Como k es b menos 1, el arreglo de conteo tiene b casillas y no crece aunque los números sean enormes. La línea 7 extrae el dígito: la división entera corre el dígito que interesa hasta las unidades y el módulo lo recorta. Y el Counting sort tiene que ser estable, si no la pasada i destruye lo que logró la i menos 1. El costo es theta de d por n más b. Con nuestra entrada: 3 por 17, o sea 51. Si d es constante, como en un entero de 32 bits, es lineal de verdad.");
  const y = pseudo(s, [
    { lines: ["RADIX-SORT(A, n, d, b)"], lead: "Firma.", note: "d es el número de dígitos del valor máximo, no la cantidad de elementos; b es la base." },
    { lines: ["  para i = 1 hasta d", "    COUNTING-SORT-ESTABLE(A, n, b-1, i)"], lead: "Líneas 2–3.", note: "Una pasada por dígito, del menos al más significativo. Cada una es un Counting sort con k = b − 1: el conteo tiene b casillas y no crece con los valores." },
    { lines: ["  devolver A"], lead: "Línea 4.", note: "Tras la última pasada, A está ordenado por todos los dígitos." },
    { lines: ["digito(x, i, b)", "  devolver piso(x / b^(i-1)) mod b"], lead: "Líneas 5–6.", note: "La división entera lleva el dígito i a las unidades y el módulo lo recorta: dos operaciones, O(1), sin recorrer nada." },
  ], 1.85);
  txt(s, [{ text: "Variante estable. ", options: { bold: true, color: RED } }, { text: "COUNTING-SORT-ESTABLE es el de la lámina 5, pero cuenta C[digito(A[j], i, b)] en vez de C[A[j]]. Con un Counting sort no estable, la pasada i destruye el orden que logró la i − 1.", options: {} }], { x: 0.6, y: y + 0.25, w: 7.2, h: 1.0, fontSize: 14 });
  txt(s, "Θ(d · (n + b))", { x: 8.3, y: y + 0.2, w: 4.4, h: 0.6, fontFace: H, fontSize: 34, color: RED });
  txt(s, "Con d = 3 y b = 10: 3 · (7 + 10) = 51. Si d es constante, como en un entero de 32 bits, es Θ(n) de verdad.", { x: 8.3, y: y + 0.82, w: 4.4, h: 0.8, fontSize: 13, color: INK });
}

// =============== 9. RADIX: TRAZA ===============
{
  const s = slide("Radix sort", "§ 03 · RADIX SORT", "Radix sort: tres pasadas en base 10",
    "Aquí está la traza real de las tres pasadas, ejecutada en Python. Pasada 1, por unidades: 50, 100, 20 y 40 terminan en 0 y caen juntos en la cubeta 0, en ese orden; el 1 va a la cubeta 1; el 19 y el 99 a la cubeta 9. Al leer las cubetas de izquierda a derecha el arreglo queda 50, 100, 20, 40, 1, 19, 99. Pasada 2, por decenas: el 100 y el 1 tienen dígito de decenas 0 y caen en la cubeta 0, y salen en ese orden porque así venían de la pasada anterior. Quedan 100, 1, 19, 20, 40, 50, 99. Pasada 3, por centenas: todos menos el 100 tienen centena 0 y el 100 va a la cubeta 1. Esta pasada separa al 100 del 1 y deja todo ordenado. Resaltado en rojo está el par que muestra la estabilidad: el 1 y el 100 empatan en la pasada 2, y el orden de entrada se respeta.");
  const names = ["unidades", "decenas", "centenas"];
  let y = 1.85;
  T.passes.forEach((p, pi) => {
    const maxN = Math.max(...Object.values(p.buckets).map((b) => b.length));
    const hh = maxN * 0.2 + 0.16, X0 = 2.55, BW = 0.66;
    txt(s, [{ text: "Pasada " + p.p, options: { fontFace: H, fontSize: 22, bold: true, breakLine: true } }, { text: names[pi], options: { fontFace: M, fontSize: 11, color: MUTE } }], { x: 0.6, y: y + 0.24, w: 1.8, h: hh, valign: "middle" });
    for (let d = 0; d < 10; d++) {
      txt(s, String(d), { x: X0 + d * BW, y, w: BW, h: 0.22, fontFace: M, fontSize: 10, color: MUTE, align: "center" });
      const items = p.buckets[d] || [];
      const runs = items.length ? items.map((v, i) => ({ text: String(v), options: { breakLine: i < items.length - 1, color: (pi === 1 && d === 0) ? RED : INK, bold: pi === 1 && d === 0 } })) : [{ text: "·", options: { color: RULE } }];
      s.addText(runs, { shape: pres.ShapeType.rect, isTextBox: true, x: X0 + d * BW, y: y + 0.24, w: BW, h: hh, fill: { color: items.length ? CREAM : TINT }, line: { color: INK, width: 0.5 }, fontFace: M, fontSize: 11, align: "center", valign: "top", margin: [0.06, 0, 0, 0], lineSpacing: 14.4 });
    }
    s.addShape(pres.ShapeType.line, { x: 9.3, y: y + 0.24 + hh / 2, w: 0.35, h: 0, line: { color: INK, width: 1, endArrowType: "triangle" } });
    txt(s, pi === 2 ? "ORDENADO" : "QUEDA", { x: 9.8, y, w: 2.9, h: 0.22, fontFace: M, fontSize: 10, color: pi === 2 ? GREEN : MUTE });
    strip(s, p.out, 9.8, y + 0.24 + hh / 2 - 0.24, 0.42, 0.48, { fontSize: 11, color: pi === 2 ? GREEN : INK, bold: pi === 2 });
    y += 0.24 + hh + 0.3;
  });
  hair(s, 0.6, y - 0.1, 12.13, RULE);
  txt(s, [{ text: "Fíjense en la pasada 2: ", options: { bold: true, color: RED } }, { text: "100 y 1 caen juntos en la cubeta 0 y salen en ese orden porque así estaban. La pasada 3 los separa. Esa es la estabilidad trabajando.", options: {} }], { x: 0.6, y: y + 0.02, w: 12.13, h: 0.5, fontSize: 13 });
  console.log("radix bottom", (y + 0.5).toFixed(2));
}

// =============== 10. BUCKET: CÓMO FUNCIONA ===============
{
  const s = slide("Bucket sort", "§ 04 · BUCKET SORT", "Bucket sort: cómo funciona",
    "Bucket sort es el híbrido del grupo: reparte sin comparar y después ordena comparando, pero solo dentro de grupos que espera encontrar diminutos. Primero repartimos: elegimos m cubetas, lo habitual es m igual a n, y cada elemento va a la cubeta que le toca según su valor. Con nuestra entrada, m es 7, el mínimo es 1 y el máximo 100: el 50 va a la cubeta 3 y el 100 a la 6. Segundo, ordenamos cada cubeta con inserción, que es lentísima en general pero inmejorable con uno o dos elementos. Solo la cubeta 1 y la cubeta 6 necesitaron ordenarse, y las dos tenían dos elementos. Tercero, concatenamos de la cubeta 0 a la m menos 1. Todo depende de que el reparto sea parejo: si casi todo cae en una cubeta, la inserción interna recibe n elementos y el costo es theta de n al cuadrado. Es el único de los tres cuyo peor caso es peor que n log n.");
  twoCol(s, "Es el híbrido del grupo: reparte sin comparar y después ordena comparando, pero solo dentro de grupos que espera encontrar diminutos.", [
    ["Repartir.", "Se eligen m cubetas (lo habitual es m = n) y cada elemento va a la que le corresponde según su valor, mediante una cuenta: j = ⌊m · (x − min) / (max − min + 1)⌋. Ninguna comparación; Θ(n). Con nuestra entrada m = 7, min = 1 y max = 100: el 50 va a B[3] y el 100 a B[6]."],
    ["Ordenar cada cubeta.", "Con inserción, que es lentísima en general pero inmejorable en listas de uno o dos elementos. Si el reparto salió parejo, cada cubeta tiene en promedio un elemento y ordenarla casi no cuesta. Aquí solo B[1] = [20, 19] y B[6] = [100, 99] necesitaron un intercambio."],
  ], [
    ["Concatenar.", "De la cubeta 0 a la m − 1. Cada cubeta ya está ordenada y todas las claves de la cubeta j son menores que las de la j + 1, así que el resultado sale ordenado: Θ(n + m)."],
    ["El supuesto que lo sostiene.", "Todo depende de que los datos se repartan parejo. Si la distribución está sesgada y casi todo cae en una cubeta, la inserción interna recibe un arreglo de tamaño n y el costo es Θ(n²). Es el único de los tres cuyo peor caso es peor que n log n: sin garantía de uniformidad, mergesort es la apuesta segura."],
  ]);
}

// =============== 11. BUCKET: PSEUDOCÓDIGO ===============
{
  const s = slide("Bucket sort", "§ 04 · BUCKET SORT", "Bucket sort: pseudocódigo",
    "Leamos Bucket sort por bloques. La línea 2 busca el mínimo y el máximo: hace falta el rango real. Si las claves fueran reales en el intervalo cero a uno, como en el libro de Cormen, este paso se omite. La línea 5 es el reparto: el más uno en el denominador evita que el máximo dé índice n, que quedaría fuera de B. Es un error de índice clásico. Con el 100 de nuestra entrada, sin el más uno daría 7 por 99 sobre 99, o sea 7, fuera del arreglo. La línea 8 usa inserción y no quicksort porque en cubetas de uno o dos elementos es más rápida y además mantiene la estabilidad. La línea 9 concatena en orden de cubeta. El costo promedio es theta de n y el peor caso theta de n al cuadrado, cuando todo cae en una cubeta.");
  pseudo(s, [
    { lines: ["BUCKET-SORT(A, n)", "  min = minimo(A); max = maximo(A)"], lead: "Línea 2.", note: "Hace falta el rango real. Con claves reales en [0, 1), como en Cormen, este paso se omite y el índice se reduce a ⌊n · A[i]⌋." },
    { lines: ["  sea B[0..n-1] listas vacías"], lead: "Línea 3.", note: "m = n cubetas: así se espera un elemento por cubeta y el trabajo interno es casi nulo." },
    { lines: ["  para i = 0 hasta n-1", "    j = piso(n·(A[i]-min) / (max-min+1))", "    añadir A[i] al final de B[j]"], lead: "Líneas 4–6.", note: "El reparto, sin comparar: Θ(n). El +1 del denominador evita que el máximo dé j = n, fuera de B (con el 100: 7·99/100 = 6, no 7). Añadir al final conserva el orden de llegada." },
    { lines: ["  para j = 0 hasta n-1", "    INSERTION-SORT(B[j])"], lead: "Líneas 7–8.", note: "Inserción y no quicksort: en listas de uno o dos elementos es más rápida, y es estable. Si todo cae en una sola cubeta, aquí está el Θ(n²)." },
    { lines: ["  devolver B[0] + B[1] + ... + B[n-1]"], lead: "Línea 9.", note: "Concatenar respeta el orden de las cubetas, que por construcción ya está ordenado por rango. Promedio Θ(n), peor caso Θ(n²)." },
  ], 1.85);
}

// =============== 12. BUCKET: TRAZA ===============
{
  const s = slide("Bucket sort", "§ 04 · BUCKET SORT", "Bucket sort: siete cubetas para siete números",
    "Esta es la traza real. Arriba, cada elemento con el índice de cubeta que le calcula la fórmula: 50 va a B[3], 1 a B[0], 100 a B[6], 20 y 19 a B[1], 40 a B[2], y 99 a B[6]. Cada cubeta cubre un rango de valores de unos catorce números: B[0] de 1 a 15, B[1] de 16 a 29, y así. Después del reparto, B[1] contiene 20 y 19, en ese orden, y B[6] contiene 100 y 99. Esas dos, resaltadas en rojo, son las únicas que necesitaron inserción, y las dos con un solo intercambio. B[4] y B[5] quedaron vacías: con n igual a 7 el reparto no alcanza a ser perfectamente uniforme. Abajo, la concatenación da la salida ordenada.");
  const X0 = 2.45, CW = 0.95;
  txt(s, "ENTRADA A", { x: 0.6, y: 1.85, w: 1.8, h: 0.45, fontFace: M, fontSize: 11, color: MUTE, valign: "middle" });
  strip(s, T.A, X0, 1.85, CW, 0.45, { fontSize: 15 });
  txt(s, "ÍNDICE j", { x: 0.6, y: 2.33, w: 1.8, h: 0.3, fontFace: M, fontSize: 11, color: MUTE, valign: "middle" });
  T.idx.forEach(([x, j], i) => txt(s, "→ B[" + j + "]", { x: X0 + i * CW, y: 2.33, w: CW, h: 0.3, fontFace: M, fontSize: 11, color: RED, align: "center", valign: "middle" }));
  txt(s, [{ text: "j = ⌊7 · (x − 1) / 100⌋", options: { fontFace: M, fontSize: 13, bold: true, breakLine: true } }, { text: "m = 7, min = 1, max = 100", options: { fontFace: M, fontSize: 11, color: MUTE } }], { x: 9.35, y: 1.85, w: 3.4, h: 0.8 });
  const LX = 0.6, LW = 1.3, CWB = 1.54, y0 = 3.0;
  const sortedFlag = T.bk.map((b) => b.length > 1);
  const rowLab = [["cubeta", 0.55], ["repartido", 0.65], ["ordenado", 0.65]];
  let yy = y0;
  rowLab.forEach(([l, h]) => { txt(s, l, { x: LX, y: yy, w: LW, h, fontFace: M, fontSize: 11, color: MUTE, valign: "middle" }); yy += h; });
  T.bk.forEach((b, j) => {
    const x = LX + LW + j * CWB;
    s.addText([{ text: `B[${j}]`, options: { bold: true, breakLine: true, fontSize: 13 } }, { text: `${T.rng[j][0]} – ${T.rng[j][1]}`, options: { fontSize: 11, color: MUTE } }],
      { shape: pres.ShapeType.rect, isTextBox: true, x, y: y0, w: CWB, h: 0.55, fill: { color: TINT }, line: { color: INK, width: 0.75 }, fontFace: M, align: "center", valign: "middle", margin: 0 });
    const hot = sortedFlag[j];
    cell(s, b.length ? b.join("  ") : "vacía", x, y0 + 0.55, CWB, 0.65, { fontSize: 15, color: hot ? RED : (b.length ? INK : MUTE), bold: hot, italic: !b.length, fontFace: b.length ? M : B, fill: { color: b.length ? CREAM : TINT } });
    const sb = [...b].sort((p, q) => p - q);
    cell(s, sb.length ? sb.join("  ") : "—", x, y0 + 1.2, CWB, 0.65, { fontSize: 15, color: b.length ? GREEN : MUTE, bold: !!b.length, fill: { color: b.length ? CREAM : TINT } });
  });
  const flat = T.bk.map((b) => [...b].sort((p, q) => p - q)).flat();
  txt(s, "CONCATENADO", { x: LX, y: 5.2, w: LW, h: 0.5, fontFace: M, fontSize: 11, color: MUTE, valign: "middle" });
  strip(s, flat, LX + LW, 5.2, CW, 0.5, { fontSize: 15, color: GREEN, bold: true });
  txt(s, [{ text: "Solo B[1] y B[6] (en rojo) necesitaron ordenamiento interno, y ambas con dos elementos. ", options: { bold: true } }, { text: "B[4] y B[5] quedaron vacías: con n = 7 el reparto no alcanza a ser perfectamente uniforme.", options: {} }], { x: 0.6, y: 6.05, w: 12.13, h: 0.75, fontSize: 14 });
}

// =============== 13. COSTOS ===============
{
  const s = slide("Balance", "§ 05 · BALANCE", "Costos en tiempo de ejecución",
    "Esta es la tabla de costos. Counting sort es theta de n más k en los tres casos: no hay mejor ni peor caso, porque nunca miramos la relación entre dos elementos. Radix sort es theta de d por n más b, también sin casos. Bucket sort es el único con diferencia: theta de n en promedio, pero theta de n al cuadrado en el peor caso. La última columna aplica las fórmulas a nuestra entrada, en unidades de trabajo sin constantes, es decir órdenes de magnitud, no ciclos de CPU: Counting sort, n más k más 1, es 108; Radix, 3 por 17, es 51; Bucket, n más m, es 14. Abajo, qué significan los símbolos. k es el rango: con n igual a mil y k igual a un millón, Counting sort es unas cien veces peor que mergesort. d y b son dígitos y base. m es el número de cubetas.");
  const mono = (t, c) => ({ t, o: c ? { color: c, bold: true } : {} });
  table(s, [
    ["Algoritmo", "Mejor", "Promedio", "Peor", "Memoria", "Estable", "Con n = 7"],
    ["Counting sort", "Θ(n+k)", "Θ(n+k)", "Θ(n+k)", "Θ(n+k)", "Sí", mono("n+k+1 = 108", RED)],
    ["Radix sort (LSD)", "Θ(d(n+b))", "Θ(d(n+b))", "Θ(d(n+b))", "Θ(n+b)", "Sí", mono("d(n+b) = 51", INK)],
    ["Bucket sort", "Θ(n)", "Θ(n)", mono("Θ(n²)", RED), "Θ(n+m)", "Sí", mono("n+m = 14", GREEN)],
  ], { x: 0.6, y: 1.85, w: 12.13, colW: [2.1, 1.6, 1.7, 1.7, 1.4, 1.25, 2.38], rowH: [0.5, 0.55, 0.55, 0.55], mono: [1, 2, 3, 4, 6] });
  txt(s, "Counting y Radix no tienen mejor ni peor caso: hacen el mismo trabajo con datos ordenados, invertidos o al azar, porque nunca miran la relación entre dos elementos. Columna «Con n = 7»: unidades de trabajo sin constantes (k = 100, d = 3, b = 10, m = 7).", { x: 0.6, y: 4.3, w: 12.13, h: 0.65, fontSize: 13, color: MUTE });
  const defs = [["k · el rango.", "Valor máximo de la clave. Counting sort es lineal solo si k = O(n); con n = 1 000 y k = 10⁶ es unas cien veces peor que mergesort."], ["d y b · dígitos y base.", "d = ⌈log_b(k+1)⌉. Para un tipo de dato fijo d es constante, y ahí Radix sort sí es Θ(n)."], ["m · las cubetas.", "Se toma m = n para que cada cubeta tenga en promedio un elemento. El Θ(n²) aparece cuando la distribución está sesgada."]];
  defs.forEach(([a, b], i) => {
    const y = 5.05 + i * 0.62; hair(s, 0.6, y, 12.13);
    txt(s, [{ text: a + " ", options: { bold: true, color: RED } }, { text: b, options: {} }], { x: 0.6, y: y + 0.06, w: 12.13, h: 0.52, fontSize: 13 });
  });
}

// =============== 14. FRENTE A LOS TRADICIONALES ===============
{
  const s = slide("Balance", "§ 05 · BALANCE", "Frente a los algoritmos tradicionales",
    "Comparemos con los clásicos. Burbuja e inserción son theta de n al cuadrado: simples y buenos solo para n pequeño. Quicksort es n log n en promedio pero n cuadrado en el peor caso. Mergesort y heapsort garantizan n log n; mergesort es estable pero usa memoria lineal, heapsort es in-place. Los tres de hoy no comparan. Counting y Radix son lineales bajo sus condiciones, y Bucket es lineal solo con reparto uniforme. Donde sí se nota la diferencia: un millón de enteros de 32 bits. Mergesort hace unos 20 millones de comparaciones, porque n log n con n igual a un millón da casi 20 millones. Radix sort en base 256 hace 4 pasadas de un millón de lecturas cada una y no compara nada. Counting sort sería imposible: necesitaría 2 a la 32 casillas, unos 4 mil millones.");
  const cmpC = (t) => ({ t, o: { color: t === "Sí" ? INK : GREEN, bold: t !== "Sí" } });
  table(s, [
    ["Algoritmo", "Peor caso", "Promedio", "Memoria", "¿Compara?", "Nota"],
    ["Burbuja, inserción", "Θ(n²)", "Θ(n²)", "O(1)", cmpC("Sí"), "simple, para n pequeño"],
    ["Quicksort", { t: "Θ(n²)", o: { color: RED, bold: true } }, "Θ(n log n)", "O(log n)", cmpC("Sí"), "rápido en la práctica"],
    ["Mergesort", "Θ(n log n)", "Θ(n log n)", "Θ(n)", cmpC("Sí"), "estable y óptimo"],
    ["Heapsort", "Θ(n log n)", "Θ(n log n)", "O(1)", cmpC("Sí"), "in-place garantizado"],
    ["Counting sort", "Θ(n+k)", "Θ(n+k)", "Θ(n+k)", cmpC("No"), "lineal si k = O(n)"],
    ["Radix sort", "Θ(d(n+b))", "Θ(d(n+b))", "Θ(n+b)", cmpC("No"), "lineal si d es fijo"],
    ["Bucket sort", { t: "Θ(n²)", o: { color: RED, bold: true } }, "Θ(n)", "Θ(n+m)", cmpC("Mixto"), "exige uniformidad"],
  ], { x: 0.6, y: 1.8, w: 12.13, colW: [2.4, 1.7, 1.7, 1.4, 1.7, 3.23], rowH: [0.45, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42], mono: [1, 2, 3] });
  hair(s, 0.6, 5.5, 12.13, INK, 0.75);
  txt(s, "DONDE SÍ SE NOTA: UN MILLÓN DE ENTEROS DE 32 BITS", { x: 0.6, y: 5.62, w: 12.13, h: 0.25, fontFace: M, fontSize: 11, color: RED, charSpacing: 1 });
  txt(s, [{ text: "Mergesort hace unos ", options: {} }, { text: "20 millones", options: { bold: true, color: RED } }, { text: " de comparaciones. Radix sort en base 256 hace ", options: {} }, { text: "4 pasadas", options: { bold: true, color: GREEN } }, { text: " de un millón de lecturas, sin comparar nada. Counting sort sería imposible: necesitaría 2³² ≈ 4 mil millones de casillas.", options: {} }], { x: 0.6, y: 5.95, w: 12.13, h: 0.8, fontSize: 15 });
}

// =============== 15. n², n log n, n  ===============
{
  const s = slide("Balance", "§ 05 · BALANCE", "n², n log n y n: lineal no significa rápido",
    "Esta lámina tiene dos mitades. A la izquierda, la curva clásica: n al cuadrado se dispara, n log n crece despacio y n es una recta. Para n grande, los tres algoritmos de hoy ganan. Pero a la derecha está el punto que normalmente se omite, y es con nuestra entrada: n igual a 7 y k igual a 100. Counting sort tiene que tocar las 101 casillas de C y los 7 elementos: unas 108 operaciones. Mergesort necesita unas 20, es decir n log n; cuando lo ejecuté sobre estos datos hizo 13 comparaciones reales. Hasta burbuja, con n(n−1)/2 igual a 21 comparaciones, le gana a Counting sort en esta entrada. Radix sort con 3 pasadas hace unas 51, y Bucket sort unas 14. Aclaro: son órdenes de magnitud sin constantes, no ciclos de CPU, pero la conclusión no cambia. Lineal no es sinónimo de rápido: es lineal en n más k, y si k es mucho mayor que n, esa k manda.");
  const ns = Array.from({ length: 20 }, (_, i) => i + 1);
  const fo = { fontFace: "+mn-lt", fontSize: 11 };
  s.addChart(pres.charts.LINE, [
    { name: "n² (burbuja, inserción)", labels: ns, values: ns.map((n) => n * n) },
    { name: "n log₂ n (mergesort, heapsort)", labels: ns, values: ns.map((n) => +(n * Math.log2(n)).toFixed(1)) },
    { name: "n (el objetivo de los tres)", labels: ns, values: ns },
  ], { x: 0.6, y: 2.05, w: 6.3, h: 3.3, chartColors: [RED, GREEN, INK], lineSize: 2.25, lineDataSymbol: "none",
    showLegend: true, legendPos: "b", legendFontFace: "+mn-lt", legendFontSize: 11, legendColor: INK,
    catAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: 10, catAxisLabelColor: MUTE, valAxisLabelFontFace: "+mn-lt", valAxisLabelFontSize: 10, valAxisLabelColor: MUTE,
    valGridLine: { color: "D9D2C2", size: 0.5 }, catGridLine: { style: "none" }, valAxisMaxVal: 400, valAxisMajorUnit: 100 });
  txt(s, "OPERACIONES RELATIVAS SEGÚN n", { x: 0.6, y: 1.8, w: 6.3, h: 0.22, fontFace: M, fontSize: 11, color: MUTE, charSpacing: 1 });
  txt(s, "CON NUESTRA ENTRADA: n = 7, k = 100", { x: 7.3, y: 1.8, w: 5.4, h: 0.22, fontFace: M, fontSize: 11, color: RED, charSpacing: 1 });
  s.addChart(pres.charts.BAR, [{ name: "Operaciones", labels: ["Counting sort", "Radix sort", "Burbuja", "Mergesort", "Bucket sort"], values: [108, 51, 21, 20, 14] }],
    { x: 7.3, y: 2.05, w: 5.43, h: 3.3, barDir: "bar", chartColors: [RED, "9A9184", "9A9184", GREEN, "9A9184"], showLegend: false, showValue: true, dataLabelPosition: "outEnd", dataLabelFontFace: "+mn-lt", dataLabelFontSize: 11, dataLabelColor: INK,
      catAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: 11, catAxisLabelColor: INK, catAxisOrientation: "maxMin", valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, barGapWidthPct: 45 });
  hair(s, 0.6, 5.5, 12.13, INK, 0.75);
  txt(s, [{ text: "n = 7 y k = 100: ", options: { bold: true, color: RED } }, { text: "Counting sort ejecuta unas ", options: {} }, { text: "108 operaciones", options: { bold: true, color: RED } }, { text: " (101 casillas de C + 7 elementos); mergesort, unas ", options: {} }, { text: "20", options: { bold: true, color: GREEN } }, { text: " (n log₂ n; medidas: 13 comparaciones). Lineal no significa rápido: es lineal en n + k.", options: {} }], { x: 0.6, y: 5.62, w: 12.13, h: 1.25, fontFace: H, fontSize: 21 });
}

// =============== 16. CUÁNDO CONVIENE ===============
{
  const s = slide("Balance", "§ 05 · BALANCE", "Cuándo conviene cada uno",
    "Para cerrar, una guía práctica. Counting sort conviene cuando el rango es chico y conocido: edades, notas del 0 al 20, días del año, categorías. Evítenlo cuando k es mucho mayor que n: con nuestros siete números y k igual a 100 ya pierde. Radix sort conviene cuando hay muchos datos de longitud fija: enteros, DNI, códigos, fechas, placas, cadenas del mismo largo. Es el de uso práctico más amplio de los tres. Bucket sort conviene cuando se sabe que los datos se reparten parejo, típicamente decimales generados al azar. Sin esa garantía, mergesort es la apuesta segura. La idea final: estos tres no son mejores que mergesort o quicksort. Cambian tiempo por memoria y por supuestos sobre los datos, y solo ganan cuando esos supuestos se cumplen. Gracias; quedo atento a preguntas.");
  txt(s, "CONVIENE CUANDO", { x: 3.6, y: 1.85, w: 5, h: 0.25, fontFace: M, fontSize: 11, color: GREEN, charSpacing: 1 });
  txt(s, "EVÍTELO SI", { x: 9.0, y: 1.85, w: 3.7, h: 0.25, fontFace: M, fontSize: 11, color: RED, charSpacing: 1 });
  const rows = [
    ["Counting sort", GREEN, "El rango es chico y conocido: edades, notas del 0 al 20, días del año, categorías.", "k ≫ n. Con n = 7 y k = 100 ya pierde."],
    ["Radix sort", GREEN, "Hay muchos datos de longitud fija: enteros, DNI, códigos, fechas, placas, cadenas del mismo largo. Es el de uso práctico más amplio.", "Las claves son de largo muy variable, o d crece con n."],
    ["Bucket sort", GREEN, "Se sabe que los datos se reparten parejo, típicamente decimales generados al azar.", "La distribución está sesgada: Θ(n²). Sin garantía, mergesort."],
  ];
  rows.forEach(([name, c, a, b], i) => {
    const y = 2.2 + i * 1.3; hair(s, 0.6, y, 12.13, i === 0 ? INK : RULE, i === 0 ? 0.75 : 0.5);
    txt(s, name, { x: 0.6, y: y + 0.18, w: 2.8, h: 0.6, fontFace: H, fontSize: 28, color: INK });
    txt(s, a, { x: 3.6, y: y + 0.2, w: 5.0, h: 1.0, fontSize: 14 });
    txt(s, b, { x: 9.0, y: y + 0.2, w: 3.7, h: 1.0, fontSize: 14, color: INK });
  });
  hair(s, 0.6, 6.1, 12.13, INK, 0.75);
  txt(s, "Estos tres no son «mejores» que mergesort o quicksort: cambian tiempo por memoria y por supuestos sobre los datos. Solo ganan cuando esos supuestos se cumplen.", { x: 0.6, y: 6.2, w: 12.13, h: 0.7, fontFace: H, fontSize: 20, italic: true });
}

// =============== 17. REFERENCIAS ===============
{
  const s = slide("Cierre", "FUENTES", "Referencias",
    "Estas son las fuentes. Cormen, Leiserson, Rivest y Stein, capítulo 8, es la referencia principal: ahí están el árbol de decisión, la cota inferior y los tres algoritmos con sus demostraciones. Sedgewick y Wayne explican Radix sort sobre cadenas. Knuth, en el volumen 3, trata la ordenación por distribución. Las trazas de las láminas 6, 9 y 12 no son estimadas: las verifiqué ejecutando los tres algoritmos en Python sobre la entrada 50, 1, 100, 20, 19, 40, 99. Las cifras de la lámina 15 son órdenes de magnitud sin constantes. Gracias.");
  const refs = [
    ["1", "Cormen, T. H., Leiserson, C. E., Rivest, R. L., y Stein, C. (2022). ", "Introduction to Algorithms", " (4.ª ed.), cap. 8: «Sorting in Linear Time». MIT Press."],
    ["2", "Sedgewick, R., y Wayne, K. (2011). ", "Algorithms", " (4.ª ed.), sec. 5.1: «String Sorts». Addison-Wesley. Sitio del libro: algs4.cs.princeton.edu"],
    ["3", "Knuth, D. E. (1998). ", "The Art of Computer Programming, vol. 3: Sorting and Searching", " (2.ª ed.), sec. 5.2.5: «Sorting by Distribution». Addison-Wesley."],
  ];
  refs.forEach(([n, a, t, c], i) => {
    const y = 2.0 + i * 1.05; hair(s, 0.6, y, 12.13, i === 0 ? INK : RULE, i === 0 ? 0.75 : 0.5);
    txt(s, n, { x: 0.6, y: y + 0.12, w: 0.6, h: 0.6, fontFace: H, fontSize: 30, color: RED });
    txt(s, [{ text: a, options: {} }, { text: t, options: { italic: true } }, { text: c, options: {} }], { x: 1.4, y: y + 0.2, w: 11.3, h: 0.8, fontSize: 15 });
  });
  hair(s, 0.6, 5.2, 12.13, INK, 0.75);
  txt(s, "Las trazas de las láminas 6, 9 y 12 se obtuvieron ejecutando los tres algoritmos en Python sobre la entrada [50, 1, 100, 20, 19, 40, 99]. Los costos de la lámina 15 son órdenes de magnitud (sin constantes), no tiempos medidos.", { x: 0.6, y: 5.35, w: 12.13, h: 0.8, fontSize: 14, color: MUTE });
}

(async () => {
  await pres.writeFile({ fileName: "../ordenar-sin-comparar.pptx" });
  await applyTheme("../ordenar-sin-comparar.pptx", THEME);
  console.log("ok");
})();
