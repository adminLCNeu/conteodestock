/* ============================================================
   Armado del Excel de conteo (formato prolijo).
   Usa ExcelJS (exceljs.min.js). Genera hasta 4 pestañas:
   Resumen, Stock, Desperdicio (si hubo) y Envíos (si hubo).

   buildStockWorkbook(data, stockItems, wasteItems, envioItems, deps)
     deps = {
       ExcelJS,                 // librería
       categories,              // CATEGORIES (para ordenar)
       priceOf(p),              // precio unitario del producto o null
       priceLabel,              // texto "Precio 1 (...)" / "Precio 2 (BCN)"
       fmtDate(iso),            // fecha larga
       envioGroupLabel(grupo),  // etiqueta del tipo de envío
     }
   ============================================================ */

function buildStockWorkbook(data, stockItems, wasteItems, envioItems, deps) {
  const { ExcelJS, categories, priceOf, priceLabel, fmtDate, envioGroupLabel } = deps;
  envioItems = envioItems || [];

  // ---------- Estilo ----------
  const C = {
    navy: 'FF1F3A5F', navyLight: 'FFE6ECF4',
    red: 'FF9B2F2F', redLight: 'FFF7E4E4',
    teal: 'FF1F6F78', tealLight: 'FFE0F0F1',
    zebra: 'FFF4F7FB', white: 'FFFFFFFF',
    line: 'FFD3DAE6', lineStrong: 'FF8FA3BF',
    text: 'FF1B2430', muted: 'FF6B778A',
    warn: 'FFFFF1CC', warnText: 'FF8A5A00',
    totalFill: 'FFDCE6F2',
  };
  const EUR = '#,##0.00 "€"';
  const FONT = 'Calibri';
  const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const thin = { style: 'thin', color: { argb: C.line } };
  const boxThin = { top: thin, bottom: thin, left: thin, right: thin };

  const modeLabel = data.mode === 'semanal' ? 'Semanal' : 'Mensual';
  const round2 = (n) => Math.round(n * 100) / 100;
  const valueOf = (p) => { const pr = priceOf(p); return pr === null ? 0 : Number(p.qty) * pr; };
  const totalOf = (items) => round2(items.reduce((s, p) => s + valueOf(p), 0));
  const missingOf = (items) => items.filter(p => priceOf(p) === null).length;
  const sortByCategory = (items) => items
    .map((p, i) => ({ p, i }))
    .sort((a, b) => (categories.indexOf(a.p.category) - categories.indexOf(b.p.category)) || (a.i - b.i))
    .map(x => x.p);

  const wb = new ExcelJS.Workbook();
  wb.creator = data.generatedBy || "Lucciano's";
  wb.created = new Date();

  // ---------- Piezas reutilizables ----------
  function setupSheet(name, tabColor, ncols, label, widths) {
    const ws = wb.addWorksheet(name, {
      properties: { tabColor: { argb: tabColor }, defaultRowHeight: 18 },
      views: [{ showGridLines: false }],
    });
    widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

    ws.mergeCells(1, 1, 1, ncols);
    const t = ws.getCell(1, 1);
    t.value = `${data.location}  ·  ${label}`;
    t.font = { name: FONT, size: 17, bold: true, color: { argb: C.white } };
    t.fill = fill(tabColor);
    t.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    ws.getRow(1).height = 32;

    ws.mergeCells(2, 1, 2, ncols);
    const s = ws.getCell(2, 1);
    s.value = `Conteo ${modeLabel}   ·   ${fmtDate(data.finishedAt)}   ·   Generado por ${data.generatedBy || '—'}`;
    s.font = { name: FONT, size: 11, color: { argb: C.text } };
    s.fill = fill(C.navyLight);
    s.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    ws.getRow(2).height = 22;

    ws.mergeCells(3, 1, 3, ncols);
    const v = ws.getCell(3, 1);
    v.value = `Valorizado con ${priceLabel}`;
    v.font = { name: FONT, size: 10, italic: true, color: { argb: C.muted } };
    v.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    ws.getRow(3).height = 18;
    return ws; // las filas 1-3 ya están ocupadas por el encabezado
  }

  function addBannerLine(ws, row, ncols, text, kind) {
    ws.mergeCells(row, 1, row, ncols);
    const c = ws.getCell(row, 1);
    c.value = text;
    const warn = kind === 'warn';
    c.font = { name: FONT, size: 10.5, bold: warn, color: { argb: warn ? C.warnText : C.text } };
    c.fill = fill(warn ? C.warn : C.zebra);
    c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1, wrapText: true };
    ws.getRow(row).height = text.length > 110 ? 34 : 20;
    return row + 1;
  }

  function addHeaderRow(ws, rowNum, headers, color, aligns) {
    const r = ws.getRow(rowNum);
    headers.forEach((h, i) => {
      const c = r.getCell(i + 1);
      c.value = h;
      c.font = { name: FONT, size: 11, bold: true, color: { argb: C.white } };
      c.fill = fill(color);
      c.alignment = { vertical: 'middle', horizontal: aligns[i], wrapText: true, indent: aligns[i] === 'center' ? 0 : 1 };
      c.border = boxThin;
    });
    r.height = 34;
  }

  function setupPrint(ws, ncols, headerRow, landscape) {
    ws.pageSetup = {
      paperSize: 9,
      orientation: landscape ? 'landscape' : 'portrait',
      fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      horizontalCentered: true,
      margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.6, header: 0.3, footer: 0.3 },
    };
    if (headerRow) ws.pageSetup.printTitlesRow = `${headerRow}:${headerRow}`;
    ws.headerFooter.oddFooter = `&L&8Lucciano's · ${String(data.location).replace(/&/g, '&&')}&R&8Página &P de &N`;
  }

  function styleBodyCell(cell, col, fillArgb) {
    cell.font = { name: FONT, size: 11, bold: !!col.bold, color: { argb: col.muted ? C.muted : C.text } };
    cell.alignment = { vertical: 'middle', horizontal: col.al, wrapText: !!col.wrap, indent: col.al === 'center' ? 0 : 1 };
    cell.border = boxThin;
    if (fillArgb) cell.fill = fill(fillArgb);
    if (col.fmt) cell.numFmt = col.fmt;
  }

  function styleTotalRow(ws, rowNum, ncols, color) {
    for (let i = 1; i <= ncols; i++) {
      const c = ws.getCell(rowNum, i);
      c.fill = fill(C.totalFill);
      c.font = { name: FONT, size: 12, bold: true, color: { argb: C.text } };
      c.border = { top: { style: 'medium', color: { argb: color } }, bottom: { style: 'medium', color: { argb: color } } };
    }
    ws.getRow(rowNum).height = 26;
  }

  function markMissingPrice(cell) {
    cell.value = 'sin precio';
    cell.font = { name: FONT, size: 10, italic: true, color: { argb: C.warnText } };
    cell.fill = fill(C.warn);
    cell.numFmt = 'General';
  }

  // ---------- Pestaña de ítems (Stock / Desperdicio) ----------
  function addItemsSheet(name, color, label, items, isWaste) {
    const cols = [
      { h: 'Código', w: 9, al: 'center' },
      { h: 'Producto', w: 40, al: 'left', wrap: true },
      { h: 'Categoría', w: 26, al: 'left' },
      { h: 'Unidad de medida', w: 14, al: 'center' },
      { h: isWaste ? 'Cantidad desperdiciada' : 'Cantidad contada', w: 14, al: 'center', bold: true },
      { h: 'Precio unitario (€)', w: 14, al: 'right', fmt: EUR },
      { h: 'Subtotal (€)', w: 15, al: 'right', fmt: EUR, bold: true },
      { h: 'Sucursal', w: 17, al: 'left', muted: true },
    ];
    if (isWaste) cols.push({ h: 'Motivo', w: 30, al: 'left', wrap: true });
    const n = cols.length;

    const ws = setupSheet(name, color, n, label, cols.map(c => c.w));
    let row = 4;
    const miss = missingOf(items);
    if (miss > 0) row = addBannerLine(ws, row, n, `⚠ ${miss} producto(s) sin precio cargado: no se incluyen en el total.`, 'warn');
    if (data.generalNote && data.generalNote.trim()) row = addBannerLine(ws, row, n, `Nota: ${data.generalNote.trim()}`, 'note');

    const headerRow = row;
    addHeaderRow(ws, headerRow, cols.map(c => c.h), color, cols.map(c => c.al));
    row++;

    const first = row;
    let block = -1, prevCat = null;
    sortByCategory(items).forEach(p => {
      if (p.category !== prevCat) { block++; }
      const newBlock = p.category !== prevCat && prevCat !== null;
      prevCat = p.category;
      const zebra = block % 2 === 1 ? C.zebra : C.white;

      const price = priceOf(p);
      const qty = Number(p.qty);
      const values = [
        p.code, p.name, p.category, p.unit, qty,
        price === null ? null : price,
        price === null ? null : round2(qty * price),
        data.location,
      ];
      if (isWaste) values.push(p.note || '');

      const r = ws.getRow(row);
      values.forEach((val, i) => {
        const c = r.getCell(i + 1);
        c.value = val;
        styleBodyCell(c, cols[i], zebra);
        if (newBlock) c.border = { ...boxThin, top: { style: 'thin', color: { argb: C.lineStrong } } };
      });
      if (price === null) { markMissingPrice(r.getCell(6)); r.getCell(6).border = boxThin; r.getCell(6).alignment = { vertical: 'middle', horizontal: 'right' }; }
      r.height = String(p.name).length > 42 ? 32 : 22;
      row++;
    });
    const last = row - 1;

    // Fila de total
    const totalRow = row;
    ws.getCell(totalRow, 5).value = 'TOTAL';
    ws.getCell(totalRow, 5).alignment = { horizontal: 'right', vertical: 'middle' };
    const sumCell = ws.getCell(totalRow, 7);
    sumCell.value = last >= first
      ? { formula: `SUM(G${first}:G${last})`, result: totalOf(items) }
      : 0;
    sumCell.numFmt = EUR;
    sumCell.alignment = { horizontal: 'right', vertical: 'middle' };
    styleTotalRow(ws, totalRow, n, color);

    // Filas fijas arriba + filtros + impresión
    ws.views = [{ state: 'frozen', ySplit: headerRow, xSplit: 2, showGridLines: false }];
    if (last >= first) ws.autoFilter = { from: { row: headerRow, column: 1 }, to: { row: last, column: n } };
    setupPrint(ws, n, headerRow, true);
    return ws;
  }

  // ---------- Pestaña Envíos ----------
  function addEnviosSheet() {
    const cols = [
      { h: 'Código', w: 9, al: 'center' },
      { h: 'Producto', w: 40, al: 'left', wrap: true },
      { h: 'Categoría', w: 26, al: 'left' },
      { h: 'Tipo', w: 20, al: 'left' },
      { h: 'Unidad de medida', w: 14, al: 'center' },
      { h: 'Cantidad enviada', w: 14, al: 'center', bold: true },
      { h: 'Precio unitario (€)', w: 14, al: 'right', fmt: EUR },
      { h: 'Subtotal (€)', w: 15, al: 'right', fmt: EUR, bold: true },
      { h: 'Origen', w: 17, al: 'left', muted: true },
      { h: 'Destino', w: 17, al: 'left', bold: true },
    ];
    const n = cols.length;
    const ws = setupSheet('Envíos', C.teal, n, 'Envíos a otras sucursales', cols.map(c => c.w));
    let row = 4;
    const miss = missingOf(envioItems);
    if (miss > 0) row = addBannerLine(ws, row, n, `⚠ ${miss} producto(s) sin precio cargado: no se incluyen en el total.`, 'warn');

    const headerRow = row;
    addHeaderRow(ws, headerRow, cols.map(c => c.h), C.teal, cols.map(c => c.al));
    row++;

    const subtotalCells = [];
    const dests = [...new Set(envioItems.map(p => p.destino))];
    dests.forEach(dest => {
      const items = sortByCategory(envioItems.filter(p => p.destino === dest));

      // Banda con el nombre del destino
      ws.mergeCells(row, 1, row, n);
      const g = ws.getCell(row, 1);
      g.value = `→ ${dest}   ·   ${items.length} producto${items.length === 1 ? '' : 's'}`;
      g.font = { name: FONT, size: 12, bold: true, color: { argb: C.teal } };
      g.fill = fill(C.tealLight);
      g.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      ws.getRow(row).height = 24;
      row++;

      const first = row;
      items.forEach((p, idx) => {
        const price = priceOf(p);
        const qty = Number(p.qty);
        const vals = [p.code, p.name, p.category, envioGroupLabel(p.grupo), p.unit, qty,
          price === null ? null : price, price === null ? null : round2(qty * price), data.location, dest];
        const r = ws.getRow(row);
        vals.forEach((val, i) => {
          const c = r.getCell(i + 1);
          c.value = val;
          styleBodyCell(c, cols[i], idx % 2 === 1 ? C.zebra : C.white);
        });
        if (price === null) { markMissingPrice(r.getCell(7)); r.getCell(7).border = boxThin; r.getCell(7).alignment = { vertical: 'middle', horizontal: 'right' }; }
        r.height = String(p.name).length > 42 ? 32 : 22;
        row++;
      });
      const last = row - 1;

      ws.getCell(row, 6).value = `TOTAL → ${dest}`;
      ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };
      const sc = ws.getCell(row, 8);
      sc.value = { formula: `SUM(H${first}:H${last})`, result: totalOf(items) };
      sc.numFmt = EUR;
      sc.alignment = { horizontal: 'right', vertical: 'middle' };
      styleTotalRow(ws, row, n, C.teal);
      subtotalCells.push(`H${row}`);
      row += 2; // espacio entre destinos
    });

    // Total general
    ws.getCell(row, 6).value = 'TOTAL ENVÍOS';
    ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };
    const gc = ws.getCell(row, 8);
    gc.value = { formula: subtotalCells.join('+'), result: totalOf(envioItems) };
    gc.numFmt = EUR;
    gc.alignment = { horizontal: 'right', vertical: 'middle' };
    styleTotalRow(ws, row, n, C.teal);
    ws.getRow(row).height = 30;

    ws.views = [{ state: 'frozen', ySplit: headerRow, xSplit: 2, showGridLines: false }];
    setupPrint(ws, n, headerRow, true);
  }

  // ---------- Pestaña Resumen ----------
  function addResumenSheet() {
    const n = 5;
    const ws = setupSheet('Resumen', C.navy, n, 'Resumen del conteo', [36, 14, 18, 13, 17]);
    let row = 4;
    const missAll = missingOf(stockItems) + missingOf(wasteItems) + missingOf(envioItems);
    if (missAll > 0) row = addBannerLine(ws, row, n, `⚠ ${missAll} producto(s) sin precio cargado: no se incluyen en los valores.`, 'warn');
    if (data.generalNote && data.generalNote.trim()) row = addBannerLine(ws, row, n, `Nota: ${data.generalNote.trim()}`, 'note');
    row++; // aire

    // Cifras principales
    const kpis = [
      ['Productos contados (Stock)', stockItems.length, null, false],
      ['Valor total del Stock', totalOf(stockItems), EUR, true],
    ];
    if (wasteItems.length) {
      kpis.push(['Productos con desperdicio', wasteItems.length, null, false]);
      kpis.push(['Valor del Desperdicio', totalOf(wasteItems), EUR, true]);
    }
    if (envioItems.length) {
      kpis.push(['Productos enviados', envioItems.length, null, false]);
      kpis.push(['Valor de los Envíos', totalOf(envioItems), EUR, true]);
    }
    kpis.forEach(([label, val, fmt, big]) => {
      ws.mergeCells(row, 1, row, 2);
      const a = ws.getCell(row, 1);
      a.value = label;
      a.font = { name: FONT, size: 11.5, bold: true, color: { argb: C.text } };
      a.fill = fill(C.navyLight);
      a.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      a.border = boxThin; ws.getCell(row, 2).border = boxThin;
      const b = ws.getCell(row, 3);
      b.value = val;
      if (fmt) b.numFmt = fmt;
      b.font = { name: FONT, size: big ? 14 : 12, bold: true, color: { argb: big ? C.navy : C.text } };
      b.alignment = { vertical: 'middle', horizontal: 'right' };
      b.border = boxThin;
      ws.getRow(row).height = big ? 26 : 22;
      row++;
    });
    row++;

    // Tabla por categoría
    function categoryTable(title, items, color, lightColor) {
      if (!items.length) return;
      ws.mergeCells(row, 1, row, n);
      const t = ws.getCell(row, 1);
      t.value = title;
      t.font = { name: FONT, size: 13, bold: true, color: { argb: color } };
      t.alignment = { vertical: 'middle', horizontal: 'left' };
      ws.getRow(row).height = 24;
      row++;

      addHeaderRow(ws, row, ['Categoría', 'Unidad de medida', 'Cantidad total', 'Productos', 'Valor (€)'], color,
        ['left', 'center', 'center', 'center', 'right']);
      ws.getCell(row, 1).alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      row++;

      const groups = {};
      items.forEach(p => {
        const key = p.category + '||' + p.unit;
        const g = groups[key] || (groups[key] = { category: p.category, unit: p.unit, qty: 0, count: 0, value: 0 });
        g.qty += Number(p.qty); g.count += 1; g.value += valueOf(p);
      });
      const list = Object.values(groups)
        .sort((a, b) => categories.indexOf(a.category) - categories.indexOf(b.category));

      const first = row;
      list.forEach((g, idx) => {
        const vals = [g.category, g.unit, round2(g.qty), g.count, round2(g.value)];
        const defs = [
          { al: 'left' }, { al: 'center' }, { al: 'center', bold: true }, { al: 'center' }, { al: 'right', fmt: EUR, bold: true },
        ];
        vals.forEach((val, i) => {
          const c = ws.getCell(row, i + 1);
          c.value = val;
          styleBodyCell(c, defs[i], idx % 2 === 1 ? C.zebra : C.white);
        });
        ws.getRow(row).height = 22;
        row++;
      });
      const last = row - 1;

      ws.getCell(row, 1).value = 'TOTAL';
      ws.getCell(row, 1).alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
      ws.getCell(row, 4).value = { formula: `SUM(D${first}:D${last})`, result: items.length };
      ws.getCell(row, 4).alignment = { horizontal: 'center', vertical: 'middle' };
      ws.getCell(row, 5).value = { formula: `SUM(E${first}:E${last})`, result: totalOf(items) };
      ws.getCell(row, 5).numFmt = EUR;
      ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
      styleTotalRow(ws, row, n, color);
      row += 2;
    }

    categoryTable('Stock por categoría', stockItems, C.navy);
    categoryTable('Desperdicio por categoría', wasteItems, C.red);

    // Envíos por destino
    if (envioItems.length) {
      ws.mergeCells(row, 1, row, n);
      const t = ws.getCell(row, 1);
      t.value = 'Envíos por sucursal destino';
      t.font = { name: FONT, size: 13, bold: true, color: { argb: C.teal } };
      t.alignment = { vertical: 'middle', horizontal: 'left' };
      ws.getRow(row).height = 24;
      row++;
      addHeaderRow(ws, row, ['Destino', '', '', 'Productos', 'Valor (€)'], C.teal, ['left', 'center', 'center', 'center', 'right']);
      ws.mergeCells(row, 1, row, 3);
      ws.getCell(row, 1).alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      row++;
      const first = row;
      [...new Set(envioItems.map(p => p.destino))].forEach((dest, idx) => {
        const items = envioItems.filter(p => p.destino === dest);
        ws.mergeCells(row, 1, row, 3);
        const z = idx % 2 === 1 ? C.zebra : C.white;
        styleBodyCell(ws.getCell(row, 1), { al: 'left', bold: true }, z);
        ws.getCell(row, 1).value = dest;
        ws.getCell(row, 2).border = boxThin; ws.getCell(row, 3).border = boxThin;
        ws.getCell(row, 4).value = items.length;
        styleBodyCell(ws.getCell(row, 4), { al: 'center' }, z);
        ws.getCell(row, 5).value = totalOf(items);
        styleBodyCell(ws.getCell(row, 5), { al: 'right', fmt: EUR, bold: true }, z);
        ws.getRow(row).height = 22;
        row++;
      });
      const last = row - 1;
      ws.getCell(row, 1).value = 'TOTAL';
      ws.getCell(row, 1).alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
      ws.getCell(row, 4).value = { formula: `SUM(D${first}:D${last})`, result: envioItems.length };
      ws.getCell(row, 4).alignment = { horizontal: 'center', vertical: 'middle' };
      ws.getCell(row, 5).value = { formula: `SUM(E${first}:E${last})`, result: totalOf(envioItems) };
      ws.getCell(row, 5).numFmt = EUR;
      ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
      styleTotalRow(ws, row, n, C.teal);
    }

    ws.views = [{ showGridLines: false }];
    setupPrint(ws, n, null, false);
  }

  // ---------- Armado ----------
  addResumenSheet();
  addItemsSheet('Stock', C.navy, 'Stock', stockItems, false);
  if (wasteItems.length > 0) addItemsSheet('Desperdicio', C.red, 'Desperdicio', wasteItems, true);
  if (envioItems.length > 0) addEnviosSheet();
  return wb;
}

if (typeof module !== 'undefined') { module.exports = { buildStockWorkbook }; }
