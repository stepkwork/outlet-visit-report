import JSZip from 'jszip';

export type VisitWorkbookData = {
  managerName: string;
  shopName: string;
  visitDate: string;
  outletCount: string;
  outletVisitAddress: string;
  counts: Record<string, Record<string, number>>;
};

type DisplayRow = {
  brand: string;
  section: 'REF' | 'LAD';
  category: string;
  units: number;
  coverage: number;
};

const FIRST_DATA_ROW = 10;
const LAST_STYLED_DATA_ROW = 45;
const LAST_TEMPLATE_ROW = 64;
const REFRIGERATOR_CATEGORIES = new Set(['SD', '2D', 'MD', 'SBS', 'CF', 'FFC']);
const CELL_COLUMNS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];

function xmlEscape(value: string) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function excelDateSerial(date: string): number | null {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(date);
  if (!match) return null;
  const [, year, month, day] = match;
  const utc = Date.UTC(Number(year), Number(month) - 1, Number(day));
  const check = new Date(utc);
  if (
    check.getUTCFullYear() !== Number(year) ||
    check.getUTCMonth() !== Number(month) - 1 ||
    check.getUTCDate() !== Number(day)
  ) {
    return null;
  }
  return utc / 86_400_000 + 25_569;
}

function displayRows(
  counts: VisitWorkbookData['counts'],
  brands: string[],
  categories: string[],
): DisplayRow[] {
  const rows: DisplayRow[] = [];
  for (const brand of brands) {
    for (const category of categories) {
      const units = Math.max(0, Number(counts[brand]?.[category]) || 0);
      if (units === 0) continue;
      rows.push({
        brand,
        section: REFRIGERATOR_CATEGORIES.has(category) ? 'REF' : 'LAD',
        category,
        units,
        coverage: 0,
      });
    }
  }

  const totals = rows.reduce(
    (sum, row) => ({ ...sum, [row.section]: sum[row.section] + row.units }),
    { REF: 0, LAD: 0 },
  );
  for (const row of rows) {
    row.coverage = totals[row.section] > 0 ? row.units / totals[row.section] : 0;
  }

  if (rows.length === 0) {
    return [{ brand: '', section: 'REF', category: '', units: 0, coverage: 0 }];
  }
  return rows.sort((a, b) => {
    if (a.section !== b.section) return a.section === 'REF' ? -1 : 1;
    const brandOrder = brands.indexOf(a.brand) - brands.indexOf(b.brand);
    if (brandOrder !== 0) return brandOrder;
    return categories.indexOf(a.category) - categories.indexOf(b.category);
  });
}

function updateCell(rowXml: string, reference: string, value: string | number | null) {
  const cellPattern = new RegExp(
    `<c\\b(?=[^>]*\\br="${reference}")[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`,
  );
  const existing = rowXml.match(cellPattern)?.[0];
  if (!existing) {
    throw new Error(`Template is missing the expected cell ${reference}.`);
  }

  const openingTag = existing.match(/^<c\b[^>]*?(?:\/>|>)/)?.[0];
  if (!openingTag) throw new Error(`Could not read template cell ${reference}.`);

  const attributes = openingTag
    .slice(2, openingTag.endsWith('/>') ? -2 : -1)
    .replace(/\s+t="[^"]*"/g, '')
    .trimEnd();

  let replacement: string;
  if (value === null || value === '') {
    replacement = `<c ${attributes}/>`;
  } else if (typeof value === 'string') {
    replacement = `<c ${attributes} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
  } else {
    if (!Number.isFinite(value)) throw new Error(`Invalid number for ${reference}.`);
    replacement = `<c ${attributes}><v>${value}</v></c>`;
  }

  return rowXml.replace(cellPattern, replacement);
}

function cloneDataRow(templateRow: string, sourceNumber: number, targetNumber: number) {
  let row = templateRow.replace(
    new RegExp(`(\\br="${sourceNumber}")`),
    `r="${targetNumber}"`,
  );
  for (const column of CELL_COLUMNS) {
    row = row.replaceAll(
      `r="${column}${sourceNumber}"`,
      `r="${column}${targetNumber}"`,
    );
  }
  return row;
}

function fillDataRow(
  rowXml: string,
  rowNumber: number,
  data: DisplayRow | null,
  visit: VisitWorkbookData,
  recordNumber: number,
) {
  const values: (string | number | null)[] = data
    ? [
        recordNumber,
        excelDateSerial(visit.visitDate),
        visit.shopName,
        visit.outletCount.trim() === '' || !Number.isFinite(Number(visit.outletCount))
          ? null
          : Number(visit.outletCount),
        visit.outletVisitAddress,
        data.brand,
        data.brand ? data.section : null,
        data.category,
        data.units,
        data.coverage,
      ]
    : Array.from({ length: CELL_COLUMNS.length }, () => null);

  let updated = rowXml;
  CELL_COLUMNS.forEach((column, index) => {
    updated = updateCell(updated, `${column}${rowNumber}`, values[index]);
  });
  return updated;
}

function removeCalculationChain(zip: JSZip) {
  zip.remove('xl/calcChain.xml');

  const relationships = zip.file('xl/_rels/workbook.xml.rels');
  const contentTypes = zip.file('[Content_Types].xml');
  if (!relationships || !contentTypes) {
    throw new Error('The uploaded Excel template is missing workbook metadata.');
  }

  return Promise.all([
    relationships.async('string').then((xml) => {
      zip.file(
        'xl/_rels/workbook.xml.rels',
        xml.replace(
          /<Relationship\b(?=[^>]*\bType="[^"]*\/calcChain")[^>]*\/>/,
          '',
        ),
      );
    }),
    contentTypes.async('string').then((xml) => {
      zip.file(
        '[Content_Types].xml',
        xml.replace(
          /<Override\b(?=[^>]*\bPartName="\/xl\/calcChain\.xml")[^>]*\/>/,
          '',
        ),
      );
    }),
  ]);
}

export async function buildVisitWorkbook(
  templateBytes: ArrayBuffer,
  visit: VisitWorkbookData,
  brands: string[],
  categories: string[],
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(templateBytes);
  const worksheetFile = zip.file('xl/worksheets/sheet1.xml');
  if (!worksheetFile) throw new Error('The uploaded Excel template has no display sheet.');

  let worksheet = await worksheetFile.async('string');
  const rowsToExport = displayRows(visit.counts, brands, categories);
  const sheetDataMatch = worksheet.match(/<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/);
  if (!sheetDataMatch) throw new Error('The uploaded Excel template has no worksheet data.');

  const originalRows = new Map<number, string>();
  for (const match of sheetDataMatch[1].matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g)) {
    originalRows.set(Number(match[1]), match[0]);
  }

  const managerRow = originalRows.get(5);
  if (!managerRow) throw new Error('The uploaded Excel template has no manager field.');
  originalRows.set(
    5,
    updateCell(
      managerRow,
      'A5',
      `Manager: ${visit.managerName.trim()}`,
    ),
  );

  let lastRow = LAST_TEMPLATE_ROW;
  const requiredLastRow = FIRST_DATA_ROW + rowsToExport.length - 1;
  if (requiredLastRow > LAST_TEMPLATE_ROW) lastRow = requiredLastRow;

  const styledTemplateRow = originalRows.get(FIRST_DATA_ROW);
  if (!styledTemplateRow) throw new Error('The uploaded Excel template has no styled data rows.');

  for (let rowNumber = FIRST_DATA_ROW; rowNumber <= LAST_TEMPLATE_ROW; rowNumber += 1) {
    const dataIndex = rowNumber - FIRST_DATA_ROW;
    const data = rowsToExport[dataIndex] ?? null;
    const currentRow = originalRows.get(rowNumber);
    if (!currentRow) continue;
    const rowToEdit =
      data && rowNumber > LAST_STYLED_DATA_ROW
        ? cloneDataRow(styledTemplateRow, FIRST_DATA_ROW, rowNumber)
        : currentRow;
    originalRows.set(
      rowNumber,
      fillDataRow(rowToEdit, rowNumber, data, visit, dataIndex + 1),
    );
  }

  for (let rowNumber = LAST_TEMPLATE_ROW + 1; rowNumber <= lastRow; rowNumber += 1) {
    const dataIndex = rowNumber - FIRST_DATA_ROW;
    const newRow = cloneDataRow(styledTemplateRow, FIRST_DATA_ROW, rowNumber);
    originalRows.set(
      rowNumber,
      fillDataRow(newRow, rowNumber, rowsToExport[dataIndex] ?? null, visit, dataIndex + 1),
    );
  }

  const newSheetData = [...originalRows.entries()]
    .sort(([a], [b]) => a - b)
    .map(([, row]) => row)
    .join('');
  worksheet = worksheet.replace(
    /(<sheetData\b[^>]*>)[\s\S]*?(<\/sheetData>)/,
    `$1${newSheetData}$2`,
  );
  worksheet = worksheet.replace(
    /(<dimension\b[^>]*\bref=")[^"]+(")/,
    `$1A1:W${lastRow}$2`,
  );
  if (lastRow > LAST_TEMPLATE_ROW) {
    worksheet = worksheet.replaceAll('sqref="J10:J64"', `sqref="J10:J${lastRow}"`);
  }

  zip.file('xl/worksheets/sheet1.xml', worksheet);
  await removeCalculationChain(zip);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
