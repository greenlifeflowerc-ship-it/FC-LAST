import * as XLSX from 'xlsx';

export type ExcelProductRow = {
  item_number: string;
  price: string;
  excel_description: string;
  sheet_name: string;
};

/**
 * Safe string conversion helper
 * Converts any value to a trimmed string, never crashes
 */
export function safeString(value: unknown): string {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

/**
 * Safe lowercase conversion
 * Never crashes on undefined/null values
 */
export function safeLower(value: unknown): string {
  return safeString(value).toLowerCase();
}

/**
 * Safe uppercase conversion
 * Never crashes on undefined/null values
 */
export function safeUpper(value: unknown): string {
  return safeString(value).toUpperCase();
}

/**
 * Normalize item number from any value (Excel cell, filename, etc.)
 * - Removes file extensions and whitespace
 * - If value starts with "AR-", replace with "AT-"
 * - If value starts with "AT-", keep as is
 * - Otherwise add "AT-" prefix
 * - Never crashes on undefined/null values
 */
export function normalizeItemNumberFromAny(value: unknown): string {
  try {
    const raw = safeUpper(value)
      .replace(/\.[^/.]+$/, '') // Remove file extension
      .replace(/\s+/g, ''); // Remove all whitespace

    if (!raw) return '';

    // AR- prefix: convert to AT-
    if (raw.startsWith('AR-')) {
      return raw.replace(/^AR-/, 'AT-');
    }

    // AT- prefix: keep as is
    if (raw.startsWith('AT-')) {
      return raw;
    }

    // Add AT- prefix
    return `AT-${raw}`;
  } catch (error) {
    console.error('normalizeItemNumberFromAny error:', error, 'value:', value);
    return '';
  }
}

/**
 * Legacy function - now uses normalizeItemNumberFromAny
 */
export function normalizeItemNumber(value: string): string {
  return normalizeItemNumberFromAny(value);
}

/**
 * Legacy function - now uses normalizeItemNumberFromAny
 */
export function normalizeItemNumberFromFilename(filename: string): string {
  return normalizeItemNumberFromAny(filename);
}

/**
 * Normalize header name to match expected column names
 * Removes line breaks, underscores, trims, lowercases, collapses multiple spaces
 * Never crashes on undefined/null
 */
function normalizeHeader(header: unknown): string {
  try {
    const normalized = safeLower(header)
      .replace(/\r?\n|\r/g, '') // Remove line breaks
      .replace(/_/g, ' ') // Convert underscores to spaces
      .replace(/\s+/g, ' '); // Collapse multiple spaces

    return normalized;
  } catch (error) {
    console.error('normalizeHeader error:', error, 'header:', header);
    return '';
  }
}

/**
 * Get compact version of normalized header (no spaces)
 */
function compactHeader(header: string): string {
  return header.replace(/\s+/g, '');
}

/**
 * Import products from Excel file
 * Reads ALL sheets and extracts product data from row 1 headers
 * Safe against undefined/null values, never crashes
 */
export async function importProductsFromExcel(file: File): Promise<ExcelProductRow[]> {
  try {
    // Validate file
    if (!file || !file.name) {
      throw new Error('Invalid file: file is missing or has no name');
    }

    // Validate file size (max 10MB)
    const maxSize = 10 * 1024 * 1024;
    if (file.size > maxSize) {
      throw new Error('Excel file is too large. Maximum size is 10MB.');
    }

    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          if (!data) {
            throw new Error('Failed to read file');
          }

          // Read workbook
          const workbook = XLSX.read(data, { type: 'binary' });

          if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
            throw new Error('Excel file has no sheets');
          }

          const allProducts: ExcelProductRow[] = [];

          // Process ALL sheets
          for (const sheetName of workbook.SheetNames) {
            try {
              const safeSheetName = safeString(sheetName) || 'Unknown Sheet';
              const worksheet = workbook.Sheets[sheetName];

              if (!worksheet) {
                console.warn(`Skipping sheet "${safeSheetName}": worksheet is undefined`);
                continue;
              }

              // Convert to JSON with header row
              const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

              if (!jsonData || jsonData.length < 2) {
                console.warn(`Skipping sheet "${safeSheetName}": no data rows`);
                continue; // Skip sheets with no data
              }

              // Extract headers from FIRST row only - safe version
              const headerRow = jsonData[0];
              if (!Array.isArray(headerRow)) {
                console.warn(`Skipping sheet "${safeSheetName}": invalid header row`);
                continue;
              }

              const normalizedHeaders = headerRow.map((h: any) => normalizeHeader(h));
              const compactHeaders = normalizedHeaders.map((h: string) => compactHeader(h));

              // Find column indices - safe version with both normalized and compact matching
              const treeNoIndex = normalizedHeaders.findIndex((h, idx) => {
                if (!h) return false;
                const compact = compactHeaders[idx];
                return (
                  h === 'tree no' ||
                  compact === 'treeno' ||
                  h === 'tree_no'
                );
              });

              const priceIndex = normalizedHeaders.findIndex((h) => {
                if (!h) return false;
                return h === 'r' || h === 'price';
              });

              const descriptionIndex = normalizedHeaders.findIndex((h, idx) => {
                if (!h) return false;
                const compact = compactHeaders[idx];
                return compact === 'description' || h.includes('descript');
              });

              if (treeNoIndex === -1) {
                console.warn(`Skipping sheet "${safeSheetName}": no "tree no" column found`);
                continue; // Skip sheets without tree no column
              }

              // Process data rows (starting from row 2)
              for (let i = 1; i < jsonData.length; i++) {
                try {
                  const row = jsonData[i];

                  // Skip empty rows
                  if (!row || !Array.isArray(row) || row.length === 0) continue;

                  // Safe cell extraction
                  const treeNoCell = treeNoIndex !== -1 ? row[treeNoIndex] : undefined;
                  const treeNoStr = safeString(treeNoCell);
                  if (!treeNoStr) continue; // Skip rows without tree number

                  const itemNumber = normalizeItemNumberFromAny(treeNoCell);
                  if (!itemNumber) continue;

                  const priceCell = priceIndex !== -1 ? row[priceIndex] : '';
                  const price = safeString(priceCell);

                  const descCell = descriptionIndex !== -1 ? row[descriptionIndex] : '';
                  const description = safeString(descCell);

                  allProducts.push({
                    item_number: itemNumber,
                    price,
                    excel_description: description,
                    sheet_name: safeSheetName,
                  });
                } catch (rowError) {
                  console.error(
                    `Error processing row ${i + 1} in sheet "${safeSheetName}":`,
                    rowError
                  );
                  // Continue processing other rows
                }
              }
            } catch (sheetError) {
              console.error(`Error processing sheet "${safeString(sheetName)}":`, sheetError);
              // Continue processing other sheets
            }
          }

          if (allProducts.length === 0) {
            throw new Error(
              'No valid product data found in Excel file. Make sure the file has "tree no", "R", and "DESCRIPTION" columns in row 1.'
            );
          }

          console.log(`Successfully imported ${allProducts.length} products from Excel`);
          resolve(allProducts);
        } catch (error) {
          console.error('Excel import error:', error);
          reject(error);
        }
      };

      reader.onerror = () => {
        const error = new Error('Failed to read Excel file');
        console.error('FileReader error:', error);
        reject(error);
      };

      reader.readAsBinaryString(file);
    });
  } catch (error) {
    console.error('importProductsFromExcel error:', error);
    throw error;
  }
}
