import { useState } from 'react';
import { importProductsFromExcel, ExcelProductRow, normalizeItemNumberFromAny } from '../utils/excelImport';

/**
 * Standalone Excel Import Demo
 * Demonstrates the fixed runtime-safe Excel parsing
 * No external dependencies - fully self-contained
 */
export const ExcelImportDemo = () => {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [data, setData] = useState<ExcelProductRow[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [testValue, setTestValue] = useState('');
  const [normalized, setNormalized] = useState('');

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    try {
      const selectedFile = e.target.files?.[0];
      if (!selectedFile) return;

      setFile(selectedFile);
      setLoading(true);
      setError(null);
      setData([]);

      // This now safely handles undefined/null values
      const result = await importProductsFromExcel(selectedFile);
      setData(result);
      setLoading(false);
    } catch (err) {
      console.error('Import error:', err);
      setError(err instanceof Error ? err.message : 'Failed to import Excel');
      setLoading(false);
    }
  };

  const handleNormalizeTest = () => {
    // Test the safe normalization function
    const result = normalizeItemNumberFromAny(testValue);
    setNormalized(result);
  };

  return (
    <div style={{ padding: '20px', maxWidth: '1200px', margin: '0 auto' }}>
      <h1 style={{ color: '#D6AA32', marginBottom: '20px' }}>
        Excel Import - Runtime Safety Demo
      </h1>

      {/* File Upload Section */}
      <div style={{ marginBottom: '30px', padding: '20px', background: '#111', borderRadius: '8px' }}>
        <h2 style={{ color: '#F7EFE0', fontSize: '18px', marginBottom: '15px' }}>
          Upload Excel File
        </h2>
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileChange}
          disabled={loading}
          style={{
            padding: '10px',
            background: '#222',
            border: '1px solid #D6AA32',
            borderRadius: '4px',
            color: '#F7EFE0',
            marginBottom: '10px',
            display: 'block',
          }}
        />
        <p style={{ fontSize: '12px', color: '#999' }}>
          Supports: .xlsx, .xls | Max 10MB | Reads ALL sheets
        </p>

        {file && (
          <div style={{ marginTop: '10px', padding: '10px', background: '#222', borderRadius: '4px' }}>
            <strong style={{ color: '#D6AA32' }}>Selected:</strong>{' '}
            <span style={{ color: '#F7EFE0' }}>{file.name}</span>
          </div>
        )}

        {loading && (
          <div style={{ textAlign: 'center', padding: '20px', color: '#D6AA32' }}>
            Loading...
          </div>
        )}

        {error && (
          <div style={{
            marginTop: '10px',
            padding: '10px',
            background: '#300',
            border: '1px solid #800',
            borderRadius: '4px',
            color: '#f88',
          }}>
            <strong>Error:</strong> {error}
          </div>
        )}
      </div>

      {/* Item Number Normalization Tester */}
      <div style={{ marginBottom: '30px', padding: '20px', background: '#111', borderRadius: '8px' }}>
        <h2 style={{ color: '#F7EFE0', fontSize: '18px', marginBottom: '15px' }}>
          Item Number Normalization Tester
        </h2>
        <p style={{ fontSize: '14px', color: '#999', marginBottom: '10px' }}>
          Test safe normalization with any value (filename, Excel cell, undefined, etc.)
        </p>
        <div style={{ display: 'flex', gap: '10px', marginBottom: '10px' }}>
          <input
            type="text"
            value={testValue}
            onChange={(e) => setTestValue(e.target.value)}
            placeholder="Enter value (e.g., '9130', 'AR-9130', 'AT-9130.png', empty, etc.)"
            style={{
              flex: 1,
              padding: '10px',
              background: '#222',
              border: '1px solid #D6AA32',
              borderRadius: '4px',
              color: '#F7EFE0',
            }}
          />
          <button
            onClick={handleNormalizeTest}
            style={{
              padding: '10px 20px',
              background: '#D6AA32',
              border: 'none',
              borderRadius: '4px',
              color: '#000',
              fontWeight: 'bold',
              cursor: 'pointer',
            }}
          >
            Normalize
          </button>
        </div>
        {normalized !== '' && (
          <div style={{ padding: '10px', background: '#222', borderRadius: '4px' }}>
            <strong style={{ color: '#D6AA32' }}>Result:</strong>{' '}
            <span style={{ color: '#8f8', fontFamily: 'monospace' }}>
              {normalized || '(empty string)'}
            </span>
          </div>
        )}
        <div style={{ marginTop: '10px', fontSize: '12px', color: '#666' }}>
          <strong>Examples:</strong><br />
          • 9130 → AT-9130<br />
          • AR-9130 → AT-9130<br />
          • AT-9130 → AT-9130<br />
          • 9102FU → AT-9102FU<br />
          • AT-9130.png → AT-9130<br />
          • undefined/null → (empty string)
        </div>
      </div>

      {/* Results Table */}
      {data.length > 0 && (
        <div style={{ padding: '20px', background: '#111', borderRadius: '8px' }}>
          <h2 style={{ color: '#F7EFE0', fontSize: '18px', marginBottom: '15px' }}>
            Imported Data ({data.length} rows)
          </h2>
          <div style={{ overflowX: 'auto' }}>
            <table style={{
              width: '100%',
              borderCollapse: 'collapse',
              fontSize: '14px',
            }}>
              <thead>
                <tr style={{ background: '#222' }}>
                  <th style={{ padding: '10px', textAlign: 'left', color: '#D6AA32', borderBottom: '2px solid #D6AA32' }}>
                    Sheet
                  </th>
                  <th style={{ padding: '10px', textAlign: 'left', color: '#D6AA32', borderBottom: '2px solid #D6AA32' }}>
                    Item Number
                  </th>
                  <th style={{ padding: '10px', textAlign: 'left', color: '#D6AA32', borderBottom: '2px solid #D6AA32' }}>
                    Price
                  </th>
                  <th style={{ padding: '10px', textAlign: 'left', color: '#D6AA32', borderBottom: '2px solid #D6AA32' }}>
                    Description
                  </th>
                </tr>
              </thead>
              <tbody>
                {data.map((row, index) => (
                  <tr key={index} style={{
                    background: index % 2 === 0 ? '#0a0a0a' : '#111',
                    borderBottom: '1px solid #333',
                  }}>
                    <td style={{ padding: '10px', color: '#F7EFE0' }}>
                      {row.sheet_name}
                    </td>
                    <td style={{ padding: '10px', color: '#8f8', fontFamily: 'monospace' }}>
                      {row.item_number}
                    </td>
                    <td style={{ padding: '10px', color: '#F7EFE0' }}>
                      {row.price || '-'}
                    </td>
                    <td style={{ padding: '10px', color: '#ccc', fontSize: '12px' }}>
                      {row.excel_description || '-'}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Safety Features Info */}
      <div style={{ marginTop: '30px', padding: '20px', background: '#111', borderRadius: '8px' }}>
        <h2 style={{ color: '#D6AA32', fontSize: '18px', marginBottom: '15px' }}>
          ✅ Runtime Safety Features
        </h2>
        <ul style={{ color: '#F7EFE0', lineHeight: '1.8' }}>
          <li>✓ Safe string operations - never crashes on undefined/null</li>
          <li>✓ Safe .includes(), .startsWith(), .endsWith() with proper checks</li>
          <li>✓ Safe filename parsing from File objects</li>
          <li>✓ Safe Excel cell value extraction</li>
          <li>✓ Safe header normalization (handles line breaks, null values)</li>
          <li>✓ Graceful error handling with console.error and stack traces</li>
          <li>✓ Skips empty/invalid rows without crashing</li>
          <li>✓ Clear UI error messages instead of runtime crashes</li>
          <li>✓ Processes ALL sheets in workbook</li>
          <li>✓ Flexible column header detection</li>
        </ul>
      </div>
    </div>
  );
};
