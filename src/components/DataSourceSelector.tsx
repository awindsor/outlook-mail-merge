import React, { useState } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import '../styles/DataSourceSelector.css';

interface DataSourceSelectorProps {
  onDataLoaded: (recipients: any[]) => void;
  onSourceChange: (source: 'csv' | 'xlsx' | 'manual' | null) => void;
  toTemplate: string;
  onToTemplateChange: (template: string) => void;
  availableFields: string[];
}

export const DataSourceSelector: React.FC<DataSourceSelectorProps> = ({
  onDataLoaded,
  onSourceChange,
  toTemplate,
  onToTemplateChange,
  availableFields
}) => {
  const [sourceType, setSourceType] = useState<'csv' | 'xlsx' | 'manual' | null>(null);
  const [manualData, setManualData] = useState<string>('');
  const [fileName, setFileName] = useState<string>('');
  const [loadedData, setLoadedData] = useState<any[]>([]);
  const [error, setError] = useState<string>('');

  const handleCSVUpload = (file: File) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        if (results.data && results.data.length > 0) {
          setLoadedData(results.data as any[]);
          onDataLoaded(results.data as any[]);
          setFileName(file.name);
          setSourceType('csv');
          setError('');
        } else {
          setError('CSV file is empty or invalid');
        }
      },
      error: (error) => {
        setError(`Error parsing CSV: ${error.message}`);
      }
    });
  };

  const handleXLSXUpload = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        if (jsonData && jsonData.length > 0) {
          setLoadedData(jsonData);
          onDataLoaded(jsonData);
          setFileName(file.name);
          setSourceType('xlsx');
          setError('');
        } else {
          setError('Excel file is empty or invalid');
        }
      } catch (err) {
        setError(`Error parsing Excel: ${err instanceof Error ? err.message : 'Unknown error'}`);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleManualEntry = () => {
    try {
      const parsed = JSON.parse(manualData);
      const dataArray = Array.isArray(parsed) ? parsed : [parsed];
      
      if (dataArray.length > 0 && typeof dataArray[0] === 'object') {
        setLoadedData(dataArray);
        onDataLoaded(dataArray);
        setSourceType('manual');
        setError('');
      } else {
        setError('Invalid JSON format. Must be an object or array of objects.');
      }
    } catch (err) {
      setError(`JSON Parse Error: ${err instanceof Error ? err.message : 'Unknown error'}`);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const fileName = file.name.toLowerCase();
    if (fileName.endsWith('.csv')) {
      handleCSVUpload(file);
    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
      handleXLSXUpload(file);
    } else {
      setError('Unsupported file type. Please use CSV or Excel.');
    }
  };

  return (
    <div className="data-source-selector">
      <div className="selector-header">
        <h2>Load Recipients</h2>
        <p>Upload your recipient data from CSV, Excel, or enter manually</p>
      </div>

      <div className="upload-options">
        <div className="option-group">
          <label htmlFor="file-upload" className="upload-label">
            📁 Upload File (CSV or Excel)
          </label>
          <input
            id="file-upload"
            type="file"
            accept=".csv,.xlsx,.xls"
            onChange={handleFileUpload}
            className="file-input"
          />
          {fileName && <p className="file-name">✓ Loaded: {fileName}</p>}
        </div>

        <div className="divider">OR</div>

        <div className="option-group">
          <label htmlFor="manual-entry">
            ✏️ Enter JSON Data Manually
          </label>
          <textarea
            id="manual-entry"
            placeholder={`[
  { "FirstName": "John", "LastName": "Doe", "Email": "john@example.com" },
  { "FirstName": "Jane", "LastName": "Smith", "Email": "jane@example.com" }
]`}
            value={manualData}
            onChange={(e) => setManualData(e.target.value)}
            className="manual-textarea"
            rows={8}
          />
          <button 
            onClick={handleManualEntry}
            className="load-button"
          >
            Load JSON Data
          </button>
        </div>
      </div>

      {error && <div className="error-banner">{error}</div>}

      {loadedData.length > 0 && (
        <div className="data-preview">
          <h3>Loaded Data ({loadedData.length} recipients)</h3>
          <div className="table-container">
            <table className="recipients-table">
              <thead>
                <tr>
                  {Object.keys(loadedData[0]).map((key) => (
                    <th key={key}>{key}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {loadedData.slice(0, 5).map((row, idx) => (
                  <tr key={idx}>
                    {Object.values(row).map((value: any, cidx: number) => (
                      <td key={cidx}>{String(value)}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {loadedData.length > 5 && (
            <p className="data-count">... and {loadedData.length - 5} more recipients</p>
          )}
        </div>
      )}

      <div className="data-tips">
        <h4>Tips:</h4>
        <ul>
          <li>CSV files should have headers in the first row</li>
          <li>Column headers must match template variables (e.g., FirstName, Email)</li>
          <li>Excel files use the first sheet by default</li>
          <li>Manual entry uses JSON object format</li>
        </ul>
      </div>
    </div>
  );
};
