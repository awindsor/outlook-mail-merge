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
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheet];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        if (jsonData.length > 0) {
          setLoadedData(jsonData);
          onDataLoaded(jsonData);
          setFileName(file.name);
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

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (sourceType === 'csv') {
      handleCSVUpload(file);
    } else if (sourceType === 'xlsx') {
      handleXLSXUpload(file);
    }
  };

  const handleManualDataChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setManualData(e.target.value);
  };

  const parseManualData = () => {
    try {
      const data = JSON.parse(`[${manualData}]`);
      setLoadedData(data);
      onDataLoaded(data);
      setError('');
    } catch (err) {
      setError('Invalid JSON format');
    }
  };

  return (
    <div className="data-source-selector">
      <div className="source-type-selector">
        <h3>Select Data Source</h3>
        <div className="source-buttons">
          <button
            className={`source-button ${sourceType === 'csv' ? 'active' : ''}`}
            onClick={() => {
              setSourceType('csv');
              onSourceChange('csv');
              setLoadedData([]);
              setFileName('');
              setError('');
            }}
          >
            📄 CSV File
          </button>
          <button
            className={`source-button ${sourceType === 'xlsx' ? 'active' : ''}`}
            onClick={() => {
              setSourceType('xlsx');
              onSourceChange('xlsx');
              setLoadedData([]);
              setFileName('');
              setError('');
            }}
          >
            📊 Excel File
          </button>
          <button
            className={`source-button ${sourceType === 'manual' ? 'active' : ''}`}
            onClick={() => {
              setSourceType('manual');
              onSourceChange('manual');
              setLoadedData([]);
              setFileName('');
              setError('');
            }}
          >
            ✏️ Manual Entry
          </button>
        </div>
      </div>

      {sourceType === 'csv' && (
        <div className="upload-section">
          <label htmlFor="csv-file">Upload CSV File</label>
          <input
            id="csv-file"
            type="file"
            accept=".csv"
            onChange={handleFileInput}
            className="file-input"
          />
          {fileName && <p className="file-name">Loaded: {fileName}</p>}
        </div>
      )}

      {sourceType === 'xlsx' && (
        <div className="upload-section">
          <label htmlFor="xlsx-file">Upload Excel File</label>
          <input
            id="xlsx-file"
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFileInput}
            className="file-input"
          />
          {fileName && <p className="file-name">Loaded: {fileName}</p>}
        </div>
      )}

      {sourceType === 'manual' && (
        <div className="manual-section">
          <label htmlFor="manual-data">Enter JSON Data (one record per line)</label>
          <textarea
            id="manual-data"
            value={manualData}
            onChange={handleManualDataChange}
            placeholder='{"FirstName":"John","LastName":"Doe","Email":"john@example.com"}&#10;{"FirstName":"Jane","LastName":"Smith","Email":"jane@example.com"}'
            className="manual-textarea"
            rows={6}
          />
          <button onClick={parseManualData} className="parse-button">
            Parse Data
          </button>
        </div>
      )}

      {error && <div className="error-message">{error}</div>}

      {loadedData.length > 0 && (
        <div className="data-preview">
          <h3>Loaded Recipients ({loadedData.length})</h3>
          
          <div className="email-field-selector">
            <label htmlFor="to-template">
              <strong>Email "To" Field Template:</strong>
            </label>
            <input
              id="to-template"
              type="text"
              placeholder="e.g., {{Name}} <{{Email}}> or just {{Email}}"
              value={toTemplate}
              onChange={(e) => onToTemplateChange(e.target.value)}
              className="to-template-input"
            />
            <p className="helper-text">Use variables with double curly braces (e.g., {{Name}}, {{Email}}). Examples: "{{Email}}" or "{{FirstName}} {{LastName}} <{{Email}}>"</p>
            
            <div className="quick-variables">
              <span className="quick-label">Quick insert:</span>
              {availableFields.map((field) => (
                <button
                  key={field}
                  className="quick-var-btn"
                  onClick={() => onToTemplateChange(toTemplate + `{{${field}}}`)}
                >
                  {{`{${field}}}`}
                </button>
              ))}
            </div>
          </div>

          <div className="table-container">
            <table className="recipients-table">
              <thead>
                <tr>
                  {Object.keys(loadedData[0]).map((key) => (
                    <th key={key}>
                      {key}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {loadedData.slice(0, 5).map((row, idx) => (
                  <tr key={idx}>
                    {Object.entries(row).map(([key, value]: [string, any], cidx) => (
                      <td kevalues(row).map((value: any, cidx) => (
                      <td key={cidx
                      </td>
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
