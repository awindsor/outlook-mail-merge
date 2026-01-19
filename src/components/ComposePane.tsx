import React, { useState } from 'react';
import '../styles/ComposePane.css';

interface ComposePaneProps {
  subject: string;
  body: string;
  onSubjectChange: (subject: string) => void;
  onBodyChange: (body: string) => void;
  availableFields: string[];
  toTemplate: string;
  onToTemplateChange: (template: string) => void;
  onLoadFromOutlook: () => void;
  isLoading: boolean;
}

export const ComposePane: React.FC<ComposePaneProps> = ({
  subject,
  body,
  onSubjectChange,
  onBodyChange,
  availableFields,
  toTemplate,
  onToTemplateChange,
  onLoadFromOutlook,
  isLoading
}) => {
  const [showVariables, setShowVariables] = useState(true);

  const insertVariable = (variable: string, field: 'subject' | 'body' | 'to') => {
    const formattedVar = variable.startsWith('{{') ? variable : `{{${variable}}}`;
    
    if (field === 'subject') {
      onSubjectChange(subject + formattedVar);
    } else if (field === 'body') {
      onBodyChange(body + formattedVar);
    } else if (field === 'to') {
      onToTemplateChange(toTemplate + formattedVar);
    }
  };

  return (
    <div className="compose-pane">
      <div className="compose-header">
        <h2>Compose Message with Variables</h2>
        <p>Use variables like {{FirstName}} that will be replaced for each recipient.</p>
        <button 
          className="load-button"
          onClick={onLoadFromOutlook}
          disabled={isLoading}
        >
          {isLoading ? 'Loading...' : 'Load from Outlook Compose Pane'}
        </button>
      </div>

      {availableFields.length > 0 && (
        <div className="variables-reference">
          <h3>Available Variables from Your Data</h3>
          <p className="variable-hint">Click a variable to insert it:</p>
          <div className="variable-grid">
            {availableFields.map((field) => (
              <button
                key={field}
                className="variable-chip"
                onClick={() => insertVariable(field, 'body')}
              >
                {`{{${field}}}`}
              </button>
            ))}
          </div>
        </div>
      )}

      <div className="compose-section">
        <label htmlFor="to-template">
          <strong>Recipient Email Template:</strong>
        </label>
        <div className="template-input-group">
          <input
            id="to-template"
            type="text"
            placeholder="e.g., {{Email}} or {{FirstName}} {{LastName}} <{{Email}}>"
            value={toTemplate}
            onChange={(e) => onToTemplateChange(e.target.value)}
            className="to-template-input"
          />
          <div className="quick-variables">
            {availableFields.map((field) => (
              <button
                key={`to-${field}`}
                className="quick-var-btn"
                onClick={() => insertVariable(field, 'to')}
                title={`Insert {{${field}}}`}
              >
                {field}
              </button>
            ))}
          </div>
        </div>
        <p className="helper-text">Examples: "{{Email}}" or "{{FirstName}} {{LastName}} &lt;{{Email}}&gt;"</p>
      </div>

      <div className="compose-section">
        <label htmlFor="subject">Subject Line</label>
        <input
          id="subject"
          type="text"
          placeholder="e.g., Hello {{FirstName}}, special offer for {{Company}}"
          value={subject}
          onChange={(e) => onSubjectChange(e.target.value)}
          className="compose-input"
        />
      </div>

      <div className="compose-section">
        <label htmlFor="body">Email Body</label>
        <textarea
          id="body"
          placeholder="Compose your email here with variables like {{FirstName}}, {{Email}}, etc.&#10;&#10;Example:&#10;Dear {{FirstName}} {{LastName}},&#10;&#10;We have a special opportunity at {{Company}}...&#10;&#10;Best regards"
          value={body}
          onChange={(e) => onBodyChange(e.target.value)}
          className="compose-textarea"
          rows={15}
        />
      </div>

      <div className="compose-tips">
        <h4>Tips:</h4>
        <ul>
          <li><strong>Use double curly braces:</strong> {{VariableName}} will be replaced with data</li>
          <li><strong>Variable names must match your data columns exactly:</strong> {{FirstName}}, {{Email}}, {{Company}}, etc.</li>
          <li><strong>Click variable buttons</strong> above to quickly insert them</li>
          <li><strong>Preview your message</strong> in the next step before sending</li>
          <li><strong>Subject and "To" fields also support variables</strong></li>
        </ul>
      </div>
    </div>
  );
};
