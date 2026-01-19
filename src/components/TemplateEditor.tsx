import React, { useState } from 'react';
import '../styles/TemplateEditor.css';

interface Template {
  subject: string;
  body: string;
}

interface TemplateEditorProps {
  template: Template;
  onTemplateChange: (template: Template) => void;
  availableFields: string[];
}

export const TemplateEditor: React.FC<TemplateEditorProps> = ({
  template,
  onTemplateChange,
  availableFields
}) => {
  const [showVariables, setShowVariables] = useState(true);

  const insertVariable = (variable: string, field: 'subject' | 'body') => {
    const formattedVar = variable.startsWith('{{') ? variable : `{{${variable}}}`;
    if (field === 'subject') {
      onTemplateChange({
        ...template,
        subject: template.subject + formattedVar
      });
    } else {
      onTemplateChange({
        ...template,
        body: template.body + formattedVar
      });
    }
  };

  return (
    <div className="template-editor">
      {availableFields.length > 0 && (
        <div className="variables-reference">
          <h3>Available Variables from Your Data</h3>
          <p className="variable-hint">Click a variable to insert it into your template:</p>
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

      <div className="template-section">
        <label htmlFor="subject">Subject Line</label>
        <input
          id="subject"
          type="text"
          placeholder="e.g., Hello {{FirstName}}, we have an opportunity for you"
          value={template.subject}
          onChange={(e) => onTemplateChange({ ...template, subject: e.target.value })}
          className="template-input"
        />
      </div>

      <div className="template-section">
        <label htmlFor="body">Email Body</label>
        <textarea
          id="body"
          placeholder="Dear {{FirstName}} {{LastName}},&#10;&#10;We would like to offer you an opportunity at {{Company}}.&#10;&#10;Best regards"
          value={template.body}
          onChange={(e) => onTemplateChange({ ...template, body: e.target.value })}
          className="template-textarea"
          rows={10}
        />
      </div>

      {showVariables && (
        <div className="variables-reference">
          <h3>Available Variables</h3>
          <p className="variable-hint">Use double curly braces around variable names: {`{`}{`{`}Variable{`}`}{`}`}</p>
          <div className="variable-grid">
            {COMMON_VARIABLES.map((variable) => (
              <button
                key={variable}
                className="variable-chip"
                onClick={() => insertVariable(variable, 'body')}
              >
                {variable}
              </button>
            ))}
          </div>
          <details>
            <summary>Advanced Variable Syntax</summary>
            <pre>{`{{name}}
Replace with value of 'name' field

{{name|if|then}}
If name equals 'if', replace with 'then'

{{name|if|then|else}}
If name equals 'if', replace with 'then', else with 'else'

{{name|*|if|then|else}}
If name includes 'if', replace with 'then', else with 'else'

{{name|^|if|then|else}}
If name starts with 'if', replace with 'then', else with 'else'`}</pre>
          </details>
        </div>
      )}

      <div className="template-tips">
        <h4>Tips:</h4>
        <ul>
          <li>Use {`{{Variable}}`} format for placeholders</li>
          <li>Variable names should match your data source column headers</li>
          <li>Subject line also supports variables</li>
          <li>HTML formatting is supported in body (use &lt;br&gt; for line breaks)</li>
        </ul>
      </div>
    </div>
  );
};
