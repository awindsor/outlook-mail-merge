import React, { useState } from 'react';
import '../styles/TemplateEditor.css';

interface Template {
  subject: string;
  body: string;
}

interface TemplateEditorProps {
  template: Template;
  onTemplateChange: (template: Template) => void;
}

const COMMON_VARIABLES = [
  '{{FirstName}}',
  '{{LastName}}',
  '{{Email}}',
  '{{Company}}',
  '{{Title}}'
];

export const TemplateEditor: React.FC<TemplateEditorProps> = ({
  template,
  onTemplateChange
}) => {
  const [showVariables, setShowVariables] = useState(false);

  const insertVariable = (variable: string, field: 'subject' | 'body') => {
    if (field === 'subject') {
      onTemplateChange({
        ...template,
        subject: template.subject + variable
      });
    } else {
      onTemplateChange({
        ...template,
        body: template.body + variable
      });
    }
  };

  return (
    <div className="template-editor">
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
        <button
          className="variable-button"
          onClick={() => insertVariable(' {{FirstName}}', 'subject')}
        >
          + Variable
        </button>
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
        <div className="template-actions">
          <button
            className="variable-button"
            onClick={() => insertVariable(' {{FirstName}}', 'body')}
          >
            + Variable
          </button>
          <button
            className="toggle-variables"
            onClick={() => setShowVariables(!showVariables)}
          >
            {showVariables ? 'Hide' : 'Show'} Variables
          </button>
        </div>
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
