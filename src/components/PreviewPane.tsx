import React, { useState } from 'react';
import { TemplateEngine } from '../lib/TemplateEngine';
import '../styles/PreviewPane.css';

interface PreviewPaneProps {
  template: {
    subject: string;
    body: string;
  };
  recipients: any[];
  toTemplate: string;
}

export const PreviewPane: React.FC<PreviewPaneProps> = ({
  template,
  recipients,
  toTemplate
}) => {
  const [currentIndex, setCurrentIndex] = useState(0);
  const engine = new TemplateEngine();

  if (recipients.length === 0) {
    return (
      <div className="preview-pane empty">
        <p>No recipients loaded. Go to step 2 to load data.</p>
      </div>
    );
  }

  if (template.subject === '' && template.body === '') {
    return (
      <div className="preview-pane empty">
        <p>No template created. Go to step 1 to create a template.</p>
      </div>
    );
  }

  const currentRecipient = recipients[currentIndex];
  const renderedSubject = engine.render(template.subject, currentRecipient);
  const renderedBody = engine.render(template.body, currentRecipient);

  return (
    <div className="preview-pane">
      <div className="preview-controls">
        <button
          onClick={() => setCurrentIndex(Math.max(0, currentIndex - 1))}
          disabled={currentIndex === 0}
          className="nav-button"
        >
          ← Previous
        </button>
        <span className="preview-counter">
          Recipient {currentIndex + 1} of {recipients.length}
        </span>
        <button
          onClick={() => setCurrentIndex(Math.min(recipients.length - 1, currentIndex + 1))}
          disabled={currentIndex === recipients.length - 1}
          className="nav-button"
        >
          Next →
        </button>
      </div>

      <div className="preview-info">
        <h4>Recipient Details:</h4>
        <div className="recipient-details">
          {Object.entries(currentRecipient).map(([key, value]) => (
            <div key={key} className="detail-row">
              <span className="detail-key">{key}:</span>
              <span className="detail-value">{String(value)}</span>
            </div>
          ))}
        </div>
      </div>

      <div className="email-preview">
        <div className="preview-field">
          <label>To:</label>
          <div className="preview-value">{engine.render(toTemplate, currentRecipient)}</div>
        </div>

        <div className="preview-field">
          <label>Subject:</label>
          <div className="preview-value subject">{renderedSubject}</div>
        </div>

        <div className="preview-field">
          <label>Body:</label>
          <div 
            className="preview-value body"
            dangerouslySetInnerHTML={{ __html: renderedBody }}
          />
        </div>
      </div>

      <div className="preview-tips">
        <h4>Notes:</h4>
        <ul>
          <li>Navigate through recipients to preview how variables are replaced</li>
          <li>If a field shows 'undefined', check that your variable name matches the data column</li>
          <li>HTML formatting in the body will be rendered as-is</li>
        </ul>
      </div>
    </div>
  );
};
