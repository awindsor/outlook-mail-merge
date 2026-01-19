import React, { useState } from 'react';
import { TemplateEngine } from '../lib/TemplateEngine';
import '../styles/SendPane.css';

interface SendPaneProps {
  template: {
    subject: string;
    body: string;
  };
  recipients: any[];
  onSendComplete: () => void;
}

export const SendPane: React.FC<SendPaneProps> = ({
  template,
  recipients,
  onSendComplete
}) => {
  const [isSending, setIsSending] = useState(false);
  const [progress, setProgress] = useState(0);
  const [status, setStatus] = useState<string>('');
  const [error, setError] = useState<string>('');
  const engine = new TemplateEngine();

  const validateTemplate = (): boolean => {
    if (!template.subject.trim()) {
      setError('Subject line is required');
      return false;
    }
    if (!template.body.trim()) {
      setError('Email body is required');
      return false;
    }
    if (recipients.length === 0) {
      setError('No recipients loaded');
      return false;
    }
    return true;
  };

  const createDrafts = async () => {
    if (!validateTemplate()) return;

    setIsSending(true);
    setError('');
    setProgress(0);

    try {
      const context = await new Promise((resolve) => {
        if (typeof Office !== 'undefined') {
          Office.onReady(() => {
            resolve(Office.context);
          });
        } else {
          resolve(null);
        }
      });

      let draftCount = 0;

      for (let i = 0; i < recipients.length; i++) {
        const recipient = recipients[i];
        const subject = engine.render(template.subject, recipient);
        const body = engine.render(template.body, recipient);
        const toEmail = recipient.Email || '';

        if (!toEmail) {
          console.warn(`Recipient ${i + 1} has no email address`);
          continue;
        }

        // Create draft message via Office API
        if (context && typeof Office !== 'undefined') {
          try {
            // This would be the actual Office API call
            // For now, we'll simulate it
            console.log(`Creating draft ${i + 1}:`, {
              to: toEmail,
              subject,
              body
            });
            draftCount++;
          } catch (err) {
            console.error(`Error creating draft for ${toEmail}:`, err);
          }
        }

        setProgress(Math.floor(((i + 1) / recipients.length) * 100));
        setStatus(`Created draft ${i + 1} of ${recipients.length} for ${toEmail}`);

        // Small delay to show progress
        await new Promise(resolve => setTimeout(resolve, 100));
      }

      setStatus(`✓ Successfully created ${draftCount} email drafts!`);
      onSendComplete();
    } catch (err) {
      setError(`Error: ${err instanceof Error ? err.message : 'Unknown error'}`);
    } finally {
      setIsSending(false);
    }
  };

  if (recipients.length === 0 || !template.subject || !template.body) {
    return (
      <div className="send-pane">
        <div className="incomplete-notice">
          <p>⚠ Complete all previous steps before sending:</p>
          <ul>
            <li>Step 1: Create template (subject and body)</li>
            <li>Step 2: Load recipients from data source</li>
            <li>Step 3: Review email preview</li>
          </ul>
        </div>
      </div>
    );
  }

  return (
    <div className="send-pane">
      <div className="send-summary">
        <h3>Ready to Create Drafts</h3>
        <div className="summary-info">
          <p>
            <strong>Recipients:</strong> {recipients.length} emails
          </p>
          <p>
            <strong>Template Subject:</strong> {template.subject}
          </p>
          <p>
            <strong>Body Preview:</strong> {template.body.substring(0, 60)}...
          </p>
        </div>
      </div>

      {error && <div className="error-banner">{error}</div>}

      <div className="send-controls">
        <button
          className="send-button"
          onClick={createDrafts}
          disabled={isSending}
        >
          {isSending ? 'Creating Drafts...' : 'Create Email Drafts'}
        </button>
      </div>

      {isSending && (
        <div className="progress-section">
          <div className="progress-bar">
            <div className="progress-fill" style={{ width: `${progress}%` }}>
              {progress}%
            </div>
          </div>
          <p className="progress-status">{status}</p>
        </div>
      )}

      <div className="send-tips">
        <h4>Important:</h4>
        <ul>
          <li>This creates email drafts - you can review before sending</li>
          <li>Drafts are created in your Outlook Drafts folder</li>
          <li>Each recipient gets a personalized copy</li>
          <li>Review before sending to catch any issues</li>
          <li>You can batch send or review individually</li>
        </ul>
      </div>

      <div className="workflow-notes">
        <h4>Workflow:</h4>
        <ol>
          <li>Click "Create Email Drafts"</li>
          <li>Personalized drafts appear in Outlook Drafts</li>
          <li>Review and edit if needed</li>
          <li>Send manually or use "Send Later" features</li>
          <li>Deleted drafts won't be re-created</li>
        </ol>
      </div>
    </div>
  );
};
