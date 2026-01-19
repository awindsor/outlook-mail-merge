import React, { useState } from 'react';
import { TemplateEngine } from '../lib/TemplateEngine';
import '../styles/SendPane.css';

interface SendPaneProps {
  template: {
    subject: string;
    body: string;
  };
  recipients: any[];
  toTemplate: string;
  onLoadFromOutlook: () => void;
  isLoadingMessage: boolean;
  messageError?: string;
  onSendComplete: () => void;
}

export const SendPane: React.FC<SendPaneProps> = ({
  template,
  recipients,
  toTemplate,
  onLoadFromOutlook,
  isLoadingMessage,
  messageError,
  onSendComplete
}) => {
  const [isSending, setIsSending] = useState(false);
  const [progress, setProgress] = useState(0);
  const [status, setStatus] = useState<string>('');
  const [error, setError] = useState<string>('');
  const engine = new TemplateEngine();

  React.useEffect(() => {
    console.log('SendPane rendered with:', {
      subject: template.subject,
      bodyLength: template.body.length,
      recipients: recipients.length,
      toTemplate,
      messageError,
      isLoadingMessage
    });
  }, [template.subject, template.body, recipients.length, toTemplate, messageError, isLoadingMessage]);

  const handleLoadClick = () => {
    console.log('Load button clicked');
    try {
      onLoadFromOutlook();
    } catch (e) {
      console.error('Error calling onLoadFromOutlook:', e);
      setError(`Load error: ${e instanceof Error ? e.message : 'Unknown error'}`);
    }
  };

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
      // Check if Office.js is available
      if (typeof Office === 'undefined') {
        setError('Office.js not available. This add-in must run in Outlook.');
        setIsSending(false);
        return;
      }

      let draftCount = 0;

      for (let i = 0; i < recipients.length; i++) {
        const recipient = recipients[i];
        const subject = engine.render(template.subject, recipient);
        const body = engine.render(template.body, recipient);
        const toEmail = engine.render(toTemplate, recipient);

        if (!toEmail) {
          console.warn(`Recipient ${i + 1} has no email address from template "${toTemplate}"`);
          continue;
        }

        try {
          // Use Office.context.mailbox.displayNewMessageForm to create draft
          await new Promise<void>((resolve, reject) => {
            const officeMailbox = Office.context.mailbox as any;
            officeMailbox.displayNewMessageForm({
              toRecipients: [toEmail],
              subject: subject,
              htmlBody: body
            });
            
            // Give time for the form to open
            setTimeout(() => {
              draftCount++;
              resolve();
            }, 500);
          });
          
        } catch (err) {
          console.error(`Error creating draft for ${toEmail}:`, err);
          setError(`Error creating draft for ${toEmail}: ${err instanceof Error ? err.message : 'Unknown error'}`);
        }

        setProgress(Math.floor(((i + 1) / recipients.length) * 100));
        setStatus(`Processing ${i + 1} of ${recipients.length} for ${toEmail}`);
      }

      setStatus(`✓ Opened ${draftCount} new message forms!`);
      onSendComplete();
    } catch (err) {
      setError(`Error: ${err instanceof Error ? err.message : 'Unknown error'}`);
    } finally {
      setIsSending(false);
    }
  };

  if (recipients.length === 0) {
    return (
      <div className="send-pane">
        <div className="incomplete-notice">
          <p>⚠ Complete all previous steps before sending:</p>
          <ul>
            <li>Step 1: Load recipients from your data source (CSV, Excel, or manual entry)</li>
            <li>Step 2: Load your message from Outlook and create drafts</li>
          </ul>
        </div>
      </div>
    );
  }

  return (
    <div className="send-pane">
      <div className="send-header">
        <h2>Merge & Send</h2>
        <p>Load your composed message from Outlook, then create personalized drafts for each recipient</p>
        
        <button 
          className="load-button"
          onClick={handleLoadClick}
          disabled={isLoadingMessage}
        >
          {isLoadingMessage ? 'Loading...' : 'Load Message from Outlook Draft'}
        </button>
      </div>

      <div className="message-preview">
        <div className="preview-section">
          <h3>Subject</h3>
          <p className="preview-text">{template.subject || '(No subject loaded)'}</p>
        </div>
        <div className="preview-section">
          <h3>Body Preview</h3>
          <p className="preview-text">{template.body ? template.body.substring(0, 100) + '...' : '(No body loaded)'}</p>
        </div>
        <div className="preview-section">
          <h3>Recipients</h3>
          <p className="preview-text">{recipients.length} recipients will receive this message</p>
        </div>
      </div>

      {messageError && <div className="error-banner">{messageError}</div>}
      {error && <div className="error-banner">{error}</div>}

      <div className="send-controls">
        <button
          className="send-button"
          onClick={createDrafts}
          disabled={isSending || !template.subject || !template.body}
        >
          {isSending ? 'Creating Drafts...' : `Create ${recipients.length} Email Drafts`}
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
          <li>This opens a new message form for each recipient</li>
          <li>Each message will be pre-filled with personalized content</li>
          <li>You can review and edit before sending</li>
          <li>Save as draft or send immediately</li>
        </ul>
      </div>

      <div className="workflow-notes">
        <h4>Workflow:</h4>
        <ol>
          <li>Click "Create Email Drafts"</li>
          <li>New message windows will open for each recipient</li>
          <li>Review the personalized content</li>
          <li>Save as draft or click Send for each message</li>
        </ol>
      </div>
    </div>
  );
};
