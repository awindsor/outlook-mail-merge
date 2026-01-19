import React, { useState } from 'react';
import '../styles/SendPane.css';

// Inline template rendering to avoid any dependencies that might cause CSP issues
const renderTemplate = (template: string, data: any): string => {
  let result = template;
  Object.keys(data).forEach(key => {
    const regex = new RegExp(`{{\\s*${key}\\s*}}`, 'g');
    result = result.replace(regex, String(data[key] || ''));
  });
  return result;
};

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
  // Ensure template has default values and is actually strings
  const safeTemplate = {
    subject: typeof template?.subject === 'string' ? template.subject : '',
    body: typeof template?.body === 'string' ? template.body : ''
  };
  
  const [isSending, setIsSending] = useState(false);
  const [progress, setProgress] = useState(0);
  const [status, setStatus] = useState<string>('');
  const [error, setError] = useState<string>('');
  const [previewIndex, setPreviewIndex] = useState(0);

  React.useEffect(() => {
    console.log('SendPane rendered with:', {
      subject: safeTemplate.subject,
      bodyLength: safeTemplate.body.length,
      recipients: recipients?.length || 0,
      toTemplate,
      messageError,
      isLoadingMessage
    });
  }, [safeTemplate.subject, safeTemplate.body, recipients?.length, toTemplate, messageError, isLoadingMessage]);

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
    if (!safeTemplate.subject.trim()) {
      setError('Subject line is required');
      return false;
    }
    if (!safeTemplate.body.trim()) {
      setError('Email body is required');
      return false;
    }
    if (!recipients || recipients.length === 0) {
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
      const errors: string[] = [];

      for (let i = 0; i < recipients.length; i++) {
        const recipient = recipients[i];
        const subject = renderTemplate(safeTemplate.subject, recipient);
        const body = renderTemplate(safeTemplate.body, recipient);
        const toEmail = renderTemplate(toTemplate, recipient);

        if (!toEmail) {
          console.warn(`Recipient ${i + 1} has no email address from template "${toTemplate}"`);
          continue;
        }

        try {
          console.log(`Creating draft for ${toEmail}...`);
          
          // Get access token for REST API
          const officeMailbox = Office.context.mailbox as any;
          
          await new Promise<void>((resolve, reject) => {
            officeMailbox.getCallbackTokenAsync({ isRest: true }, (result: any) => {
              if (result.status === 'succeeded') {
                const token = result.value;
                
                // Create draft using Outlook REST API
                const restUrl = officeMailbox.restUrl || 'https://outlook.office.com/api/v2.0';
                const messagePayload = {
                  Subject: subject,
                  Body: {
                    ContentType: 'Text',
                    Content: body
                  },
                  ToRecipients: [
                    {
                      EmailAddress: {
                        Address: toEmail
                      }
                    }
                  ]
                };

                fetch(`${restUrl}/me/messages`, {
                  method: 'POST',
                  headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                  },
                  body: JSON.stringify(messagePayload)
                })
                  .then(response => {
                    if (response.ok) {
                      draftCount++;
                      console.log(`Created draft ${draftCount} for ${toEmail}`);
                      resolve();
                    } else {
                      return response.text().then(text => {
                        throw new Error(`API error: ${response.status} - ${text}`);
                      });
                    }
                  })
                  .catch(err => {
                    console.error(`Error creating draft for ${toEmail}:`, err);
                    const errorMsg = err instanceof Error ? err.message : String(err);
                    errors.push(`${toEmail}: ${errorMsg}`);
                    resolve(); // Continue with next recipient
                  });
              } else {
                reject(new Error('Failed to get access token: ' + result.error?.message));
              }
            });
          });
          
        } catch (err) {
          console.error(`Error creating draft for ${toEmail}:`, err);
          const errorMsg = err instanceof Error ? err.message : String(err);
          errors.push(`${toEmail}: ${errorMsg}`);
        }

        setProgress(Math.floor(((i + 1) / recipients.length) * 100));
        setStatus(`Created ${draftCount} of ${recipients.length} drafts (processing ${i + 1}/${recipients.length})...`);
      }

      if (errors.length > 0) {
        setError(`Created ${draftCount} drafts with ${errors.length} errors. Check console for details.`);
        console.error('Errors:', errors);
      } else {
        setStatus(`✓ Successfully created ${draftCount} draft messages in your Drafts folder!`);
      }
      onSendComplete();
    } catch (err) {
      console.error('Error in createDrafts:', err);
      setError(`Error: ${err instanceof Error ? err.message : 'Unknown error'}`);
    } finally {
      setIsSending(false);
    }
  };

  if (!recipients || recipients.length === 0) {
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
        {recipients && recipients.length > 0 ? (
          <>
            <div className="preview-navigation">
              <button
                onClick={() => setPreviewIndex(Math.max(0, previewIndex - 1))}
                disabled={previewIndex === 0}
                className="nav-button"
              >
                ← Previous
              </button>
              <span className="preview-counter">
                Preview {previewIndex + 1} of {recipients.length}
              </span>
              <button
                onClick={() => setPreviewIndex(Math.min(recipients.length - 1, previewIndex + 1))}
                disabled={previewIndex === recipients.length - 1}
                className="nav-button"
              >
                Next →
              </button>
            </div>
            <div className="preview-section">
              <h3>To:</h3>
              <p className="preview-text">{renderTemplate(toTemplate, recipients[previewIndex])}</p>
            </div>
            <div className="preview-section">
              <h3>Subject</h3>
              <p className="preview-text">{renderTemplate(safeTemplate.subject, recipients[previewIndex]) || '(No subject loaded)'}</p>
            </div>
            <div className="preview-section">
              <h3>Body Preview</h3>
              <div className="preview-text" style={{ whiteSpace: 'pre-wrap' }}>
                {renderTemplate(safeTemplate.body, recipients[previewIndex]) || '(No body loaded)'}
              </div>
            </div>
          </>
        ) : (
          <div className="preview-section">
            <p className="preview-text">Load recipients first to see preview</p>
          </div>
        )}
      </div>

      {messageError && <div className="error-banner">{typeof messageError === 'string' ? messageError : JSON.stringify(messageError)}</div>}
      {error && <div className="error-banner">{error}</div>}

      <div className="send-controls">
        <button
          className="send-button"
          onClick={createDrafts}
          disabled={isSending || !safeTemplate.subject || !safeTemplate.body}
        >
          {isSending ? 'Creating Drafts...' : `Create ${recipients?.length || 0} Email Drafts`}
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
          <li>Drafts will be saved to your Drafts folder</li>
          <li>Each message will be personalized with recipient data</li>
          <li>Review and edit drafts before sending</li>
          <li>Send drafts manually when ready</li>
        </ul>
      </div>

      <div className="workflow-notes">
        <h4>Workflow:</h4>
        <ol>
          <li>Click "Create Email Drafts"</li>
          <li>Drafts will be created in your Drafts folder</li>
          <li>Open Drafts folder to review personalized messages</li>
          <li>Send each draft individually or use Outlook's batch send</li>
        </ol>
      </div>
    </div>
  );
};
