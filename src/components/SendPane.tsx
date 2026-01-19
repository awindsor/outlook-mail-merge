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

  const exportMessages = () => {
    if (!recipients || recipients.length === 0) {
      setError('No recipients loaded');
      return;
    }

    let exportText = 'PERSONALIZED EMAIL MESSAGES\n';
    exportText += '='.repeat(50) + '\n\n';

    recipients.forEach((recipient, index) => {
      const subject = renderTemplate(safeTemplate.subject, recipient);
      const body = renderTemplate(safeTemplate.body, recipient);
      const toEmail = renderTemplate(toTemplate, recipient);

      exportText += `MESSAGE ${index + 1} of ${recipients.length}\n`;
      exportText += '-'.repeat(50) + '\n';
      exportText += `To: ${toEmail}\n`;
      exportText += `Subject: ${subject}\n\n`;
      exportText += `${body}\n\n`;
      exportText += '='.repeat(50) + '\n\n';
    });

    // Copy to clipboard
    navigator.clipboard.writeText(exportText).then(() => {
      setStatus(`✓ Copied ${recipients.length} personalized messages to clipboard! You can paste them into a text editor.`);
    }).catch(() => {
      // Fallback: download as file
      const blob = new Blob([exportText], { type: 'text/plain' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'mail-merge-messages.txt';
      a.click();
      URL.revokeObjectURL(url);
      setStatus(`✓ Downloaded ${recipients.length} personalized messages as text file!`);
    });
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

      const officeMailbox = Office.context.mailbox as any;
      
      // Check if we have EWS support (desktop) or display form (either)
      const hasEws = typeof officeMailbox.makeEwsRequestAsync === 'function';
      const hasDisplayForm = typeof officeMailbox.displayNewMessageForm === 'function';
      
      if (!hasEws && !hasDisplayForm) {
        setError('This Outlook version has limited API support. Please use the "Export Messages" button below to copy/paste messages manually.');
        setIsSending(false);
        return;
      }

      let draftCount = 0;
      const errors: string[] = [];
      let ewsFailedOnce = false;

      // Try EWS first if available (desktop), but fall back to displayNewMessageForm if it fails
      let useEws = hasEws;

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
          console.log(`Creating message for ${toEmail}...`);
          
          if (useEws) {
            // Use EWS for desktop - creates drafts directly in Drafts folder
            const ewsRequest = `<?xml version="1.0" encoding="utf-8"?>
              <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                            xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                            xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                            xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Header>
                  <t:RequestServerVersion Version="Exchange2013" />
                </soap:Header>
                <soap:Body>
                  <m:CreateItem MessageDisposition="SaveOnly">
                    <m:Items>
                      <t:Message>
                        <t:Subject>${subject.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</t:Subject>
                        <t:Body BodyType="Text">${body.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</t:Body>
                        <t:ToRecipients>
                          <t:Mailbox>
                            <t:EmailAddress>${toEmail}</t:EmailAddress>
                          </t:Mailbox>
                        </t:ToRecipients>
                      </t:Message>
                    </m:Items>
                  </m:CreateItem>
                </soap:Body>
              </soap:Envelope>`;

            await new Promise<void>((resolve) => {
              officeMailbox.makeEwsRequestAsync(ewsRequest, (result: any) => {
                if (result.status === 'succeeded') {
                  draftCount++;
                  console.log(`Created draft ${draftCount} for ${toEmail}`);
                } else {
                  // EWS failed - log details and switch to fallback
                  const errorMsg = result.error?.message || 'EWS request failed';
                  console.error('EWS error:', errorMsg, 'Full result:', result);
                  
                  // On first EWS failure, switch to displayNewMessageForm for remaining messages
                  if (!ewsFailedOnce && hasDisplayForm) {
                    ewsFailedOnce = true;
                    useEws = false;
                    console.log('Switching to displayNewMessageForm fallback for remaining messages');
                  }
                  
                  errors.push(`${toEmail}: ${errorMsg}`);
                }
                resolve();
              });
            });
          } else if (hasDisplayForm) {
            // Fallback to displayNewMessageForm (opens windows)
            officeMailbox.displayNewMessageForm({
              toRecipients: [toEmail],
              subject: subject,
              htmlBody: body
            });
            
            draftCount++;
            console.log(`Opened message form ${draftCount} for ${toEmail}`);
            
            // Small delay between opening windows
            await new Promise(resolve => setTimeout(resolve, 500));
          } else {
            errors.push(`${toEmail}: No API method available to create message`);
          }
          
        } catch (err) {
          console.error(`Error creating message for ${toEmail}:`, err);
          const errorMsg = err instanceof Error ? err.message : String(err);
          errors.push(`${toEmail}: ${errorMsg}`);
        }

        setProgress(Math.floor(((i + 1) / recipients.length) * 100));
        setStatus(`Created ${draftCount} of ${recipients.length} drafts (processing ${i + 1}/${recipients.length})...`);
      }

      if (errors.length > 0) {
        const errorDetails = errors.slice(0, 3).join('\n'); // Show first 3 errors
        const moreErrors = errors.length > 3 ? `\n...and ${errors.length - 3} more errors` : '';
        const fallbackMsg = ewsFailedOnce && hasDisplayForm ? '\n\nNote: EWS failed, automatically switched to opening message windows.' : '';
        setError(`${useEws ? 'Created' : 'Opened'} ${draftCount} ${useEws ? 'drafts' : 'message forms'} with ${errors.length} errors:\n${errorDetails}${moreErrors}${fallbackMsg}`);
        console.error('All errors:', errors);
      } else {
        setStatus(useEws 
          ? `✓ Successfully created ${draftCount} draft messages in your Drafts folder!`
          : `✓ Opened ${draftCount} message forms! You can now save them as drafts or send them.`
        );
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
        
        <button
          className="send-button"
          onClick={exportMessages}
          disabled={!safeTemplate.subject || !safeTemplate.body || !recipients || recipients.length === 0}
          style={{ marginTop: '10px', backgroundColor: '#0078d4' }}
        >
          Export Messages to Clipboard
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
          <li>A new message window will open for each recipient</li>
          <li>Each message will be pre-filled with personalized content</li>
          <li>Review each message before sending</li>
          <li>Save as draft or send immediately from each window</li>
        </ul>
      </div>

      <div className="workflow-notes">
        <h4>Workflow:</h4>
        <ol>
          <li>Click "Create Email Drafts"</li>
          <li>New compose windows will open for each recipient</li>
          <li>Review the personalized message in each window</li>
          <li>Click Send in each window or save as draft</li>
        </ol>
      </div>
    </div>
  );
};
