import React, { useState, useCallback, useRef, useEffect } from 'react';
import { DataSourceSelector } from './components/DataSourceSelector';
import { SendPane } from './components/SendPane';
import { ErrorBoundary } from './components/ErrorBoundary';
import './App.css';

type Tab = 'data' | 'send';

export const App: React.FC = () => {
  const [activeTab, setActiveTab] = useState<Tab>('data');
  const [recipients, setRecipients] = useState<any[]>([]);
  const [availableFields, setAvailableFields] = useState<string[]>([]);
  const [toTemplate, setToTemplate] = useState<string>('');
  const [messageSubject, setMessageSubject] = useState<string>('');
  const [messageBody, setMessageBody] = useState<string>('');
  const [isLoadingMessage, setIsLoadingMessage] = useState(false);
  const [messageError, setMessageError] = useState<string>('');
  const isMountedRef = useRef(true);

  useEffect(() => {
    return () => {
      isMountedRef.current = false;
    };
  }, []);

  const handleDataLoaded = (data: any[]) => {
    setRecipients(data);
    if (data.length > 0) {
      const fields = Object.keys(data[0]);
      setAvailableFields(fields);
      const emailFieldName = fields.find(f => 
        f.toLowerCase().includes('email') || f.toLowerCase() === 'to'
      ) || fields[0];
      setToTemplate(`{{${emailFieldName}}}`);
    }
  };

  const loadMessageFromOutlook = useCallback(() => {
    if (!isMountedRef.current) return;

    // Start loading state on the next tick to avoid React rendering conflicts
    setTimeout(() => {
      if (isMountedRef.current) {
        setIsLoadingMessage(true);
        setMessageError('');
      }
    }, 0);
    
    try {
      const officeObj = (window as any).Office;
      if (!officeObj?.context?.mailbox?.item) {
        throw new Error('Office context not available - are you in a message compose window?');
      }

      const item = officeObj.context.mailbox.item;
      console.log('Loading message from Outlook item...');
      
      // Get subject - in compose mode, subject might be an object with getAsync
      if (item.subject && typeof item.subject.getAsync === 'function') {
        item.subject.getAsync((result: any) => {
          if (!isMountedRef.current) return;
          if (result?.status === 'succeeded' && result.value) {
            setTimeout(() => {
              if (isMountedRef.current) setMessageSubject(String(result.value));
            }, 0);
          }
        });
      } else {
        // Fallback for read mode or if subject is already a string
        const subject = typeof item.subject === 'string' ? item.subject : '';
        setTimeout(() => {
          if (!isMountedRef.current) return;
          setMessageSubject(subject);
        }, 0);
      }
      
      // Get recipients
      let toAddresses = '';
      if (item.to && Array.isArray(item.to) && item.to.length > 0) {
        toAddresses = item.to.map((r: any) => r.emailAddress).join(', ');
      }
      if (toAddresses) {
        setTimeout(() => {
          if (!isMountedRef.current) return;
          setToTemplate(toAddresses);
        }, 0);
      }
      
      // Get body asynchronously
      if (item.body && typeof item.body.getAsync === 'function') {
        item.body.getAsync('text', (result: any) => {
          if (!isMountedRef.current) return;
          try {
            if (result?.status === 'succeeded' && result.value) {
              const value = result.value;
              setTimeout(() => {
                if (isMountedRef.current) setMessageBody(value);
              }, 0);
            } else if (result?.status === 'failed') {
              const msg = result.error?.message || 'Unknown';
              setTimeout(() => {
                if (isMountedRef.current) setMessageError(`Failed to load body: ${msg}`);
              }, 0);
            }
          } catch (e) {
            const msg = e instanceof Error ? e.message : 'Unknown';
            setTimeout(() => {
              if (isMountedRef.current) setMessageError(`Error processing body: ${msg}`);
            }, 0);
          } finally {
            setTimeout(() => {
              if (isMountedRef.current) setIsLoadingMessage(false);
            }, 0);
          }
        });
      } else {
        setTimeout(() => {
          if (isMountedRef.current) setIsLoadingMessage(false);
        }, 0);
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      setTimeout(() => {
        if (isMountedRef.current) {
          setMessageError(`Error: ${msg}`);
          setIsLoadingMessage(false);
        }
      }, 0);
    }
  }, []);

  return (
    <ErrorBoundary>
      <div className="app">
      <header className="app-header">
        <h1>Mail Merge for Outlook</h1>
        <p>Send personalized emails to your recipients from an Outlook draft</p>
      </header>

      <div className="tabs">
        <button
          className={`tab-button ${activeTab === 'data' ? 'active' : ''}`}
          onClick={() => setActiveTab('data')}
        >
          1. Load Recipients
        </button>
        <button
          className={`tab-button ${activeTab === 'send' ? 'active' : ''}`}
          onClick={() => setActiveTab('send')}
          disabled={recipients.length === 0}
        >
          2. Merge & Send
        </button>
      </div>

      <div className="tab-content">
        {activeTab === 'data' && (
          <DataSourceSelector
            onDataLoaded={handleDataLoaded}
            onSourceChange={() => {}}
            toTemplate={toTemplate}
            onToTemplateChange={setToTemplate}
            availableFields={availableFields}
          />
        )}
        {activeTab === 'send' && (
          <SendPane
            template={{
              subject: messageSubject,
              body: messageBody
            }}
            recipients={recipients}
            toTemplate={toTemplate}
            onLoadFromOutlook={loadMessageFromOutlook}
            isLoadingMessage={isLoadingMessage}
            messageError={messageError}
            onSendComplete={() => {
              console.log('Drafts created successfully!');
              // Stay on the send tab to show completion message
            }}
          />
        )}
      </div>
      </div>
    </ErrorBoundary>
  );
};

export default App;
