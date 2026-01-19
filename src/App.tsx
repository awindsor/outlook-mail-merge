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
    
    setIsLoadingMessage(true);
    setMessageError('');
    
    try {
      const officeObj = (window as any).Office;
      if (!officeObj?.context?.mailbox?.item) {
        throw new Error('Office context not available - are you in a message compose window?');
      }

      const item = officeObj.context.mailbox.item;
      console.log('Loading message from Outlook item...');
      
      // Get subject synchronously
      const subject = item.subject || '';
      console.log('Subject:', subject);
      
      // Get To field synchronously
      let toAddresses = '';
      if (item.to && Array.isArray(item.to) && item.to.length > 0) {
        toAddresses = item.to.map((r: any) => r.emailAddress).join(', ');
        console.log('To addresses:', toAddresses);
      }
      
      // Update state once with sync data
      if (isMountedRef.current) {
        setMessageSubject(subject);
        if (toAddresses) setToTemplate(toAddresses);
      }
      
      // Get body asynchronously
      if (item.body && typeof item.body.getAsync === 'function') {
        item.body.getAsync('text', (result: any) => {
          if (!isMountedRef.current) return;
          
          try {
            if (result?.status === 'succeeded' && result.value) {
              console.log('Body loaded, length:', result.value.length);
              if (isMountedRef.current) {
                setMessageBody(result.value);
              }
            } else if (result?.status === 'failed') {
              console.error('Failed to get body:', result.error);
              if (isMountedRef.current) {
                setMessageError(`Failed to load body: ${result.error?.message || 'Unknown'}`);
              }
            }
          } catch (e) {
            console.error('Error in body callback:', e);
            if (isMountedRef.current) {
              setMessageError(`Error processing body: ${e instanceof Error ? e.message : 'Unknown'}`);
            }
          } finally {
            if (isMountedRef.current) {
              setIsLoadingMessage(false);
            }
          }
        });
      } else {
        console.warn('Body getAsync not available');
        if (isMountedRef.current) {
          setIsLoadingMessage(false);
        }
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error('Error loading message:', err);
      if (isMountedRef.current) {
        setMessageError(`Error: ${msg}`);
        setIsLoadingMessage(false);
      }
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
            onSendComplete={() => alert('Drafts created successfully!')}
          />
        )}
      </div>
      </div>
    </ErrorBoundary>
  );
};

export default App;
