import React, { useState } from 'react';
import { DataSourceSelector } from './components/DataSourceSelector';
import { SendPane } from './components/SendPane';
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

  const loadMessageFromOutlook = async () => {
    setIsLoadingMessage(true);
    setMessageError('');
    try {
      const officeObj = (window as any).Office;
      if (!officeObj) {
        throw new Error('Office.js not available');
      }
      
      const context = officeObj.context;
      if (!context) {
        throw new Error('Office context not available');
      }

      const mailbox = context.mailbox;
      if (!mailbox) {
        throw new Error('Mailbox not available');
      }

      const item = mailbox.item;
      if (!item) {
        throw new Error('No item in mailbox - are you in a message compose window?');
      }

      console.log('Loading message from Outlook item...');
      
      // Get subject
      try {
        const subject = item.subject || '';
        console.log('Subject:', subject);
        setMessageSubject(subject);
      } catch (e) {
        console.error('Error getting subject:', e);
      }
      
      // Get To field
      try {
        if (item.to && Array.isArray(item.to) && item.to.length > 0) {
          const toAddresses = item.to.map((recipient: any) => recipient.emailAddress).join(', ');
          console.log('To addresses:', toAddresses);
          setToTemplate(toAddresses);
        } else {
          console.warn('No To field found');
        }
      } catch (e) {
        console.error('Error getting To field:', e);
      }
      
      // Get body - this requires a callback
      try {
        if (item.body && typeof item.body.getAsync === 'function') {
          item.body.getAsync('html', (result: any) => {
            if (result.status === 'succeeded' && result.value) {
              console.log('Body loaded successfully');
              setMessageBody(result.value);
            } else if (result.status === 'failed') {
              console.error('Failed to get body:', result.error);
              setMessageError(`Failed to load body: ${result.error?.message || 'Unknown error'}`);
            }
          });
        } else {
          console.warn('Body getAsync not available');
        }
      } catch (e) {
        console.error('Error setting up body getAsync:', e);
      }
      
      console.log('Message load initiated');
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      console.error('Error loading message from Outlook:', err);
      setMessageError(`Error: ${errorMsg}`);
    } finally {
      setIsLoadingMessage(false);
    }
  };

  return (
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
  );
};

export default App;
