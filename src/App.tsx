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
    try {
      const officeObj = (window as any).Office;
      if (officeObj && officeObj.context && officeObj.context.mailbox && officeObj.context.mailbox.item) {
        const item = officeObj.context.mailbox.item;
        
        // Get subject
        const subject = item.subject || '';
        setMessageSubject(subject);
        
        // Get To field
        if (item.to && Array.isArray(item.to) && item.to.length > 0) {
          const toAddresses = item.to.map((recipient: any) => recipient.emailAddress).join(', ');
          setToTemplate(toAddresses);
        }
        
        // Get body
        if (item.body && typeof item.body.getAsync === 'function') {
          item.body.getAsync('html', (result: any) => {
            if (result && result.value) {
              setMessageBody(result.value);
            }
          });
        }
      }
    } catch (err) {
      console.error('Error loading message from Outlook:', err);
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
            onSendComplete={() => alert('Drafts created successfully!')}
          />
        )}
      </div>
    </div>
  );
};

export default App;
