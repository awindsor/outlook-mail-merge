import React, { useState } from 'react';
import { DataSourceSelector } from './components/DataSourceSelector';
import { ComposePane } from './components/ComposePane';
import { SendPane } from './components/SendPane';
import './App.css';

type Tab = 'data' | 'compose' | 'send';

export const App: React.FC = () => {
  const [activeTab, setActiveTab] = useState<Tab>('data');
  const [recipients, setRecipients] = useState<any[]>([]);
  const [dataSource, setDataSource] = useState<'csv' | 'xlsx' | 'manual' | null>(null);
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
        const subject = item.subject || '';
        setMessageSubject(subject);
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
        <p>Create personalized bulk emails from templates</p>
      </header>

      <div className="tabs">
        <button
          className={`tab-button ${activeTab === 'data' ? 'active' : ''}`}
          onClick={() => setActiveTab('data')}
        >
          1. Load Recipients
        </button>
        <button
          className={`tab-button ${activeTab === 'compose' ? 'active' : ''}`}
          onClick={() => setActiveTab('compose')}
          disabled={recipients.length === 0}
        >
          2. Message & Send
        </button>
        <button
          className={`tab-button ${activeTab === 'send' ? 'active' : ''}`}
          onClick={() => setActiveTab('send')}
          disabled={recipients.length === 0 || !messageSubject || !messageBody}
        >
          3. Merge & Send
        </button>
      </div>

      <div className="tab-content">
        {activeTab === 'data' && (
          <DataSourceSelector
            onDataLoaded={handleDataLoaded}
            onSourceChange={setDataSource}
            toTemplate={toTemplate}
            onToTemplateChange={setToTemplate}
            availableFields={availableFields}
          />
        )}
        {activeTab === 'compose' && (
          <ComposePane
            subject={messageSubject}
            body={messageBody}
            onSubjectChange={setMessageSubject}
            onBodyChange={setMessageBody}
            availableFields={availableFields}
            toTemplate={toTemplate}
            onToTemplateChange={setToTemplate}
            onLoadFromOutlook={loadMessageFromOutlook}
            isLoading={isLoadingMessage}
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
            onSendComplete={() => alert('Drafts created successfully!')}
          />
        )}
      </div>
    </div>
  );
};

export default App;
