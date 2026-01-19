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
      // Set default "To" template with email field
      const emailFieldName = fields.find(f => 
        f.toLowerCase().includes('email') || f.toLowerCase() === 'to'
      ) || fields[0];
      setToTemplate(`{{${emailFieldName}}}`);
    }
  };

  const loadMessageFromOutlook = async () => {
    setIsLoadingMessage(true);
    try {
      if (typeof Office !== 'undefined' && Office.context?.mailbox?.item) {
        const item = Office.context.mailbox.item;
        const subject = item.subject || '';
        const body = item.body?.getAsync?.(Office.CoercionType.Html, (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            setMessageBody(result.value);
          }
        });
        setMessageSubject(subject);
      }Load Recipients
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
          3. Merge &tton>
        <button
          className={`tab-button ${activeTab === 'template' ? 'active' : ''}`}
          onClick={() => setActiveTab('template')}
          disabled={recipients.length === 0}
        >
          2. Template
        </button>
        <button
          className={`tab-button ${activeTab === 'preview' ? 'active' : ''}`}
          onClick={() => setActiveTab('preview')}
          disabled={recipients.length === 0}
        >
          3. Preview
        </button>
        <button
          className={`tab-button ${activeTab === 'send' ? 'active' : ''}`}
          onClick={() => setActiveTab('send')}
          disabled={recipients.length === 0}
        >
          4. Send
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
        {activeTab === 'template' && (
          <TemplateEditor
            template={template}
            onTemplateChange={setTemplate}
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
            }
  );
};

export default App;
