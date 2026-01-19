import React, { useState } from 'react';
import { DataSourceSelector } from './components/DataSourceSelector';
import { TemplateEditor } from './components/TemplateEditor';
import { PreviewPane } from './components/PreviewPane';
import { SendPane } from './components/SendPane';
import './App.css';

type Tab = 'template' | 'data' | 'preview' | 'send';

export const App: React.FC = () => {
  const [activeTab, setActiveTab] = useState<Tab>('template');
  const [template, setTemplate] = useState({
    subject: '',
    body: ''
  });
  const [recipients, setRecipients] = useState<any[]>([]);
  const [dataSource, setDataSource] = useState<'csv' | 'xlsx' | 'manual' | null>(null);

  return (
    <div className="app">
      <header className="app-header">
        <h1>Mail Merge for Outlook</h1>
        <p>Create personalized bulk emails from templates</p>
      </header>

      <div className="tabs">
        <button
          className={`tab-button ${activeTab === 'template' ? 'active' : ''}`}
          onClick={() => setActiveTab('template')}
        >
          1. Template
        </button>
        <button
          className={`tab-button ${activeTab === 'data' ? 'active' : ''}`}
          onClick={() => setActiveTab('data')}
        >
          2. Data Source
        </button>
        <button
          className={`tab-button ${activeTab === 'preview' ? 'active' : ''}`}
          onClick={() => setActiveTab('preview')}
        >
          3. Preview
        </button>
        <button
          className={`tab-button ${activeTab === 'send' ? 'active' : ''}`}
          onClick={() => setActiveTab('send')}
        >
          4. Send
        </button>
      </div>

      <div className="tab-content">
        {activeTab === 'template' && (
          <TemplateEditor
            template={template}
            onTemplateChange={setTemplate}
          />
        )}
        {activeTab === 'data' && (
          <DataSourceSelector
            onDataLoaded={setRecipients}
            onSourceChange={setDataSource}
          />
        )}
        {activeTab === 'preview' && (
          <PreviewPane
            template={template}
            recipients={recipients}
          />
        )}
        {activeTab === 'send' && (
          <SendPane
            template={template}
            recipients={recipients}
            onSendComplete={() => alert('Drafts created successfully!')}
          />
        )}
      </div>
    </div>
  );
};

export default App;
