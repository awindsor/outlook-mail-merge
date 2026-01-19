import React, { ReactNode } from 'react';

interface Props {
  children: ReactNode;
}

interface State {
  hasError: boolean;
  error: Error | null;
}

export class ErrorBoundary extends React.Component<Props, State> {
  constructor(props: Props) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: Error): State {
    console.error('Error caught by boundary:', error);
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo) {
    console.error('Error details:', error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div style={{
          padding: '20px',
          backgroundColor: '#fee',
          border: '1px solid #f99',
          borderRadius: '4px',
          color: '#c00',
          fontFamily: 'monospace'
        }}>
          <h3>⚠️ Component Error</h3>
          <p>An unexpected error occurred:</p>
          <code style={{ display: 'block', whiteSpace: 'pre-wrap', marginTop: '10px', fontSize: '12px' }}>
            {this.state.error?.message}
          </code>
          <button onClick={() => window.location.reload()} style={{ marginTop: '10px' }}>
            Reload Add-in
          </button>
        </div>
      );
    }

    return this.props.children;
  }
}
