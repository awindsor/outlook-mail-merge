/**
 * Global type definitions for Office JavaScript API
 */

declare namespace Office {
  function onReady(callback: (info: OfficeContextInfo) => void): void;
  
  namespace context {
    interface Mailbox {
      item?: any;
    }
    const mailbox: Mailbox;
  }
  const context: {
    mailbox: Mailbox;
  };
}

interface OfficeContextInfo {
  host: string;
  platform: string;
}

interface Mailbox {
  item?: any;
  displayNewMessageForm?: (data: string) => void;
}

declare const Office: typeof Office & {
  onReady: (callback: (info: OfficeContextInfo) => void) => void;
  context: {
    mailbox: Mailbox;
  };
};
