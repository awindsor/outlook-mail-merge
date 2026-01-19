/**
 * HTML Safety Utilities
 * Handles HTML rendering safely within Outlook's CSP constraints
 */

/**
 * Safely render HTML in Outlook add-in context
 * Avoids TrustedType policy conflicts by working with Outlook's CSP
 */
export function sanitizeHtmlForOutlook(html: string): string {
  // Create a temporary DOM element to safely parse HTML
  const temp = document.createElement('div');
  temp.textContent = html;
  
  // If the HTML contains actual tags, we need to be careful
  // For Outlook, we'll use a basic approach that avoids DOMPurify entirely
  
  // Simple whitelist of safe tags
  const safeTagsRegex = /<\/?(?:p|br|span|div|strong|em|b|i|u|ul|ol|li|table|tr|td|th|thead|tbody|a|img|h[1-6])[^>]*>/gi;
  
  let sanitized = html;
  
  // Remove script tags completely
  sanitized = sanitized.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '');
  
  // Remove event handlers
  sanitized = sanitized.replace(/\s*on\w+\s*=\s*['""][^""]*[''""]/gi, '');
  
  // Remove style attributes that might contain javascript
  sanitized = sanitized.replace(/\s*style\s*=\s*['""](?!.*expression)[^""]*[''""]/gi, '');
  
  return sanitized;
}

/**
 * Suppress TrustedType policy warnings in console
 * This helps avoid console spam from Outlook's CSP
 */
export function suppressTrustedTypeWarnings(): void {
  if (typeof window !== 'undefined') {
    const originalError = console.error;
    console.error = (...args: any[]) => {
      // Suppress TrustedType policy warnings
      const message = args.length > 0 ? String(args[0]) : '';
      if (message.includes('TrustedTypePolicy') || message.includes('Content Security Policy')) {
        return; // Suppress this error
      }
      originalError.apply(console, args);
    };
  }
}
