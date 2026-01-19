/**
 * HTML Safety Utilities
 * Handles HTML rendering safely within Outlook's CSP constraints
 */

/**
 * Safely render HTML in Outlook add-in context
 * Avoids TrustedType policy conflicts by working with Outlook's CSP
 */
export function sanitizeHtmlForOutlook(html: string): string {
  if (!html) return '';
  
  let sanitized = html;
  
  // Remove script tags completely
  sanitized = sanitized.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '');
  
  // Remove event handlers
  sanitized = sanitized.replace(/\s*on\w+\s*=\s*['""][^""]*[''""]/gi, '');
  
  // Remove javascript: URLs
  sanitized = sanitized.replace(/javascript:/gi, '');
  
  return sanitized;
}
