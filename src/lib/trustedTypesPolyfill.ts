/**
 * TrustedTypes API Polyfill for Outlook
 * Prevents errors from libraries trying to create TrustedTypePolicy objects
 * that aren't in Outlook's CSP allowlist
 */

export function initTrustedTypesPolyfill(): void {
  // Only run if trustedTypes exists (browser supports the API)
  if (!window.trustedTypes) {
    return;
  }

  // Store the original createPolicy function
  const originalCreatePolicy = window.trustedTypes.createPolicy.bind(window.trustedTypes);

  // Override createPolicy to catch and suppress errors for disallowed policies
  window.trustedTypes.createPolicy = function(policyName: string, rules: any) {
    try {
      return originalCreatePolicy(policyName, rules);
    } catch (error) {
      // Silently handle CSP violations for TrustedTypePolicy creation
      // These policies are blocked by Outlook's CSP and we can't do anything about it
      if (error instanceof Error && error.message.includes('TrustedTypePolicy')) {
        // Return a no-op policy object to prevent downstream errors
        return {
          createHTML: (html: string) => html as any,
          createScript: (script: string) => script as any,
          createScriptURL: (url: string) => url as any,
          createURL: (url: string) => url as any
        };
      }
      // Re-throw other errors
      throw error;
    }
  } as any;
}
