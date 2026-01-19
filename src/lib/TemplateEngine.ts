/**
 * TemplateEngine
 * Renders email templates with variable substitution
 * Supports Thunderbird Mail Merge-compatible syntax
 */

export class TemplateEngine {
  /**
   * Render a template string with data
   * Supports variable types:
   * - {{name}} - simple variable substitution
   * - {{name|if|then}} - conditional (equals)
   * - {{name|if|then|else}} - conditional with else
   * - {{name|*|if|then|else}} - conditional (includes)
   * - {{name|^|if|then|else}} - conditional (starts with)
   */
  render(template: string, data: Record<string, any>): string {
    if (!template) return '';

    let result = template;

    // Find all variables in format {{...}}
    const variableRegex = /\{\{([^}]+)\}\}/g;
    let match;

    while ((match = variableRegex.exec(template)) !== null) {
      const fullMatch = match[0]; // {{...}}
      const content = match[1]; // ...
      const replacement = this.evaluateVariable(content, data);
      result = result.replace(fullMatch, replacement);
    }

    return result;
  }

  private evaluateVariable(content: string, data: Record<string, any>): string {
    const parts = content.split('|');

    // Simple variable: {{name}}
    if (parts.length === 1) {
      return this.getValue(content, data);
    }

    // Conditional variable
    const fieldName = parts[0].trim();
    const fieldValue = this.getValue(fieldName, data);

    // {{name|if|then}}
    if (parts.length === 3) {
      const condition = parts[1].trim();
      const thenValue = parts[2].trim();
      return fieldValue === condition ? thenValue : '';
    }

    // {{name|if|then|else}}
    if (parts.length === 4) {
      const condition = parts[1].trim();
      const thenValue = parts[2].trim();
      const elseValue = parts[3].trim();
      return fieldValue === condition ? thenValue : elseValue;
    }

    // {{name|*|if|then|else}} (includes)
    if (parts.length === 5 && parts[1].trim() === '*') {
      const condition = parts[2].trim();
      const thenValue = parts[3].trim();
      const elseValue = parts[4].trim();
      return String(fieldValue).includes(condition) ? thenValue : elseValue;
    }

    // {{name|^|if|then|else}} (starts with)
    if (parts.length === 5 && parts[1].trim() === '^') {
      const condition = parts[2].trim();
      const thenValue = parts[3].trim();
      const elseValue = parts[4].trim();
      return String(fieldValue).startsWith(condition) ? thenValue : elseValue;
    }

    return content;
  }

  private getValue(fieldName: string, data: Record<string, any>): string {
    // Handle nested fields like "Address.City"
    const keys = fieldName.split('.');
    let value: any = data;

    for (const key of keys) {
      if (value && typeof value === 'object') {
        value = value[key];
      } else {
        return '';
      }
    }

    // Handle undefined/null
    if (value === null || value === undefined) {
      return '';
    }

    return String(value);
  }
}
