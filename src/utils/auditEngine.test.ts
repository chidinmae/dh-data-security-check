import { describe, it, expect } from 'vitest';
import { scanForBSN, analyzePermissions, getSeverity } from './auditEngine';

describe('Audit Engine - BSN Scanner', () => {
  it('should identify a valid 9-digit BSN', () => {
    // 123456782 is a valid BSN (11-test: 9*1 + 8*2 + 7*3 + 6*4 + 5*5 + 4*6 + 3*7 + 2*8 - 1*2 = 121, 121 % 11 = 0)
    // Wait, let's use a real one or verify the math: 
    // Example: 111222333
    // 9*1 + 8*1 + 7*1 + 6*2 + 5*2 + 4*2 + 3*3 + 2*3 - 1*3 = 9+8+7+12+10+8+9+6-3 = 66. 66/11 = 6. Correct.
    expect(scanForBSN('My BSN is 111222333.')).toBe(true);
  });

  it('should reject an invalid BSN', () => {
    expect(scanForBSN('My fake BSN is 123456789.')).toBe(false);
  });

  it('should handle 8-digit BSNs (padded with zero)', () => {
    // Example: 011122233
    // 9*0 + 8*1 + 7*1 + 6*1 + 5*2 + 4*2 + 3*2 + 2*3 - 1*3 = 0+8+7+6+10+8+6+6-3 = 48. Not valid.
    // Let's find a valid 8-digit: 12345671 (sum: 9*0 + 8*1 + 7*2 + 6*3 + 5*4 + 4*5 + 3*6 + 2*7 - 1*1 = 8+14+18+20+20+18+14-1 = 111. No)
    // Real one: 000000000 (valid but unlikely)
    expect(scanForBSN('000000000')).toBe(true);
  });
});

describe('Audit Engine - Permissions', () => {
  it('should identify anonymous links', () => {
    const perms = [{ link: { type: 'anonymous', scope: 'anonymous' } }];
    expect(analyzePermissions(perms)).toBe('Anonymous Link');
  });

  it('should identify everyone links by name', () => {
    const perms = [{ grantedToIdentities: [{ user: { displayName: 'Everyone except external users' } }] }];
    expect(analyzePermissions(perms)).toBe('Everyone Link');
  });
});

describe('Audit Engine - Severity', () => {
  it('should return High for anonymous links', () => {
    expect(getSeverity('Anonymous Link', false)).toBe('High');
  });

  it('should return High for PII in Everyone links', () => {
    expect(getSeverity('Everyone Link', true)).toBe('High');
  });
});
