/**
 * Audit Engine for identifying security vulnerabilities in M365 Data.
 */

export interface AuditResult {
  id: string;
  file: string;
  site: string;
  type: 'Anonymous Link' | 'Everyone Link' | 'Internal Sharing';
  severity: 'High' | 'Medium' | 'Low';
  bsnFound: boolean;
  linkUrl?: string;
}

/**
 * Scans content for Dutch BSN (Citizen Service Number).
 * A BSN is a 9-digit number that satisfies the '11-test'.
 */
export const scanForBSN = (content: string): boolean => {
  // Regex for 8 or 9 digits (BSNs can be 8 digits with leading zero, though usually 9)
  const bsnRegex = /\b\d{8,9}\b/g;
  const matches = content.match(bsnRegex);

  if (!matches) return false;

  for (const match of matches) {
    const bsn = match.padStart(9, '0');
    if (isValidBSN(bsn)) return true;
  }

  return false;
};

/**
 * Validates a BSN using the 11-test.
 */
const isValidBSN = (bsn: string): boolean => {
  if (bsn.length !== 9) return false;
  
  let sum = 0;
  for (let i = 0; i < 8; i++) {
    sum += parseInt(bsn[i]) * (9 - i);
  }
  sum += parseInt(bsn[8]) * -1;

  return sum % 11 === 0;
};

/**
 * Analyzes file permissions to identify over-privileged links.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export const analyzePermissions = (permissions: any[]): AuditResult['type'] | null => {
  for (const perm of permissions) {
    if (perm.link) {
      if (perm.link.type === 'anonymous' || perm.link.scope === 'anonymous') {
        return 'Anonymous Link';
      }
      if (perm.link.scope === 'organization' && perm.roles?.includes('read')) {
        // Technically "Everyone" in the org
        return 'Everyone Link';
      }
    }
    
    // Check for "Everyone" or "Everyone except external users" in identities
    if (perm.grantedToIdentities) {
      for (const identity of perm.grantedToIdentities) {
        const displayName = identity.user?.displayName || '';
        if (displayName.toLowerCase().includes('everyone')) {
          return 'Everyone Link';
        }
      }
    }
  }

  return null;
};

/**
 * Determines severity based on violation type and content findings.
 */
export const getSeverity = (type: AuditResult['type'], bsnFound: boolean): AuditResult['severity'] => {
  if (bsnFound && (type === 'Anonymous Link' || type === 'Everyone Link')) return 'High';
  if (type === 'Anonymous Link') return 'High';
  if (bsnFound) return 'Medium';
  if (type === 'Everyone Link') return 'Medium';
  return 'Low';
};
