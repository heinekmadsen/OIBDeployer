// Incremental consent sends the user to Entra ID and back, reloading the
// page. These helpers remember which wizard flow to return to.

const RESUME_KEY = 'oib-consent-resume';
const MAX_AGE_MS = 15 * 60 * 1000;

/**
 * @param state.deploymentType        'new' | 'existing' | 'device' | …
 * @param state.osTypes               OS buckets to reload
 * @param state.allowPrerelease
 * @param state.autoLoadStatus        reload assignment status on return
 * @param state.policyTypesSelection  New Deployment: the policy-types step result
 * @param state.enableGroups          turn assignment groups back on
 * @param state.groupSettings         { target, include, exclude } to restore
 */
export const saveConsentResume = ({
  deploymentType,
  osTypes,
  allowPrerelease,
  autoLoadStatus = false,
  policyTypesSelection = null,
  enableGroups = false,
  groupSettings = null,
}) => {
  try {
    sessionStorage.setItem(RESUME_KEY, JSON.stringify({
      deploymentType,
      osTypes,
      allowPrerelease: !!allowPrerelease,
      autoLoadStatus,
      policyTypesSelection,
      enableGroups,
      groupSettings,
      timestamp: Date.now(),
    }));
  } catch {
    // Storage unavailable — the user just re-runs the wizard.
  }
};

// Read and clear the saved state. Returns null when absent or stale.
export const takeConsentResume = () => {
  try {
    const resume = JSON.parse(sessionStorage.getItem(RESUME_KEY) || 'null');
    sessionStorage.removeItem(RESUME_KEY);
    if (!resume?.deploymentType || !resume?.osTypes?.length) return null;
    if (Date.now() - (resume.timestamp || 0) > MAX_AGE_MS) return null;
    return resume;
  } catch {
    return null;
  }
};
