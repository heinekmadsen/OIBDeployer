// Incremental consent sends the user to Entra ID and back, reloading the
// page. These helpers remember which wizard flow to return to.

const RESUME_KEY = 'oib-consent-resume';
const MAX_AGE_MS = 15 * 60 * 1000;

export const saveConsentResume = ({ deploymentType, osTypes, allowPrerelease, autoLoadStatus = false }) => {
  try {
    sessionStorage.setItem(RESUME_KEY, JSON.stringify({
      deploymentType,
      osTypes,
      allowPrerelease: !!allowPrerelease,
      autoLoadStatus,
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
