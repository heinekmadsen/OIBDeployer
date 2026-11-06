// Policy assignment & device-status overview for the Existing Deployment
// dashboard. Pure functions only — Graph calls live in graph-api.js.

import { mapReportPolicyStatus, DEVICE_STATUS } from './device-validation.js';

const num = (value) => {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
};

const withTotals = (counts) => {
  const total = counts.succeeded + counts.error + counts.conflict + counts.nonCompliant +
    counts.pending + counts.notApplicable;
  return {
    ...counts,
    total,
    issues: counts.error + counts.conflict + counts.nonCompliant,
  };
};

const EMPTY_COUNTS = withTotals({
  succeeded: 0, error: 0, conflict: 0, nonCompliant: 0, pending: 0, notApplicable: 0,
});

/**
 * Normalise a raw summary from graphAPI.getPolicyStatusSummaries() into
 * { succeeded, error, conflict, nonCompliant, pending, notApplicable, total, issues }.
 * Returns null when no per-device report exists for the policy type, and
 * { loadError } when the report failed ("error" is the error-device count).
 */
export const normalisePolicyStatusSummary = (raw) => {
  if (!raw) return null;
  if (raw.error) return { loadError: raw.error };

  if (raw.source === 'configurationReport') {
    // No row: the policy exists but no device has reported on it yet.
    if (!raw.row) return EMPTY_COUNTS;
    const row = raw.row;
    return withTotals({
      succeeded: num(row.NumberOfCompliantDevices),
      error: num(row.NumberOfErrorDevices),
      conflict: num(row.NumberOfConflictDevices),
      nonCompliant: num(row.NumberOfNonCompliantDevices),
      pending: num(row.NumberOfInProgressDevices),
      notApplicable: num(row.NumberOfNotApplicableDevices),
    });
  }

  if (raw.source === 'complianceOverview') {
    const o = raw.overview || {};
    return withTotals({
      succeeded: num(o.successCount),
      error: num(o.errorCount),
      conflict: num(o.conflictCount),
      nonCompliant: num(o.failedCount),
      pending: num(o.pendingCount),
      notApplicable: num(o.notApplicableCount) + num(o.notApplicablePlatformCount),
    });
  }

  return null;
};

// Compliance report PolicyStatus_loc text → normalised device status.
export const mapComplianceReportStatus = (text) => {
  switch ((text || '').toLowerCase().replace(/[^a-z]/g, '')) {
    case 'compliant':
    case 'remediated':     return 'succeeded';
    case 'notcompliant':
    case 'noncompliant':
    case 'ingraceperiod':  return 'nonCompliant';
    case 'error':          return 'error';
    case 'conflict':       return 'conflict';
    case 'notapplicable':  return 'notApplicable';
    default:               return 'pending';
  }
};

/**
 * Normalise rows from graphAPI.getPolicyDeviceStatuses() into
 * [{ deviceName, upn, status, rawStatus, lastReported, deviceId }],
 * worst status first.
 */
export const normalisePolicyDeviceRows = (result) => {
  const rows = (result?.rows || []).map(row => {
    if (result.source === 'complianceReport') {
      return {
        deviceName: row.DeviceName || '—',
        upn: row.UPN || '',
        status: mapComplianceReportStatus(row.PolicyStatus_loc),
        rawStatus: row.PolicyStatus_loc || row.PolicyStatus,
        lastReported: row.LastContact || null,
        deviceId: row.DeviceId,
      };
    }
    return {
      deviceName: row.DeviceName || '—',
      upn: row.UPN || '',
      status: mapReportPolicyStatus(row.PolicyStatus),
      rawStatus: row.PolicyStatus,
      lastReported: row.PspdpuLastModifiedTimeUtc || null,
      deviceId: row.IntuneDeviceId,
    };
  });

  return rows.sort((a, b) =>
    DEVICE_STATUS[a.status].severity - DEVICE_STATUS[b.status].severity ||
    a.deviceName.localeCompare(b.deviceName)
  );
};

// Tenant policies matched to OIB policies in a ComparisonDashboard result.
export const getMatchedTenantPolicies = (comparisonData) => {
  const byId = new Map();
  Object.values(comparisonData?.byOS || {}).forEach(osData => {
    [...(osData.current || []), ...(osData.outdated || []), ...(osData.newer || [])]
      .forEach(entry => {
        if (entry.existingPolicy?.id) byId.set(entry.existingPolicy.id, entry.existingPolicy);
      });
    (osData.duplicates || []).forEach(entry =>
      (entry.matchedPolicies || []).forEach(p => { if (p?.id) byId.set(p.id, p); })
    );
  });
  return [...byId.values()];
};

// Every group ID referenced by a set of summarizeAssignments() results.
export const collectGroupIds = (assignmentSummaries) => {
  const ids = new Set();
  assignmentSummaries.forEach(summary => {
    if (!summary || summary.error) return;
    summary.includeGroupIds.forEach(id => id && ids.add(id.toLowerCase()));
    summary.excludeGroupIds.forEach(id => id && ids.add(id.toLowerCase()));
  });
  return [...ids];
};

// Display name for a group ID given the (optional) resolved names map.
export const groupLabel = (groupId, groupNames) => {
  const key = (groupId || '').toLowerCase();
  if (groupNames?.has(key)) return groupNames.get(key) ?? 'Deleted group';
  return `Group ${key.slice(0, 8)}…`;
};

/**
 * Roll-up for the dashboard: how many matched tenant policies are unassigned
 * and how many have devices in error/conflict.
 */
export const summariseAssignmentStatus = (tenantPolicies, assignmentsById, statusById) => {
  let unassigned = 0;
  let withIssues = 0;
  let assigned = 0;
  tenantPolicies.forEach(policy => {
    const assignment = assignmentsById?.get(policy.id);
    if (assignment && !assignment.error) {
      if (assignment.hasIncludes) assigned += 1;
      else unassigned += 1;
    }
    const status = statusById?.get(policy.id);
    if (status && !status.loadError && status.issues > 0) withIssues += 1;
  });
  return { total: tenantPolicies.length, assigned, unassigned, withIssues };
};
