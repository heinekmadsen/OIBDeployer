// Device Validation: compare the OIB policies matched in the tenant against
// what Intune reports for a single managed device.
//
// Pure functions only — Graph calls live in graph-api.js.

// Normalised per-device statuses, in order of severity (worst first).
// When a policy reports more than once for a device (device + user context),
// the worst status wins.
export const DEVICE_STATUS = {
  error:         { label: 'Error',                       badge: 'error',    row: 'error',     severity: 0 },
  conflict:      { label: 'Conflict',                    badge: 'newer',    row: 'error',     severity: 1 },
  nonCompliant:  { label: 'Applied – not compliant',     badge: 'outdated', row: 'drifted',   severity: 2 },
  pending:       { label: 'Pending',                     badge: 'pending',  row: 'drifted',   severity: 3 },
  // Policy is assigned in the tenant but the device has no report for it.
  notReported:   { label: 'Assigned – not reported',     badge: 'pending',  row: 'drifted',   severity: 4 },
  // No report and assignments couldn't be read.
  notTargeted:   { label: 'Not reported by device',      badge: 'missing',  row: 'missing',   severity: 5 },
  // Policy exists in the tenant but has no assignments at all.
  unassigned:    { label: 'No assignments in tenant',    badge: 'missing',  row: 'missing',   severity: 6 },
  succeeded:     { label: 'Applied',                     badge: 'current',  row: 'compliant', severity: 7 },
  notApplicable: { label: 'Not applicable',              badge: 'neutral',  row: '',          severity: 8 },
  notInTenant:   { label: 'Not deployed in tenant',      badge: 'neutral',  row: 'missing',   severity: 9 },
  notReportable: { label: 'No per-device report',        badge: 'neutral',  row: '',          severity: 10 },
};

// Statuses that mean "an OIB policy deployed in this tenant hasn't landed
// cleanly on this device". Policies missing from the tenant entirely are a
// tenant-level gap (see Existing Deployment) and are counted separately.
export const ISSUE_STATUSES = ['error', 'conflict', 'nonCompliant', 'pending', 'notReported', 'notTargeted', 'unassigned'];

// Summarise a policy's Graph assignments (any Intune policy type — they all
// share the deviceAndAppManagementAssignmentTarget shape).
export const summarizeAssignments = (assignments = []) => {
  const summary = { allDevices: false, allUsers: false, groups: 0, exclusions: 0 };
  assignments.forEach(assignment => {
    const type = (assignment?.target?.['@odata.type'] || '').toLowerCase();
    if (type.includes('alldevices')) summary.allDevices = true;
    else if (type.includes('alllicensedusers')) summary.allUsers = true;
    else if (type.includes('exclusiongroup')) summary.exclusions += 1;
    else if (type.includes('groupassignment')) summary.groups += 1;
  });
  summary.hasIncludes = summary.allDevices || summary.allUsers || summary.groups > 0;
  const parts = [
    summary.allDevices && 'All devices',
    summary.allUsers && 'All users',
    summary.groups > 0 && `${summary.groups} group${summary.groups > 1 ? 's' : ''}`,
  ].filter(Boolean);
  summary.label = (parts.join(', ') || 'Nothing') +
    (summary.exclusions ? ` (${summary.exclusions} excluded group${summary.exclusions > 1 ? 's' : ''})` : '');
  return summary;
};

// PolicyStatus integers from reports/getConfigurationPoliciesReportForDevice.
// Microsoft does not document this enum; 1/2/5/6 are confirmed by community
// tooling and match the portal. Anything else is surfaced as pending with the
// raw code so it can be checked against the portal.
export const mapReportPolicyStatus = (code) => {
  switch (Number(code)) {
    case 1: return 'notApplicable';
    case 2: return 'succeeded';
    case 5: return 'error';
    case 6: return 'conflict';
    default: return 'pending';
  }
};

// deviceCompliancePolicyState.state (complianceStatus enum)
export const mapComplianceState = (state) => {
  switch ((state || '').toLowerCase()) {
    case 'compliant':
    case 'remediated':    return 'succeeded';
    case 'noncompliant':  return 'nonCompliant';
    case 'error':         return 'error';
    case 'conflict':      return 'conflict';
    case 'notapplicable': return 'notApplicable';
    case 'notassigned':   return 'notTargeted';
    default:              return 'pending'; // unknown
  }
};

// Map an Intune managedDevice to the OIB OS buckets (release manifest
// platform folders) its baselines live in. BYOD holds App Protection
// policies, which are per app/user and never reported per device.
export const getOIBOSTypesForDevice = (device) => {
  const os = (device?.operatingSystem || '').toLowerCase();
  if (os.startsWith('windows')) {
    const isCloudPC = /cloud pc/i.test(device?.model || '');
    return isCloudPC ? ['WINDOWS', 'WINDOWS365'] : ['WINDOWS'];
  }
  if (os === 'macos') return ['MACOS'];
  if (os.startsWith('ios') || os.startsWith('ipados')) return ['IOS'];
  if (os.startsWith('android')) return ['ANDROID'];
  return [];
};

const worstOf = (statuses) =>
  statuses.reduce((worst, s) =>
    DEVICE_STATUS[s].severity < DEVICE_STATUS[worst].severity ? s : worst
  );

// Group raw report rows by tenant policy ID (lower-cased).
const indexConfigurationRows = (rows) => {
  const byId = new Map();
  rows.forEach(row => {
    const id = (row.PolicyId || '').toLowerCase();
    if (!id) return;
    const status = mapReportPolicyStatus(row.PolicyStatus);
    const entry = byId.get(id) || { policyId: row.PolicyId, policyName: row.PolicyName, contexts: [] };
    entry.contexts.push({
      status,
      rawStatus: row.PolicyStatus,
      upn: row.UPN || '',
      lastReported: row.PspdpuLastModifiedTimeUtc || null,
      baseType: row.PolicyBaseTypeName,
    });
    byId.set(id, entry);
  });
  byId.forEach(entry => { entry.status = worstOf(entry.contexts.map(c => c.status)); });
  return byId;
};

const indexComplianceStates = (states) => {
  // One state is returned per context (primary user, "System account"…);
  // group them into one entry per policy.
  const grouped = new Map();
  states.forEach(state => {
    const key = (state.id || state.displayName || '').toLowerCase();
    const entry = grouped.get(key) || { policyId: state.id, policyName: state.displayName, contexts: [] };
    entry.contexts.push({
      status: mapComplianceState(state.state),
      rawStatus: state.state,
      upn: state.userPrincipalName || '',
      lastReported: null,
    });
    grouped.set(key, entry);
  });
  const entries = [...grouped.values()].map(entry => ({
    ...entry,
    status: worstOf(entry.contexts.map(c => c.status)),
  }));

  // deviceCompliancePolicyState.id is the policy ID in practice, but has been
  // seen suffixed with a user ID, so match on "contains" and fall back to name.
  return {
    entries,
    find: (tenantPolicy) => {
      const id = (tenantPolicy.id || '').toLowerCase();
      const name = (tenantPolicy.displayName || '').toLowerCase();
      const matches = entries.filter(e =>
        (id && (e.policyId || '').toLowerCase().includes(id)) ||
        (name && (e.policyName || '').toLowerCase() === name)
      );
      if (matches.length === 0) return null;
      return {
        ...matches[0],
        matchedIds: matches.map(m => m.policyId),
        status: worstOf(matches.map(m => m.status)),
        contexts: matches.flatMap(m => m.contexts),
      };
    },
  };
};

// Policy types Intune does not report per device through these APIs.
const isNotReportable = (tenantPolicy) =>
  tenantPolicy?.type === 'appProtection' || tenantPolicy?.type === 'updatePolicy';

/**
 * Build the per-device comparison.
 *
 * @param {object} comparisonData  output of ComparisonDashboard (byOS buckets)
 * @param {object} deviceStatus    output of graphAPI.getDevicePolicyStatus()
 * @param {string[]} osTypes       OIB OS buckets that apply to the device
 * @param {Map} [assignmentsById]  tenant policy id → summarizeAssignments()
 *                                 result (or { error }) for policies the
 *                                 device didn't report
 * @returns {{ items, otherPolicies, totals }}
 */
export const buildDeviceValidation = (comparisonData, deviceStatus, osTypes, assignmentsById = null) => {
  const configIndex = indexConfigurationRows(deviceStatus.configurationRows || []);
  const complianceIndex = indexComplianceStates(deviceStatus.complianceStates || []);
  const claimedConfigIds = new Set();
  const claimedComplianceIds = new Set();

  const lookup = (tenantPolicy) => {
    if (!tenantPolicy) return null;
    if (tenantPolicy.type === 'compliance') {
      const hit = complianceIndex.find(tenantPolicy);
      if (hit) hit.matchedIds.forEach(id => claimedComplianceIds.add(id));
      return hit;
    }
    const hit = configIndex.get((tenantPolicy.id || '').toLowerCase()) || null;
    if (hit) claimedConfigIds.add(hit.policyId.toLowerCase());
    return hit;
  };

  const items = [];

  osTypes.forEach(osType => {
    const osData = comparisonData?.byOS?.[osType];
    if (!osData) return;

    const matched = [
      ...(osData.current || []),
      ...(osData.outdated || []),
      ...(osData.newer || []),
    ];

    matched.forEach(policy => {
      const report = lookup(policy.existingPolicy);
      let status;
      let assignments = null;
      if (report) {
        status = report.status;
      } else if (isNotReportable(policy.existingPolicy)) {
        status = 'notReportable';
      } else {
        const summary = assignmentsById?.get(policy.existingPolicy?.id);
        if (summary && !summary.error) {
          assignments = summary;
          status = summary.hasIncludes ? 'notReported' : 'unassigned';
        } else {
          status = 'notTargeted';
        }
      }

      items.push({ policy, osType, tenantPolicy: policy.existingPolicy, report, status, assignments });
    });

    // Ambiguous tenant matches: report whichever copy reached the device.
    (osData.duplicates || []).forEach(policy => {
      const hits = (policy.matchedPolicies || [])
        .map(tp => ({ tenantPolicy: tp, report: lookup(tp) }))
        .filter(h => h.report);
      const best = hits.sort((a, b) =>
        DEVICE_STATUS[a.report.status].severity - DEVICE_STATUS[b.report.status].severity
      )[0];
      items.push({
        policy,
        osType,
        tenantPolicy: best?.tenantPolicy || null,
        report: best?.report || null,
        status: best ? best.report.status : 'notTargeted',
        duplicate: true,
        duplicateHits: hits.length,
      });
    });

    (osData.missing || []).forEach(policy => {
      items.push({ policy, osType, tenantPolicy: null, report: null, status: 'notInTenant' });
    });
  });

  // Policies that reached the device but aren't part of the OIB baseline —
  // useful when chasing conflicts.
  const otherPolicies = [
    ...[...configIndex.values()]
      .filter(entry => !claimedConfigIds.has(entry.policyId.toLowerCase()))
      .map(entry => ({ ...entry, kind: 'Configuration' })),
    ...complianceIndex.entries
      .filter(entry => !claimedComplianceIds.has(entry.policyId))
      .map(entry => ({ ...entry, kind: 'Compliance' })),
  ].sort((a, b) =>
    DEVICE_STATUS[a.status].severity - DEVICE_STATUS[b.status].severity ||
    (a.policyName || '').localeCompare(b.policyName || '')
  );

  const totals = Object.keys(DEVICE_STATUS).reduce((acc, key) => ({ ...acc, [key]: 0 }), {});
  items.forEach(item => { totals[item.status] += 1; });
  totals.total = items.length;
  totals.issues = items.filter(item => ISSUE_STATUSES.includes(item.status)).length;

  return { items, otherPolicies, totals };
};

// Tenant policies the device didn't report on, whose assignments are worth
// fetching to tell "unassigned" apart from "assigned but not reported".
export const getPoliciesNeedingAssignments = (validation) =>
  (validation?.items || [])
    .filter(item => item.status === 'notTargeted' && item.tenantPolicy?.id && !item.duplicate)
    .map(item => item.tenantPolicy);
