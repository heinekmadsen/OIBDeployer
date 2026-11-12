import React, { useState, useMemo } from 'react';
import {
  ArrowLeft, Search, Monitor, Laptop, Smartphone, RefreshCw, Loader, AlertTriangle,
  CheckCircle, XCircle, Clock, MinusCircle, Download, ChevronDown, ChevronRight, KeyRound
} from 'lucide-react';
import graphAPI from '../utils/graph-api';
import authService from '../utils/auth';
import { CONFIG } from '../utils/config';
import { saveConsentResume } from '../utils/consent-resume';
import {
  DEVICE_STATUS, ISSUE_STATUSES, buildDeviceValidation, getOIBOSTypesForDevice,
  getPoliciesNeedingAssignments, summarizeAssignments, getPoliciesNeedingFilterEvaluation, filterExcludes
} from '../utils/device-validation';

// "include filter "Corp" didn't match" / "exclude filter "Kiosk" matched"
const describeEvaluation = (e) =>
  `${e.mode === 'exclude' ? 'exclude' : 'include'} filter "${e.name || e.filterId}" ${e.result === 'match' ? 'matched' : e.result === 'notmatch' ? "didn't match" : `returned "${e.result}"`}`;


const formatDate = (value) => {
  if (!value) return '—';
  const date = new Date(value);
  if (Number.isNaN(date.getTime()) || date.getFullYear() < 2000) return '—';
  return date.toLocaleString();
};

const daysSince = (value) => {
  const date = new Date(value);
  if (!value || Number.isNaN(date.getTime())) return null;
  return Math.floor((Date.now() - date.getTime()) / 86400000);
};

const DeviceIcon = ({ os, size = 18 }) => {
  const value = (os || '').toLowerCase();
  if (value.startsWith('windows')) return <Monitor size={size} />;
  if (value === 'macos') return <Laptop size={size} />;
  return <Smartphone size={size} />;
};

const StatusBadge = ({ status, rawStatus }) => {
  const meta = DEVICE_STATUS[status];
  const showRaw = status === 'pending' && rawStatus != null && typeof rawStatus === 'number';
  return (
    <span className={`status-badge ${meta.badge}`}>
      {meta.label}{showRaw ? ` (status ${rawStatus})` : ''}
    </span>
  );
};

const FILTERS = [
  { key: 'all',         label: 'All' },
  { key: 'issues',      label: 'Needs attention' },
  { key: 'succeeded',   label: 'Applied' },
  { key: 'failed',      label: 'Error / Conflict' },
  { key: 'pending',     label: 'Pending / not reported' },
  { key: 'unassigned',  label: 'No assignments' },
  { key: 'notInTenant', label: 'Not in tenant' },
];

const matchesFilter = (item, filter) => {
  switch (filter) {
    case 'issues':      return ISSUE_STATUSES.includes(item.status);
    case 'failed':      return item.status === 'error' || item.status === 'conflict' || item.status === 'nonCompliant';
    case 'pending':     return item.status === 'pending' || item.status === 'notReported';
    case 'unassigned':  return item.status === 'unassigned' || item.status === 'notTargeted';
    case 'all':         return true;
    default:            return item.status === filter;
  }
};

// Fetch assignments with a small concurrency limit to stay clear of Graph throttling.
const fetchAssignments = async (tenantPolicies) => {
  const result = new Map();
  const queue = [...tenantPolicies];
  const worker = async () => {
    while (queue.length) {
      const policy = queue.shift();
      try {
        result.set(policy.id, summarizeAssignments(await graphAPI.getPolicyAssignments(policy)));
      } catch (err) {
        if (err?.name === 'ConsentRequiredError') throw err;
        result.set(policy.id, { error: err.message });
      }
    }
  };
  await Promise.all(Array.from({ length: Math.min(4, queue.length) }, worker));
  return result;
};

// Filter evaluations per policy for one device. Failures are left out, so
// the policy just stays "Assigned – not reported".
const fetchFilterEvaluations = async (entries, device) => {
  const result = new Map();
  const queue = [...entries];
  const worker = async () => {
    while (queue.length) {
      const { policy, filterIds } = queue.shift();
      try {
        result.set(policy.id, await graphAPI.getFilterEvaluation({
          managedDeviceId: device.id,
          payloadId: policy.id,
          userId: device.userId || null,
          filterIds,
        }));
      } catch (err) {
        console.warn(`Filter evaluation for ${policy.id} failed:`, err.message);
      }
    }
  };
  await Promise.all(Array.from({ length: Math.min(4, queue.length) }, worker));
  return result;
};

/**
 * DeviceValidationDashboard
 *
 * Pick an Intune managed device and check which OIB policies actually reached
 * it, using the same per-device report as the Intune portal.
 *
 * Props:
 *  - comparisonData  : ComparisonDashboard output (OIB ↔ tenant matches by OS)
 *  - selectedOSTypes : OIB OS buckets loaded for this session
 *  - allowPrerelease : persisted when the user is redirected for consent
 *  - selectedVersion : label of the OIB release(s) compared against
 *  - onBack          : () => void
 */
const DeviceValidationDashboard = ({
  comparisonData,
  selectedOSTypes = [],
  allowPrerelease = false,
  selectedVersion,
  onBack,
}) => {
  const [query, setQuery] = useState('');
  const [devices, setDevices] = useState(null);
  const [isSearching, setIsSearching] = useState(false);
  const [selectedDevice, setSelectedDevice] = useState(null);
  const [deviceStatus, setDeviceStatus] = useState(null);
  const [assignmentsById, setAssignmentsById] = useState(null);
  const [filterEvalsById, setFilterEvalsById] = useState(null);
  const [isChecking, setIsChecking] = useState(false);
  const [error, setError] = useState(null);
  const [needsConsent, setNeedsConsent] = useState(false);
  const [filter, setFilter] = useState('all');
  const [expanded, setExpanded] = useState(new Set());
  const [showOther, setShowOther] = useState(false);

  const handleError = (err) => {
    if (err?.name === 'ConsentRequiredError') {
      setNeedsConsent(true);
      setError(null);
    } else {
      setError(err?.message || String(err));
    }
  };

  const handleSearch = async (event) => {
    event?.preventDefault();
    setIsSearching(true);
    setError(null);
    try {
      const results = await graphAPI.searchManagedDevices(query);
      setDevices(results);
      setNeedsConsent(false);
    } catch (err) {
      handleError(err);
    } finally {
      setIsSearching(false);
    }
  };

  const runCheck = async (device) => {
    setSelectedDevice(device);
    setDeviceStatus(null);
    setAssignmentsById(null);
    setFilterEvalsById(null);
    setIsChecking(true);
    setError(null);
    setFilter('all');
    setExpanded(new Set());
    try {
      const status = await graphAPI.getDevicePolicyStatus(device.id);
      const currentDevice = status.device || device;

      // For policies the device didn't report on, read their assignments so
      // "never assigned" can be told apart from "assigned, not reported yet".
      const osTypes = getOIBOSTypesForDevice(currentDevice).filter(os => selectedOSTypes.includes(os));
      const preliminary = buildDeviceValidation(comparisonData, status, osTypes);
      const assignments = await fetchAssignments(getPoliciesNeedingAssignments(preliminary));

      // Assigned but not reported, with an assignment filter: ask Intune how
      // the filter was evaluated for this device at its last check-in.
      const withAssignments = buildDeviceValidation(comparisonData, status, osTypes, assignments);
      const filterEvals = await fetchFilterEvaluations(
        getPoliciesNeedingFilterEvaluation(withAssignments), currentDevice
      );

      setDeviceStatus(status);
      setAssignmentsById(assignments);
      setFilterEvalsById(filterEvals);
      setSelectedDevice(currentDevice);
    } catch (err) {
      handleError(err);
    } finally {
      setIsChecking(false);
    }
  };

  const handleGrantConsent = async () => {
    saveConsentResume({ deploymentType: 'device', osTypes: selectedOSTypes, allowPrerelease });
    try {
      await authService.requestAdditionalScopes(CONFIG.msal.deviceScopes);
    } catch (err) {
      setError('Could not start the consent flow: ' + err.message);
    }
  };

  const deviceOSTypes = useMemo(() => getOIBOSTypesForDevice(selectedDevice), [selectedDevice]);
  const applicableOSTypes = deviceOSTypes.filter(os => selectedOSTypes.includes(os));

  const validation = useMemo(() => {
    if (!deviceStatus || !comparisonData) return null;
    return buildDeviceValidation(comparisonData, deviceStatus, applicableOSTypes, assignmentsById, filterEvalsById);
    // applicableOSTypes is derived from selectedDevice + selectedOSTypes
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [deviceStatus, comparisonData, selectedDevice, selectedOSTypes, assignmentsById, filterEvalsById]);

  const filteredItems = (validation?.items || []).filter(item => matchesFilter(item, filter));

  const toggleExpanded = (key) => {
    setExpanded(prev => {
      const next = new Set(prev);
      next.has(key) ? next.delete(key) : next.add(key);
      return next;
    });
  };

  const csvEscape = (value) => {
    const str = String(value ?? '');
    return /[",\n]/.test(str) ? `"${str.replace(/"/g, '""')}"` : str;
  };

  const handleExportCsv = () => {
    const headers = ['Device', 'OIB Policy', 'OS', 'Policy Type', 'Tenant Policy', 'Device Status', 'Assigned To', 'Contexts', 'Last Reported'];
    const rows = validation.items.map(item => [
      selectedDevice.deviceName,
      item.policy.name.replace('.json', ''),
      item.osType,
      item.policy.policyType || '',
      item.tenantPolicy?.displayName || '',
      DEVICE_STATUS[item.status].label,
      item.assignments?.label || '',
      (item.report?.contexts || []).map(c => `${c.upn || 'device'}: ${DEVICE_STATUS[c.status].label}`).join('; '),
      (item.report?.contexts || []).map(c => c.lastReported).filter(Boolean).sort().pop() || '',
    ]);
    const csv = [headers, ...rows].map(r => r.map(csvEscape).join(',')).join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const safeName = (selectedDevice.deviceName || 'device').replace(/[^a-z0-9-_]+/gi, '_');
    link.download = `oib-device-validation-${safeName}-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  // ── Render helpers ───────────────────────────────────────────────────────

  const renderConsent = () => (
    <div className="device-consent">
      <KeyRound size={28} />
      <div>
        <h3>Additional permission needed</h3>
        <p>
          Device Validation reads managed devices and their per-device policy status, which needs the
          delegated Microsoft Graph permission <code>DeviceManagementManagedDevices.Read.All</code>.
          It is read-only — nothing is changed on your devices or tenant.
        </p>
        <p className="device-consent-note">
          You'll be redirected to Microsoft to approve it and then brought back here. If your account
          can't consent, ask an administrator to grant admin consent for this app registration.
        </p>
        <button className="btn btn-primary" onClick={handleGrantConsent}>Grant access</button>
      </div>
    </div>
  );

  const renderDevicePicker = () => (
    <div className="device-picker">
      <form className="device-search" onSubmit={handleSearch}>
        <div className="device-search-input">
          <Search size={16} />
          <input
            type="text"
            value={query}
            onChange={e => setQuery(e.target.value)}
            placeholder="Device name, primary user UPN or serial number"
            aria-label="Search devices"
          />
        </div>
        <button type="submit" className="btn btn-primary" disabled={isSearching}>
          {isSearching ? <><Loader size={14} className="spinning" /> Searching…</> : 'Search'}
        </button>
      </form>
      <p className="device-search-hint">
        Matches the start of the device name or UPN, or an exact serial number. Leave empty to list the first 50 devices.
      </p>

      {devices && devices.length === 0 && (
        <div className="empty-state">
          <AlertTriangle size={24} />
          <p>No devices found.</p>
        </div>
      )}

      {devices && devices.length > 0 && (
        <ul className="device-results">
          {devices.map(device => {
            const osTypes = getOIBOSTypesForDevice(device);
            const supported = osTypes.some(os => selectedOSTypes.includes(os));
            const isSelected = selectedDevice?.id === device.id;
            return (
              <li key={device.id}>
                <button
                  type="button"
                  className={`device-result ${isSelected ? 'selected' : ''}`}
                  disabled={!supported || isChecking}
                  onClick={() => runCheck(device)}
                  title={supported ? 'Check this device' : `No ${device.operatingSystem} baseline loaded — go back and select it`}
                >
                  <span className="device-result-icon"><DeviceIcon os={device.operatingSystem} /></span>
                  <span className="device-result-main">
                    <span className="device-result-name">{device.deviceName}</span>
                    <span className="device-result-meta">
                      {device.operatingSystem} {device.osVersion} · {device.userPrincipalName || 'No primary user'}
                    </span>
                  </span>
                  <span className="device-result-sync">
                    {supported ? `Last sync ${formatDate(device.lastSyncDateTime)}` : 'Baseline not loaded'}
                  </span>
                </button>
              </li>
            );
          })}
        </ul>
      )}
    </div>
  );

  const renderDeviceSummary = () => {
    const syncAge = daysSince(selectedDevice.lastSyncDateTime);
    return (
      <div className="device-summary">
        <div className="device-summary-icon"><DeviceIcon os={selectedDevice.operatingSystem} size={32} /></div>
        <div className="device-summary-main">
          <h3>{selectedDevice.deviceName}</h3>
          <div className="device-summary-grid">
            <span><b>OS:</b> {selectedDevice.operatingSystem} {selectedDevice.osVersion}</span>
            <span><b>Primary user:</b> {selectedDevice.userPrincipalName || '—'}</span>
            <span><b>Model:</b> {[selectedDevice.manufacturer, selectedDevice.model].filter(Boolean).join(' ') || '—'}</span>
            <span><b>Serial:</b> {selectedDevice.serialNumber || '—'}</span>
            <span><b>Last sync:</b> {formatDate(selectedDevice.lastSyncDateTime)}</span>
            <span><b>Compliance:</b> {selectedDevice.complianceState || '—'}</span>
          </div>
          {syncAge !== null && syncAge >= 7 && (
            <p className="validation-scope-note">
              <AlertTriangle size={14} />
              {` This device hasn't synced for ${syncAge} days — statuses below may be out of date.`}
            </p>
          )}
          {applicableOSTypes.length === 0 && (
            <p className="validation-scope-note">
              <AlertTriangle size={14} />
              {` No OIB baseline loaded for ${selectedDevice.operatingSystem}. Go back and include it in the OS selection.`}
            </p>
          )}
        </div>
        <button
          className="btn btn-secondary btn-refresh"
          onClick={() => runCheck(selectedDevice)}
          disabled={isChecking}
        >
          <RefreshCw size={16} className={isChecking ? 'spinning' : ''} />
          Re-check
        </button>
      </div>
    );
  };

  const renderStats = () => {
    const t = validation.totals;
    const cards = [
      { cls: 'current',  icon: <CheckCircle size={24} />, value: t.succeeded, label: 'Applied' },
      { cls: 'error',    icon: <XCircle size={24} />,     value: t.error + t.conflict + t.nonCompliant, label: 'Error / Conflict' },
      { cls: 'outdated', icon: <Clock size={24} />,       value: t.pending + t.notReported, label: 'Pending / not reported' },
      { cls: 'missing',  icon: <MinusCircle size={24} />, value: t.unassigned + t.notTargeted, label: 'No assignments' },
      { cls: 'duplicate', icon: <AlertTriangle size={24} />, value: t.notInTenant, label: 'Not in tenant' },
    ];
    return (
      <div className="comparison-stats">
        {cards.map(card => (
          <div key={card.label} className={`stat-card ${card.cls}`}>
            <div className="stat-icon">{card.icon}</div>
            <div className="stat-content">
              <div className="stat-number">{card.value}</div>
              <div className="stat-label">{card.label}</div>
            </div>
          </div>
        ))}
      </div>
    );
  };

  const renderItem = (item) => {
    const key = `${item.osType}-${item.policy.name}`;
    const contexts = item.report?.contexts || [];
    const isExpanded = expanded.has(key);
    const lastReported = contexts.map(c => c.lastReported).filter(Boolean).sort().pop();
    const firstRaw = contexts.find(c => c.status === item.status)?.rawStatus;

    return (
      <div key={key} className={`policy-comparison-item ${DEVICE_STATUS[item.status].row}`}>
        <div className="policy-info">
          <div className="policy-name">
            {item.policy.name.replace('.json', '')}
            {item.policy.skuRequirements === 'Enterprise' && (
              <span className="req-tag req-tag--enterprise" title="Requires Windows Enterprise SKU">Enterprise</span>
            )}
            {item.policy.licenseRequirements === 'MDE' && (
              <span className="req-tag req-tag--mde" title="Requires Microsoft Defender for Endpoint licence">MDE</span>
            )}
          </div>
          {item.tenantPolicy && (
            <div className="matched-policy-name">
              <span className="matched-label">Tenant policy:</span>
              <span className="matched-name">{item.tenantPolicy.displayName || item.tenantPolicy.name}</span>
            </div>
          )}
          <div className="policy-status">
            <StatusBadge status={item.status} rawStatus={firstRaw} />
            {item.policy.status === 'outdated' && (
              <span className="status-badge outdated">Older OIB version in tenant</span>
            )}
            {item.duplicate && (
              <span className="status-badge duplicate">
                {`Multiple tenant matches${item.duplicateHits > 1 ? ` – ${item.duplicateHits} reached this device` : ''}`}
              </span>
            )}
            {lastReported && (
              <span className="validation-score">Last reported {formatDate(lastReported)}</span>
            )}
          </div>
          {item.status === 'filteredOut' && (
            <div className="device-item-hint">
              {`Assigned to ${item.assignments.label}, but at its last check-in the `}
              {item.filterEvaluation.filter(filterExcludes).map(describeEvaluation).join(' and the ')}
              {', so this device is left out on purpose.'}
            </div>
          )}
          {item.status === 'notReported' && (
            <div className="device-item-hint">
              {`Assigned to ${item.assignments.label}. `}
              {item.assignments.filters?.length > 0
                ? (item.filterEvaluation?.length > 0
                  ? `Filter evaluation at the last check-in: ${item.filterEvaluation.map(describeEvaluation).join('; ')}. `
                  : "The assignment uses a filter that Intune hasn't evaluated for this device yet. ")
                : ''}
              {item.assignments.allDevices || item.assignments.allUsers
                ? 'This device should receive it — sync the device, and check it again if the policy was assigned recently.'
                : "The device (or its user) may not be in the assigned groups, or hasn't checked in since the policy was assigned."}
            </div>
          )}
          {item.status === 'unassigned' && (
            <div className="device-item-hint">
              {item.assignments?.exclusions
                ? `The policy only has exclusions (${item.assignments.label}), so no device receives it. Add an assignment in Intune.`
                : 'The policy exists in the tenant but has no assignments, so no device receives it. Assign it in Intune per OIB guidance.'}
            </div>
          )}
          {item.status === 'notTargeted' && (
            <div className="device-item-hint">
              The device has no report for this policy, and its assignments couldn't be read. It may not be assigned to this device or its user, or the device hasn't checked in since.
            </div>
          )}
          {item.status === 'notReportable' && (
            <div className="device-item-hint">
              Intune doesn't report this policy type per device here (App Protection is per app/user; feature, quality and driver update profiles use Windows Update reports).
            </div>
          )}
        </div>

        {contexts.length > 1 && (
          <div className="policy-actions">
            <button className="btn btn-ghost btn-small" onClick={() => toggleExpanded(key)}>
              {isExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
              {`${contexts.length} contexts`}
            </button>
          </div>
        )}

        {isExpanded && (
          <div className="validation-details">
            <table className="validation-table">
              <thead>
                <tr><th>Context</th><th>Status</th><th>Last reported</th></tr>
              </thead>
              <tbody>
                {contexts.map((c, i) => (
                  <tr key={i}>
                    <td>{c.upn || 'Device'}</td>
                    <td><StatusBadge status={c.status} rawStatus={c.rawStatus} /></td>
                    <td>{formatDate(c.lastReported)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    );
  };

  const renderResults = () => {
    const byOS = {};
    filteredItems.forEach(item => {
      (byOS[item.osType] ||= {});
      const typeKey = item.policy.policyType || 'Other';
      (byOS[item.osType][typeKey] ||= []).push(item);
    });

    const sourceErrors = Object.entries(deviceStatus.errors || {}).filter(([, msg]) => msg);

    return (
      <>
        {sourceErrors.length > 0 && (
          <div className="deprecation-notice">
            <AlertTriangle size={14} />
            <span>Some device data couldn't be loaded — results may be incomplete:</span>
            {sourceErrors.map(([source, msg]) => (
              <span key={source} className="deprecation-replacement" title={msg}>
                {source === 'configuration' ? 'Configuration report' : source === 'compliance' ? 'Compliance states' : 'Device details'}
                {`: ${msg.slice(0, 160)}${msg.length > 160 ? '…' : ''}`}
              </span>
            ))}
          </div>
        )}

        {renderStats()}

        <div className="filter-tabs">
          {FILTERS.map(f => {
            const count = (validation.items || []).filter(item => matchesFilter(item, f.key)).length;
            return (
              <button
                key={f.key}
                className={`filter-tab ${filter === f.key ? 'active' : ''}`}
                onClick={() => setFilter(f.key)}
              >
                {`${f.label} (${count})`}
              </button>
            );
          })}
        </div>

        <div className="policy-comparison-list">
          {filteredItems.length === 0 ? (
            <div className="empty-state">
              <CheckCircle size={24} />
              <p>No OIB policies in this view.</p>
            </div>
          ) : (
            Object.entries(byOS).map(([osType, types]) => (
              <div key={osType} className="os-policy-section">
                <div className="os-section-header">
                  <h3 className="os-section-title">{osType}</h3>
                </div>
                <div className="os-policies">
                  {Object.entries(types).map(([policyType, items]) => (
                    <div key={`${osType}-${policyType}`} className="policy-type-section">
                      <div className="policy-type-header">
                        <h4 className="policy-type-title">{policyType.replace(/([A-Z])/g, ' $1').trim()}</h4>
                        <span className="policy-type-count">{items.length} policies</span>
                      </div>
                      <div className="policy-type-policies">
                        {items
                          .sort((a, b) =>
                            DEVICE_STATUS[a.status].severity - DEVICE_STATUS[b.status].severity ||
                            a.policy.name.localeCompare(b.policy.name)
                          )
                          .map(renderItem)}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ))
          )}
        </div>

        {validation.otherPolicies.length > 0 && (
          <div className="device-other-policies">
            <button className="validation-section-toggle" onClick={() => setShowOther(v => !v)}>
              {showOther ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
              {`Other (non-OIB) policies on this device (${validation.otherPolicies.length})`}
              {validation.otherPolicies.some(p => p.status === 'error' || p.status === 'conflict') && (
                <span className="status-badge error">has errors/conflicts</span>
              )}
            </button>
            {showOther && (
              <table className="validation-table">
                <thead>
                  <tr><th>Policy</th><th>Kind</th><th>Status</th></tr>
                </thead>
                <tbody>
                  {validation.otherPolicies.map(p => (
                    <tr key={`${p.kind}-${p.policyId}`}>
                      <td>{p.policyName}</td>
                      <td>{p.kind}</td>
                      <td><StatusBadge status={p.status} rawStatus={p.contexts?.[0]?.rawStatus} /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        )}
      </>
    );
  };

  // ── Main render ──────────────────────────────────────────────────────────

  return (
    <div className="wizard-container comparison-dashboard">
      <div className="comparison-content">
        <div className="wizard-header">
          <h2>Device Validation</h2>
          <p>Pick a device to see whether each OIB policy has actually been delivered to it.</p>
          <p>Status comes from Intune's per-device report — the same data as <i>Devices › device › Device configuration</i> in the Intune admin center. Read-only.</p>
          {selectedVersion && (
            <div className="version-info">
              <span className="version-badge">Comparing against: {selectedVersion}</span>
            </div>
          )}
        </div>

        {error && (
          <div className="error-banner">
            <AlertTriangle className="error-icon" />
            <span>{error}</span>
            <button onClick={() => setError(null)} className="error-dismiss">×</button>
          </div>
        )}

        {needsConsent ? renderConsent() : (
          <>
            {!selectedDevice || !deviceStatus ? renderDevicePicker() : null}

            {selectedDevice && (isChecking || deviceStatus) && renderDeviceSummary()}

            {isChecking && (
              <div className="loading-state">
                <RefreshCw className="spinner" size={32} />
                <p>Reading policy status for {selectedDevice?.deviceName}…</p>
              </div>
            )}

            {!isChecking && deviceStatus && validation && renderResults()}
          </>
        )}
      </div>

      <div className="comparison-navigation-sticky">
        <div className="wizard-navigation">
          <button className="btn-secondary" onClick={onBack}>
            <ArrowLeft size={16} />
            Back
          </button>

          {deviceStatus && !isChecking && (
            <button
              className="btn-secondary"
              onClick={() => { setSelectedDevice(null); setDeviceStatus(null); }}
            >
              <Search size={16} />
              Choose another device
            </button>
          )}

          {validation && validation.items.length > 0 && !isChecking && (
            <button className="btn-secondary" onClick={handleExportCsv}>
              <Download size={16} />
              Export CSV
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

export default DeviceValidationDashboard;
