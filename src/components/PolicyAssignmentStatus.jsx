import React, { useState } from 'react';
import { Users, Monitor, Filter, ChevronDown, ChevronRight, Loader, AlertTriangle, Ban } from 'lucide-react';
import graphAPI from '../utils/graph-api';
import { DEVICE_STATUS } from '../utils/device-validation';
import { normalisePolicyDeviceRows, groupLabel } from '../utils/policy-status';

const formatDate = (value) => {
  if (!value) return '—';
  const date = new Date(value);
  if (Number.isNaN(date.getTime()) || date.getFullYear() < 2000) return '—';
  return date.toLocaleString();
};

const COUNT_PARTS = [
  { key: 'succeeded',     label: 'succeeded',      cls: 'ok' },
  { key: 'error',         label: 'error',          cls: 'bad' },
  { key: 'conflict',      label: 'conflict',       cls: 'bad' },
  { key: 'nonCompliant',  label: 'not compliant',  cls: 'warn' },
  { key: 'pending',       label: 'pending',        cls: 'warn' },
  { key: 'notApplicable', label: 'not applicable', cls: 'muted' },
];

/**
 * Assignment targets + device status counts for one tenant policy, with a
 * lazily-loaded per-device list. Rendered inside a ComparisonDashboard row.
 *
 * Props:
 *  - tenantPolicy : normalised tenant policy ({ id, type, displayName })
 *  - assignment   : summarizeAssignments() result, { error }, or undefined
 *  - status       : normalisePolicyStatusSummary() result, { loadError }, null
 *                   (type has no per-device report) or undefined (not loaded)
 *  - groupNames   : Map<groupId, name|null> or null
 *  - filters      : Map<filterId, { displayName }> or null
 */
const PolicyAssignmentStatus = ({ tenantPolicy, assignment, status, groupNames, filters }) => {
  const [showDevices, setShowDevices] = useState(false);
  const [devices, setDevices] = useState(null);
  const [devicesError, setDevicesError] = useState(null);
  const [loadingDevices, setLoadingDevices] = useState(false);

  const toggleDevices = async () => {
    const next = !showDevices;
    setShowDevices(next);
    if (!next || devices || loadingDevices) return;
    setLoadingDevices(true);
    setDevicesError(null);
    try {
      const result = await graphAPI.getPolicyDeviceStatuses(tenantPolicy);
      setDevices({ rows: normalisePolicyDeviceRows(result), total: result.total, truncated: result.truncated });
    } catch (error) {
      setDevicesError(error.message);
    } finally {
      setLoadingDevices(false);
    }
  };

  const filterLabel = (filter) => {
    const name = filters?.get(filter.id.toLowerCase())?.displayName || `Filter ${filter.id.slice(0, 8)}…`;
    return `${filter.mode === 'exclude' ? 'Exclude' : 'Include'} filter: ${name}`;
  };

  const renderAssignments = () => {
    if (assignment === undefined) return null;
    if (assignment.error) {
      return (
        <div className="pas-row pas-muted">
          <AlertTriangle size={13} /> Couldn't read assignments: {assignment.error.slice(0, 140)}
        </div>
      );
    }
    if (!assignment.hasIncludes) {
      return (
        <div className="pas-row">
          <span className="status-badge missing"><Ban size={12} /> Not assigned</span>
          {assignment.exclusions > 0 && (
            <span className="pas-muted">only exclusions configured — no device receives this policy</span>
          )}
        </div>
      );
    }
    return (
      <div className="pas-row pas-chips">
        <span className="pas-label">Assigned to</span>
        {assignment.allDevices && <span className="pas-chip"><Monitor size={12} /> All devices</span>}
        {assignment.allUsers && <span className="pas-chip"><Users size={12} /> All users</span>}
        {assignment.includeGroupIds.map(id => (
          <span key={`inc-${id}`} className="pas-chip" title={id}><Users size={12} /> {groupLabel(id, groupNames)}</span>
        ))}
        {assignment.excludeGroupIds.map(id => (
          <span key={`exc-${id}`} className="pas-chip pas-chip-exclude" title={id}>
            <Ban size={12} /> Excluded: {groupLabel(id, groupNames)}
          </span>
        ))}
        {assignment.filters.map((filter, i) => (
          <span key={`flt-${i}`} className="pas-chip pas-chip-filter" title={filter.target ? `On ${filter.target}` : undefined}>
            <Filter size={12} /> {filterLabel(filter)}
          </span>
        ))}
      </div>
    );
  };

  const renderStatus = () => {
    if (status === undefined) return null;
    if (status === null) {
      return <div className="pas-row pas-muted">No per-device status report for this policy type.</div>;
    }
    if (status.loadError) {
      return (
        <div className="pas-row pas-muted">
          <AlertTriangle size={13} /> Couldn't read device status: {status.loadError.slice(0, 140)}
        </div>
      );
    }
    if (status.total === 0) {
      return <div className="pas-row pas-muted">No devices have reported on this policy yet.</div>;
    }
    return (
      <div className="pas-row pas-counts">
        <span className="pas-label">{`${status.total} device${status.total === 1 ? '' : 's'}`}</span>
        {COUNT_PARTS.filter(p => status[p.key] > 0 || p.key === 'succeeded').map(part => (
          <span key={part.key} className={`pas-count pas-count-${part.cls}`}>
            {status[part.key]} {part.label}
          </span>
        ))}
        <button type="button" className="btn btn-ghost btn-small" onClick={toggleDevices}>
          {showDevices ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
          {showDevices ? 'Hide devices' : 'Show devices'}
        </button>
      </div>
    );
  };

  const renderDevices = () => {
    if (!showDevices) return null;
    if (loadingDevices) {
      return <div className="pas-row pas-muted"><Loader size={13} className="spinning" /> Loading devices…</div>;
    }
    if (devicesError) {
      return <div className="pas-row pas-muted"><AlertTriangle size={13} /> {devicesError.slice(0, 200)}</div>;
    }
    if (!devices) return null;
    return (
      <div className="pas-devices">
        <table className="validation-table">
          <thead>
            <tr><th>Device</th><th>User</th><th>Status</th><th>Last reported</th></tr>
          </thead>
          <tbody>
            {devices.rows.map((row, i) => (
              <tr key={`${row.deviceId}-${row.upn}-${i}`}>
                <td>{row.deviceName}</td>
                <td>{row.upn || '—'}</td>
                <td>
                  <span className={`status-badge ${DEVICE_STATUS[row.status].badge}`}>
                    {DEVICE_STATUS[row.status].label}
                    {row.status === 'pending' && typeof row.rawStatus === 'number' ? ` (status ${row.rawStatus})` : ''}
                  </span>
                </td>
                <td>{formatDate(row.lastReported)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {devices.truncated && (
          <div className="pas-muted pas-truncated">{`Showing ${devices.rows.length} of ${devices.total} rows.`}</div>
        )}
      </div>
    );
  };

  if (assignment === undefined && status === undefined) return null;

  return (
    <div className="policy-assignment-status">
      {renderAssignments()}
      {renderStatus()}
      {renderDevices()}
    </div>
  );
};

export default PolicyAssignmentStatus;
