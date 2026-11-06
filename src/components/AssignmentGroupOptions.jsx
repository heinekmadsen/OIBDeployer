import React, { useEffect, useState } from 'react';
import { Users, KeyRound, AlertTriangle, Loader, CheckCircle } from 'lucide-react';
import graphAPI from '../utils/graph-api';
import {
  ASSIGN_TARGET_LABELS, buildGroupName, builtInFor, builtInLabel, groupKindsFor, validateGroupNaming
} from '../utils/assignment-groups';

const TARGET_HELP = {
  group: 'Groups are assigned (static) security groups. A group that already exists with the same name is reused. ' +
    'New groups start empty, so the policy reaches no one until you add devices or users to the include group.',
  allDevices: 'Every device the policy applies to gets it at its next check-in. Put devices, not users, in the exclude ' +
    'group: Intune doesn\'t support excluding user groups from a device assignment.',
  allUsers: 'Every licensed user gets the policy at their next check-in. Put users, not devices, in the exclude group: ' +
    'Intune doesn\'t support excluding device groups from a user assignment.',
  byScope: 'OIB marks each policy D (device) or U (user): D policies go to All devices and U policies to All users. ' +
    'Policies without a marker go to All devices. Put devices in the exclude groups of D policies and users in those of U policies.',
};

/**
 * Assignment options shown next to the Deploy button: new include/exclude
 * groups, or All devices / All users with an optional exclude group.
 * State lives in App so it survives navigation and is used by handleDeployment.
 *
 * Props:
 *  - options        { enabled, target, include, exclude, naming }
 *  - onChange(next)
 *  - policyNames    names of the selected policies (preview + validation)
 *  - onAccessChange(state)  'checking' | 'granted' | 'consent' | 'error'
 *  - onGrantAccess  start the consent redirect
 */
const AssignmentGroupOptions = ({ options, onChange, policyNames = [], onAccessChange, onGrantAccess }) => {
  const [access, setAccess] = useState(null);
  const [accessError, setAccessError] = useState(null);

  const target = options.target || 'group';
  const builtInMode = target !== 'group';
  const kinds = groupKindsFor(options);
  const needsGroups = options.enabled && kinds.length > 0;

  // Group permissions only matter when groups will be created.
  useEffect(() => {
    if (!needsGroups) return;
    let cancelled = false;
    const update = (state, error = null) => {
      if (cancelled) return;
      setAccess(state);
      setAccessError(error);
      onAccessChange?.(state);
    };
    update('checking');
    graphAPI.checkGroupManageAccess()
      .then(() => update('granted'))
      .catch(error => update(error?.name === 'ConsentRequiredError' ? 'consent' : 'error', error.message));
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [needsGroups]);

  const set = (patch) => onChange({ ...options, ...patch });
  const setNaming = (patch) => onChange({ ...options, naming: { ...options.naming, ...patch } });

  const errors = options.enabled
    ? validateGroupNaming(options.naming, policyNames, {
      include: kinds.includes('include'), exclude: kinds.includes('exclude'), target,
    })
    : [];
  const sample = policyNames[0];

  // How the selection splits between All devices and All users.
  const builtInCounts = builtInMode
    ? policyNames.reduce((counts, name) => {
      const kind = builtInFor(target, name);
      counts[kind] = (counts[kind] || 0) + 1;
      return counts;
    }, {})
    : {};

  const kindField = (kind, label) => (
    <label className="group-options-kind">
      <input type="checkbox" checked={options[kind]} onChange={e => set({ [kind]: e.target.checked })} />
      <span>{label}</span>
      <input
        type="text"
        value={kind === 'include' ? options.naming.includeLabel : options.naming.excludeLabel}
        disabled={!options[kind]}
        aria-label={`${kind === 'include' ? 'Include' : 'Exclude'} group label`}
        onChange={e => setNaming(kind === 'include' ? { includeLabel: e.target.value } : { excludeLabel: e.target.value })}
      />
    </label>
  );

  return (
    <div className={`group-options ${options.enabled ? 'group-options--open' : ''}`}>
      <label className="group-options-toggle">
        <input
          type="checkbox"
          checked={options.enabled}
          onChange={e => set({ enabled: e.target.checked })}
        />
        <Users size={15} />
        <span>Assign each deployed policy</span>
      </label>

      {options.enabled && (
        <div className="group-options-body">
          <label className="group-options-target">
            <span>Assign to</span>
            <select value={target} onChange={e => set({ target: e.target.value })}>
              {Object.entries(ASSIGN_TARGET_LABELS).map(([value, label]) => (
                <option key={value} value={value}>{label}</option>
              ))}
            </select>
          </label>

          <div className="group-options-fields">
            {!builtInMode && kindField('include', 'Include group, label')}
            {kindField('exclude', builtInMode ? 'Also create an exclude group, label' : 'Exclude group, label')}
            {kinds.length > 0 && (
              <>
                <label>
                  <span>Prefix</span>
                  <input type="text" value={options.naming.prefix} placeholder="e.g. SG-Intune-"
                    onChange={e => setNaming({ prefix: e.target.value })} />
                </label>
                <label>
                  <span>Suffix</span>
                  <input type="text" value={options.naming.suffix} placeholder="e.g. -Pilot"
                    onChange={e => setNaming({ suffix: e.target.value })} />
                </label>
              </>
            )}
          </div>

          {sample && (builtInMode || kinds.length > 0) && (
            <div className="group-options-preview">
              <span className="pas-label">Preview</span>
              {builtInMode && (
                <span className="group-options-builtin">
                  {Object.entries(builtInCounts)
                    .map(([kind, count]) => `${builtInLabel(kind)}${policyNames.length > 1 ? ` (${count})` : ''}`)
                    .join(' · ')}
                </span>
              )}
              {kinds.map(kind => <code key={kind}>{buildGroupName(sample, kind, options.naming)}</code>)}
              {kinds.length > 0 && policyNames.length > 1 && (
                <span className="pas-muted">…and the same for {policyNames.length - 1} more polic{policyNames.length === 2 ? 'y' : 'ies'}</span>
              )}
            </div>
          )}

          <div className="group-options-help pas-muted">
            {TARGET_HELP[target]}
            {builtInMode && (
              <>
                {' '}App protection policies and Windows feature, quality and driver update profiles can&apos;t be assigned
                this way and are left unassigned. To pilot a new OIB version first, choose an include group, then use
                &quot;Migrate devices…&quot; in Existing Deployment.
              </>
            )}
          </div>

          {errors.map(message => (
            <div key={message} className="group-options-note group-options-note--error">
              <AlertTriangle size={14} /> {message}
            </div>
          ))}

          {needsGroups && access === 'checking' && (
            <div className="group-options-note"><Loader size={14} className="spinning" /> Checking group permissions…</div>
          )}
          {needsGroups && access === 'granted' && (
            <div className="group-options-note group-options-note--ok"><CheckCircle size={14} /> Group permissions granted</div>
          )}
          {needsGroups && access === 'consent' && (
            <div className="group-options-note">
              <KeyRound size={14} />
              <span>
                Creating groups needs <code>Group.ReadWrite.All</code>, <code>Device.Read.All</code> and <code>User.ReadBasic.All</code> (admin consent).
              </span>
              {onGrantAccess && (
                <button type="button" className="btn btn-outline btn-small" onClick={onGrantAccess}>Grant access</button>
              )}
            </div>
          )}
          {needsGroups && access === 'error' && (
            <div className="group-options-note group-options-note--error">
              <AlertTriangle size={14} /> Couldn't check group permissions: {String(accessError).slice(0, 160)}
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default AssignmentGroupOptions;
