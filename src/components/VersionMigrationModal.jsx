import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  ArrowRight, ArrowLeft, X, Loader, AlertTriangle, KeyRound, Users, Monitor, User,
  Search, CheckCircle, GitMerge, Info
} from 'lucide-react';
import graphAPI from '../utils/graph-api';
import { CONFIG } from '../utils/config';
import { summarizeAssignments } from '../utils/device-validation';
import { groupLabel } from '../utils/policy-status';
import { planVersionMigration, planHandOverTargets, groupTarget, loadGroupNaming } from '../utils/assignment-groups';

const nameOf = (policy) => policy?.displayName || policy?.name || '';

const MemberIcon = ({ type }) => {
  if (type === 'device') return <Monitor size={13} />;
  if (type === 'user') return <User size={13} />;
  return <Users size={13} />;
};

const describeAssignment = (summary, groupNames) => {
  if (!summary?.hasIncludes) return 'Not assigned';
  const parts = [
    summary.allDevices && 'All devices',
    summary.allUsers && 'All users',
    ...summary.includeGroupIds.map(id => groupLabel(id, groupNames)),
  ].filter(Boolean);
  const excluded = summary.excludeGroupIds.map(id => groupLabel(id, groupNames));
  return parts.join(', ') + (excluded.length ? ` — excluded: ${excluded.join(', ')}` : '');
};

const MemberRow = ({ member, checked, onToggle, extra }) => (
  <label className="mig-member">
    <input type="checkbox" checked={checked} onChange={onToggle} />
    <MemberIcon type={member.type} />
    <span className="mig-member-name">{member.name || `${member.type} ${member.id.slice(0, 8)}…`}</span>
    {member.detail && <span className="mig-member-detail">{member.detail}</span>}
    {extra}
  </label>
);

/**
 * Move devices/users from an old OIB policy version to the new one, a few at
 * a time, then hand the remaining assignment over. See planVersionMigration
 * for the two strategies.
 *
 * Props:
 *  - pair        { name, current, legacy: [] } from getMigrationPairs
 *  - naming      group naming settings
 *  - onClose(changed)
 *  - onGrantAccess  start the consent redirect for the group permissions
 */
const VersionMigrationModal = ({ pair, naming: namingProp, onClose, onGrantAccess }) => {
  // Fixed for the life of the dialog, so load() doesn't re-run on re-renders.
  const [naming] = useState(() => namingProp || loadGroupNaming());
  const [legacyId, setLegacyId] = useState(pair.legacy[0].id);
  const legacyPolicy = pair.legacy.find(p => p.id === legacyId) || pair.legacy[0];
  const currentPolicy = pair.current;

  const [data, setData] = useState({ loading: true });
  const [selectedOld, setSelectedOld] = useState(() => new Set());
  const [selectedNew, setSelectedNew] = useState(() => new Set());
  const [filter, setFilter] = useState('');
  const [searchQuery, setSearchQuery] = useState('');
  const [searchResults, setSearchResults] = useState(null);
  const [searching, setSearching] = useState(false);
  const [busy, setBusy] = useState(null);
  const [log, setLog] = useState([]);
  const [confirmFinish, setConfirmFinish] = useState(false);
  const [changed, setChanged] = useState(false);

  // Lines logged during one action are kept together, in the order they
  // happened; the newest action is shown first.
  const batchRef = useRef(null);
  const pushLogGroup = (lines) => setLog(prev =>
    [{ id: `${Date.now()}-${Math.random()}`, lines }, ...prev].slice(0, 8));
  const addLog = (kind, text) => {
    if (batchRef.current) batchRef.current.push({ kind, text });
    else pushLogGroup([{ kind, text }]);
  };

  const load = useCallback(async () => {
    setData(prev => ({ ...prev, loading: true, error: null }));
    try {
      await graphAPI.checkGroupManageAccess();
    } catch (error) {
      const consent = error?.name === 'ConsentRequiredError';
      setData({ loading: false, consent, error: consent ? null : error.message });
      return;
    }

    try {
      const [legacyRaw, currentRaw] = await Promise.all([
        graphAPI.getAssignmentsFor(legacyPolicy),
        graphAPI.getAssignmentsFor(currentPolicy),
      ]);
      const legacyAssignment = summarizeAssignments(legacyRaw);
      const currentAssignment = summarizeAssignments(currentRaw);
      const groupIds = [
        ...legacyAssignment.includeGroupIds, ...legacyAssignment.excludeGroupIds,
        ...currentAssignment.includeGroupIds, ...currentAssignment.excludeGroupIds,
      ];
      const groupNames = groupIds.length > 0
        ? await graphAPI.getGroupNames(groupIds, CONFIG.msal.groupManageScopes)
        : new Map();
      const plan = planVersionMigration({
        legacyPolicy, currentPolicy, legacyAssignment, currentAssignment, groupNames, naming,
      });

      const membersOf = (groupId) => (groupId ? graphAPI.getGroupMembers(groupId) : Promise.resolve([]));
      const [sourceLists, movedMembers, excludedMembers] = await Promise.all([
        Promise.all(plan.sourceGroups.map(group =>
          membersOf(group.id).then(members => members.map(m => ({ ...m, sourceGroupId: group.id }))))),
        membersOf(plan.currentIncludeGroup.id),
        plan.strategy === 'exclude' ? membersOf(plan.legacyExcludeGroup?.id) : Promise.resolve([]),
      ]);

      // A member can sit in several of the old version's groups.
      const sourceById = new Map();
      sourceLists.flat().forEach(member => {
        const entry = sourceById.get(member.id) || { ...member, sourceGroupIds: [] };
        entry.sourceGroupIds.push(member.sourceGroupId);
        sourceById.set(member.id, entry);
      });

      setData({
        loading: false,
        legacyRaw,
        legacyAssignment,
        currentAssignment,
        groupNames,
        plan,
        sourceMembers: [...sourceById.values()],
        movedMembers,
        excludedIds: new Set(excludedMembers.map(m => m.id)),
      });
      setSelectedOld(new Set());
      setSelectedNew(new Set());
    } catch (error) {
      setData({ loading: false, error: error.message });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [legacyPolicy?.id, currentPolicy?.id, naming]);

  useEffect(() => { load(); }, [load]);

  const { plan } = data;
  const movedIds = useMemo(() => new Set((data.movedMembers || []).map(m => m.id)), [data.movedMembers]);

  // Members still on the old version: its group members (minus moved or
  // excluded ones), plus search hits when it targets All devices/users.
  const remaining = useMemo(() => {
    const byId = new Map();
    [...(data.sourceMembers || []), ...(searchResults || [])].forEach(member => {
      if (movedIds.has(member.id) || data.excludedIds?.has(member.id)) return;
      if (!byId.has(member.id)) byId.set(member.id, member);
    });
    return [...byId.values()];
  }, [data.sourceMembers, data.excludedIds, searchResults, movedIds]);

  const matchesFilter = (member) => {
    const q = filter.trim().toLowerCase();
    if (!q) return true;
    return [member.name, member.detail, member.id].some(v => String(v || '').toLowerCase().includes(q));
  };
  const visibleRemaining = remaining.filter(matchesFilter);

  const toggle = (setter) => (id) => setter(prev => {
    const next = new Set(prev);
    if (next.has(id)) next.delete(id); else next.add(id);
    return next;
  });
  const toggleOld = toggle(setSelectedOld);
  const toggleNew = toggle(setSelectedNew);

  // Create the new version's include group (and, for the exclude strategy,
  // the old version's exclude group) the first time something is moved.
  const ensureGroupsReady = async () => {
    let currentIncludeId = plan.currentIncludeGroup.id;
    if (!currentIncludeId) {
      const group = await graphAPI.ensureGroup(
        plan.currentIncludeGroup.name,
        `Include group for the Intune policy "${plan.currentName}". Created by OIB Deployer.`
      );
      if (group.dynamic) throw new Error(`"${group.displayName}" is a dynamic group — members can't be added to it.`);
      await graphAPI.addAssignmentTargets(currentPolicy, [groupTarget(group.id, 'include')]);
      currentIncludeId = group.id;
      addLog('ok', `${group.created ? 'Created' : 'Reused'} "${group.displayName}" and assigned it to the new version.`);
    }

    let legacyExcludeId = plan.legacyExcludeGroup?.id ?? null;
    if (plan.strategy === 'exclude' && !legacyExcludeId) {
      const group = await graphAPI.ensureGroup(
        plan.legacyExcludeGroup.name,
        `Exclude group for the Intune policy "${plan.legacyName}". Created by OIB Deployer.`
      );
      if (group.dynamic) throw new Error(`"${group.displayName}" is a dynamic group — members can't be added to it.`);
      await graphAPI.addAssignmentTargets(legacyPolicy, [groupTarget(group.id, 'exclude')]);
      legacyExcludeId = group.id;
      addLog('ok', `${group.created ? 'Created' : 'Reused'} "${group.displayName}" and excluded it from the old version.`);
    }
    return { currentIncludeId, legacyExcludeId };
  };

  const reportResults = (verb, results) => {
    const failed = results.filter(r => !r.ok);
    if (results.length - failed.length > 0) addLog('ok', `${verb} ${results.length - failed.length} member(s).`);
    if (failed.length > 0) addLog('error', `${failed.length} member(s) failed: ${failed[0].error}`);
    return results.filter(r => r.ok).map(r => r.id);
  };

  // New version first, then old: a member briefly gets both rather than neither.
  const doMove = async (members) => {
    const { currentIncludeId, legacyExcludeId } = await ensureGroupsReady();
    const okIds = reportResults('Added to the new version:',
      await graphAPI.addGroupMembers(currentIncludeId, members.map(m => m.id)));
    if (okIds.length === 0) return;
    const ok = new Set(okIds);

    if (plan.strategy === 'move') {
      const bySource = new Map();
      members.filter(m => ok.has(m.id)).forEach(member => {
        (member.sourceGroupIds || []).forEach(groupId => {
          bySource.set(groupId, [...(bySource.get(groupId) || []), member.id]);
        });
      });
      for (const [groupId, ids] of bySource) {
        reportResults(`Removed from "${groupLabel(groupId, data.groupNames)}":`, await graphAPI.removeGroupMembers(groupId, ids));
      }
    } else {
      reportResults('Excluded from the old version:', await graphAPI.addGroupMembers(legacyExcludeId, okIds));
    }
  };

  const run = async (label, action) => {
    setBusy(label);
    batchRef.current = [];
    try {
      await action();
    } catch (error) {
      addLog('error', error.message.slice(0, 300));
    } finally {
      const lines = batchRef.current;
      batchRef.current = null;
      if (lines.length > 0) pushLogGroup(lines);
      setBusy(null);
      setChanged(true);
      await load();
    }
  };

  const moveSelected = () => run('Moving…', () =>
    doMove(remaining.filter(m => selectedOld.has(m.id))));

  const moveBack = () => run('Moving back…', async () => {
    const ids = [...selectedNew];
    if (plan.strategy === 'move') {
      const source = plan.sourceGroups[0];
      reportResults(`Returned to "${groupLabel(source.id, data.groupNames)}":`, await graphAPI.addGroupMembers(source.id, ids));
    } else if (plan.legacyExcludeGroup?.id) {
      reportResults('No longer excluded from the old version:', await graphAPI.removeGroupMembers(plan.legacyExcludeGroup.id, ids));
    }
    reportResults('Removed from the new version:', await graphAPI.removeGroupMembers(plan.currentIncludeGroup.id, ids));
  });

  // What the new version takes over from the old one when the migration is
  // finished: everything except the old version's own exclude group.
  const handOverTargets = () => planHandOverTargets(data.legacyRaw, [plan.legacyExcludeGroup?.id]);

  const finish = () => run('Finishing…', async () => {
    setConfirmFinish(false);
    if (plan.strategy === 'move') {
      if (remaining.length > 0) await doMove(remaining);
    } else {
      const targets = handOverTargets();
      const { added } = await graphAPI.addAssignmentTargets(currentPolicy, targets);
      addLog('ok', `Gave the new version the old version's ${targets.length} assignment target(s) (${added} new).`);
    }
    await graphAPI.setAssignmentsFor(legacyPolicy, []);
    addLog('ok', `Removed all assignments from "${plan.legacyName}". You can delete it in Intune when you're ready.`);
  });

  const runSearch = async (event) => {
    event.preventDefault();
    if (!searchQuery.trim()) return;
    setSearching(true);
    try {
      setSearchResults(await graphAPI.searchDirectory(searchQuery, plan.searchTypes));
    } catch (error) {
      addLog('error', `Search failed: ${error.message.slice(0, 200)}`);
    } finally {
      setSearching(false);
    }
  };

  const close = () => onClose(changed);

  const renderBody = () => {
    if (data.loading && !plan) {
      return <div className="mig-state"><Loader size={16} className="spinning" /> Loading assignments and group members…</div>;
    }
    if (data.consent) {
      return (
        <div className="mig-state">
          <KeyRound size={16} />
          <span>
            Migrating needs <code>Group.ReadWrite.All</code>, <code>Device.Read.All</code> and <code>User.ReadBasic.All</code> (admin consent).
          </span>
          {onGrantAccess && <button className="btn btn-outline btn-small" onClick={onGrantAccess}>Grant access</button>}
        </div>
      );
    }
    if (data.error) {
      return <div className="mig-state mig-state--error"><AlertTriangle size={16} /> {data.error.slice(0, 400)}</div>;
    }

    const groupStatus = (group) => group?.id ? 'exists' : 'will be created on the first move';

    return (
      <>
        <div className="mig-versions">
          <div className="mig-version">
            <span className="pas-label">Old version</span>
            {pair.legacy.length > 1 ? (
              <select value={legacyId} onChange={e => setLegacyId(e.target.value)} disabled={!!busy}>
                {pair.legacy.map(p => <option key={p.id} value={p.id}>{nameOf(p)}</option>)}
              </select>
            ) : <strong>{nameOf(legacyPolicy)}</strong>}
            <span className="mig-assigned">Assigned to: {describeAssignment(data.legacyAssignment, data.groupNames)}</span>
          </div>
          <ArrowRight size={18} className="mig-arrow" />
          <div className="mig-version">
            <span className="pas-label">New version</span>
            <strong>{nameOf(currentPolicy)}</strong>
            <span className="mig-assigned">Assigned to: {describeAssignment(data.currentAssignment, data.groupNames)}</span>
          </div>
        </div>

        <div className="mig-plan">
          <Info size={14} />
          <div>
            <div>{plan.reason}</div>
            {plan.strategy !== 'none' && (
              <ul>
                <li>New version's include group: <code>{plan.currentIncludeGroup.name}</code> ({groupStatus(plan.currentIncludeGroup)})</li>
                {plan.strategy === 'exclude' && (
                  <li>Old version's exclude group: <code>{plan.legacyExcludeGroup.name}</code> ({groupStatus(plan.legacyExcludeGroup)})</li>
                )}
              </ul>
            )}
          </div>
        </div>

        {plan.strategy !== 'none' && (
          <div className="mig-columns">
            <div className="mig-column">
              <div className="mig-column-header">
                <h4>Still on the old version ({remaining.length})</h4>
                {/* With All devices/users there is no list to filter, only search. */}
                {plan.sourceGroups.length > 0 && (
                  <input type="search" placeholder="Filter…" value={filter} onChange={e => setFilter(e.target.value)} />
                )}
              </div>
              {plan.searchTypes.length > 0 && (
                <form className="mig-search" onSubmit={runSearch}>
                  <input
                    type="search"
                    value={searchQuery}
                    onChange={e => setSearchQuery(e.target.value)}
                    placeholder={`Find ${plan.searchTypes.map(t => (t === 'device' ? 'devices' : 'users')).join(' or ')} by name…`}
                  />
                  <button type="submit" className="btn btn-secondary btn-small" disabled={searching || !searchQuery.trim()}>
                    {searching ? <Loader size={13} className="spinning" /> : <Search size={13} />} Search
                  </button>
                </form>
              )}
              <div className="mig-list">
                {visibleRemaining.length === 0 && (
                  <div className="pas-muted mig-empty">
                    {plan.searchTypes.length > 0 && !searchResults
                      ? `The old version targets ${plan.searchTypes.includes('device') ? 'All devices' : 'All users'} — search to pick who moves first.`
                      : 'Nothing left to move.'}
                  </div>
                )}
                {visibleRemaining.map(member => (
                  <MemberRow key={member.id} member={member} checked={selectedOld.has(member.id)} onToggle={() => toggleOld(member.id)} />
                ))}
              </div>
              <button
                className="btn btn-primary btn-small"
                disabled={!!busy || selectedOld.size === 0}
                onClick={moveSelected}
              >
                Move {selectedOld.size || ''} to new version <ArrowRight size={14} />
              </button>
            </div>

            <div className="mig-column">
              <div className="mig-column-header">
                <h4>On the new version ({data.movedMembers.length})</h4>
              </div>
              <div className="mig-list">
                {data.movedMembers.length === 0 && <div className="pas-muted mig-empty">No one has been moved yet.</div>}
                {data.movedMembers.map(member => (
                  <MemberRow key={member.id} member={member} checked={selectedNew.has(member.id)} onToggle={() => toggleNew(member.id)} />
                ))}
              </div>
              <button
                className="btn btn-secondary btn-small"
                disabled={!!busy || selectedNew.size === 0}
                onClick={moveBack}
              >
                <ArrowLeft size={14} /> Move {selectedNew.size || ''} back
              </button>
            </div>
          </div>
        )}

        {plan.strategy !== 'none' && (
          <div className="mig-finish">
            {!confirmFinish ? (
              <button className="btn btn-outline btn-small" disabled={!!busy} onClick={() => setConfirmFinish(true)}>
                <GitMerge size={14} /> Finish migration…
              </button>
            ) : (
              <div className="mig-confirm">
                <AlertTriangle size={16} />
                <div>
                  <strong>Finish the migration?</strong>
                  {plan.strategy === 'move' ? (
                    <p>
                      {remaining.length > 0
                        ? <>Moves the remaining {remaining.length} member(s) to <code>{plan.currentIncludeGroup.name}</code>, then removes</>
                        : 'Everyone has been moved. Removes'}
                      {' '}every assignment from <em>{plan.legacyName}</em>.
                    </p>
                  ) : (
                    <p>
                      Also assigns <em>{plan.currentName}</em> to{' '}
                      <strong>{describeAssignment(summarizeAssignments(handOverTargets().map(target => ({ target }))), data.groupNames)}</strong>
                      , as the old version is today, then removes every assignment from <em>{plan.legacyName}</em>.
                      The new version keeps the groups it already has.
                    </p>
                  )}
                  <p className="pas-muted">The old policy itself is not deleted.</p>
                  <div className="mig-confirm-actions">
                    <button className="btn btn-secondary btn-small" onClick={() => setConfirmFinish(false)}>Cancel</button>
                    <button className="btn btn-primary btn-small" onClick={finish}>Finish migration</button>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {log.length > 0 && (
          <ul className="mig-log">
            {log.map(group => (
              <li key={group.id} className="mig-log-group">
                {group.lines.map((line, index) => (
                  <div key={index} className={`mig-log-line mig-log-${line.kind}`}>
                    {line.kind === 'ok' ? <CheckCircle size={13} /> : <AlertTriangle size={13} />} {line.text}
                  </div>
                ))}
              </li>
            ))}
          </ul>
        )}
      </>
    );
  };

  return (
    <div className="modal-overlay" onClick={busy ? undefined : close} role="presentation">
      <div
        className="modal-content migration-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="migration-modal-title"
        onClick={event => event.stopPropagation()}
      >
        <div className="modal-header">
          <h2 id="migration-modal-title"><GitMerge size={20} /> Migrate to the new version</h2>
          {/* In the header, so the body doesn't shift while an action runs. */}
          <div className="mig-header-actions">
            {busy && <span className="mig-busy" role="status"><Loader size={14} className="spinning" /> {busy}</span>}
            <button className="btn btn-ghost btn-small" onClick={close} disabled={!!busy} aria-label="Close"><X size={18} /></button>
          </div>
        </div>
        <div className="modal-body">
          {renderBody()}
        </div>
      </div>
    </div>
  );
};

export default VersionMigrationModal;
