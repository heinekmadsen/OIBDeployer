import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  ArrowRight, ArrowLeft, X, Loader, AlertTriangle, KeyRound, Users, Monitor, User,
  Search, CheckCircle, GitMerge, Info
} from 'lucide-react';
import graphAPI from '../utils/graph-api';
import { CONFIG } from '../utils/config';
import { summarizeAssignments } from '../utils/device-validation';
import { groupLabel } from '../utils/policy-status';
import {
  planVersionMigration, planFinishMigration, groupTarget, loadGroupNaming, builtInsOf, builtInLabel, isDedicatedGroup,
  mergeAssignmentTargets,
  defaultFinishModel, supportsBuiltInTargets, withoutBuiltIns, policyPlatform, PLATFORM_LABELS
} from '../utils/assignment-groups';

const builtInList = (kinds) => kinds.map(builtInLabel).join(' and ');

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
  const [finishModel, setFinishModel] = useState(null);
  const [confirmPilot, setConfirmPilot] = useState(false);
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
      const [sourceLists, movedMembers, excludedMembers, includeMatches] = await Promise.all([
        Promise.all(plan.sourceGroups.map(group =>
          membersOf(group.id).then(members => members.map(m => ({ ...m, sourceGroupId: group.id }))))),
        membersOf(plan.currentPilotGroup.id),
        membersOf(plan.legacyExcludeGroup?.id),
        // An include group with the conventional name may exist without being
        // assigned yet; finishing with groups then merges into it.
        plan.strategy !== 'none' && !plan.currentIncludeGroup.id
          ? graphAPI.findGroupsByName(plan.currentIncludeGroup.name)
          : Promise.resolve([]),
      ]);
      const existingInclude = includeMatches.find(g => g.securityEnabled) || includeMatches[0] || null;

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
        currentRaw,
        legacyAssignment,
        currentAssignment,
        groupNames,
        plan,
        sourceMembers: [...sourceById.values()],
        movedMembers,
        excludedIds: new Set(excludedMembers.map(m => m.id)),
        existingIncludeGroup: existingInclude
          ? { id: existingInclude.id, dynamic: (existingInclude.groupTypes || []).includes('DynamicMembership') }
          : null,
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

  // The new version is already on All devices/users: it reaches everyone, so
  // piloting is impossible and everyone the old version reaches gets both.
  const currentBuiltIns = builtInsOf(data.currentAssignment);
  const coversAll = !!plan && plan.strategy !== 'none' && currentBuiltIns.length > 0;
  const builtInAllowed = supportsBuiltInTargets(currentPolicy?.endpoint);

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

  // Create the new version's pilot group (and, for the exclude strategy,
  // the old version's exclude group) the first time something is moved.
  // Create (or reuse) a group named after the new version and assign it.
  const ensureAssignedGroup = async (name, description) => {
    const group = await graphAPI.ensureGroup(name, description);
    if (group.dynamic) throw new Error(`"${group.displayName}" is a dynamic group — members can't be added to it.`);
    await graphAPI.addAssignmentTargets(currentPolicy, [groupTarget(group.id, 'include')]);
    addLog('ok', `${group.created ? 'Created' : 'Reused'} "${group.displayName}" and assigned it to the new version.`);
    return group.id;
  };

  const includeDescription = () => `Include group for the Intune policy "${plan.currentName}". Created by OIB Deployer.`;

  const ensurePilotGroup = async () => plan.currentPilotGroup.id || ensureAssignedGroup(
    plan.currentPilotGroup.name,
    `Pilot group for the Intune policy "${plan.currentName}" during a version migration. Created by OIB Deployer.`
  );

  const ensureGroupsReady = async () => {
    const currentIncludeId = await ensurePilotGroup();

    let legacyExcludeId = plan.legacyExcludeGroup?.id ?? null;
    if (!legacyExcludeId) {
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

  // New version first, then old: a member briefly gets both rather than
  // neither. The old version's include groups are left as they are.
  const doMove = async (members) => {
    const { currentIncludeId, legacyExcludeId } = await ensureGroupsReady();
    const okIds = reportResults('Added to the pilot group:',
      await graphAPI.addGroupMembers(currentIncludeId, members.map(m => m.id)));
    if (okIds.length === 0) return;
    reportResults('Excluded from the old version:', await graphAPI.addGroupMembers(legacyExcludeId, okIds));
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

  // Old version first, then new. A member the old version wouldn't reach
  // (not in any of its groups, e.g. moved before exclusions were used) goes
  // back into the old version's own group, or stays on the new version.
  const moveBack = () => run('Moving back…', async () => {
    const ids = [...selectedNew];
    const broad = data.legacyAssignment?.allDevices || data.legacyAssignment?.allUsers;
    const inOldGroups = new Set((data.sourceMembers || []).map(m => m.id));
    const unreached = broad ? [] : ids.filter(id => !inOldGroups.has(id));
    const stay = new Set();
    if (unreached.length > 0) {
      const own = plan.sourceGroups.find(g => g.name && isDedicatedGroup(g.name, plan.legacyName));
      if (own) {
        const returned = new Set(reportResults(`Returned to "${own.name}":`, await graphAPI.addGroupMembers(own.id, unreached)));
        unreached.filter(id => !returned.has(id)).forEach(id => stay.add(id));
      } else {
        unreached.forEach(id => stay.add(id));
        addLog('error', `${unreached.length} member(s) aren't in any of the old version's groups, so they stay on the new version.`);
      }
    }
    const back = ids.filter(id => !stay.has(id));
    if (back.length === 0) return;
    if (plan.legacyExcludeGroup?.id) {
      reportResults('No longer excluded from the old version:', await graphAPI.removeGroupMembers(plan.legacyExcludeGroup.id, back));
    }
    reportResults('Removed from the pilot group:', await graphAPI.removeGroupMembers(plan.currentPilotGroup.id, back));
  });

  const finishPlan = (model) => planFinishMigration({
    model,
    plan,
    legacyAssignments: data.legacyRaw,
    currentAssignments: data.currentRaw,
    legacyAssignment: data.legacyAssignment,
    groupNames: data.groupNames,
    existingIncludeGroupId: data.existingIncludeGroup?.id ?? null,
  });

  const openFinish = () => {
    setFinishModel(coversAll ? 'builtIn' : defaultFinishModel(data.legacyAssignment, currentPolicy?.endpoint));
    setConfirmPilot(false);
    setConfirmFinish(true);
  };

  // Take every member out of the pilot group once it's no longer assigned.
  const emptyPilotGroup = async () => {
    const ids = (data.movedMembers || []).map(m => m.id);
    if (!plan.currentPilotGroup.id || ids.length === 0) return;
    reportResults(`Removed from the pilot group "${plan.currentPilotGroup.name}" (the empty group is kept):`,
      await graphAPI.removeGroupMembers(plan.currentPilotGroup.id, ids));
  };

  // The new version's include group for the group model: the pilot group
  // renamed, an existing include group, or a new one.
  const prepareIncludeGroup = async (fin) => {
    const includeName = plan.currentIncludeGroup.name;
    switch (fin.includeAction) {
      case 'rename':
        await graphAPI.renameGroup(fin.pilotGroupId, includeName, includeDescription());
        addLog('ok', `Renamed "${plan.currentPilotGroup.name}" to "${includeName}".`);
        return fin.pilotGroupId;
      case 'merge':
        if (!fin.includeAssigned) {
          if (data.existingIncludeGroup?.dynamic) throw new Error(`"${includeName}" is a dynamic group — members can't be added to it.`);
          await graphAPI.addAssignmentTargets(currentPolicy, [groupTarget(fin.includeGroupId, 'include')]);
          addLog('ok', `Assigned the existing group "${includeName}" to the new version.`);
        }
        return fin.includeGroupId;
      case 'use':
        return fin.includeGroupId;
      default:
        return ensureAssignedGroup(includeName, includeDescription());
    }
  };

  // Group model: everyone left on the old version joins the new version's
  // include group. If anyone can't be added, stop before the old version is
  // unassigned, so no one is left without the policy.
  const finishWithGroups = async (fin) => {
    const includeId = await prepareIncludeGroup(fin);
    setBusy('Collecting members…');
    const [excludedLists, audience] = await Promise.all([
      Promise.all(fin.excludedGroupIds.map(id => graphAPI.getGroupMembers(id))),
      fin.audience.length > 0 ? graphAPI.listAudience(fin.audience, policyPlatform(currentPolicy)) : Promise.resolve([]),
    ]);
    const excluded = new Set(excludedLists.flat().map(m => m.id));
    const ownGroups = new Set(fin.ownIncludeGroupIds.map(id => id.toLowerCase()));
    const fromOwnGroups = (data.sourceMembers || [])
      .filter(m => (m.sourceGroupIds || []).some(id => ownGroups.has(id.toLowerCase())));
    const merging = fin.includeAction === 'merge';
    // Pilot members are excluded from the old version, so they are added
    // without that filter; after a rename they're already in the group.
    const ids = [...new Set([
      ...(merging ? [...movedIds] : []),
      ...[...fromOwnGroups, ...audience].map(m => m.id).filter(id => !movedIds.has(id) && !excluded.has(id)),
    ])];

    if (ids.length > 0) {
      setBusy(`Adding ${ids.length} member(s)…`);
      const results = await graphAPI.addGroupMembers(includeId, ids);
      reportResults(`Added to "${plan.currentIncludeGroup.name}":`, results);
      const failed = results.filter(r => !r.ok).length;
      if (failed > 0) {
        throw new Error(`Not finished: ${failed} member(s) couldn't be added to the new version's group, so the old version is still assigned. Try again.`);
      }
    }
    if (merging) {
      const pilotId = fin.pilotGroupId.toLowerCase();
      const fresh = await graphAPI.getAssignmentsFor(currentPolicy);
      await graphAPI.setAssignmentsFor(currentPolicy, fresh
        .filter(a => String(a?.target?.groupId ?? '').toLowerCase() !== pilotId)
        .map(a => ({ target: { ...a.target } })));
      addLog('ok', `Removed the pilot group "${plan.currentPilotGroup.name}" from the new version.`);
      await emptyPilotGroup();
    }
    if (fin.handOver.length > 0) {
      const { added } = await graphAPI.addAssignmentTargets(currentPolicy, fin.handOver);
      addLog('ok', `Gave the new version ${fin.sharedIncludes} shared group(s) and ${fin.carriedExclusions} exclusion(s) from the old version (${added} new).`);
    }
  };

  // Members excluded from the old version outside the migration (they never
  // joined the pilot group) stay excluded: they move to the new version's
  // exclude group before the old one is emptied.
  const carryExclusions = async (ids) => {
    let groupId = plan.currentExcludeGroup.id;
    if (!groupId) {
      const group = await graphAPI.ensureGroup(
        plan.currentExcludeGroup.name,
        `Exclude group for the Intune policy "${plan.currentName}". Created by OIB Deployer.`
      );
      if (group.dynamic) throw new Error(`"${group.displayName}" is a dynamic group — members can't be added to it.`);
      groupId = group.id;
      addLog('ok', `${group.created ? 'Created' : 'Reused'} "${group.displayName}".`);
    }
    const results = await graphAPI.addGroupMembers(groupId, ids);
    reportResults(`Still excluded, now from the new version through "${plan.currentExcludeGroup.name}":`, results);
    if (results.some(r => !r.ok)) {
      throw new Error('Not finished: not every exclusion could be carried over to the new version, so the old version is still assigned. Try again.');
    }
    return groupTarget(groupId, 'exclude');
  };

  // Once the old version is unassigned its own groups serve no purpose, so
  // everyone is taken out. The empty groups are kept (security groups can't
  // be restored once deleted) and shared groups are never touched.
  const emptyOldGroups = async (groupIds) => {
    let emptied = 0;
    for (const id of groupIds) {
      const name = groupLabel(id, data.groupNames);
      try {
        const members = await graphAPI.getGroupMembers(id);
        if (members.length === 0) continue;
        reportResults(`Removed from "${name}":`, await graphAPI.removeGroupMembers(id, members.map(m => m.id)));
        emptied += 1;
      } catch (error) {
        addLog('error', `Couldn't empty "${name}": ${error.message.slice(0, 200)}`);
      }
    }
    if (emptied > 0) addLog('ok', 'The old version\'s groups are kept, empty. Delete them in Entra ID when you\'re ready.');
  };

  // New version first, then old: everyone briefly gets both rather than neither.
  const finish = () => run('Finishing…', async () => {
    setConfirmFinish(false);
    const fin = finishPlan(finishModel);
    // While the new version is on All devices, members excluded from the old
    // one are the ones already migrated, not exclusions to keep.
    const heldBack = coversAll ? [] : alreadyMoved;
    const carried = heldBack.length > 0 ? [await carryExclusions(heldBack)] : [];
    if (fin.model === 'builtIn') {
      await graphAPI.setAssignmentsFor(currentPolicy, mergeAssignmentTargets(fin.assignments, carried).assignments);
      addLog('ok', `Assigned "${plan.currentName}" to ${builtInList(fin.builtIns)}.`);
      if (fin.removesPilotGroup) {
        addLog('ok', `Removed the pilot group "${plan.currentPilotGroup.name}" from it.`);
      }
      if (fin.carriedExclusions > 0) addLog('ok', `Took over ${fin.carriedExclusions} exclusion(s) from the old version.`);
      await emptyPilotGroup();
    } else {
      if (carried.length > 0) await graphAPI.addAssignmentTargets(currentPolicy, carried);
      await finishWithGroups(fin);
    }
    await graphAPI.setAssignmentsFor(legacyPolicy, []);
    addLog('ok', `Removed all assignments from "${plan.legacyName}". You can delete it in Intune when you're ready.`);
    await emptyOldGroups(fin.ownLegacyGroupIds);
  });

  // Members already excluded from the old version only have the new one
  // through All devices/users; they must keep it once that's removed.
  const alreadyMoved = useMemo(
    () => [...(data.excludedIds ?? [])].filter(id => !movedIds.has(id)),
    [data.excludedIds, movedIds]
  );

  // Take the new version off All devices/users so it can be piloted. Members
  // already excluded from the old version join the pilot group first.
  const pilotInstead = () => run('Removing…', async () => {
    setConfirmPilot(false);
    if (alreadyMoved.length > 0) {
      const pilotId = await ensurePilotGroup();
      const results = await graphAPI.addGroupMembers(pilotId, alreadyMoved);
      reportResults('Kept on the new version:', results);
      if (results.some(r => !r.ok)) {
        throw new Error(`Stopped: not everyone already moved could be added to the pilot group, so ${builtInList(currentBuiltIns)} is still assigned. Try again.`);
      }
    }
    // Read again: the pilot group may just have been assigned.
    await graphAPI.setAssignmentsFor(currentPolicy, withoutBuiltIns(await graphAPI.getAssignmentsFor(currentPolicy)));
    addLog('ok', `Removed ${builtInList(currentBuiltIns)} from "${plan.currentName}". Everyone else stays on the old version until you move them.`);
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

  // Finish: pick how the new version is assigned once the old one is gone.
  const renderFinishChoice = () => {
    const builtInFin = finishPlan('builtIn');
    const groupsFin = finishPlan('groups');
    const includeName = plan.currentIncludeGroup.name;
    const pilotName = plan.currentPilotGroup.name;
    const pilotCount = (data.movedMembers || []).length;
    const includeStep = {
      rename: <>Renames the pilot group <code>{pilotName}</code> to <code>{includeName}</code>, so its {pilotCount} member(s) stay on the new version.</>,
      merge: <>
        The pilot group&apos;s {pilotCount} member(s) join <code>{includeName}</code>
        {!groupsFin.includeAssigned && ' (an existing group, assigned now)'}; the pilot group is unassigned and emptied, and the empty group is kept.
      </>,
      use: <>Keeps it on <code>{includeName}</code>.</>,
      create: <>Assigns it to a new group <code>{includeName}</code>.</>,
    }[groupsFin.includeAction];
    const ownGroups = new Set(groupsFin.ownIncludeGroupIds.map(id => id.toLowerCase()));
    const ownLeft = remaining.filter(m => (m.sourceGroupIds || []).some(id => ownGroups.has(id.toLowerCase()))).length;
    const platformLabel = PLATFORM_LABELS[policyPlatform(currentPolicy)];
    const audienceText = [
      groupsFin.audience.includes('device') && `every ${platformLabel ? `${platformLabel} ` : ''}device in Entra ID`,
      groupsFin.audience.includes('user') && 'every member user (guests are skipped)',
    ].filter(Boolean).join(' and ');
    const chosenAllowed = finishModel === 'builtIn' ? builtInAllowed : finishModel === 'groups' && !coversAll;

    const modelOption = (value, title, disabled, body) => (
      <label className={`mig-model ${finishModel === value ? 'mig-model--on' : ''} ${disabled ? 'mig-model--off' : ''}`}>
        <input
          type="radio"
          name="finish-model"
          value={value}
          checked={finishModel === value}
          disabled={disabled}
          onChange={() => setFinishModel(value)}
        />
        <div>
          <strong>{title}</strong>
          {body}
        </div>
      </label>
    );

    return (
      <div className="mig-confirm">
        <GitMerge size={16} />
        <div>
          <strong>Finish the migration</strong>
          <p>How should <em>{plan.currentName}</em> be assigned once the old version is removed?</p>
          <div className="mig-models" role="radiogroup" aria-label="Assignment after the migration">
            {modelOption('builtIn', builtInList(builtInFin.builtIns), !builtInAllowed, builtInAllowed ? (
              <>
                <span>
                  Assigns it to {builtInList(builtInFin.builtIns)}
                  {builtInFin.removesPilotGroup && (
                    <> and removes the pilot group <code>{pilotName}</code> from it{pilotCount > 0 && `, taking its ${pilotCount} member(s) out`} (the empty group is kept)</>
                  )}
                  . New devices get it automatically.
                  {builtInFin.carriedExclusions > 0 && ` It takes over ${builtInFin.carriedExclusions} exclusion(s) from the old version.`}
                </span>
                {builtInFin.widensScope && !coversAll && (
                  <span className="mig-model-warn">The old version was only assigned to groups, so this reaches more than it did.</span>
                )}
              </>
            ) : (
              <span>Intune can&apos;t assign this policy type to All devices or All users.</span>
            ))}
            {modelOption('groups', 'Include/exclude groups', coversAll, coversAll ? (
              <span>The new version is on {builtInList(currentBuiltIns)}. Use &quot;Pilot instead&quot; first to switch it to groups.</span>
            ) : (
              <>
                <span>
                  {includeStep}
                  {ownLeft > 0 && ` The ${ownLeft} member(s) left in the old version's own group join it too.`}
                  {groupsFin.sharedIncludes > 0 && ` ${groupsFin.sharedIncludes} shared group(s) are assigned to it as they are.`}
                  {groupsFin.carriedExclusions > 0 && ` It takes over ${groupsFin.carriedExclusions} exclusion(s) from the old version.`}
                </span>
                {groupsFin.audience.length > 0 && (
                  <span className="mig-model-warn">
                    The old version reaches everyone through {builtInList(builtInsOf(data.legacyAssignment))}, so {audienceText}{' '}
                    is added to the group, except excluded ones. This is a snapshot: {groupsFin.audience.includes('device') ? 'devices enrolled' : 'users added'}
                    {' '}later must be added by hand.
                  </span>
                )}
              </>
            ))}
          </div>
          {!coversAll && alreadyMoved.length > 0 && (
            <p>
              The {alreadyMoved.length} member(s) excluded from the old version outside the migration stay excluded: they join{' '}
              <code>{plan.currentExcludeGroup.name}</code> on the new version.
            </p>
          )}
          <p className="pas-muted">
            Then every assignment is removed from <em>{plan.legacyName}</em>
            {builtInFin.ownLegacyGroupIds.length > 0 && (
              <>
                , and everyone is taken out of its own group(s){' '}
                {builtInFin.ownLegacyGroupIds.map((id, index) => (
                  <React.Fragment key={id}>{index > 0 && ', '}<code>{groupLabel(id, data.groupNames)}</code></React.Fragment>
                ))}
                . The empty groups are kept; shared groups aren&apos;t changed
              </>
            )}
            . The old policy itself is not deleted.
          </p>
          <div className="mig-confirm-actions">
            <button className="btn btn-secondary btn-small" onClick={() => setConfirmFinish(false)}>Cancel</button>
            <button className="btn btn-primary btn-small" disabled={!!busy || !chosenAllowed} onClick={finish}>Finish migration</button>
          </div>
        </div>
      </div>
    );
  };

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

        {coversAll ? (
          <div className="mig-confirm">
            <AlertTriangle size={16} />
            <div>
              <strong>The new version is already assigned to {builtInList(currentBuiltIns)}</strong>
              <p>
                It already reaches everyone, so everyone the old version still reaches gets both versions, and there is
                nothing to pilot. Either take {builtInList(currentBuiltIns)} off the new version and pilot it on a few
                devices first, or finish the migration right away.
              </p>
              {!confirmPilot ? (
                <div className="mig-confirm-actions">
                  <button className="btn btn-outline btn-small" disabled={!!busy} onClick={() => { setConfirmPilot(true); setConfirmFinish(false); }}>
                    Pilot instead…
                  </button>
                </div>
              ) : (
                <>
                  <p>
                    Removes {builtInList(currentBuiltIns)} from <em>{plan.currentName}</em>; its other assignments stay.
                    {alreadyMoved.length > 0 && (
                      <> The {alreadyMoved.length} member(s) already excluded from the old version first join the pilot group{' '}
                        <code>{plan.currentPilotGroup.name}</code>, so they keep the new version.</>
                    )}
                    {' '}Everyone else keeps the old version until you move them over.
                  </p>
                  <div className="mig-confirm-actions">
                    <button className="btn btn-secondary btn-small" onClick={() => setConfirmPilot(false)}>Cancel</button>
                    <button className="btn btn-primary btn-small" disabled={!!busy} onClick={pilotInstead}>
                      Remove {builtInList(currentBuiltIns)}
                    </button>
                  </div>
                </>
              )}
            </div>
          </div>
        ) : (
          <div className="mig-plan">
            <Info size={14} />
            <div>
              <div>{plan.reason}</div>
              {plan.strategy !== 'none' && (
                <ul>
                  <li>New version&apos;s pilot group: <code>{plan.currentPilotGroup.name}</code> ({groupStatus(plan.currentPilotGroup)})</li>
                  {plan.strategy === 'exclude' && (
                    <li>Old version's exclude group: <code>{plan.legacyExcludeGroup.name}</code> ({groupStatus(plan.legacyExcludeGroup)})</li>
                  )}
                </ul>
              )}
            </div>
          </div>
        )}

        {plan.strategy !== 'none' && !coversAll && (
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
              <button className="btn btn-outline btn-small" disabled={!!busy} onClick={openFinish}>
                <GitMerge size={14} /> Finish migration…
              </button>
            ) : renderFinishChoice()}
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
