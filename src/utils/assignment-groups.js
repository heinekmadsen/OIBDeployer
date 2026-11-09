// Assignment groups: naming, assignment payloads and version-migration
// planning. Pure functions only — Graph calls live in graph-api.js.

export const DEFAULT_GROUP_NAMING = Object.freeze({
  prefix: '',
  suffix: '',
  includeLabel: 'Include',
  excludeLabel: 'Exclude',
  // Version migrations pilot a new version through its own group, which
  // becomes the include group (or is emptied) when the migration finishes.
  pilotLabel: 'Pilot',
});

export const MAX_GROUP_NAME_LENGTH = 256;
const STORAGE_KEY = 'oib-group-naming';

// Naming is a per-viewer convenience, so browser storage is fine; it may be
// unavailable (private mode, blocked storage), hence the try/catch.
export function loadGroupNaming() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return { ...DEFAULT_GROUP_NAMING };
    const parsed = JSON.parse(raw);
    return {
      prefix: typeof parsed.prefix === 'string' ? parsed.prefix : '',
      suffix: typeof parsed.suffix === 'string' ? parsed.suffix : '',
      includeLabel: typeof parsed.includeLabel === 'string' ? parsed.includeLabel : DEFAULT_GROUP_NAMING.includeLabel,
      excludeLabel: typeof parsed.excludeLabel === 'string' ? parsed.excludeLabel : DEFAULT_GROUP_NAMING.excludeLabel,
      pilotLabel: typeof parsed.pilotLabel === 'string' && parsed.pilotLabel.trim() ? parsed.pilotLabel : DEFAULT_GROUP_NAMING.pilotLabel,
    };
  } catch {
    return { ...DEFAULT_GROUP_NAMING };
  }
}

export function saveGroupNaming(naming) {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(naming));
  } catch {
    // Storage unavailable — naming just won't be remembered.
  }
}

export function policyDisplayName(name) {
  return String(name ?? '').replace(/\.json$/i, '').trim();
}

/**
 * Group name for a policy: {prefix}{policy name} - {label}{suffix}.
 * Prefix and suffix are used verbatim, so the user controls separators.
 */
export function buildGroupName(policyName, kind, naming = DEFAULT_GROUP_NAMING) {
  const labels = {
    include: naming.includeLabel,
    exclude: naming.excludeLabel,
    pilot: naming.pilotLabel || DEFAULT_GROUP_NAMING.pilotLabel,
  };
  const label = String(labels[kind] ?? naming.includeLabel ?? '').trim();
  const base = policyDisplayName(policyName);
  return `${naming.prefix ?? ''}${base}${label ? ` - ${label}` : ''}${naming.suffix ?? ''}`.trim();
}

/**
 * Problems with a naming setup, checked against the policies it will be
 * used for. Returns a list of messages (empty when fine).
 */
export function validateGroupNaming(naming, policyNames = [], { include = true, exclude = true, target = 'group' } = {}) {
  const errors = [];
  // With All devices/All users the exclude group is optional.
  if (target === 'group' && !include && !exclude) errors.push('Choose at least one of include or exclude groups.');

  const inc = String(naming.includeLabel ?? '').trim().toLowerCase();
  const exc = String(naming.excludeLabel ?? '').trim().toLowerCase();
  if (include && exclude && inc === exc) {
    errors.push('Include and exclude labels must be different, or both groups get the same name.');
  }

  const tooLong = policyNames
    .flatMap(name => [include && buildGroupName(name, 'include', naming), exclude && buildGroupName(name, 'exclude', naming)])
    .filter(name => name && name.length > MAX_GROUP_NAME_LENGTH);
  if (tooLong.length > 0) {
    errors.push(`${tooLong.length} group name(s) would be longer than ${MAX_GROUP_NAME_LENGTH} characters.`);
  }
  return errors;
}

/**
 * mailNickname is required even for security groups: ASCII, no spaces or
 * @()\[]";:<>, and at most 64 characters. A random tail keeps it unique.
 */
export function buildMailNickname(displayName, randomHex = randomSuffix()) {
  const slug = String(displayName ?? '')
    .normalize('NFKD')
    .replace(/[^A-Za-z0-9_-]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 50)
    .replace(/-$/, '');
  return `${slug || 'oib-group'}-${randomHex}`;
}

function randomSuffix() {
  const bytes = new Uint8Array(4);
  (globalThis.crypto ?? window.crypto).getRandomValues(bytes);
  return [...bytes].map(b => b.toString(16).padStart(2, '0')).join('');
}

// ── Assignment payloads ────────────────────────────────────────────────────

const INCLUDE_TYPE = '#microsoft.graph.groupAssignmentTarget';
const EXCLUDE_TYPE = '#microsoft.graph.exclusionGroupAssignmentTarget';

export function groupTarget(groupId, kind = 'include') {
  return { '@odata.type': kind === 'exclude' ? EXCLUDE_TYPE : INCLUDE_TYPE, groupId };
}

// ── Assign to: a new group or a built-in (virtual) group ───────────────────

export const ASSIGN_TARGET_LABELS = Object.freeze({
  group: 'A new include group per policy',
  allDevices: 'All devices',
  allUsers: 'All users',
  byScope: 'All devices or All users, by the policy\'s D/U scope',
});

const BUILT_IN_LABELS = { allDevices: 'All devices', allUsers: 'All users' };

export function builtInTarget(kind) {
  return {
    '@odata.type': kind === 'allUsers'
      ? '#microsoft.graph.allLicensedUsersAssignmentTarget'
      : '#microsoft.graph.allDevicesAssignmentTarget',
  };
}

export const builtInLabel = (kind) => BUILT_IN_LABELS[kind] ?? kind;

// OIB policy names carry their scope: "… - D - …" (device) or "… - U - …" (user).
export function policyScope(policyName) {
  const match = policyDisplayName(policyName).match(/\s-\s([DU])\s-\s/);
  if (!match) return null;
  return match[1] === 'U' ? 'user' : 'device';
}

// The built-in group a policy gets for a target choice; policies without a
// D/U marker count as device policies.
export function builtInFor(target, policyName) {
  if (target === 'allDevices' || target === 'allUsers') return target;
  if (target === 'byScope') return policyScope(policyName) === 'user' ? 'allUsers' : 'allDevices';
  return null;
}

// App protection policies and Windows feature/quality/driver update profiles
// can't be assigned to the All devices / All users virtual groups.
export function supportsBuiltInTargets(endpoint) {
  const path = String(endpoint ?? '').toLowerCase();
  if (path.includes('/deviceappmanagement/')) return false;
  return !/\/windows(feature|quality|driver)update(profiles|policies)\b/.test(path);
}

const normaliseOptions = (options = {}) => ({
  target: ASSIGN_TARGET_LABELS[options.target] ? options.target : 'group',
  include: options.include !== false,
  exclude: options.exclude !== false,
});

// Which groups the options create, before policy-type exceptions.
export function groupKindsFor(options) {
  const { target, include, exclude } = normaliseOptions(options);
  if (target === 'group') return [include && 'include', exclude && 'exclude'].filter(Boolean);
  return exclude ? ['exclude'] : [];
}

/**
 * What a deployed policy is assigned to.
 *
 * @param policy  { displayName, endpoint }
 * @param options { target, include, exclude }
 * @returns { builtIn: 'allDevices'|'allUsers'|null, groups: ['include'|'exclude'], skipped: string|null }
 *   skipped explains why nothing is assigned (the policy type can't use the
 *   chosen built-in group).
 */
export function planPolicyAssignment(policy, options) {
  const normalised = normaliseOptions(options);
  const builtIn = builtInFor(normalised.target, policy?.displayName);
  if (builtIn && !supportsBuiltInTargets(policy?.endpoint)) {
    return {
      builtIn: null,
      groups: [],
      skipped: `Intune can't assign this policy type to ${builtInLabel(builtIn)}, so it was left unassigned. Assign it to a group in Intune.`,
    };
  }
  return { builtIn, groups: groupKindsFor(normalised), skipped: null };
}

export function targetKey(target) {
  const type = String(target?.['@odata.type'] ?? '').toLowerCase();
  return `${type}|${String(target?.groupId ?? '').toLowerCase()}`;
}

// Assignments read back from Graph carry ids and source info that /assign
// rejects or ignores; keep only the target (including any filter on it).
export function toAssignInput(assignment) {
  return { target: { ...assignment.target } };
}

/**
 * /assign replaces a policy's whole assignment list, so new targets are
 * merged into the existing ones. Targets already present are skipped.
 * Returns { assignments, added } where added counts new targets.
 */
export function mergeAssignmentTargets(existing = [], additions = []) {
  const assignments = existing.map(toAssignInput);
  const keys = new Set(assignments.map(a => targetKey(a.target)));
  let added = 0;
  additions.forEach(target => {
    const key = targetKey(target);
    if (keys.has(key)) return;
    keys.add(key);
    assignments.push({ target: { ...target } });
    added += 1;
  });
  return { assignments, added };
}

// ── Directory objects ──────────────────────────────────────────────────────

export function normaliseDirectoryObject(obj) {
  const odataType = String(obj?.['@odata.type'] ?? '').toLowerCase();
  const type = odataType.endsWith('.device') ? 'device'
    : odataType.endsWith('.user') ? 'user'
    : odataType.endsWith('.group') ? 'group'
    : 'other';
  // Without Device.Read.All / User.ReadBasic.All Graph returns only the id.
  const name = obj?.displayName || obj?.userPrincipalName || null;
  const detail = type === 'device'
    ? [obj?.operatingSystem, obj?.operatingSystemVersion].filter(Boolean).join(' ')
    : type === 'user' ? (obj?.userPrincipalName || obj?.mail || '') : '';
  return { id: obj?.id, type, name, detail };
}

// ── Version migration ──────────────────────────────────────────────────────

const VERSION_RE = /\bv(\d+(?:\.\d+)*)\s*$/i;

export function extractPolicyVersion(name) {
  const match = policyDisplayName(name).match(VERSION_RE);
  return match ? match[1] : null;
}

export function compareVersions(a, b) {
  const pa = String(a).split('.').map(Number);
  const pb = String(b).split('.').map(Number);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const diff = (pa[i] || 0) - (pb[i] || 0);
    if (diff !== 0) return diff > 0 ? 1 : -1;
  }
  return 0;
}

const nameOf = (policy) => policy?.displayName || policy?.name || '';

/**
 * Policies where an older OIB version is still deployed next to the current
 * one. OIBID matches put old versions in `legacyTenantPolicies`; name-based
 * matches land in `duplicates`, so those are split by version here.
 *
 * Returns [{ key, name, current, legacy: [tenantPolicy…] }]
 */
export function getMigrationPairs(comparisonData) {
  const pairs = [];
  Object.values(comparisonData?.byOS ?? {}).forEach(osData => {
    (osData.current ?? []).forEach(item => {
      if (item.existingPolicy && item.legacyTenantPolicies?.length > 0) {
        pairs.push({
          key: item.existingPolicy.id,
          name: item.name,
          current: item.existingPolicy,
          legacy: item.legacyTenantPolicies,
        });
      }
    });

    (osData.duplicates ?? []).forEach(item => {
      if (item.matchMethod !== 'name' || !(item.matchedPolicies?.length > 1)) return;
      const available = extractPolicyVersion(item.name);
      if (!available) return;
      const current = item.matchedPolicies.filter(p => extractPolicyVersion(nameOf(p)) === available);
      const legacy = item.matchedPolicies.filter(p => {
        const version = extractPolicyVersion(nameOf(p));
        return version && compareVersions(version, available) < 0;
      });
      if (current.length === 1 && legacy.length > 0) {
        pairs.push({ key: current[0].id, name: item.name, current: current[0], legacy });
      }
    });
  });
  return pairs;
}

const normalise = (value) => String(value ?? '').trim().toLowerCase();

// A group "belongs" to a policy version when its name contains the policy's
// full display name (which includes the version), e.g.
// "Win - OIB - … - TPM - v3.7 - Include" for "Win - OIB - … - TPM - v3.7".
export function isDedicatedGroup(groupName, policyName) {
  const group = normalise(groupName);
  const policy = normalise(policyDisplayName(policyName));
  return !!group && !!policy && group.includes(policy);
}

// Pick the policy's own group of a kind from its assignment: the exact
// conventional name first, then any dedicated group. Returns { id, name }
// with id null when the group still has to be created.
function findOwnGroup(policyName, groupIds, groupNames, kind, naming) {
  const wanted = buildGroupName(policyName, kind, naming);
  const named = groupIds.map(id => ({ id, name: groupNames?.get(id.toLowerCase()) ?? null }));
  const pilotLabel = normalise(`- ${naming.pilotLabel || DEFAULT_GROUP_NAMING.pilotLabel}`);
  const isPilot = (g) => !!g.name && normalise(g.name).includes(pilotLabel);
  const exact = named.find(g => normalise(g.name) === normalise(wanted));
  if (exact) return { id: exact.id, name: exact.name, existing: true };
  // Any other group named after the policy (e.g. after the prefix changed),
  // telling pilot groups and include groups apart by the pilot label.
  const dedicated = named.find(g => g.name && isDedicatedGroup(g.name, policyName)
    && (kind === 'pilot' ? isPilot(g) : !isPilot(g)));
  if (dedicated) return { id: dedicated.id, name: dedicated.name, existing: true };
  return { id: null, name: wanted, existing: false };
}

/**
 * How devices (or users) move from a legacy policy version to the current one.
 *
 * - strategy 'exclude': moved members are added to the old version's exclude
 *   group and the new version's pilot group. The old version's include
 *   groups are never changed: an exclusion wins even when a member is in
 *   them through a nested or dynamic group, and shared groups stay intact
 *   for the other policies that use them.
 * - strategy 'none': the old version isn't assigned; nothing to migrate.
 *
 * @param legacyAssignment / currentAssignment  summarizeAssignments() results
 * @param groupNames  Map<lowercased groupId, name|null>
 */
export function planVersionMigration({ legacyPolicy, currentPolicy, legacyAssignment, currentAssignment, groupNames, naming = DEFAULT_GROUP_NAMING }) {
  const legacyName = nameOf(legacyPolicy);
  const currentName = nameOf(currentPolicy);
  const currentIncludeIds = currentAssignment?.includeGroupIds ?? [];
  const legacyIncludeIds = legacyAssignment?.includeGroupIds ?? [];

  const base = {
    legacyName,
    currentName,
    // Moved members go to the pilot group; when the migration finishes with
    // groups it becomes (or joins) the include group.
    currentPilotGroup: findOwnGroup(currentName, currentIncludeIds, groupNames, 'pilot', naming),
    currentIncludeGroup: findOwnGroup(currentName, currentIncludeIds, groupNames, 'include', naming),
    currentExcludeGroup: findOwnGroup(currentName, currentAssignment?.excludeGroupIds ?? [], groupNames, 'exclude', naming),
    sourceGroups: [],
    legacyExcludeGroup: null,
    searchTypes: [],
  };

  if (!legacyAssignment?.hasIncludes) {
    return { ...base, strategy: 'none', reason: 'The old version is not assigned to anything, so there is nothing to move.' };
  }

  const sourceGroups = legacyIncludeIds.map(id => ({ id, name: groupNames?.get(id.toLowerCase()) ?? null }));
  const broad = legacyAssignment.allDevices || legacyAssignment.allUsers;
  const allDedicated = sourceGroups.length > 0 && sourceGroups.every(g => g.name && isDedicatedGroup(g.name, legacyName));
  const how = 'moved members are added to its exclude group and to the new version\'s pilot group';

  const searchTypes = [
    ...(legacyAssignment.allDevices ? ['device'] : []),
    ...(legacyAssignment.allUsers ? ['user'] : []),
  ];
  return {
    ...base,
    strategy: 'exclude',
    sourceGroups,
    searchTypes,
    legacyExcludeGroup: findOwnGroup(legacyName, legacyAssignment.excludeGroupIds ?? [], groupNames, 'exclude', naming),
    reason: broad
      ? `The old version is assigned to ${[legacyAssignment.allDevices && 'All devices', legacyAssignment.allUsers && 'All users'].filter(Boolean).join(' and ')}, so ${how}.`
      : allDedicated
        ? `The old version is assigned through its own group: ${how}. Its include group is left as it is.`
        : `The old version is assigned to groups shared with other policies. Those groups are left alone: ${how}.`,
  };
}

// ── Finishing a migration ──────────────────────────────────────────────────

const targetType = (target) => String(target?.['@odata.type'] ?? '').toLowerCase();
const isBuiltInTarget = (target) => /alldevices|alllicensedusers/.test(targetType(target));
const isExclusionTarget = (target) => targetType(target).includes('exclusiongroup');
const isIncludeGroupTarget = (target) => targetType(target).includes('groupassignmenttarget') && !isExclusionTarget(target);

// Built-in groups a summarizeAssignments() result includes.
export const builtInsOf = (summary) =>
  [summary?.allDevices && 'allDevices', summary?.allUsers && 'allUsers'].filter(Boolean);

// The built-in group(s) the new version gets in the All devices model: what
// the old version had, or the new version's D/U scope.
export function finishBuiltIns(legacyAssignment, currentName) {
  const kinds = builtInsOf(legacyAssignment);
  return kinds.length > 0 ? kinds : [builtInFor('byScope', currentName)];
}

// Same model as the old version had, where the policy type allows it.
export function defaultFinishModel(legacyAssignment, currentEndpoint) {
  if (!supportsBuiltInTargets(currentEndpoint)) return 'groups';
  return builtInsOf(legacyAssignment).length > 0 ? 'builtIn' : 'groups';
}

// Drop All devices / All users from an assignment list, keeping the rest.
export function withoutBuiltIns(assignments = []) {
  return assignments.filter(a => !isBuiltInTarget(a?.target)).map(toAssignInput);
}

/**
 * What finishing a migration does to the new version. The old version loses
 * all its assignments in both models.
 *
 * - 'builtIn': the new version is assigned to All devices / All users and
 *   keeps its other targets, except its pilot group, which All devices
 *   makes redundant; the caller empties that group.
 * - 'groups': the new version ends up on its include group. The pilot group
 *   is renamed to it ('rename'), or, when an include group already exists,
 *   its members join that one and the pilot group is unassigned and emptied
 *   ('merge'). Everyone still on the old version's own include groups joins
 *   too, as does everyone the old version reached through All devices / All
 *   users (`audience`, listed by the caller). Shared include groups are
 *   handed over as they are.
 *
 * In both models the old version's exclusions carry over, except its own
 * exclude group: during a migration that holds the devices already moved.
 *
 * @param plan  from planVersionMigration
 * @param legacyAssignments / currentAssignments  raw assignments from Graph
 * @param legacyAssignment  summarizeAssignments(legacyAssignments)
 * @param groupNames  Map<lowercased groupId, name|null>
 * @param existingIncludeGroupId  a directory group with the include group's
 *   name that isn't assigned to the new version yet, if any
 */
export function planFinishMigration({
  model, plan, legacyAssignments = [], currentAssignments = [], legacyAssignment, groupNames, existingIncludeGroupId = null,
}) {
  const nameOfGroup = (id) => groupNames?.get(String(id).toLowerCase()) ?? null;
  const ownLegacyGroup = (id) => {
    if (id && plan.legacyExcludeGroup?.id && id.toLowerCase() === plan.legacyExcludeGroup.id.toLowerCase()) return true;
    const name = nameOfGroup(id);
    return !!name && isDedicatedGroup(name, plan.legacyName);
  };
  const legacyTargets = legacyAssignments.map(a => a?.target).filter(Boolean);
  const carriedExclusions = legacyTargets.filter(t => isExclusionTarget(t) && !ownLegacyGroup(t.groupId));
  const excludedGroupIds = legacyTargets.filter(isExclusionTarget).map(t => t.groupId);
  // The old version's own include/exclude groups: emptied once it's unassigned.
  // Shared groups are never touched.
  const ownLegacyGroupIds = [...new Set(legacyTargets
    .filter(t => t.groupId && !isBuiltInTarget(t) && ownLegacyGroup(t.groupId))
    .map(t => t.groupId))];

  if (model === 'builtIn') {
    const builtIns = finishBuiltIns(legacyAssignment, plan.currentName);
    const pilotId = String(plan.currentPilotGroup?.id ?? '').toLowerCase();
    const kept = currentAssignments.filter(a =>
      !(pilotId && isIncludeGroupTarget(a?.target) && String(a.target.groupId).toLowerCase() === pilotId));
    const { assignments } = mergeAssignmentTargets(kept, [...builtIns.map(builtInTarget), ...carriedExclusions]);
    return {
      model,
      builtIns,
      assignments,
      pilotGroupId: plan.currentPilotGroup?.id ?? null,
      ownLegacyGroupIds,
      removesPilotGroup: kept.length < currentAssignments.length,
      carriedExclusions: carriedExclusions.length,
      // The old version only reached groups, so All devices/users is wider.
      widensScope: builtInsOf(legacyAssignment).length === 0,
    };
  }

  const ownIncludeGroupIds = legacyTargets.filter(t => isIncludeGroupTarget(t) && ownLegacyGroup(t.groupId)).map(t => t.groupId);
  const sharedIncludes = legacyTargets.filter(t => isIncludeGroupTarget(t) && !ownLegacyGroup(t.groupId));

  const pilotGroupId = plan.currentPilotGroup?.id ?? null;
  const includeGroupId = plan.currentIncludeGroup?.id ?? existingIncludeGroupId ?? null;
  const sameGroup = !!pilotGroupId && !!includeGroupId && pilotGroupId.toLowerCase() === includeGroupId.toLowerCase();
  // rename: the pilot group becomes the include group
  // merge:  pilot members join the existing include group; the pilot group is unassigned and emptied
  // create: no pilot group, so the include group is created (or reused)
  // use:    the include group exists and there's no separate pilot group
  const includeAction = includeGroupId
    ? (pilotGroupId && !sameGroup ? 'merge' : 'use')
    : (pilotGroupId ? 'rename' : 'create');

  return {
    model,
    includeAction,
    includeGroupId,
    includeName: plan.currentIncludeGroup?.name ?? null,
    pilotGroupId,
    includeAssigned: !!plan.currentIncludeGroup?.id,
    ownLegacyGroupIds,
    ownIncludeGroupIds,
    excludedGroupIds,
    audience: builtInsOf(legacyAssignment).map(kind => (kind === 'allUsers' ? 'user' : 'device')),
    handOver: [...sharedIncludes, ...carriedExclusions],
    sharedIncludes: sharedIncludes.length,
    carriedExclusions: carriedExclusions.length,
  };
}

// ── Platforms (for listing "All devices" as group members) ─────────────────

/**
 * Platform of a tenant policy ({ displayName, policy: raw } or the raw
 * object): 'windows' | 'macos' | 'ios' | 'android' | 'linux' | null.
 */
export function policyPlatform(policy) {
  const raw = policy?.policy ?? policy ?? {};
  const fromText = (text) => {
    const value = String(text ?? '').toLowerCase();
    if (!value) return null;
    if (value.includes('macos')) return 'macos';
    if (/\bios\b|ios[a-z]/.test(value)) return 'ios';
    if (value.includes('android') || value.includes('aosp')) return 'android';
    if (value.includes('linux')) return 'linux';
    if (value.includes('windows') || value.includes('win32')) return 'windows';
    return null;
  };
  const fromPlatforms = fromText(raw.platforms);
  if (fromPlatforms) return fromPlatforms;
  const fromType = fromText(raw['@odata.type']);
  if (fromType) return fromType;
  const name = policyDisplayName(policy?.displayName ?? raw.displayName ?? raw.name).toLowerCase();
  if (/^(win|w365)\b/.test(name)) return 'windows';
  if (/^macos\b/.test(name)) return 'macos';
  if (/^(ios|ipados)\b/.test(name)) return 'ios';
  if (/^android\b/.test(name)) return 'android';
  return null;
}

export const PLATFORM_LABELS = Object.freeze({
  windows: 'Windows', macos: 'macOS', ios: 'iOS/iPadOS', android: 'Android', linux: 'Linux',
});

// Whether an Entra device's operatingSystem belongs to a platform. Unknown
// platforms match every device.
export function deviceMatchesPlatform(operatingSystem, platform) {
  const os = String(operatingSystem ?? '').toLowerCase();
  switch (platform) {
    case 'windows': return os.startsWith('windows');
    case 'macos': return os.startsWith('mac');
    case 'ios': return /^(ios|ipados|iphone|ipad)/.test(os);
    case 'android': return os.startsWith('android');
    case 'linux': return os.startsWith('linux') || os.startsWith('ubuntu') || os.startsWith('rhel');
    default: return true;
  }
}

/**
 * Targets the new version should get when the migration is finished: the
 * old version's assignments, minus the groups that only existed for the
 * migration (the old version's own include/exclude groups).
 */
export function planHandOverTargets(legacyAssignments = [], ownGroupIds = []) {
  const skip = new Set(ownGroupIds.filter(Boolean).map(id => id.toLowerCase()));
  return legacyAssignments
    .map(a => a.target)
    .filter(target => target && !skip.has(String(target.groupId ?? '').toLowerCase()));
}
