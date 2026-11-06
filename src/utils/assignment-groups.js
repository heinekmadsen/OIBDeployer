// Assignment groups: naming, assignment payloads and version-migration
// planning. Pure functions only — Graph calls live in graph-api.js.

export const DEFAULT_GROUP_NAMING = Object.freeze({
  prefix: '',
  suffix: '',
  includeLabel: 'Include',
  excludeLabel: 'Exclude',
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
  const label = String((kind === 'exclude' ? naming.excludeLabel : naming.includeLabel) ?? '').trim();
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
  const exact = named.find(g => normalise(g.name) === normalise(wanted));
  if (exact) return { id: exact.id, name: exact.name, existing: true };
  const dedicated = named.find(g => g.name && isDedicatedGroup(g.name, policyName));
  if (dedicated) return { id: dedicated.id, name: dedicated.name, existing: true };
  return { id: null, name: wanted, existing: false };
}

/**
 * How devices (or users) move from a legacy policy version to the current one.
 *
 * - strategy 'move': the old version is only assigned through groups that
 *   belong to it, so members are moved out of that group into the new
 *   version's include group.
 * - strategy 'exclude': the old version is assigned to All devices / All
 *   users or to shared groups. Removing members from shared groups would
 *   affect other policies, so members are added to the old version's exclude
 *   group and the new version's include group instead.
 * - strategy 'none': the old version isn't assigned; nothing to migrate.
 *
 * @param legacyAssignment / currentAssignment  summarizeAssignments() results
 * @param groupNames  Map<lowercased groupId, name|null>
 */
export function planVersionMigration({ legacyPolicy, currentPolicy, legacyAssignment, currentAssignment, groupNames, naming = DEFAULT_GROUP_NAMING }) {
  const legacyName = nameOf(legacyPolicy);
  const currentName = nameOf(currentPolicy);
  const currentInclude = findOwnGroup(currentName, currentAssignment?.includeGroupIds ?? [], groupNames, 'include', naming);
  const legacyIncludeIds = legacyAssignment?.includeGroupIds ?? [];

  const base = {
    legacyName,
    currentName,
    currentIncludeGroup: currentInclude,
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

  if (!broad && allDedicated) {
    return {
      ...base,
      strategy: 'move',
      sourceGroups,
      reason: 'The old version is assigned through its own group, so members are moved from that group to the new version\'s include group.',
    };
  }

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
      ? `The old version is assigned to ${[legacyAssignment.allDevices && 'All devices', legacyAssignment.allUsers && 'All users'].filter(Boolean).join(' and ')}, so moved members are added to its exclude group and to the new version's include group.`
      : 'The old version is assigned to groups shared with other policies. Those groups are left alone: moved members are added to the old version\'s exclude group and the new version\'s include group.',
  };
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
