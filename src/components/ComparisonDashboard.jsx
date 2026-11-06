import React, { useState, useEffect } from 'react';
import { BarChart3, TrendingUp, AlertTriangle, CheckCircle, Plus, ArrowLeft, RefreshCw, Copy, Users, Loader, KeyRound } from 'lucide-react';
import graphAPI from '../utils/graph-api';
import { summarizeAssignments } from '../utils/device-validation';
import {
  getMatchedTenantPolicies, normalisePolicyStatusSummary, collectGroupIds, summariseAssignmentStatus
} from '../utils/policy-status';
import PolicyAssignmentStatus from './PolicyAssignmentStatus';

const ComparisonDashboard = ({
  existingPolicies,
  availablePolicies,
  onSelectPolicies,
  onBack,
  onRefresh,
  onComparisonReady,
  isLoading,
  selectedVersion,
  showAssignmentStatus = false,
  autoLoadAssignmentStatus = false,
  onAutoLoadConsumed,
  onRequestGroupConsent
}) => {
  const [comparisonData, setComparisonData] = useState(null);
  const [selectedForDeployment, setSelectedForDeployment] = useState([]);
  const [viewFilter, setViewFilter] = useState('all'); // 'all', 'new', 'outdated', 'missing'

  // Assignment & device-status overview (loaded on demand)
  const [assignmentsById, setAssignmentsById] = useState(null);
  const [statusById, setStatusById] = useState(null);
  const [groupNames, setGroupNames] = useState(null);
  const [groupNamesIssue, setGroupNamesIssue] = useState(null); // 'consent' | error message
  const [assignmentFilters, setAssignmentFilters] = useState(null);
  const [isLoadingStatus, setIsLoadingStatus] = useState(false);
  const [statusError, setStatusError] = useState(null);

  useEffect(() => {
    if (availablePolicies && Object.keys(availablePolicies).length > 0) {
      generateComparisonData();
    } else {
      setComparisonData(null);
    }
  }, [availablePolicies, existingPolicies]);

  const loadAssignmentStatus = async (data = comparisonData) => {
    const tenantPolicies = getMatchedTenantPolicies(data);
    if (tenantPolicies.length === 0) return;

    setIsLoadingStatus(true);
    setStatusError(null);
    try {
      const [rawAssignments, rawStatus, filters] = await Promise.all([
        graphAPI.getAssignmentsForPolicies(tenantPolicies),
        graphAPI.getPolicyStatusSummaries(tenantPolicies),
        graphAPI.getAssignmentFilters().catch(() => null)
      ]);

      const assignments = new Map();
      rawAssignments.forEach((value, id) => {
        assignments.set(id, Array.isArray(value) ? summarizeAssignments(value) : value);
      });
      const statuses = new Map();
      tenantPolicies.forEach(policy => {
        statuses.set(policy.id, normalisePolicyStatusSummary(rawStatus.get(policy.id)));
      });

      setAssignmentsById(assignments);
      setStatusById(statuses);
      setAssignmentFilters(filters);

      // Group names need an extra permission; degrade to IDs without it.
      const groupIds = collectGroupIds([...assignments.values()]);
      if (groupIds.length > 0) {
        try {
          setGroupNames(await graphAPI.getGroupNames(groupIds));
          setGroupNamesIssue(null);
        } catch (error) {
          setGroupNames(null);
          setGroupNamesIssue(error?.name === 'ConsentRequiredError' ? 'consent' : error.message);
        }
      }
    } catch (error) {
      setStatusError(error?.name === 'ConsentRequiredError'
        ? 'Sign-in needs to be refreshed — sign out and in again.'
        : error.message);
    } finally {
      setIsLoadingStatus(false);
    }
  };

  // Returning from the group-names consent redirect: load straight away.
  useEffect(() => {
    if (autoLoadAssignmentStatus && showAssignmentStatus && comparisonData && !isLoadingStatus && !statusById) {
      onAutoLoadConsumed?.();
      loadAssignmentStatus(comparisonData);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [autoLoadAssignmentStatus, showAssignmentStatus, comparisonData]);

  const isUnassigned = (policy) => {
    const assignment = policy.existingPolicy && assignmentsById?.get(policy.existingPolicy.id);
    return !!assignment && !assignment.error && !assignment.hasIncludes;
  };

  const isAssigned = (policy) => {
    const assignment = policy.existingPolicy && assignmentsById?.get(policy.existingPolicy.id);
    return !!assignment && !assignment.error && assignment.hasIncludes;
  };

  const hasDeviceIssues = (policy) => {
    const status = policy.existingPolicy && statusById?.get(policy.existingPolicy.id);
    return !!status && !status.loadError && status.issues > 0;
  };

  // Policies of one OS shown for a filter. With forSelection, only entries
  // that can be ticked for deployment (missing / outdated / newer).
  const policiesForFilter = (osData, filter, { forSelection = false } = {}) => {
    const deployed = [...(osData.current || []), ...(osData.outdated || []), ...(osData.newer || [])];
    const deployable = (list) => forSelection ? list.filter(p => p.status !== 'current') : list;
    switch (filter) {
      case 'new':          return osData.missing || [];
      case 'outdated':     return osData.outdated || [];
      case 'newer':        return osData.newer || [];
      case 'current':      return osData.current || [];
      case 'duplicate':    return forSelection ? [] : (osData.duplicates || []);
      case 'assigned':     return deployable(deployed.filter(isAssigned));
      case 'unassigned':   return deployable(deployed.filter(isUnassigned));
      case 'deviceIssues': return deployable(deployed.filter(hasDeviceIssues));
      default:
        return forSelection
          ? [...(osData.missing || []), ...(osData.outdated || []), ...(osData.newer || [])]
          : [
              ...(osData.missing || []),
              ...(osData.outdated || []),
              ...(osData.newer || []),
              ...(osData.current || []),
              ...(osData.duplicates || [])
            ];
    }
  };

  const generateComparisonData = () => {
    const comparison = {
      byOS: {},
      totals: {
        current: 0,
        outdated: 0,
        newer: 0,
        missing: 0,
        duplicates: 0,
        total: 0
      }
    };

    // Helper functions used by the name-based fallback path
    const extractVersion = (name) => {
      const versionMatch = name.match(/v(\d+(?:\.\d+)*)/i);
      return versionMatch ? versionMatch[1] : null;
    };

    const extractBaseName = (name) => {
      return name.replace(/\s*-\s*v\d+(?:\.\d+)*\s*$/i, '').trim();
    };

    const compareVersions = (v1, v2) => {
      const parts1 = v1.split('.').map(Number);
      const parts2 = v2.split('.').map(Number);
      for (let i = 0; i < Math.max(parts1.length, parts2.length); i++) {
        const part1 = parts1[i] || 0;
        const part2 = parts2[i] || 0;
        if (part1 > part2) return 1;
        if (part1 < part2) return -1;
      }
      return 0;
    };

    // Process each OS separately
    Object.entries(availablePolicies).forEach(([osType, osData]) => {
      comparison.byOS[osType] = {
        current: [],
        outdated: [],
        newer: [],
        missing: [],
        duplicates: [],
        totals: {
          current: 0,
          outdated: 0,
          newer: 0,
          missing: 0,
          duplicates: 0,
          total: 0
        }
      };

      // Flatten policies for this OS
      const osPolicies = [];
      Object.entries(osData).forEach(([policyType, policies]) => {
        if (policyType === '_deprecatedPolicies') return;
        osPolicies.push(...policies.map(p => ({ ...p, osType })));
      });

      comparison.byOS[osType].deprecatedPolicies = (osData._deprecatedPolicies || [])
        .map(deprecatedPolicy => ({
          ...deprecatedPolicy,
          tenantPolicy: existingPolicies.find(policy => policy.oibId === deprecatedPolicy.oibId) || null,
          replacements: deprecatedPolicy.replacements.map(replacement => ({
            ...replacement,
            tenantPolicy: existingPolicies.find(policy => policy.oibId === replacement.oibId) || null
          }))
        }))
        .filter(deprecatedPolicy => deprecatedPolicy.tenantPolicy);

      osPolicies.forEach(availablePolicy => {
        const policyName = availablePolicy.name.replace('.json', '');
        const availableOibId = availablePolicy.oibId?.toUpperCase();
        const previousVersionIds = (availablePolicy.previousVersions || []).map(id => id.toUpperCase());
        availablePolicy.replacements = (availablePolicy.replacements || []).map(replacement => ({
          ...replacement,
          tenantPolicy: existingPolicies.find(policy => policy.oibId === replacement.oibId) || null
        }));

        // --- Pass 1: OIBID matching (v3.8+ branches with PolicyManifest) ---
        if (availableOibId) {
          // Current version already deployed?
          const currentMatches = existingPolicies.filter(p => p.oibId === availableOibId);
          if (currentMatches.length > 1) {
            comparison.byOS[osType].duplicates.push({
              ...availablePolicy,
              matchedPolicies: currentMatches,
              status: 'duplicate',
              matchMethod: 'oibid'
            });
            return;
          }
          const matchedByCurrent = currentMatches[0] || null;

          // A previous version can coexist with the current policy and should be
          // retained as a cleanup candidate rather than silently ignored.
          const previousMatches = existingPolicies.filter(p => p.oibId && previousVersionIds.includes(p.oibId));
          let matchedByPrevious = null;
          if (!matchedByCurrent) {
            if (previousMatches.length > 1) {
              comparison.byOS[osType].duplicates.push({
                ...availablePolicy,
                matchedPolicies: previousMatches,
                status: 'duplicate',
                matchMethod: 'oibid'
              });
              return;
            }
            matchedByPrevious = previousMatches[0] || null;
          }

          if (matchedByCurrent) {
            comparison.byOS[osType].current.push({
              ...availablePolicy,
              existingPolicy: matchedByCurrent,
              legacyTenantPolicies: previousMatches,
              status: 'current',
              matchMethod: 'oibid'
            });
            return;
          } else if (matchedByPrevious) {
            comparison.byOS[osType].outdated.push({
              ...availablePolicy,
              existingPolicy: matchedByPrevious,
              status: 'outdated',
              availableVersion: extractVersion(policyName),
              existingVersion: extractVersion(matchedByPrevious.displayName || ''),
              matchMethod: 'oibid'
            });
            return;
          }
          // No OIBID match found — fall through to name-based matching below.
          // The tenant policy may exist but simply not have the OIBID in its description.
        }

        // --- Pass 2: Name-based fallback (pre-3.8 branches without PolicyManifest) ---
        const availableVersion = extractVersion(policyName);
        const availableBaseName = extractBaseName(policyName);

        const nameMatches = existingPolicies.filter(existing => {
          const existingBaseName = extractBaseName(existing.displayName || existing.name || '');
          return existingBaseName.toLowerCase() === availableBaseName.toLowerCase();
        });

        if (nameMatches.length > 1) {
          comparison.byOS[osType].duplicates.push({
            ...availablePolicy,
            matchedPolicies: nameMatches,
            status: 'duplicate',
            matchMethod: 'name'
          });
          return;
        }

        const existingPolicy = nameMatches[0] || null;

        if (existingPolicy) {
          const existingVersion = extractVersion(existingPolicy.displayName || existingPolicy.name || '');

          if (!availableVersion || !existingVersion) {
            comparison.byOS[osType].current.push({
              ...availablePolicy,
              existingPolicy,
              status: 'current',
              matchMethod: 'name'
            });
          } else {
            const versionComparison = compareVersions(availableVersion, existingVersion);

            if (versionComparison === 0) {
              comparison.byOS[osType].current.push({
                ...availablePolicy,
                existingPolicy,
                status: 'current',
                matchMethod: 'name'
              });
            } else if (versionComparison > 0) {
              comparison.byOS[osType].outdated.push({
                ...availablePolicy,
                existingPolicy,
                status: 'outdated',
                availableVersion,
                existingVersion,
                matchMethod: 'name'
              });
            } else {
              comparison.byOS[osType].newer.push({
                ...availablePolicy,
                existingPolicy,
                status: 'newer',
                availableVersion,
                existingVersion,
                matchMethod: 'name'
              });
            }
          }
        } else {
          comparison.byOS[osType].missing.push({
            ...availablePolicy,
            status: 'missing',
            matchMethod: 'name'
          });
        }
      });

      // Calculate OS totals
      comparison.byOS[osType].totals = {
        current: comparison.byOS[osType].current.length,
        outdated: comparison.byOS[osType].outdated.length,
        newer: comparison.byOS[osType].newer.length,
        missing: comparison.byOS[osType].missing.length,
        duplicates: comparison.byOS[osType].duplicates.length,
        total: osPolicies.length
      };

      // Add to overall totals
      comparison.totals.current += comparison.byOS[osType].totals.current;
      comparison.totals.outdated += comparison.byOS[osType].totals.outdated;
      comparison.totals.newer += comparison.byOS[osType].totals.newer;
      comparison.totals.missing += comparison.byOS[osType].totals.missing;
      comparison.totals.duplicates += comparison.byOS[osType].totals.duplicates;
      comparison.totals.total += comparison.byOS[osType].totals.total;
    });

    setComparisonData(comparison);
    if (onComparisonReady) {
      onComparisonReady(comparison);
    }
  };

  const getFilteredPolicies = () => {
    if (!comparisonData || !comparisonData.byOS) return [];
    
    // Collect policies from all OS types based on filter
    const allPolicies = [];
    
    Object.values(comparisonData.byOS).forEach(osData => {
      switch (viewFilter) {
        case 'new':
          allPolicies.push(...(osData.missing || []));
          break;
        case 'outdated':
          allPolicies.push(...(osData.outdated || []));
          break;
        case 'newer':
          allPolicies.push(...(osData.newer || []));
          break;
        case 'current':
          allPolicies.push(...(osData.current || []));
          break;
        default:
          allPolicies.push(
            ...(osData.missing || []),
            ...(osData.outdated || []),
            ...(osData.newer || []),
            ...(osData.current || [])
          );
          break;
      }
    });
    
    return allPolicies;
  };

  const togglePolicySelection = (policy) => {
    setSelectedForDeployment(prev => {
      const isSelected = prev.some(p => p.name === policy.name);
      if (isSelected) {
        return prev.filter(p => p.name !== policy.name);
      } else {
        return [...prev, policy];
      }
    });
  };

  const handleDeploySelected = () => {
    console.warn('ComparisonDashboard: Deploy button clicked');
    console.warn('Selected policies for deployment:', selectedForDeployment);
    if (selectedForDeployment.length > 0) {
      onSelectPolicies(selectedForDeployment);
    } else {
      console.warn('No policies selected for deployment');
    }
  };

  const handleGlobalSelectAll = () => {
    // Get all currently filtered policies across all OS types
    const allFilteredPolicies = [];
    Object.values(comparisonData.byOS).forEach(osData => {
      allFilteredPolicies.push(...policiesForFilter(osData, viewFilter, { forSelection: true }));
    });

    // Check if all filtered policies are selected
    const areAllSelected = allFilteredPolicies.length > 0 && 
      allFilteredPolicies.every(policy => selectedForDeployment.some(p => p.name === policy.name));

    if (areAllSelected) {
      // Deselect all filtered policies
      setSelectedForDeployment(prev => 
        prev.filter(selected => !allFilteredPolicies.some(filtered => filtered.name === selected.name))
      );
    } else {
      // Select all filtered policies that aren't already selected
      setSelectedForDeployment(prev => {
        const newSelections = allFilteredPolicies.filter(
          filtered => !prev.some(selected => selected.name === filtered.name)
        );
        return [...prev, ...newSelections];
      });
    }
  };

  if (!comparisonData) {
    return (
      <div className="wizard-container">
        <div className="loading-state">
          <RefreshCw className="spinner" size={32} />
          <p>{isLoading ? 'Fetching Policies...' : 'Analysing your tenant policies...'}</p>
        </div>
      </div>
    );
  }

  return (
    <div className="wizard-container comparison-dashboard">
      <div className="comparison-content">
        <div className="wizard-header">
          <h2>Policy Comparison Dashboard</h2>
          <p>Compare your current OIB deployment with the latest version</p>
          <p><b>Important:</b> Checks are done on OIBID or policy name match <i>only</i>, not settings held within!</p>
          <p>Use the Validation Dashboard to verify settings.</p>
          {selectedVersion && (
            <div className="version-info">
              <span className="version-badge">Comparing against: {selectedVersion}</span>
            </div>
          )}
          {Object.values(comparisonData.byOS).flatMap(osData => osData.deprecatedPolicies || []).length > 0 && (
            <div className="deprecation-notice">
              <AlertTriangle size={14} />
              <span>Deprecated OIB policies deployed in this tenant - review for removal:</span>
              {Object.values(comparisonData.byOS).flatMap(osData => osData.deprecatedPolicies || []).map(policy => (
                <span key={policy.tenantPolicy.id} className="deprecation-replacement">
                  {policy.tenantPolicy.displayName || policy.tenantPolicy.name}
                  {policy.replacements.some(replacement => replacement.tenantPolicy)
                    ? ' - replacement deployed'
                    : ' - replacement available in OIB'}
                </span>
              ))}
            </div>
          )}
        </div>

        {/* Summary Stats */}
        <div className="comparison-stats">
        <div className="stat-card current">
          <div className="stat-icon">
            <CheckCircle size={24} />
          </div>
          <div className="stat-content">
            <div className="stat-number">{comparisonData.totals.current}</div>
            <div className="stat-label">Current Policies</div>
          </div>
        </div>

        <div className="stat-card outdated">
          <div className="stat-icon">
            <TrendingUp size={24} />
          </div>
          <div className="stat-content">
            <div className="stat-number">{comparisonData.totals.outdated}</div>
            <div className="stat-label">Outdated Policies</div>
          </div>
        </div>

        <div className="stat-card missing">
          <div className="stat-icon">
            <Plus size={24} />
          </div>
          <div className="stat-content">
            <div className="stat-number">{comparisonData.totals.missing}</div>
            <div className="stat-label">New Policies</div>
          </div>
        </div>

        <div className="stat-card newer">
          <div className="stat-icon">
            <AlertTriangle size={24} />
          </div>
          <div className="stat-content">
            <div className="stat-number">{comparisonData.totals.newer}</div>
            <div className="stat-label">Newer Than Latest</div>
          </div>
        </div>

        <div className="stat-card duplicate">
          <div className="stat-icon">
            <Copy size={24} />
          </div>
          <div className="stat-content">
            <div className="stat-number">{comparisonData.totals.duplicates}</div>
            <div className="stat-label">Duplicate Matches</div>
          </div>
        </div>
      </div>

      {/* Filter Tabs */}
      <div className="filter-tabs">
        <button 
          className={`filter-tab ${viewFilter === 'all' ? 'active' : ''}`}
          onClick={() => setViewFilter('all')}
        >
          All Policies ({comparisonData.totals.total})
        </button>
        <button 
          className={`filter-tab ${viewFilter === 'new' ? 'active' : ''}`}
          onClick={() => setViewFilter('new')}
        >
          New ({comparisonData.totals.missing})
        </button>
        <button 
          className={`filter-tab ${viewFilter === 'outdated' ? 'active' : ''}`}
          onClick={() => setViewFilter('outdated')}
        >
          Updates Available ({comparisonData.totals.outdated})
        </button>
        <button 
          className={`filter-tab ${viewFilter === 'newer' ? 'active' : ''}`}
          onClick={() => setViewFilter('newer')}
        >
          Newer Than Latest ({comparisonData.totals.newer})
        </button>
        {comparisonData.totals.duplicates > 0 && (
          <button
            className={`filter-tab ${viewFilter === 'duplicate' ? 'active' : ''}`}
            onClick={() => setViewFilter('duplicate')}
          >
            Duplicate Matches ({comparisonData.totals.duplicates})
          </button>
        )}
        {statusById && (() => {
          const count = (filter) => Object.values(comparisonData.byOS)
            .reduce((sum, osData) => sum + policiesForFilter(osData, filter).length, 0);
          return (
            <>
              <button
                className={`filter-tab ${viewFilter === 'assigned' ? 'active' : ''}`}
                onClick={() => setViewFilter('assigned')}
              >
                Assigned ({count('assigned')})
              </button>
              <button
                className={`filter-tab ${viewFilter === 'unassigned' ? 'active' : ''}`}
                onClick={() => setViewFilter('unassigned')}
              >
                Not Assigned ({count('unassigned')})
              </button>
              <button
                className={`filter-tab ${viewFilter === 'deviceIssues' ? 'active' : ''}`}
                onClick={() => setViewFilter('deviceIssues')}
              >
                Device Errors ({count('deviceIssues')})
              </button>
            </>
          );
        })()}
      </div>

      {/* Assignment & device status overview */}
      {showAssignmentStatus && (
        <div className="assignment-status-bar">
          <div className="assignment-status-main">
            <Users size={18} />
            {statusById ? (() => {
              const summary = summariseAssignmentStatus(getMatchedTenantPolicies(comparisonData), assignmentsById, statusById);
              return (
                <span>
                  {`${summary.assigned} of ${summary.total} deployed OIB policies are assigned`}
                  {summary.unassigned > 0 && <> · <b>{summary.unassigned} not assigned</b></>}
                  {summary.withIssues > 0 && <> · <b>{summary.withIssues} with device errors or conflicts</b></>}
                </span>
              );
            })() : (
              <span>See who each deployed OIB policy is assigned to, and how many devices report success, errors or conflicts.</span>
            )}
            <button
              className="btn btn-secondary btn-small"
              onClick={() => loadAssignmentStatus()}
              disabled={isLoadingStatus || getMatchedTenantPolicies(comparisonData).length === 0}
            >
              {isLoadingStatus
                ? <><Loader size={14} className="spinning" /> Loading…</>
                : statusById ? <><RefreshCw size={14} /> Refresh status</> : 'Load assignments & device status'}
            </button>
          </div>
          {statusError && (
            <div className="assignment-status-note"><AlertTriangle size={14} /> {statusError}</div>
          )}
          {groupNamesIssue === 'consent' && (
            <div className="assignment-status-note">
              <KeyRound size={14} />
              <span>Group names need the read-only <code>GroupMember.Read.All</code> permission (admin consent). Showing group IDs until then.</span>
              {onRequestGroupConsent && (
                <button className="btn btn-outline btn-small" onClick={onRequestGroupConsent}>Grant access</button>
              )}
            </div>
          )}
          {groupNamesIssue && groupNamesIssue !== 'consent' && (
            <div className="assignment-status-note"><AlertTriangle size={14} /> Couldn't read group names: {groupNamesIssue.slice(0, 160)}</div>
          )}
        </div>
      )}

      {/* Global Selection Controls */}
      <div className="global-controls">
        <button
          onClick={handleGlobalSelectAll}
          className="btn btn-outline btn-small global-select-all"
        >
          {(() => {
            // Get all currently filtered policies
            const allFilteredPolicies = [];
            Object.values(comparisonData.byOS).forEach(osData => {
              allFilteredPolicies.push(...policiesForFilter(osData, viewFilter, { forSelection: true }));
            });

            const areAllSelected = allFilteredPolicies.length > 0 && 
              allFilteredPolicies.every(policy => selectedForDeployment.some(p => p.name === policy.name));
            
            return areAllSelected ? 'Deselect All Policies' : 'Select All Policies';
          })()}
        </button>
        <span className="selected-count">
          {selectedForDeployment.length} policies selected
        </span>
      </div>

      {/* Policy List */}
      <div className="policy-comparison-list">
        {comparisonData.byOS && Object.entries(comparisonData.byOS).map(([osType, osData]) => {
          // Get filtered policies for this OS
          const osPolicies = policiesForFilter(osData, viewFilter);
          
          if (osPolicies.length === 0) return null;

          return (
            <div key={osType} className="os-policy-section">
              <div className="os-section-header">
                <h3 className="os-section-title">{osType.toUpperCase()}</h3>
                <div className="os-section-stats">
                  {osData.totals && (
                    <>
                      {viewFilter === 'all' && (
                        <span className="os-stat-summary">
                          {osData.totals.missing} new, {osData.totals.outdated} updates, {osData.totals.current} current
                        </span>
                      )}
                      {viewFilter === 'new' && <span className="os-stat-summary">{osData.totals.missing} policies</span>}
                      {viewFilter === 'outdated' && <span className="os-stat-summary">{osData.totals.outdated} policies</span>}
                      {viewFilter === 'newer' && <span className="os-stat-summary">{osData.totals.newer} policies</span>}
                      {viewFilter === 'current' && <span className="os-stat-summary">{osData.totals.current} policies</span>}
                      {viewFilter === 'duplicate' && <span className="os-stat-summary">{osData.totals.duplicates} policies</span>}
                      {(viewFilter === 'assigned' || viewFilter === 'unassigned' || viewFilter === 'deviceIssues') && (
                        <span className="os-stat-summary">{osPolicies.length} policies</span>
                      )}
                    </>
                  )}
                </div>
              </div>
              
              <div className="os-policies">
                {(() => {
                  // Sub-group this OS's policies by policy type (Compliance, Settings Catalog, etc.)
                  const policiesByType = {};
                  osPolicies.forEach(policy => {
                    const key = policy.policyType || 'Other';
                    if (!policiesByType[key]) policiesByType[key] = [];
                    policiesByType[key].push(policy);
                  });

                  return Object.entries(policiesByType).map(([policyType, policies]) => (
                    <div key={`${osType}-${policyType}`} className="policy-type-section">
                      <div className="policy-type-header">
                        <h4 className="policy-type-title">
                          {policyType.replace(/([A-Z])/g, ' $1').trim()}
                        </h4>
                        <span className="policy-type-count">
                          {policies.length} policies
                        </span>
                      </div>

                      <div className="policy-type-policies">
                        {policies.map((policy, index) => (
                          <div key={`${policy.name}-${osType}-${index}`} className={`policy-comparison-item ${policy.status}`}>
                            <div className="policy-info">
                              <div className="policy-name">
                                {policy.name.replace('.json', '')}
                                {policy.skuRequirements === 'Enterprise' && (
                                  <span className="req-tag req-tag--enterprise" title="Requires Windows Enterprise SKU">Enterprise</span>
                                )}
                                {policy.licenseRequirements === 'MDE' && (
                                  <span className="req-tag req-tag--mde" title="Requires Microsoft Defender for Endpoint licence">MDE</span>
                                )}
                              </div>
                              {policy.existingPolicy && (
                                <div className="matched-policy-name">
                                  <span className="matched-label">Matched Policy:</span>
                                  <span className="matched-name">{policy.existingPolicy.displayName || policy.existingPolicy.name}</span>
                                </div>
                              )}
                              {showAssignmentStatus && policy.existingPolicy && statusById && (
                                <PolicyAssignmentStatus
                                  tenantPolicy={policy.existingPolicy}
                                  assignment={assignmentsById?.get(policy.existingPolicy.id)}
                                  status={statusById.get(policy.existingPolicy.id)}
                                  groupNames={groupNames}
                                  filters={assignmentFilters}
                                />
                              )}
                              {policy.lifecycleStatus === 'deprecated' && (
                                <div className="deprecation-notice">
                                  <AlertTriangle size={14} />
                                  <span>
                                    Deprecated - {policy.replacements?.some(replacement => replacement.tenantPolicy)
                                      ? 'replacement is deployed; this policy can be reviewed for removal.'
                                      : 'replacement available in OIB; deploy it before removing this policy.'}
                                  </span>
                                  {policy.replacements?.map(replacement => (
                                    <span key={replacement.oibId} className="deprecation-replacement">
                                      {replacement.tenantPolicy ? 'Deployed: ' : 'Replacement: '}{replacement.name}
                                    </span>
                                  ))}
                                </div>
                              )}
                              {policy.legacyTenantPolicies?.length > 0 && (
                                <div className="deprecation-notice">
                                  <AlertTriangle size={14} />
                                  <span>Legacy OIB version also deployed - review it for removal:</span>
                                  {policy.legacyTenantPolicies.map(legacyPolicy => (
                                    <span key={legacyPolicy.id} className="deprecation-replacement">
                                      {legacyPolicy.displayName || legacyPolicy.name}
                                    </span>
                                  ))}
                                </div>
                              )}
                              {policy.status === 'duplicate' && (
                                <div className="duplicate-matches">
                                  <span className="matched-label">Ambiguous - {policy.matchedPolicies.length} matching tenant policies:</span>
                                  <ul className="duplicate-matches-list">
                                    {policy.matchedPolicies.map(match => (
                                      <li key={match.id}>
                                        <span className="matched-name">{match.displayName || match.name}</span>
                                        <span className="duplicate-match-id">Graph Policy ID: {match.id}</span>
                                      </li>
                                    ))}
                                  </ul>
                                </div>
                              )}
                              <div className="policy-status">
                                {policy.status === 'current' && <span className="status-badge current">Up to date</span>}
                                {policy.status === 'outdated' && (
                                  <span className="status-badge outdated">
                                    Update available: v{policy.existingVersion} → v{policy.availableVersion}
                                  </span>
                                )}
                                {policy.status === 'missing' && <span className="status-badge missing">New policy</span>}
                                {policy.status === 'newer' && (
                                  <span className="status-badge newer">
                                    Newer than latest: v{policy.existingVersion} {'>'} v{policy.availableVersion}
                                  </span>
                                )}
                                {policy.status === 'duplicate' && (
                                  <span className="status-badge duplicate">Skipped - not validated or matched automatically</span>
                                )}
                                {policy.matchMethod && (
                                  <span
                                    className={`match-badge match-badge--${policy.matchMethod}`}
                                    title={policy.matchMethod === 'oibid' ? 'Matched by unique OIBID' : 'Matched by policy name (no OIBID available)'}
                                  >
                                    {policy.matchMethod === 'oibid' ? 'OIBID' : 'Name'}
                                  </span>
                                )}
                              </div>
                            </div>

                            {(policy.status === 'outdated' || policy.status === 'missing') && (
                              <div className="policy-actions">
                                <label className="checkbox-label">
                                  <input
                                    type="checkbox"
                                    checked={selectedForDeployment.some(p => p.name === policy.name)}
                                    onChange={() => togglePolicySelection(policy)}
                                  />
                                  <span className="checkbox-text">Deploy</span>
                                </label>
                              </div>
                            )}
                          </div>
                        ))}
                      </div>
                    </div>
                  ));
                })()}
              </div>
            </div>
          );
        })}
      </div>
      </div>

      {/* Sticky Navigation */}
      <div className="comparison-navigation-sticky">
        <div className="wizard-navigation">
          <button className="btn-secondary" onClick={onBack}>
            <ArrowLeft size={16} />
            Back
          </button>
          
          <button className="btn-secondary btn-refresh" onClick={onRefresh} disabled={isLoading}>
            <RefreshCw size={16} className={isLoading ? 'spinning' : ''} />
            Refresh Data
          </button>

          {selectedForDeployment.length > 0 && (
            <button className="btn-primary" onClick={handleDeploySelected}>
              Deploy Selected ({selectedForDeployment.length})
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

export default ComparisonDashboard;
