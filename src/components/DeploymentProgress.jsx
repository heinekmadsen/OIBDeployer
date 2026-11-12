import React from 'react';
import { CheckCircle, XCircle, Clock, Loader, RotateCcw, ArrowLeft, Users, Ban, AlertTriangle, Globe, Filter } from 'lucide-react';
import { builtInLabel, filterModeLabel } from '../utils/assignment-groups';

const outcomeMessage = (groups) => {
  if (groups.added === 0) return 'The policy already had this assignment.';
  if (!groups.builtIn) return 'Groups assigned to the policy.';
  return `Assigned to ${builtInLabel(groups.builtIn)}${groups.exclude ? ', with the exclude group' : ''}.`;
};

// Outcome of ensureAssignmentGroupsForPolicy for one deployed policy.
const GroupOutcome = ({ groups }) => {
  if (groups.skipped) {
    return (
      <div className="deployment-groups">
        <div className="deployment-message warning">
          <AlertTriangle size={13} /> {groups.skipped}
        </div>
      </div>
    );
  }
  // Neutral chips: here the outcome matters, not include vs exclude, and red
  // would read as a failure.
  const chip = (kind, group) => group && (
    <span key={kind} className="pas-chip" title={group.id}>
      {kind === 'exclude' ? <Ban size={12} /> : <Users size={12} />}
      <span className="pas-muted">{kind === 'exclude' ? 'Exclude:' : 'Include:'}</span>
      {group.displayName}
      <span className={group.created ? 'group-created' : 'pas-muted'}>
        · {group.created ? 'created' : 'existing group'}{group.dynamic ? ', dynamic' : ''}
      </span>
    </span>
  );
  return (
    <div className="deployment-groups">
      <div className="pas-chips">
        {groups.builtIn && (
          <span className="pas-chip">
            <Globe size={12} />
            <span className="pas-muted">Include:</span>
            {builtInLabel(groups.builtIn)}
          </span>
        )}
        {chip('include', groups.include)}
        {chip('exclude', groups.exclude)}
        {groups.filter && (
          <span className="pas-chip pas-chip-filter" title={groups.filter.id}>
            <Filter size={12} />
            <span className="pas-muted">{filterModeLabel(groups.filter.mode)}:</span>
            {groups.filter.name || groups.filter.id}
          </span>
        )}
      </div>
      {groups.error ? (
        <div className="deployment-message error">
          <AlertTriangle size={13} /> Assignment not fully set up: {String(groups.error.message || groups.error).slice(0, 300)}
        </div>
      ) : (
        <div className="deployment-message success">{outcomeMessage(groups)}</div>
      )}
    </div>
  );
};

const DeploymentProgress = ({ 
  selectedPolicies, 
  results, 
  isLoading, 
  showActions = false, 
  onNewDeployment 
}) => {
  const getStatusIcon = (result) => {
    if (result.success) {
      return <CheckCircle className="status-icon success" />;
    } else {
      return <XCircle className="status-icon error" />;
    }
  };

  const getOverallStats = () => {
    const totalPolicies = selectedPolicies.length;
    const resultsArray = Array.isArray(results) ? results : [];
    const completedResults = resultsArray.length;
    const successCount = resultsArray.filter(r => r.success).length;
    const failureCount = resultsArray.filter(r => !r.success).length;
    const pendingCount = totalPolicies - completedResults;

    return {
      total: totalPolicies,
      completed: completedResults,
      success: successCount,
      failure: failureCount,
      pending: pendingCount,
      progressPercentage: totalPolicies > 0 ? (completedResults / totalPolicies) * 100 : 0
    };
  };

  const stats = getOverallStats();

  return (
    <div className="deployment-progress">
      <div className="progress-header">
        <h2>
          {isLoading ? 'Deploying Policies...' : 'Deployment Results'}
        </h2>
        <p>
          {isLoading 
            ? `Processing ${stats.total} selected policies`
            : `Deployment completed: ${stats.success} successful, ${stats.failure} failed`
          }
        </p>
      </div>

      <div className="progress-stats">
        <div className="stat-card">
          <div className="stat-value">{stats.total}</div>
          <div className="stat-label">Total Policies</div>
        </div>
        
        <div className="stat-card success">
          <div className="stat-value">{stats.success}</div>
          <div className="stat-label">Successful</div>
        </div>
        
        <div className="stat-card error">
          <div className="stat-value">{stats.failure}</div>
          <div className="stat-label">Failed</div>
        </div>
        
        <div className="stat-card pending">
          <div className="stat-value">{stats.pending}</div>
          <div className="stat-label">Pending</div>
        </div>
      </div>

      <div className="progress-bar-container">
        <div className="progress-bar">
          <div 
            className="progress-fill"
            style={{ width: `${stats.progressPercentage}%` }}
          />
        </div>
        <div className="progress-text">
          {Math.round(stats.progressPercentage)}% Complete
        </div>
      </div>

      <div className="deployment-list">
        {selectedPolicies.map((policy, index) => {
          const resultsArray = Array.isArray(results) ? results : [];
          const result = resultsArray.find(r => r.index === index);
          const isPending = !result && isLoading;
          const isCompleted = !!result;

          return (
            <div key={`${policy.fileName}-${index}`} className="deployment-item">
              <div className="deployment-status">
                {isPending && (
                  <div className="status-pending">
                    <Loader className="status-icon spinner" />
                  </div>
                )}
                {isCompleted && getStatusIcon(result)}
                {!isLoading && !result && (
                  <div className="status-waiting">
                    <Clock className="status-icon waiting" />
                  </div>
                )}
              </div>

              <div className="deployment-info">
                <div className="policy-name">
                  {policy.actualDisplayName || policy.name}
                </div>
                <div className="policy-details">
                  <span className="policy-type">{policy.policyType}</span>
                  <span className="policy-filename">{policy.fileName}</span>
                </div>
                
                {result && (
                  <div className={`deployment-message ${result.success ? 'success' : 'error'}`}>
                    {result.message}
                  </div>
                )}
                {result?.groups && <GroupOutcome groups={result.groups} />}
                
                {isPending && (
                  <div className="deployment-message pending">
                    Deploying to tenant...
                  </div>
                )}
              </div>

              <div className="deployment-time">
                {isCompleted && (
                  <span className="completion-time">
                    {new Date().toLocaleTimeString()}
                  </span>
                )}
              </div>
            </div>
          );
        })}
      </div>

      {showActions && (
        <div className="deployment-actions">
          <button
            onClick={onNewDeployment}
            className="btn btn-primary"
          >
            <ArrowLeft className="btn-icon" />
            Deploy More Policies
          </button>
          
          <button
            onClick={() => window.location.reload()}
            className="btn btn-secondary"
          >
            <RotateCcw className="btn-icon" />
            Refresh Application
          </button>
        </div>
      )}
    </div>
  );
};

export default DeploymentProgress;
