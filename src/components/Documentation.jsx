import React from 'react';
import { CheckCircle, AlertTriangle, Info, ExternalLink, ArrowLeft, Rocket, HelpCircle, Wrench, Settings, BookOpen } from 'lucide-react';

const Documentation = ({ onBack }) => {
  return (
    <div className="documentation">
      <div className="documentation-header">
        <button onClick={onBack} className="btn btn-secondary back-btn">
          <ArrowLeft size={16} />
          Back to Main
        </button>
        <h1>OIB Deployer - Documentation & FAQ</h1>
      </div>

      <div className="documentation-content">
        {/* Quick Start Guide */}
        <section className="doc-section">
          <h2><Rocket size={22} className="section-icon" /> Quick Start Guide</h2>
          <div className="steps">
            <div className="step">
              <span className="step-number">1</span>
              <div className="step-content">
                <h3>Authentication</h3>
                <p>Click "Sign in with Microsoft" to authenticate with your Entra ID tenant using your Microsoft account.</p>
              </div>
            </div>
            <div className="step">
              <span className="step-number">2</span>
              <div className="step-content">
                <h3>Choose Deployment Type</h3>
                <p>Select "New Deployment" if you're deploying the OIB for the first time.</p>
                <p>Select "Existing Deployment" if you've already deployed the OIB and want to compare it against the latest version, then deploy any missing or outdated policies.</p>
                <p>Select "Policy Validation" to check your already-deployed OIB policies for per-setting configuration drift against the baseline &mdash; this is read-only and makes no changes to your tenant.</p>
                <p>Select "Device Validation" to pick a specific managed device and see which OIB policies have actually been delivered to it &mdash; also read-only.</p>
              </div>
            </div>
            <div className="step">
              <span className="step-number">3</span>
              <div className="step-content">
                <h3>Select Operating System(s)</h3>
                <p>Select the operating system(s) you want to deploy (e.g., Windows, macOS).</p>
              </div>
            </div>
            <div className="step">
              <span className="step-number">4</span>
              <div className="step-content">
                <h3>Deploy, Compare, or Validate</h3>
                <p><strong>New:</strong> Answer a couple of quick questions about your tenant's licensing and Windows Update strategy (so irrelevant policies can be filtered out), then select and deploy the policies you want.</p>
                <p><strong>Existing:</strong> Compare your tenant against the latest OIB version, then deploy any missing or outdated policies. Use "Load assignments &amp; device status" to see who each deployed policy is assigned to and how many devices report success, errors or conflicts.</p>
                <p><strong>Validation:</strong> Review a per-setting drift report for each matched policy against the OIB baseline.</p>
                <p><strong>Device Validation:</strong> Search for a device and see, per OIB policy, whether it applied, errored, conflicted, is pending, or isn't assigned to that device.</p>
              </div>
            </div>
          </div>
        </section>

        {/* FAQ Section */}
        <section className="doc-section">
          <h2><HelpCircle size={22} className="section-icon" /> Frequently Asked Questions</h2>
          
          <div className="faq-item">
            <h3>What is OIB Deployer?</h3>
            <p>
              OIB Deployer is a web-based tool that simplifies the deployment of OpenIntuneBaseline policies to Microsoft Intune tenants. 
              It provides a user-friendly interface to select and deploy standardized security policies without manual configuration.
            </p>
          </div>

          <div className="faq-item">
            <h3>What policies can I deploy?</h3>
            <p>You can deploy various types of policies including:</p>
            <ul>
              <li><strong>Compliance Policies:</strong> Device health, security, and password requirements</li>
              <li><strong>Settings Catalog:</strong> Advanced device configurations and security settings</li>
              <li><strong>Endpoint Security:</strong> Antivirus, firewall, encryption, and attack surface reduction</li>
              <li><strong>Windows Update:</strong> Feature and driver update policies</li>
              <li><strong>Device Configuration:</strong> Basic device settings and restrictions</li>
            </ul>
          </div>

          <div className="faq-item">
            <h3>What permissions do I need?</h3>
            <p>The OIB Deployer Enterprise Application only requires you to have access to the following Graph scope via Delegated permissions:</p>
            <ul>
              <li><code>DeviceManagementConfiguration.ReadWrite.All</code></li>
            </ul>
            <p>If your account has Intune Administrator role (or higher), this will allow the above.</p>
            <p>
              Device Validation additionally requests <code>DeviceManagementManagedDevices.Read.All</code> (read-only) the first time you use it.
              This permission requires admin consent, so you may need an administrator to approve it for the app registration.
            </p>
            <p>
              To show group names in the Existing Deployment assignment overview, the app can also request <code>GroupMember.Read.All</code> (read-only,
              admin consent). Without it, assigned groups are shown by their ID.
            </p>
            <p>
              Creating assignment groups and migrating between policy versions requests <code>Group.ReadWrite.All</code>,
              <code>Device.Read.All</code> and <code>User.ReadBasic.All</code> (admin consent) the first time you turn it on.
            </p>
          </div>

          <div className="faq-item">
            <h3>Will this duplicate existing policies?</h3>
            <p>
              The tool primarily matches existing tenant policies using a hidden OIBID reference embedded in the policy description (available on OIB v3.8+ baselines), 
              falling back to name-based matching for older baselines without this metadata. If no match is found and you select the policy anyway, a policy with a 
              duplicate name could be created.
            </p>
          </div>

          <div className="faq-item">
            <h3>Why are some policy types missing or greyed out during New Deployment?</h3>
            <p>
              New Deployment asks a few quick questions about your tenant (Business Premium vs Microsoft 365 E3/E5/E7 licensing, whether Defender for Endpoint is your 
              primary antivirus, and whether you use Autopatch) so it can filter out policies you can't use or that would conflict with your setup:
            </p>
            <ul>
              <li>Policies requiring a Windows Enterprise entitlement (e.g. Driver Update Policies) are greyed out on Business Premium tenants</li>
              <li>Policies requiring Defender for Endpoint are hidden by default if it isn't your primary antivirus</li>
              <li>Update Policies and Driver Update Policies are greyed out if you're using Autopatch, since it manages Windows Update itself</li>
            </ul>
            <p>Individually-gated policies (not whole categories) can still be revealed via the "Show them anyway" toggle if you want to override the recommendation.</p>
          </div>

          <div className="faq-item">
            <h3>What is Policy Validation?</h3>
            <p>
              Policy Validation is a read-only flow that compares each matched policy's individual settings against the OIB baseline, highlighting any drift: 
              mismatched values, settings present in the baseline but missing from the tenant, or extra settings present in the tenant but not the baseline. 
              It currently supports Settings Catalog, Endpoint Security, and Compliance policies. No changes are ever made to your tenant during validation.
            </p>
          </div>

          <div className="faq-item">
            <h3>Can I see who an OIB policy is assigned to and whether it works?</h3>
            <p>
              Yes. In Existing Deployment, click "Load assignments &amp; device status". Each deployed OIB policy then shows its assignment targets
              (All devices, All users, included/excluded groups and assignment filters) and the number of devices reporting succeeded, error, conflict,
              not compliant, pending or not applicable, with a per-device list on demand. Use the "Assigned", "Not Assigned" and "Device Errors" tabs to filter the list.
              Device counts are devices that have reported on the policy, as on the policy's overview page in Intune, not the size of the assigned groups.
            </p>
          </div>

          <div className="faq-item">
            <h3>What is Device Validation?</h3>
            <p>
              Device Validation shows whether the OIB policies in your tenant have actually reached a specific device. It reads Intune's per-device
              report (the same data as <em>Devices &rsaquo; device &rsaquo; Device configuration</em> in the Intune admin center) plus the device's compliance
              policy states, and lists each OIB policy as Applied, Error, Conflict, Pending, or Not deployed in the tenant. When the device has no
              report for a policy, the policy's assignments are read to tell "assigned, but not reported by the device yet" apart from "no assignments in the tenant".
              App Protection policies and Windows feature/quality/driver update profiles aren't reported per device by these APIs and are marked as such.
              When an assigned policy uses an assignment filter, Intune is asked how the filter was evaluated for the device at its last check-in:
              a policy whose filters leave the device out on purpose is shown as <em>Filtered out</em> rather than as a problem.
            </p>
          </div>

          <div className="faq-item">
            <h3>Will it assign any policies deployed?</h3>
            <p>
              Not by default. Next to the Deploy button you can turn on "Assign each deployed policy" and choose what to assign it to:
            </p>
            <ul>
              <li>
                <strong>A new include group per policy:</strong> the app creates (or reuses, if one with the same name exists) an include and/or exclude
                security group named <code>{'{prefix}{policy name} - Include{suffix}'}</code> (the labels, prefix and suffix are yours to choose) and assigns
                them to the policy. The groups start empty, so the policy reaches no one until you add devices or users to the include group. This is the
                way to pilot a policy first.
              </li>
              <li>
                <strong>All devices</strong> or <strong>All users</strong>, optionally with an exclude group per policy. Microsoft recommends these built-in
                groups over your own "everyone" groups when a policy should reach everyone, as they need no group processing. Put devices in the exclude
                group of an All devices assignment and users in that of an All users assignment: Intune doesn't support excluding user groups from device
                assignments, or the other way round.
              </li>
              <li>
                <strong>All devices or All users by scope:</strong> OIB policies marked D (device) go to All devices and those marked U (user) to All users.
              </li>
            </ul>
            <p>
              With any of these you can also pick an existing <strong>assignment filter</strong> per platform (Windows, macOS, …) and whether it
              includes or excludes the devices it matches. The filter goes on the include target, e.g. All devices with the include filter
              &quot;Corporate Windows 11&quot;. Filters are evaluated at check-in, so they need no group processing; exclude groups can&apos;t have filters.
              Create filters in Intune; the app only lists the ones that fit each policy&apos;s platform.
            </p>
            <p>
              App protection policies and Windows feature, quality and driver update profiles can&apos;t be assigned to All devices or All users and are left
              unassigned with that option. Without any of these options, assign the policies yourself as per OIB guidance.
            </p>
          </div>

          <div className="faq-item">
            <h3>How do I move devices from an old OIB version to the new one?</h3>
            <p>
              When an older version of a policy is still deployed next to the new one, Existing Deployment shows a "Version Migrations" tab and a
              "Migrate devices…" button on the policy. The dialog lets you pilot the new version on a few devices or users first, move them back if
              needed, and finish when you're ready:
            </p>
            <ul>
              <li>
                Moved devices or users are added to the new version&apos;s pilot group, <em>{'{prefix}{policy name} - Pilot{suffix}'}</em>, and to the
                old version&apos;s exclude group. Both are created on the first move.
              </li>
              <li>
                The old version&apos;s include groups, All devices or All users assignments are never changed during the migration. An exclusion
                always wins in Intune, also when a device is in a group through a nested or dynamic group, and shared groups stay intact for
                other policies. &quot;Move back&quot; removes the member from both groups again.
              </li>
              <li>
                "Finish migration" lets you choose how the new version is assigned afterwards:
                <ul>
                  <li>
                    <strong>All devices / All users:</strong> the pilot group is unassigned and its members are taken out (the empty group is
                    kept). New devices get the policy automatically.
                  </li>
                  <li>
                    <strong>Include/exclude groups:</strong> the pilot group is renamed to <em>… - Include</em>, or, if the new version already has an
                    include group, its members join that one and the pilot group is unassigned and emptied. Everyone still on the old version
                    joins too. If the old version was on All devices, that adds every device of the platform as a snapshot, so devices enrolled
                    later must be added by hand.
                  </li>
                </ul>
                Either way, all assignments are then removed from the old version and everyone is taken out of its own groups (e.g.{' '}
                <em>… v3.7 - Include</em> and <em>… v3.7 - Exclude</em>). The empty groups are kept for you to delete, and groups shared with other
                policies aren&apos;t changed. Devices that were excluded from the old version for other reasons than the migration stay excluded:
                they move to the new version&apos;s exclude group first. The old policy itself is not deleted.
              </li>
              <li>
                If the new version is already assigned to All devices or All users, everyone gets both versions and there is nothing to pilot.
                The dialog then offers to take the new version off All devices so you can pilot it, or to finish right away.
              </li>
              <li>
                <strong>Pilot through an assignment filter:</strong> when the old version is assigned to All devices or All users, with or without a
                filter, you can pilot through a filter instead of a group. Pilot devices are listed by name in the filter
                <em>{' {prefix}{policy name} - Pilot{suffix}'}</em>. The new version gets the same built-in target with that filter as an include filter;
                the old version gets it as an exclude filter, or, if it already had a filter, a copy of that filter without the pilot devices
                (<em>… - Not piloted</em>). Finishing gives the new version the old version&apos;s built-in target with its original filter and deletes the
                pilot filters. One filter rule holds up to 3,072 characters, roughly 100 device names.
              </li>
              <li>
                Finishing on All devices / All users always keeps the old version&apos;s assignment filter, so a filtered rollout stays filtered.
              </li>
            </ul>
          </div>

          <div className="faq-item">
            <h3>Where do I see assignment filters?</h3>
            <p>
              Existing Deployment&apos;s assignment overview, Policy Validation (&quot;Load assignments&quot;) and the migration dialog show each policy&apos;s
              targets with their filters and whether they include or exclude. Policy Validation&apos;s CSV export has an &quot;Assigned To&quot; column once
              assignments are loaded. Device Validation marks policies as <em>Filtered out</em> when Intune&apos;s filter evaluation leaves the device out.
            </p>
          </div>

          <div className="faq-item">
            <h3>Can I deploy to production tenants?</h3>
            <div className="warning-box">
              <AlertTriangle size={16} />
              <div>
                <strong>Use with caution:</strong>
                <ul>
                  <li>Always test policies in a development or test tenant first.</li>
                  <li>These are baseline security policies that may impact user experience and device functionality.</li>
                  <li>The OIB was not designed to be assigned to devices previously managed in other ways (e.g. GPO, or using other baseline configurations)</li>
                </ul>
              </div>
            </div>
          </div>

          <div className="faq-item">
            <h3>What happens if deployment fails?</h3>
            <p>
              If a policy fails to deploy, you'll see detailed error information in the results section. 
              Common causes include:
            </p>
            <ul>
              <li>Insufficient permissions</li>
              <li>Policy conflicts with existing configurations</li>
              <li>Network connectivity issues</li>
              <li>Invalid policy content or format</li>
            </ul>
          </div>

          <div className="faq-item">
            <h3>How do I customize policies before deployment?</h3>
            <p>
              Currently, the tool deploys policies as-is from the OpenIntuneBaseline repository. 
              For customization, you can:
            </p>
            <ul>
              <li>Deploy the baseline policies first</li>
              <li>Duplicate the policies in Intune, giving them a customised name</li>
              <li>Modify the duplicated policies as needed</li>
            </ul>
          </div>
        </section>

        {/* Troubleshooting */}
        <section className="doc-section">
          <h2><Wrench size={22} className="section-icon" /> Troubleshooting</h2>
          
          <div className="troubleshoot-item">
            <h3>Authentication Issues</h3>
            <ul>
              <li>Ensure your account has the necessary permissions</li>
              <li>Try logging out and logging back in</li>
              <li>Check if your tenant allows the required Graph API permissions</li>
              <li>Verify that pop-ups are not blocked in your browser</li>
            </ul>
          </div>

          <div className="troubleshoot-item">
            <h3>Policy Deployment Errors</h3>
            <ul>
              <li>Check the browser console for detailed error messages</li>
              <li>Verify your account has write permissions to Intune</li>
              <li>Ensure the policy doesn't conflict with existing configurations</li>
              <li>Try deploying policies one at a time to isolate issues</li>
            </ul>
          </div>

          <div className="troubleshoot-item">
            <h3>Network/Loading Issues</h3>
            <ul>
              <li>Check your internet connection</li>
              <li>Disable VPN if experiencing connection issues</li>
              <li>Clear browser cache and cookies</li>
              <li>Try using a different browser or incognito mode</li>
            </ul>
          </div>
        </section>

        {/* Technical Information */}
        <section className="doc-section">
          <h2><Settings size={22} className="section-icon" /> Technical Information</h2>
          
          <div className="tech-info">
            <h3>Policy Sources</h3>
            <p>
              All policies are loaded automatically from the&nbsp;
              <a href="https://github.com/SkipToTheEndpoint/OpenIntuneBaseline" target="_blank" rel="noopener noreferrer">
                OpenIntuneBaseline GitHub Repo<ExternalLink size={14} /></a> via the GitHub API.
            </p>
          </div>

          <div className="tech-info">
            <h3>Compatibility</h3>
            <p>
              The import functionality is designed to match the behavior of the&nbsp;
              <a href="https://github.com/Micke-K/IntuneManagement" target="_blank" rel="noopener noreferrer">
                IntuneManagement PowerShell module <ExternalLink size={14} />
              </a>, 
              ensuring compatibility with existing workflows and policy formats. Enhanced matching (via OIBID) and licensing-based filtering require a 
              <code>PolicyManifest.json</code> file, only present on OIB v3.8+ branches &mdash; older branches still work via name-based matching, without licensing filters.
            </p>
          </div>

          <div className="tech-info">
            <h3>Data Privacy</h3>
            <p>
              This application runs entirely in your browser. No policy data or tenant information is stored on external servers. 
              Authentication is handled directly through Microsoft's MSAL library, and all API calls go directly to Microsoft Graph.
            </p>
          </div>
        </section>

        {/* Support & Resources */}
        <section className="doc-section">
          <h2><BookOpen size={22} className="section-icon" /> Resources & Support</h2>
          
          <div className="resources">
            <div className="resource-item">
              <h3>Official Documentation</h3>
              <ul>
                <li>
                  <a href="https://github.com/SkipToTheEndpoint/OpenIntuneBaseline" target="_blank" rel="noopener noreferrer">
                    OpenIntuneBaseline Documentation <ExternalLink size={14} />
                  </a>
                </li>
                <li>
                  <a href="https://docs.microsoft.com/en-us/mem/intune/" target="_blank" rel="noopener noreferrer">
                    Microsoft Intune Documentation <ExternalLink size={14} />
                  </a>
                </li>
                <li>
                  <a href="https://docs.microsoft.com/en-us/graph/api/resources/intune-graph-overview" target="_blank" rel="noopener noreferrer">
                    Microsoft Graph Intune API <ExternalLink size={14} />
                  </a>
                </li>
              </ul>
            </div>

            <div className="resource-item">
              <h3>Community</h3>
              <ul>
                <li>
                  <a href="https://discord.com/invite/winadmins" target="_blank" rel="noopener noreferrer">
                    WinAdmins Discord <ExternalLink size={14} />
                  </a>
                </li>
                <li>
                  <a href="https://www.reddit.com/r/Intune/" target="_blank" rel="noopener noreferrer">
                    r/Intune on Reddit <ExternalLink size={14} />
                  </a>
                </li>
              </ul>
            </div>
          </div>
        </section>

        {/* Version Info */}
        <section className="doc-section version-info">
          <h2><Info size={22} className="section-icon" /> Version Information</h2>
          <div className="info-box">
            <Info size={16} />
            <div>
              <p><strong>OIB Deployer Version:</strong> v1.3.1</p>
              <p><strong>Last Updated:</strong> November 2026</p>
              <p><strong>Compatible with:</strong> OpenIntuneBaseline v3.0+ (OIBID matching and licensing-based filtering require v3.8+)</p>
            </div>
          </div>
        </section>
      </div>
    </div>
  );
};

export default Documentation;
