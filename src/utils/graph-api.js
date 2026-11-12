import { CONFIG } from './config.js';
import authService from './auth.js';
import logger from './logger.js';
import {
    buildGroupName,
    buildMailNickname,
    builtInTarget,
    deviceMatchesPlatform,
    groupTarget,
    mergeAssignmentTargets,
    normaliseDirectoryObject,
    planPolicyAssignment,
    withFilter
} from './assignment-groups.js';

class GraphAPI {
    constructor() {
        this.config = CONFIG.graph;
        this.retryAttempts = this.config.batch.maxRetries;
        this.retryDelay = this.config.batch.retryDelay;
        this.cachedPolicies = null; // Cache for existing policies
    }

    // Replace tenant-specific variables in policy content
    replaceTenantVariables(policyContent, tenantId) {
        if (!policyContent || !tenantId) {
            return policyContent;
        }

        try {
            // Convert to string for replacement if it's an object
            let contentStr = typeof policyContent === 'string' ? 
                policyContent : JSON.stringify(policyContent);
            
            // Replace %OrganizationId% with actual tenant ID (case-insensitive)
            contentStr = contentStr.replace(/%OrganizationId%/gi, tenantId);
            
            // Log if any replacements were made
            if (contentStr.includes(tenantId) && contentStr !== JSON.stringify(policyContent)) {
                logger.debug(`Tenant variable replacement: %OrganizationId% -> ${tenantId}`);
            }
            
            // Parse back to object if original was an object
            return typeof policyContent === 'string' ? 
                contentStr : JSON.parse(contentStr);
                
        } catch (error) {
            logger.error('Error during tenant variable replacement:', error);
            return policyContent; // Return original on error
        }
    }

    // Helper method to get authorization headers
    async getAuthHeaders() {
        const token = await authService.getAccessToken();
        return {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        };
    }

    // Make authenticated request to Microsoft Graph with pagination support
    async makeRequest(endpoint, method = 'GET', body = null) {
        try {
            const url = `${this.config.baseUrl}${endpoint}`;
            logger.debug(`Making Graph API request: ${method} ${endpoint}`);

            const options = {
                method,
                headers: await this.getAuthHeaders()
            };

            if (body && (method === 'POST' || method === 'PUT' || method === 'PATCH')) {
                options.body = JSON.stringify(body);
            }

            const response = await this.makeRequestWithRetry(url, options);

            // Handle error responses
            if (!response.ok) {
                const errorText = await response.text();
                logger.error(`Graph API error response: ${response.status} ${response.statusText}`, errorText);
                throw new Error(`Graph API error: ${response.status} ${response.statusText} - ${errorText}`);
            }

            // Handle no content responses
            if (response.status === 204) {
                return null;
            }

            const result = await response.json();
            logger.debug(`Graph API request successful: ${method} ${url}`);
            return result;
        } catch (error) {
            logger.error('Graph API request failed:', error);
            throw error;
        }
    }

    // Make paginated request to handle large datasets
    async makePaginatedRequest(endpoint) {
        try {
            let allResults = [];
            let nextLink = `${this.config.baseUrl}${endpoint}`;
            
            while (nextLink) {
                console.log(`Making paginated Graph API request: ${nextLink}`);
                
                const options = {
                    method: 'GET',
                    headers: await this.getAuthHeaders()
                };

                const response = await this.makeRequestWithRetry(nextLink, options);

                if (!response.ok) {
                    const errorText = await response.text();
                    console.error(`Graph API error response: ${response.status} ${response.statusText}`, errorText);
                    throw new Error(`Graph API error: ${response.status} ${response.statusText} - ${errorText}`);
                }

                const result = await response.json();
                
                // Add results to our collection
                if (result.value && Array.isArray(result.value)) {
                    allResults = allResults.concat(result.value);
                }

                // Check for next page
                nextLink = result['@odata.nextLink'] || null;
                
                console.log(`Retrieved ${result.value?.length || 0} items, total so far: ${allResults.length}`);
            }

            console.log(`Completed paginated request, total items: ${allResults.length}`);
            return { value: allResults };
        } catch (error) {
            console.error('Paginated Graph API request failed:', error);
            throw error;
        }
    }

    // Make request with retry logic
    async makeRequestWithRetry(url, options, attempt = 1) {
        try {
            const response = await fetch(url, options);
            
            // Handle rate limiting (429) and server errors (5xx)
            if ((response.status === 429 || response.status >= 500) && attempt <= this.retryAttempts) {
                const retryAfterHeader = response.headers.get('Retry-After');
                const delay = retryAfterHeader
                    ? parseInt(retryAfterHeader, 10) * 1000
                    : this.retryDelay * Math.pow(2, attempt - 1);
                console.log(`Request failed with status ${response.status}, retrying in ${delay}ms (attempt ${attempt}/${this.retryAttempts})`);
                
                await this.sleep(delay);
                return this.makeRequestWithRetry(url, options, attempt + 1);
            }

            return response;
        } catch (error) {
            if (attempt <= this.retryAttempts) {
                const delay = this.retryDelay * Math.pow(2, attempt - 1);
                console.log(`Request failed with error, retrying in ${delay}ms (attempt ${attempt}/${this.retryAttempts}):`, error.message);
                
                await this.sleep(delay);
                return this.makeRequestWithRetry(url, options, attempt + 1);
            }
            
            throw error;
        }
    }

    // Sleep utility for retry delays
    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    // Extract OIBID from a tenant policy's description field.
    // Format: "OIBID:{GUID}" appended at the end of the description (introduced in OIB v3.8).
    // Returns the GUID string (uppercased) or null if not present.
    extractOibId(policy) {
        const description = policy?.description ?? null;
        if (!description) return null;
        const match = description.match(/OIBID:([0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12})/i);
        return match ? match[1].toUpperCase() : null;
    }

    // Endpoint groups per OS type.
    // Each group is { id, endpoints[], normalise(policy) } where endpoints is a list
    // of relative Graph paths to fetch. Multiple endpoints in one group (e.g. Update
    // policies) are merged into a single result array.
    static getEndpointGroupsForOS(osTypes) {
        const windows = osTypes.includes('WINDOWS');
        const macos   = osTypes.includes('MACOS');
        const byod    = osTypes.includes('BYOD');

        const groups = [];

        // $select ensures 'description' is always returned (needed for OIBID extraction)
        const sel    = '$select=id,displayName,description';
        const selName = '$select=id,name,description'; // configurationPolicies uses 'name' not 'displayName'

        // Settings Catalog – all platforms (uses 'name' field, not 'displayName')
        groups.push({
            id: 'configurationPolicy',
            endpoints: [`/deviceManagement/configurationPolicies?${selName}&$top=100`],
            normalise: p => ({ id: p.id, displayName: p.name, oibId: null, type: 'configurationPolicy', policy: p })
        });

        // Compliance policies – all platforms
        groups.push({
            id: 'compliance',
            endpoints: [`/deviceManagement/deviceCompliancePolicies?${sel}`],
            normalise: p => ({ id: p.id, displayName: p.displayName, oibId: null, type: 'compliance', policy: p })
        });

        if (windows || macos) {
            // Device Configuration profiles
            groups.push({
                id: 'deviceConfiguration',
                endpoints: [`/deviceManagement/deviceConfigurations?${sel}`],
                normalise: p => ({ id: p.id, displayName: p.displayName, oibId: null, type: 'deviceConfiguration', policy: p })
            });

            // Admin templates (ADMX / Group Policy)
            groups.push({
                id: 'adminTemplate',
                endpoints: [`/deviceManagement/groupPolicyConfigurations?${sel}`],
                normalise: p => ({ id: p.id, displayName: p.displayName, oibId: null, type: 'adminTemplate', policy: p })
            });
        }

        if (windows) {
            // Endpoint Security (legacy intents)
            groups.push({
                id: 'endpointSecurity',
                endpoints: [`/deviceManagement/intents?${sel}`],
                normalise: p => ({ id: p.id, displayName: p.displayName, oibId: null, type: 'endpointSecurity', policy: p })
            });

            // Windows Update for Business (three endpoints merged into one group)
            groups.push({
                id: 'updatePolicy',
                endpoints: [
                    `/deviceManagement/windowsFeatureUpdateProfiles?${sel}`,
                    `/deviceManagement/windowsQualityUpdateProfiles?${sel}`,
                    `/deviceManagement/windowsDriverUpdateProfiles?${sel}`
                ],
                normalise: p => ({ id: p.id, displayName: p.displayName, oibId: null, type: 'updatePolicy', policy: p })
            });
        }

        if (byod || macos) {
            // App Protection Policies
            groups.push({
                id: 'appProtection',
                endpoints: [
                    `/deviceAppManagement/iosManagedAppProtections?${sel}`,
                    `/deviceAppManagement/androidManagedAppProtections?${sel}`
                ],
                normalise: p => ({ id: p.id, displayName: p.displayName, oibId: null, type: 'appProtection', policy: p })
            });
        }

        return groups;
    }

    // Fetch existing tenant policies scoped to the selected OS types using Graph $batch.
    // Batch limit is 20 requests per call; each endpoint is one batch item.
    // If a response page has @odata.nextLink, subsequent pages are fetched individually.
    async getPoliciesForOSTypes(osTypes) {
        const groups = GraphAPI.getEndpointGroupsForOS(osTypes);

        // Flatten to individual endpoint requests
        const requests = groups.flatMap(group =>
            group.endpoints.map((endpoint, i) => ({
                batchId: `${group.id}__${i}`,
                groupId: group.id,
                endpoint
            }))
        );

        console.log(`Fetching tenant policies for [${osTypes.join(', ')}] via Graph batch (${requests.length} endpoints)`);

        // Issue requests in batches of 20 (Graph hard limit).
        // If any batch items return 429, collect them and retry after the Retry-After delay.
        const BATCH_LIMIT = 20;
        const MAX_THROTTLE_RETRIES = 3;
        const rawResults = new Map(); // batchId -> array of policy objects

        let pendingRequests = [...requests];

        for (let throttleAttempt = 0; throttleAttempt < MAX_THROTTLE_RETRIES && pendingRequests.length > 0; throttleAttempt++) {
            const throttledRequests = [];
            let retryAfterMs = 0;

            for (let offset = 0; offset < pendingRequests.length; offset += BATCH_LIMIT) {
                const slice = pendingRequests.slice(offset, offset + BATCH_LIMIT);

                const batchBody = {
                    requests: slice.map(req => ({
                        id: req.batchId,
                        method: 'GET',
                        url: req.endpoint
                    }))
                };

                const batchUrl = `${this.config.baseUrl}${this.config.endpoints.batch}`;
                const headers = await this.getAuthHeaders();

                const response = await this.makeRequestWithRetry(batchUrl, {
                    method: 'POST',
                    headers,
                    body: JSON.stringify(batchBody)
                });

                if (!response.ok) {
                    const errorText = await response.text();
                    throw new Error(`Graph batch request failed: ${response.status} - ${errorText}`);
                }

                const batchResponse = await response.json();

                for (const res of batchResponse.responses) {
                    if (res.status === 429) {
                        // Batch item throttled — queue for retry respecting Retry-After
                        const ra = res.headers?.['Retry-After'] ?? res.headers?.['retry-after'];
                        retryAfterMs = Math.max(retryAfterMs, ra ? parseInt(ra, 10) * 1000 : 10000);
                        const original = slice.find(r => r.batchId === res.id);
                        if (original) throttledRequests.push(original);
                        continue;
                    }

                    if (res.status !== 200) {
                        console.warn(`Batch item ${res.id} returned status ${res.status} — skipping`);
                        rawResults.set(res.id, []);
                        continue;
                    }

                    const items = res.body?.value || [];
                    rawResults.set(res.id, items);

                    // Follow nextLinks outside the batch
                    let nextLink = res.body?.['@odata.nextLink'] || null;
                    while (nextLink) {
                        const pageHeaders = await this.getAuthHeaders();
                        const pageResp = await this.makeRequestWithRetry(nextLink, { method: 'GET', headers: pageHeaders });
                        if (!pageResp.ok) break;
                        const pageData = await pageResp.json();
                        rawResults.get(res.id).push(...(pageData.value || []));
                        nextLink = pageData['@odata.nextLink'] || null;
                    }

                    console.log(`Batch item ${res.id}: ${rawResults.get(res.id).length} policies`);
                }
            }

            if (throttledRequests.length > 0) {
                const waitMs = retryAfterMs || 10000;
                console.warn(`${throttledRequests.length} batch item(s) throttled (429). Waiting ${waitMs}ms before retry (attempt ${throttleAttempt + 1}/${MAX_THROTTLE_RETRIES})...`);
                await this.sleep(waitMs);
                pendingRequests = throttledRequests;
            } else {
                pendingRequests = [];
            }
        }

        if (pendingRequests.length > 0) {
            console.warn(`${pendingRequests.length} batch item(s) still throttled after ${MAX_THROTTLE_RETRIES} retries — results may be incomplete`);
            for (const req of pendingRequests) {
                if (!rawResults.has(req.batchId)) rawResults.set(req.batchId, []);
            }
        }

        // Assemble normalised policy list, grouped by groupId, then merge per group
        const groupMap = new Map(groups.map(g => [g.id, { normalise: g.normalise, items: [] }]));

        for (const req of requests) {
            const items = rawResults.get(req.batchId) || [];
            // Remember which collection each policy came from — update and
            // app protection groups merge several endpoints, and assignment
            // calls need the exact one.
            const collection = req.endpoint.split('?')[0];
            groupMap.get(req.groupId).items.push(...items.map(item => ({ item, collection })));
        }

        const allPolicies = [];
        for (const [, group] of groupMap) {
            for (const { item: policy, collection } of group.items) {
                const normalised = group.normalise(policy);
                normalised.oibId = this.extractOibId(policy);
                normalised.endpoint = collection;
                allPolicies.push(normalised);
            }
        }

        console.log(`Retrieved ${allPolicies.length} tenant policies via batch for OS: [${osTypes.join(', ')}]`);
        this.cachedPolicies = allPolicies;
        return allPolicies;
    }

    // Get all policies from tenant for comparison
    async getAllPolicies() {
        try {
            console.log('Starting to fetch all policies from tenant...');
            
            const results = await Promise.allSettled([
                this.getDeviceCompliancePolicies(),
                this.getDeviceConfigurations(),
                this.getConfigurationPolicies(),
                this.getEndpointSecurityPolicies(),
                this.getAdministrativeTemplates(),
                this.getUpdatePolicies()
            ]);

            // Process results and log any failures
            const [
                compliancePoliciesResult,
                deviceConfigurationsResult,
                configurationPoliciesResult,
                endpointSecurityPoliciesResult,
                adminTemplatesResult,
                updatePoliciesResult
            ] = results;

            const compliancePolicies = compliancePoliciesResult.status === 'fulfilled' 
                ? compliancePoliciesResult.value 
                : [];
            const deviceConfigurations = deviceConfigurationsResult.status === 'fulfilled' 
                ? deviceConfigurationsResult.value 
                : [];
            const configurationPolicies = configurationPoliciesResult.status === 'fulfilled' 
                ? configurationPoliciesResult.value 
                : [];
            const endpointSecurityPolicies = endpointSecurityPoliciesResult.status === 'fulfilled' 
                ? endpointSecurityPoliciesResult.value 
                : [];
            const adminTemplates = adminTemplatesResult.status === 'fulfilled' 
                ? adminTemplatesResult.value 
                : [];
            const updatePolicies = updatePoliciesResult.status === 'fulfilled' 
                ? updatePoliciesResult.value 
                : [];

            // Log results for debugging
            console.log(`Policy fetch results:
                - Compliance Policies: ${compliancePoliciesResult.status === 'fulfilled' ? `${compliancePolicies.length} loaded` : `Failed: ${compliancePoliciesResult.reason}`}
                - Device Configurations: ${deviceConfigurationsResult.status === 'fulfilled' ? `${deviceConfigurations.length} loaded` : `Failed: ${deviceConfigurationsResult.reason}`}
                - Configuration Policies (Settings Catalog): ${configurationPoliciesResult.status === 'fulfilled' ? `${configurationPolicies.length} loaded` : `Failed: ${configurationPoliciesResult.reason}`}
                - Endpoint Security Policies: ${endpointSecurityPoliciesResult.status === 'fulfilled' ? `${endpointSecurityPolicies.length} loaded` : `Failed: ${endpointSecurityPoliciesResult.reason}`}
                - Administrative Templates: ${adminTemplatesResult.status === 'fulfilled' ? `${adminTemplates.length} loaded` : `Failed: ${adminTemplatesResult.reason}`}
                - Update Policies: ${updatePoliciesResult.status === 'fulfilled' ? `${updatePolicies.length} loaded` : `Failed: ${updatePoliciesResult.reason}`}`);

            // Combine all policies with consistent naming
            const allPolicies = [];
            
            // Add compliance policies
            compliancePolicies.forEach(policy => {
                allPolicies.push({
                    id: policy.id,
                    displayName: policy.displayName,
                    oibId: this.extractOibId(policy),
                    type: 'compliance',
                    policy: policy
                });
            });

            // Add device configurations
            deviceConfigurations.forEach(policy => {
                allPolicies.push({
                    id: policy.id,
                    displayName: policy.displayName,
                    oibId: this.extractOibId(policy),
                    type: 'deviceConfiguration',
                    policy: policy
                });
            });

            // Add configuration policies (Settings Catalog)
            configurationPolicies.forEach(policy => {
                allPolicies.push({
                    id: policy.id,
                    displayName: policy.name, // Settings Catalog uses 'name' instead of 'displayName'
                    oibId: this.extractOibId(policy),
                    type: 'configurationPolicy',
                    policy: policy
                });
            });

            // Add endpoint security policies
            endpointSecurityPolicies.forEach(policy => {
                allPolicies.push({
                    id: policy.id,
                    displayName: policy.displayName,
                    oibId: this.extractOibId(policy),
                    type: 'endpointSecurity',
                    policy: policy
                });
            });

            // Add administrative templates
            adminTemplates.forEach(policy => {
                allPolicies.push({
                    id: policy.id,
                    displayName: policy.displayName,
                    oibId: this.extractOibId(policy),
                    type: 'adminTemplate',
                    policy: policy
                });
            });

            // Add update policies
            updatePolicies.forEach(policy => {
                allPolicies.push({
                    id: policy.id,
                    displayName: policy.displayName,
                    oibId: this.extractOibId(policy),
                    type: 'updatePolicy',
                    policy: policy
                });
            });

            console.log(`Retrieved ${allPolicies.length} existing policies from tenant`);
            
            // Cache the results
            this.cachedPolicies = allPolicies;
            
            return allPolicies;
        } catch (error) {
            console.error('Failed to get all policies:', error);
            throw new Error('Failed to retrieve existing policies from tenant');
        }
    }

    // Check if a policy exists by name - with improved matching logic
    isPolicyDeployed(policyName, policyContent, existingPolicies) {
        if (!existingPolicies || !Array.isArray(existingPolicies)) {
            return false;
        }

        // Try multiple name sources for matching
        const namesToCheck = [
            policyName, // Name from filename
            policyContent?.displayName, // displayName from policy content
            policyContent?.name, // name from policy content (Settings Catalog)
        ].filter(name => name && typeof name === 'string');

        // Remove duplicates
        const uniqueNames = [...new Set(namesToCheck)];

        // Check against existing policies
        return existingPolicies.some(existing => {
            const existingName = existing.displayName;
            return uniqueNames.some(nameToCheck => 
                nameToCheck === existingName
            );
        });
    }

    // Check if a policy exists and return the policy object if found
    async checkPolicyExists(policyName, policyType) {
        try {
            console.log(`Checking if policy exists: ${policyName} (${policyType})`);
            
            // Get all existing policies if not already loaded
            if (!this.cachedPolicies) {
                this.cachedPolicies = await this.getAllPolicies();
            }

            // Filter policies by type and name
            const existingPolicy = this.cachedPolicies.find(policy => {
                const matchesName = policy.displayName === policyName;
                const matchesType = this.getPolicyTypeFromObject(policy) === policyType;
                return matchesName && matchesType;
            });

            if (existingPolicy) {
                console.log(`Policy exists: ${policyName}`);
                return existingPolicy;
            } else {
                console.log(`Policy not found: ${policyName}`);
                return null;
            }
        } catch (error) {
            console.error(`Error checking if policy exists: ${error.message}`);
            return null;
        }
    }

    // Helper method to determine policy type from policy object
    getPolicyTypeFromObject(policy) {
        if (policy.type) {
            // Convert our internal type naming to the expected format
            switch (policy.type) {
                case 'compliance':
                    return 'CompliancePolicies';
                case 'deviceConfiguration':
                    return 'DeviceConfigurations';
                case 'configurationPolicy':
                    return 'SettingsCatalog';
                case 'endpointSecurity':
                    return 'EndpointSecurity';
                case 'adminTemplate':
                    return 'AdminTemplates';
                case 'updatePolicy':
                    return 'UpdatePolicies';
                case 'appProtection':
                    return 'AppProtection';
                default:
                    return policy.type;
            }
        }
        return 'Unknown';
    }

    // Get all device compliance policies
    async getDeviceCompliancePolicies() {
        try {
            console.log('Fetching device compliance policies...');
            const response = await this.makePaginatedRequest(this.config.endpoints.deviceCompliancePolicies);
            const policies = response.value || [];
            console.log(`Successfully fetched ${policies.length} device compliance policies`);
            return policies;
        } catch (error) {
            console.error('Failed to get compliance policies:', error);
            throw new Error(`Failed to retrieve compliance policies from tenant: ${error.message}`);
        }
    }

    // Get all device configuration policies
    async getDeviceConfigurations() {
        try {
            console.log('Fetching device configuration policies...');
            const response = await this.makePaginatedRequest(this.config.endpoints.deviceConfigurations);
            const policies = response.value || [];
            console.log(`Successfully fetched ${policies.length} device configuration policies`);
            return policies;
        } catch (error) {
            console.error('Failed to get device configurations:', error);
            throw new Error(`Failed to retrieve device configurations from tenant: ${error.message}`);
        }
    }

    // Get all configuration policies (Settings Catalog)
    async getConfigurationPolicies() {
        try {
            console.log('Fetching configuration policies (Settings Catalog)...');
            const response = await this.makePaginatedRequest(this.config.endpoints.configurationPolicies);
            const policies = response.value || [];
            console.log(`Successfully fetched ${policies.length} configuration policies (Settings Catalog)`);
            return policies;
        } catch (error) {
            console.error('Failed to get configuration policies:', error);
            throw new Error(`Failed to retrieve configuration policies from tenant: ${error.message}`);
        }
    }

    // Get all endpoint security policies (intents)
    async getEndpointSecurityPolicies() {
        try {
            console.log('Fetching endpoint security policies...');
            const response = await this.makePaginatedRequest(this.config.endpoints.intents);
            const policies = response.value || [];
            console.log(`Successfully fetched ${policies.length} endpoint security policies`);
            return policies;
        } catch (error) {
            console.error('Failed to get endpoint security policies:', error);
            throw new Error(`Failed to retrieve endpoint security policies from tenant: ${error.message}`);
        }
    }

    // Get administrative templates (Group Policy)
    async getAdministrativeTemplates() {
        try {
            console.log('Fetching administrative templates...');
            const response = await this.makePaginatedRequest(this.config.endpoints.groupPolicyConfigurations);
            const policies = response.value || [];
            console.log(`Successfully fetched ${policies.length} administrative templates`);
            return policies;
        } catch (error) {
            console.error('Failed to get administrative templates:', error);
            throw new Error(`Failed to retrieve administrative templates from tenant: ${error.message}`);
        }
    }

    // Get all update policies (Windows Update for Business)
    async getUpdatePolicies() {
        try {
            console.log('Fetching update policies...');
            const [featureUpdatesResponse, qualityUpdatesResponse, driverUpdatesResponse] = await Promise.allSettled([
                this.makePaginatedRequest(this.config.endpoints.windowsFeatureUpdateProfiles),
                this.makePaginatedRequest(this.config.endpoints.windowsQualityUpdateProfiles),
                this.makePaginatedRequest(this.config.endpoints.windowsDriverUpdateProfiles)
            ]);

            const featureUpdates = featureUpdatesResponse.status === 'fulfilled' 
                ? featureUpdatesResponse.value.value || [] 
                : [];
            const qualityUpdates = qualityUpdatesResponse.status === 'fulfilled' 
                ? qualityUpdatesResponse.value.value || [] 
                : [];
            const driverUpdates = driverUpdatesResponse.status === 'fulfilled' 
                ? driverUpdatesResponse.value.value || [] 
                : [];

            const allUpdatePolicies = [...featureUpdates, ...qualityUpdates, ...driverUpdates];
            console.log(`Successfully fetched ${allUpdatePolicies.length} update policies (${featureUpdates.length} feature, ${qualityUpdates.length} quality, ${driverUpdates.length} driver)`);
            return allUpdatePolicies;
        } catch (error) {
            console.error('Failed to get update policies:', error);
            throw new Error(`Failed to retrieve update policies from tenant: ${error.message}`);
        }
    }

    // Transform policy content for import based on type
    transformPolicyForImport(policyContent, policyType, tenantId = null) {
        if (!policyContent || typeof policyContent !== 'object') {
            throw new Error('Invalid policy content provided');
        }

        // Get tenant ID if not provided
        const currentTenantId = tenantId || authService.getTenantId();

        // Create a deep copy to avoid modifying original
        const transformedPolicy = JSON.parse(JSON.stringify(policyContent));

        // Remove properties that shouldn't be included in import
        const propsToRemove = [
            'id', 'createdDateTime', 'lastModifiedDateTime', 
            'creationSource', 'version', 'supportsScopeTags',
            'roleScopeTagIds', 'isAssigned'
        ];

        propsToRemove.forEach(prop => {
            if (transformedPolicy.hasOwnProperty(prop)) {
                delete transformedPolicy[prop];
            }
        });

        // Handle different policy types
        switch (policyType) {
            case 'ConfigurationPolicies':
            case 'SettingsCatalog':
                return this.transformConfigurationPolicyForImport(transformedPolicy, currentTenantId);
            
            case 'CompliancePolicies':
                return this.transformCompliancePolicyForImport(transformedPolicy, currentTenantId);
            
            case 'DeviceConfiguration':
            case 'DeviceConfigurations':
                return this.transformDeviceConfigurationForImport(transformedPolicy, currentTenantId);
            
            case 'EndpointSecurity':
                return this.transformEndpointSecurityPolicyForImport(transformedPolicy, currentTenantId);
            
            case 'AdminTemplates':
                return this.transformAdminTemplateForImport(transformedPolicy, currentTenantId);
            
            case 'AppProtection':
            case 'AppProtectionPolicies':
                return this.transformAppProtectionForImport(transformedPolicy, currentTenantId);

            case 'UpdatePolicies':
                // Check if this is a legacy WUfB device configuration
                if (transformedPolicy['@odata.type'] && 
                    transformedPolicy['@odata.type'].includes('windowsUpdateForBusinessConfiguration')) {
                    // Use device configuration transformation for legacy WUfB policies
                    console.log('Using device configuration transformation for legacy WUfB policy');
                    return this.transformDeviceConfigurationForImport(transformedPolicy, currentTenantId);
                } else {
                    // Use update policy transformation for modern update profiles
                    return this.transformUpdatePolicyForImport(transformedPolicy, currentTenantId);
                }
            
            default:
                console.warn(`Unknown policy type: ${policyType}, applying generic transformation`);
                return transformedPolicy;
        }
    }

    // Transform Settings Catalog/Configuration Policies - FIXED to preserve main @odata.type
    transformConfigurationPolicyForImport(policy, tenantId = null) {
        let transformed = { ...policy };

        // Apply tenant variable replacement first
        if (tenantId) {
            transformed = this.replaceTenantVariables(transformed, tenantId);
        }

        // Remove @odata annotation properties but preserve the main @odata.type
        Object.keys(transformed).forEach(key => {
            if (key.includes('@odata') && key !== '@odata.type') {
                delete transformed[key];
            }
        });

        // Clean up settings if they exist
        if (transformed.settings && Array.isArray(transformed.settings)) {
            transformed.settings = transformed.settings.map(setting => {
                const cleanSetting = { ...setting };
                
                // Remove @odata annotation properties from settings but preserve main @odata.type
                Object.keys(cleanSetting).forEach(key => {
                    if (key.includes('@odata') && key !== '@odata.type') {
                        delete cleanSetting[key];
                    }
                });

                // Clean up settingInstance if it exists
                if (cleanSetting.settingInstance) {
                    Object.keys(cleanSetting.settingInstance).forEach(key => {
                        if (key.includes('@odata') && key !== '@odata.type') {
                            delete cleanSetting.settingInstance[key];
                        }
                    });
                }

                return cleanSetting;
            });
        }

        return transformed;
    }

    // Transform Compliance Policies
    transformCompliancePolicyForImport(policy, tenantId = null) {
        let transformed = { ...policy };

        // Apply tenant variable replacement first
        if (tenantId) {
            transformed = this.replaceTenantVariables(transformed, tenantId);
        }

        // Remove compliance-specific properties but keep scheduledActionsForRule
        const compliancePropsToRemove = [
            'deviceStatuses', 'userStatuses', 'deviceStatusOverview',
            'userStatusOverview', 'assignments'
        ];

        compliancePropsToRemove.forEach(prop => {
            if (transformed.hasOwnProperty(prop)) {
                delete transformed[prop];
            }
        });

        // Handle scheduledActionsForRule - required for compliance policies
        if (transformed.scheduledActionsForRule && Array.isArray(transformed.scheduledActionsForRule)) {
            // Clean up existing scheduledActionsForRule
            transformed.scheduledActionsForRule = transformed.scheduledActionsForRule.map(rule => {
                const cleanRule = { ...rule };
                
                // Remove @odata properties except @odata.type
                Object.keys(cleanRule).forEach(key => {
                    if (key.includes('@odata') && key !== '@odata.type') {
                        delete cleanRule[key];
                    }
                });

                // Clean up scheduledActionConfigurations
                if (cleanRule.scheduledActionConfigurations && Array.isArray(cleanRule.scheduledActionConfigurations)) {
                    cleanRule.scheduledActionConfigurations = cleanRule.scheduledActionConfigurations.map(config => {
                        const cleanConfig = { ...config };
                        
                        // Remove @odata properties except @odata.type
                        Object.keys(cleanConfig).forEach(key => {
                            if (key.includes('@odata') && key !== '@odata.type') {
                                delete cleanConfig[key];
                            }
                        });

                        // Remove id and other properties that shouldn't be in creation
                        delete cleanConfig.id;

                        return cleanConfig;
                    });
                }

                // Remove id and other properties that shouldn't be in creation
                delete cleanRule.id;

                return cleanRule;
            });
        } else {
            // Add default scheduledActionsForRule if missing - required by Microsoft Graph
            console.log('Adding default scheduledActionsForRule for compliance policy');
            transformed.scheduledActionsForRule = [
                {
                    "@odata.type": "#microsoft.graph.deviceComplianceScheduledActionForRule",
                    "ruleName": null,
                    "scheduledActionConfigurations": [
                        {
                            "@odata.type": "#microsoft.graph.deviceComplianceActionItem",
                            "gracePeriodHours": 12,
                            "actionType": "block",
                            "notificationTemplateId": "00000000-0000-0000-0000-000000000000",
                            "notificationMessageCCList": []
                        }
                    ]
                }
            ];
        }

        return transformed;
    }

    // Transform Device Configuration Policies
    transformDeviceConfigurationForImport(policy, tenantId = null) {
        let transformed = { ...policy };

        // Apply tenant variable replacement first
        if (tenantId) {
            transformed = this.replaceTenantVariables(transformed, tenantId);
        }

        // Remove device configuration specific properties
        const deviceConfigPropsToRemove = [
            'deviceStatuses', 'userStatuses', 'deviceStatusOverview',
            'userStatusOverview', 'assignments', 'deviceSettingStateSummaries'
        ];

        deviceConfigPropsToRemove.forEach(prop => {
            if (transformed.hasOwnProperty(prop)) {
                delete transformed[prop];
            }
        });

        return transformed;
    }

    // Transform Endpoint Security Policies
    transformEndpointSecurityPolicyForImport(policy, tenantId = null) {
        let transformed = { ...policy };

        // Apply tenant variable replacement first
        if (tenantId) {
            transformed = this.replaceTenantVariables(transformed, tenantId);
        }

        // Remove endpoint security specific properties
        const endpointSecurityPropsToRemove = [
            'templateId', 'intentSettings', 'assignments'
        ];

        endpointSecurityPropsToRemove.forEach(prop => {
            if (transformed.hasOwnProperty(prop)) {
                delete transformed[prop];
            }
        });

        return transformed;
    }

    // Transform App Protection (MAM) Policies. Apps and assignments are not part of the
    // create payload; apps are attached afterwards via the targetApps action.
    transformAppProtectionForImport(policy, tenantId = null) {
        let transformed = { ...policy };

        if (tenantId) {
            transformed = this.replaceTenantVariables(transformed, tenantId);
        }

        // Remove OData annotations (including "<prop>@odata.type" enum hints) but keep the main @odata.type
        Object.keys(transformed).forEach(key => {
            if (key.includes('@odata') && key !== '@odata.type') {
                delete transformed[key];
            }
        });

        ['apps', 'assignments', 'deploymentSummary', 'deployedAppCount', 'isAssigned'].forEach(prop => {
            delete transformed[prop];
        });

        return transformed;
    }

    // Get the Graph endpoint for an App Protection policy based on its @odata.type
    determineAppProtectionEndpoint(policy) {
        const odataType = (policy?.['@odata.type'] || '').toLowerCase();

        if (odataType.includes('androidmanagedappprotection')) {
            return this.config.endpoints.androidManagedAppProtections;
        }
        if (odataType.includes('iosmanagedappprotection')) {
            return this.config.endpoints.iosManagedAppProtections;
        }
        if (odataType.includes('windowsmanagedappprotection')) {
            return this.config.endpoints.windowsManagedAppProtections;
        }
        throw new Error(`Unsupported App Protection policy @odata.type: ${policy?.['@odata.type'] || 'missing'}`);
    }

    // Build the targetApps payload from the apps array stored in an exported policy
    buildTargetAppsPayload(apps) {
        const targetApps = (Array.isArray(apps) ? apps : [])
            .filter(app => app?.mobileAppIdentifier)
            .map(app => ({
                '@odata.type': '#microsoft.graph.managedMobileApp',
                mobileAppIdentifier: app.mobileAppIdentifier
            }));
        return { apps: targetApps };
    }

    // Transform Administrative Template Policies
    transformAdminTemplateForImport(policy, tenantId = null) {
        let transformed = { ...policy };

        // Apply tenant variable replacement first
        if (tenantId) {
            transformed = this.replaceTenantVariables(transformed, tenantId);
        }

        // Remove admin template specific properties
        const adminTemplatePropsToRemove = [
            'assignments', 'groupAssignments'
        ];

        adminTemplatePropsToRemove.forEach(prop => {
            if (transformed.hasOwnProperty(prop)) {
                delete transformed[prop];
            }
        });

        return transformed;
    }

    // Transform Update Policies (WUfB, Driver Updates, etc.)
    transformUpdatePolicyForImport(policy, tenantId = null) {
        let transformed = { ...policy };

        // Apply tenant variable replacement first
        if (tenantId) {
            transformed = this.replaceTenantVariables(transformed, tenantId);
        }

        // Remove update policy specific properties that shouldn't be included in creation
        const updatePropsToRemove = [
            'assignments', 'groupAssignments', 'id', 'createdDateTime', 'lastModifiedDateTime'
        ];

        updatePropsToRemove.forEach(prop => {
            if (transformed.hasOwnProperty(prop)) {
                delete transformed[prop];
            }
        });

        // Handle legacy windowsUpdateForBusinessConfiguration policies
        if (transformed['@odata.type'] && transformed['@odata.type'].includes('windowsUpdateForBusinessConfiguration')) {
            // These are device configuration policies, not the newer update profile format
            // Keep the @odata.type as-is for device configuration endpoint
            console.log('Detected legacy WUfB device configuration policy');
        }

        return transformed;
    }

    // Deploy a single policy
    async deployPolicy(policyData) {
        try {
            const { policy, type, fileName } = policyData;
            let endpoint = '';

            // Detect effective type: modern ES policies have been migrated to Settings Catalog
            // format and must use the configurationPolicies endpoint, not the legacy intents endpoint.
            let effectiveType = type;
            if (type === 'EndpointSecurity' &&
                policy?.['@odata.type']?.includes('deviceManagementConfigurationPolicy')) {
                effectiveType = 'SettingsCatalog';
                console.log(`ES policy is Settings Catalog format, routing to configurationPolicies: ${fileName}`);
            }

            let transformedPolicy = this.transformPolicyForImport(policy, effectiveType);

            // Determine the correct endpoint based on policy type
            switch (effectiveType) {
                case 'CompliancePolicies':
                    endpoint = this.config.endpoints.deviceCompliancePolicies;
                    break;
                case 'SettingsCatalog':
                case 'ConfigurationPolicies':
                    endpoint = this.config.endpoints.configurationPolicies;
                    break;
                case 'DeviceConfiguration':
                case 'DeviceConfigurations':
                    endpoint = this.config.endpoints.deviceConfigurations;
                    break;
                case 'EndpointSecurity':
                    endpoint = this.config.endpoints.intents;
                    break;
                case 'AdminTemplates':
                    endpoint = this.config.endpoints.groupPolicyConfigurations;
                    break;
                case 'AppProtection':
                case 'AppProtectionPolicies':
                    endpoint = this.determineAppProtectionEndpoint(transformedPolicy);
                    break;
                case 'UpdatePolicies':
                    // Determine the correct endpoint based on policy content and @odata.type
                    endpoint = this.determineUpdatePolicyEndpoint(transformedPolicy);
                    break;
                default:
                    throw new Error(`Unsupported policy type: ${effectiveType}`);
            }

            console.log(`Deploying ${effectiveType} policy: ${fileName}`);
            const result = await this.makeRequest(endpoint, 'POST', transformedPolicy);

            if (effectiveType === 'AppProtection' || effectiveType === 'AppProtectionPolicies') {
                // "allMicrosoftApps" policies list their expanded apps in the export, but
                // targeting them explicitly would switch the policy to a custom app list.
                const targetApps = policy.appGroupType === 'allMicrosoftApps'
                    ? { apps: [] }
                    : this.buildTargetAppsPayload(policy.apps);
                if (targetApps.apps.length > 0) {
                    try {
                        await this.makeRequest(`${endpoint}/${result.id}/targetApps`, 'POST', targetApps);
                    } catch (appsError) {
                        throw new Error(`Policy created (id ${result.id}) but targeting its apps failed: ${appsError.message}`);
                    }
                }
            }
            
            return {
                success: true,
                policyId: result.id,
                // Collection the policy was created in — used to assign it afterwards.
                endpoint,
                policyType: effectiveType,
                policyName: transformedPolicy.displayName || transformedPolicy.name || fileName,
                fileName: fileName,
                message: `Successfully deployed: ${transformedPolicy.displayName || transformedPolicy.name || fileName}`,
                result: result
            };
        } catch (error) {
            console.error(`Failed to deploy policy ${policyData.fileName}:`, error);
            return {
                success: false,
                policyName: policyData.fileName || policyData.name || 'Unknown Policy',
                fileName: policyData.fileName,
                message: `Failed to deploy ${policyData.fileName}: ${error.message}`,
                error: error
            };
        }
    }

    // Deploy policies with progress tracking
    async deployPoliciesBatchWithProgress(policiesData, progressCallback) {
        const results = [];
        const total = policiesData.length;
        console.log(`Starting deployment of ${total} policies...`);
        
        for (let i = 0; i < policiesData.length; i++) {
            const policyData = policiesData[i];
            const policyName = policyData.fileName || policyData.displayName || 'Unknown Policy';
            
            console.log(`Deploying policy ${i + 1}/${total}: ${policyName}`);
            
            // Update progress
            if (progressCallback) {
                progressCallback(i, total, policyName);
            }
            
            // Deploy policy - this function returns a result object, doesn't throw
            const result = await this.deployPolicy(policyData);
            console.log(`Policy ${policyName} deployment result:`, result);
            
            // Add the result with index for UI matching
            results.push({
                ...result,
                index: i, // Add index for DeploymentProgress component matching
                message: result.message || (result.success ? 
                    `Successfully deployed: ${result.policyName}` : 
                    `Failed to deploy: ${result.policyName || policyName}`)
            });
        }
        
        // Final progress update
        if (progressCallback) {
            progressCallback(total, total, 'Deployment complete');
        }
        
        console.log(`Deployment batch completed. Total results: ${results.length}`);
        return results;
    }

    // Deploy multiple policies using batch requests
    async deployPoliciesBatch(policies) {
        const results = [];
        const batchSize = this.config.batch.maxBatchSize;

        // Process policies in batches
        for (let i = 0; i < policies.length; i += batchSize) {
            const batch = policies.slice(i, i + batchSize);
            const batchResults = await Promise.allSettled(
                batch.map(policy => this.deployPolicy(policy))
            );

            // Process batch results
            batchResults.forEach((result, index) => {
                const originalIndex = i + index;
                if (result.status === 'fulfilled') {
                    results.push({
                        index: originalIndex,
                        ...result.value
                    });
                } else {
                    results.push({
                        index: originalIndex,
                        success: false,
                        message: `Batch deployment failed: ${result.reason?.message || 'Unknown error'}`,
                        error: result.reason
                    });
                }
            });

            // Add delay between batches to avoid rate limiting
            if (i + batchSize < policies.length) {
                await this.sleep(1000);
            }
        }

        return results;
    }

    // Helper method to determine the correct endpoint for update policies
    determineUpdatePolicyEndpoint(policy) {
        const odataType = policy['@odata.type'];
        
        if (odataType) {
            // Check for driver update profile (modern format)
            if (odataType.includes('windowsDriverUpdateProfile')) {
                console.log('Detected driver update profile - using driver endpoint');
                return this.config.endpoints.windowsDriverUpdateProfiles;
            }
            
            // Check for feature update profile (modern format)
            if (odataType.includes('windowsFeatureUpdateProfile')) {
                console.log('Detected feature update profile - using feature endpoint');
                return this.config.endpoints.windowsFeatureUpdateProfiles;
            }
            
            // Check for quality update profile (modern format)
            if (odataType.includes('windowsQualityUpdateProfile')) {
                console.log('Detected quality update profile - using quality endpoint');
                return this.config.endpoints.windowsQualityUpdateProfiles;
            }
            
            // Legacy WUfB configuration (these are device configurations, not update profiles)
            if (odataType.includes('windowsUpdateForBusinessConfiguration')) {
                console.log('Detected legacy WUfB device configuration - using device configuration endpoint');
                return this.config.endpoints.deviceConfigurations;
            }
        }
        
        // Fallback logic based on display name if @odata.type is not clear
        const displayName = (policy.displayName || policy.name || '').toLowerCase();
        
        if (displayName.includes('driver')) {
            console.log('Detected driver policy by name - using driver endpoint');
            return this.config.endpoints.windowsDriverUpdateProfiles;
        }
        
        if (displayName.includes('feature') || displayName.includes('version')) {
            console.log('Detected feature policy by name - using feature endpoint');
            return this.config.endpoints.windowsFeatureUpdateProfiles;
        }
        
        // Default to device configurations for legacy WUfB policies
        console.log('Defaulting to device configuration endpoint for update policy');
        return this.config.endpoints.deviceConfigurations;
    }

    // ── Validation ────────────────────────────────────────────────────────────

    /**
     * Fetch all settings for a Settings Catalog (configurationPolicy) by ID.
     * Returns the raw settings array from Graph, handling pagination.
     * Also used for Endpoint Security policies migrated to Settings Catalog
     * format (they live on the same configurationPolicies endpoint).
     */
    async getConfigurationPolicySettings(policyId) {
        try {
            const endpoint = `/deviceManagement/configurationPolicies/${policyId}/settings?$top=1000`;
            logger.debug(`Fetching settings for configurationPolicy: ${policyId}`);

            let allSettings = [];
            let nextLink = `${this.config.baseUrl}${endpoint}`;

            while (nextLink) {
                const options = {
                    method: 'GET',
                    headers: await this.getAuthHeaders()
                };

                const response = await this.makeRequestWithRetry(nextLink, options);
                if (!response.ok) {
                    const errorText = await response.text();
                    throw new Error(`Graph API error fetching settings: ${response.status} - ${errorText}`);
                }

                const result = await response.json();
                const items  = result.value ?? [];
                allSettings  = allSettings.concat(items);
                nextLink     = result['@odata.nextLink'] ?? null;
            }

            logger.debug(`Fetched ${allSettings.length} settings for policy ${policyId}`);
            return allSettings;
        } catch (error) {
            logger.error(`Failed to fetch settings for policy ${policyId}:`, error);
            throw error;
        }
    }

    /**
     * Fetch the full compliance policy object by ID (the cached tenant list
     * only holds id/displayName/description — compliance settings are flat
     * top-level properties on the full object, not a settings sub-resource).
     */
    async getCompliancePolicyDetail(policyId) {
        try {
            const endpoint = `/deviceManagement/deviceCompliancePolicies/${policyId}`;
            logger.debug(`Fetching compliance policy detail: ${policyId}`);

            const options = {
                method: 'GET',
                headers: await this.getAuthHeaders()
            };

            const response = await this.makeRequestWithRetry(`${this.config.baseUrl}${endpoint}`, options);
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Graph API error fetching compliance policy: ${response.status} - ${errorText}`);
            }

            return await response.json();
        } catch (error) {
            logger.error(`Failed to fetch compliance policy detail ${policyId}:`, error);
            throw error;
        }
    }

    /**
     * Validate a matched policy's settings against the OIB baseline.
     *
     * @param {object} matchedPolicy  - entry from ComparisonDashboard with
     *                                  .existingPolicy (tenant) and .downloadUrl
     *                                  or .content (OIB)
     * @param {string} tenantId       - current tenant ID for %OrganizationId% substitution
     * @returns {object}  validation result (see settings-comparator.js)
     */
    async validatePolicySettings(matchedPolicy, tenantId) {
        const { compareSettings, compareCompliancePolicy } = await import('./settings-comparator.js');

        const policyName = matchedPolicy.name ?? matchedPolicy.existingPolicy?.displayName ?? 'Unknown';
        const isCompliance = matchedPolicy.policyType === 'CompliancePolicies';

        try {
            // 1. Load OIB content (use cache if already downloaded)
            const { githubAPI } = await import('./github-api.js');
            const oibContent = await githubAPI.getPolicyContent(matchedPolicy);

            // Endpoint Security policies were historically deployed via the legacy
            // /deviceManagement/intents API (a completely different schema — no
            // `settings` array). Modern OIB ES policies have been migrated to
            // Settings Catalog format (@odata.type: deviceManagementConfigurationPolicy),
            // which compares the same way as plain Settings Catalog. Guard against
            // stray legacy-schema exports so they don't get silently mis-compared.
            if (matchedPolicy.policyType === 'EndpointSecurity' &&
                !oibContent?.['@odata.type']?.includes('deviceManagementConfigurationPolicy')) {
                throw new Error('Legacy Endpoint Security (intent) format is not supported for validation');
            }

            const tenantPolicyId = matchedPolicy.existingPolicy?.id;
            if (!tenantPolicyId) {
                throw new Error('No tenant policy ID available for validation');
            }

            let result;
            if (isCompliance) {
                // Compliance policies are a flat property bag — apply tenant
                // variable substitution to the whole object, fetch the full
                // tenant object (list cache only has id/displayName/description),
                // and diff key-by-key.
                const substitutedOib = tenantId
                    ? this.replaceTenantVariables(oibContent, tenantId)
                    : oibContent;
                const tenantPolicy = await this.getCompliancePolicyDetail(tenantPolicyId);
                result = compareCompliancePolicy(substitutedOib, tenantPolicy);
            } else {
                // Settings Catalog (and Settings-Catalog-format Endpoint Security).
                let oibSettings = oibContent?.settings ?? [];
                if (tenantId && oibSettings.length > 0) {
                    const substituted = this.replaceTenantVariables({ settings: oibSettings }, tenantId);
                    oibSettings = substituted.settings;
                }
                const tenantSettings = await this.getConfigurationPolicySettings(tenantPolicyId);
                result = compareSettings(oibSettings, tenantSettings);
            }

            return {
                ...result,
                policyName,
                status: result.compliant ? 'compliant' : 'drifted',
            };
        } catch (error) {
            logger.error(`Validation failed for ${policyName}:`, error);
            return {
                policyName,
                status: 'error',
                error:       error.message,
                totalOib:    0,
                totalTenant: 0,
                matched:     0,
                mismatches:  [],
                oibOnly:     [],
                tenantOnly:  [],
                compliant:   false,
            };
        }
    }

    // ── Device Validation ─────────────────────────────────────────────────────
    // These calls need DeviceManagementManagedDevices.Read.All on top of the
    // base scope. Tokens are acquired silently only; a ConsentRequiredError
    // bubbles up so the UI can offer a "grant access" step.

    async makeDeviceRequest(endpoint, method = 'GET', body = null) {
        return this.makeScopedRequest(CONFIG.msal.deviceScopes, endpoint, method, body);
    }

    // Graph request with a silently-acquired token for specific scopes.
    // Throws ConsentRequiredError instead of redirecting.
    async makeScopedRequest(scopes, endpoint, method = 'GET', body = null) {
        const token = await authService.getAccessTokenForScopes(scopes);
        const url = endpoint.startsWith('https://') ? endpoint : `${this.config.baseUrl}${endpoint}`;
        const options = {
            method,
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        };
        if (body) options.body = JSON.stringify(body);

        const response = await this.makeRequestWithRetry(url, options);
        if (!response.ok) {
            const errorText = await response.text();
            const error = new Error(`Graph API error: ${response.status} ${response.statusText} - ${errorText}`);
            error.status = response.status;
            throw error;
        }
        if (response.status === 204) return null;

        // Report actions can come back as application/octet-stream, so parse
        // the text ourselves rather than relying on response.json().
        const text = await response.text();
        return text ? JSON.parse(text) : null;
    }

    static escapeODataString(value) {
        return String(value).replace(/'/g, "''");
    }

    /**
     * Search Intune managed devices by device name or primary user UPN prefix.
     * An empty query returns the first page of devices.
     */
    async searchManagedDevices(query = '') {
        const select = '$select=id,deviceName,operatingSystem,osVersion,model,manufacturer,serialNumber,' +
            'userPrincipalName,userDisplayName,lastSyncDateTime,complianceState,managementAgent,azureADDeviceId';
        const base = '/deviceManagement/managedDevices';
        const trimmed = query.trim();

        if (!trimmed) {
            const result = await this.makeDeviceRequest(`${base}?${select}&$top=50`);
            return result?.value ?? [];
        }

        const escaped = encodeURIComponent(GraphAPI.escapeODataString(trimmed));
        const filters = [
            `startswith(deviceName,'${escaped}')`,
            `startswith(userPrincipalName,'${escaped}')`,
            `serialNumber eq '${escaped}'`
        ];

        const results = await Promise.allSettled(
            filters.map(filter => this.makeDeviceRequest(`${base}?$filter=${filter}&${select}&$top=50`))
        );

        // Surface consent problems rather than silently returning nothing.
        const consentFailure = results.find(r => r.status === 'rejected' && r.reason?.name === 'ConsentRequiredError');
        if (consentFailure) throw consentFailure.reason;
        if (results.every(r => r.status === 'rejected')) throw results[0].reason;

        const byId = new Map();
        results.forEach(r => {
            if (r.status !== 'fulfilled') return;
            (r.value?.value ?? []).forEach(device => byId.set(device.id, device));
        });
        return [...byId.values()].sort((a, b) => (a.deviceName || '').localeCompare(b.deviceName || ''));
    }

    async getManagedDevice(deviceId) {
        return this.makeDeviceRequest(`/deviceManagement/managedDevices/${encodeURIComponent(deviceId)}`);
    }

    /**
     * Per-device configuration status — the same report the Intune portal
     * shows under Devices > {device} > Device configuration. Covers Settings
     * Catalog, Endpoint Security (intents), Admin Templates and legacy device
     * configuration profiles. Returns one row per policy per user/device context.
     */
    async getDeviceConfigurationReport(deviceId) {
        if (!/^[0-9a-f-]{36}$/i.test(deviceId)) {
            throw new Error(`Invalid Intune device ID: ${deviceId}`);
        }

        const baseTypes = [
            'Microsoft.Management.Services.Api.DeviceConfiguration',
            'DeviceManagementConfigurationPolicy',
            'DeviceConfigurationAdmxPolicy',
            'Microsoft.Management.Services.Api.DeviceManagementIntent'
        ];
        const filter = `(${baseTypes.map(t => `(PolicyBaseTypeName eq '${t}')`).join(' or ')}) and (IntuneDeviceId eq '${deviceId}')`;
        const select = [
            'IntuneDeviceId', 'PolicyBaseTypeName', 'PolicyId', 'PolicyStatus', 'UPN', 'UserId',
            'PspdpuLastModifiedTimeUtc', 'PolicyName', 'UnifiedPolicyType'
        ];

        const { rows } = await this.fetchReportRows(
            CONFIG.msal.deviceScopes,
            'getConfigurationPoliciesReportForDevice',
            { select, filter, orderBy: ['PolicyName'] }
        );
        return rows;
    }

    /**
     * Page through an Intune report action (POST /deviceManagement/reports/{action})
     * and turn its Schema/Values arrays into row objects.
     * Returns { rows, total, truncated }.
     */
    async fetchReportRows(scopes, action, body, { maxRows = 5000, pageSize = 50, paged = true } = {}) {
        const rows = [];
        let schema = null;
        let total = null;

        for (let skip = 0; skip < maxRows; skip += pageSize) {
            const page = await this.makeScopedRequest(
                scopes,
                `/deviceManagement/reports/${action}`,
                'POST',
                paged ? { ...body, skip, top: pageSize } : body
            );

            schema = schema || page?.Schema || page?.schema;
            const values = page?.Values || page?.values || [];
            if (total === null && page?.TotalRowCount != null) total = Number(page.TotalRowCount);

            const columns = (schema || []).map(col => col.Column ?? col.column);
            values.forEach(valueRow => {
                const row = {};
                columns.forEach((col, i) => { row[col] = valueRow[i]; });
                rows.push(row);
            });

            if (!paged || values.length < pageSize || (total !== null && rows.length >= total)) break;
        }

        return { rows, total: total ?? rows.length, truncated: total !== null && rows.length < total };
    }

    /**
     * Per-device compliance policy states. Compliance policies are not part of
     * the configuration report, so they are read from the managed device.
     */
    async getDeviceCompliancePolicyStates(deviceId) {
        // Preferred: the compliance report the portal uses. In testing,
        // managedDevices/{id}/deviceCompliancePolicyStates lagged behind it
        // and omitted policies the device had already evaluated.
        try {
            const { rows } = await this.fetchReportRows(
                CONFIG.msal.deviceScopes,
                'getDevicePoliciesComplianceReport',
                { filter: `(DeviceId eq '${GraphAPI.escapeODataString(deviceId)}')`, select: [], orderBy: [] }
            );
            // Same shape as deviceCompliancePolicyState so callers don't care.
            return rows.map(row => ({
                id: row.PolicyId,
                displayName: row.PolicyName,
                state: String(row.PolicyStatus_loc ?? row.PolicyStatus ?? '').replace(/\s+/g, ''),
                userPrincipalName: row.UPN || ''
            }));
        } catch (error) {
            if (error?.name === 'ConsentRequiredError') throw error;
            logger.warn('Compliance report failed, falling back to deviceCompliancePolicyStates:', error);
        }

        const result = await this.makeDeviceRequest(
            `/deviceManagement/managedDevices/${encodeURIComponent(deviceId)}/deviceCompliancePolicyStates`
        );
        return result?.value ?? [];
    }

    /**
     * Assignments for a normalised tenant policy (see getEndpointGroupsForOS).
     * Used to explain why a device has no report for a policy.
     */
    async getPolicyAssignments(tenantPolicy) {
        const endpoints = {
            configurationPolicy: '/deviceManagement/configurationPolicies',
            compliance: '/deviceManagement/deviceCompliancePolicies',
            deviceConfiguration: '/deviceManagement/deviceConfigurations',
            adminTemplate: '/deviceManagement/groupPolicyConfigurations',
            endpointSecurity: '/deviceManagement/intents'
        };
        const base = endpoints[tenantPolicy?.type];
        if (!base) throw new Error(`Assignments not supported for policy type ${tenantPolicy?.type}`);

        // Base scope is enough for assignments.
        const result = await this.makeScopedRequest(
            CONFIG.msal.scopes,
            `${base}/${encodeURIComponent(tenantPolicy.id)}/assignments`
        );
        return result?.value ?? [];
    }

    /**
     * Fetch everything Device Validation needs for one device. Each source is
     * fetched independently so a failure in one (e.g. a retired API) still
     * lets the others render. Consent errors are re-thrown.
     */
    async getDevicePolicyStatus(deviceId) {
        const [device, configuration, compliance] = await Promise.allSettled([
            this.getManagedDevice(deviceId),
            this.getDeviceConfigurationReport(deviceId),
            this.getDeviceCompliancePolicyStates(deviceId)
        ]);

        const consentFailure = [device, configuration, compliance]
            .find(r => r.status === 'rejected' && r.reason?.name === 'ConsentRequiredError');
        if (consentFailure) throw consentFailure.reason;

        return {
            device: device.status === 'fulfilled' ? device.value : null,
            configurationRows: configuration.status === 'fulfilled' ? configuration.value : [],
            complianceStates: compliance.status === 'fulfilled' ? compliance.value : [],
            errors: {
                device: device.status === 'rejected' ? device.reason.message : null,
                configuration: configuration.status === 'rejected' ? configuration.reason.message : null,
                compliance: compliance.status === 'rejected' ? compliance.reason.message : null
            },
            fetchedAt: new Date().toISOString()
        };
    }

    // ── Policy assignments & device status (Existing Deployment) ────────────
    // Uses the base DeviceManagementConfiguration scope, except group names
    // which need GroupMember.Read.All (requested incrementally).

    // Follow @odata.nextLink for a GET collection with a scoped token.
    async getAllPagesScoped(scopes, endpoint) {
        const items = [];
        let next = endpoint;
        while (next) {
            const page = await this.makeScopedRequest(scopes, next);
            items.push(...(page?.value ?? []));
            next = page?.['@odata.nextLink'] ?? null;
        }
        return items;
    }

    /**
     * Assignments for every tenant policy of the given normalised types,
     * fetched with $expand=assignments on the list endpoints (a handful of
     * calls instead of one per policy). Intents don't support $expand, so
     * those are fetched per policy.
     *
     * @param {object[]} tenantPolicies  normalised tenant policies
     * @returns {Map<string, object[]|{error:string}>}  policy id → assignments
     */
    async getAssignmentsForPolicies(tenantPolicies) {
        const scopes = CONFIG.msal.scopes;
        const types = new Set(tenantPolicies.map(p => p.type));
        const sources = [
            { type: 'configurationPolicy', urls: ['/deviceManagement/configurationPolicies?$select=id&$expand=assignments&$top=100'] },
            { type: 'compliance', urls: ['/deviceManagement/deviceCompliancePolicies?$select=id&$expand=assignments'] },
            { type: 'deviceConfiguration', urls: ['/deviceManagement/deviceConfigurations?$select=id&$expand=assignments'] },
            { type: 'adminTemplate', urls: ['/deviceManagement/groupPolicyConfigurations?$select=id&$expand=assignments'] },
            { type: 'updatePolicy', urls: [
                '/deviceManagement/windowsFeatureUpdateProfiles?$select=id&$expand=assignments',
                '/deviceManagement/windowsQualityUpdateProfiles?$select=id&$expand=assignments',
                '/deviceManagement/windowsDriverUpdateProfiles?$select=id&$expand=assignments'
            ] },
            { type: 'appProtection', urls: [
                '/deviceAppManagement/iosManagedAppProtections?$select=id&$expand=assignments',
                '/deviceAppManagement/androidManagedAppProtections?$select=id&$expand=assignments'
            ] }
        ].filter(source => types.has(source.type));

        const byId = new Map();
        const failedTypes = new Map();

        await Promise.all(sources.flatMap(source => source.urls.map(async url => {
            let items;
            try {
                items = await this.getAllPagesScoped(scopes, url);
            } catch (error) {
                if (error?.name === 'ConsentRequiredError') throw error;
                // Some list endpoints reject $select together with $expand.
                try {
                    items = await this.getAllPagesScoped(scopes, url.replace('$select=id&', ''));
                } catch (retryError) {
                    if (retryError?.name === 'ConsentRequiredError') throw retryError;
                    failedTypes.set(source.type, retryError.message);
                    return;
                }
            }
            items.forEach(item => byId.set(item.id.toLowerCase(), item.assignments ?? []));
        })));

        // Legacy Endpoint Security intents: per-policy fetch.
        const intents = tenantPolicies.filter(p => p.type === 'endpointSecurity');
        for (const intent of intents) {
            try {
                byId.set(intent.id.toLowerCase(), await this.getPolicyAssignments(intent));
            } catch (error) {
                if (error?.name === 'ConsentRequiredError') throw error;
                byId.set(intent.id.toLowerCase(), { error: error.message });
            }
        }

        const result = new Map();
        tenantPolicies.forEach(policy => {
            const key = policy.id.toLowerCase();
            if (byId.has(key)) result.set(policy.id, byId.get(key));
            else if (failedTypes.has(policy.type)) result.set(policy.id, { error: failedTypes.get(policy.type) });
            else result.set(policy.id, { error: 'Policy not found when reading assignments' });
        });
        return result;
    }

    // Assignment filters, Map<lowercased id, filter> with displayName,
    // platform, rule, description and management type.
    async getAssignmentFilters() {
        const items = await this.getAllPagesScoped(
            CONFIG.msal.scopes,
            '/deviceManagement/assignmentFilters?$select=id,displayName,platform,rule,description,assignmentFilterManagementType'
        );
        return new Map(items.map(f => [f.id.toLowerCase(), f]));
    }

    async createAssignmentFilter({ displayName, description, platform, rule }) {
        return this.makeScopedRequest(CONFIG.msal.scopes, '/deviceManagement/assignmentFilters', 'POST', {
            '@odata.type': '#microsoft.graph.deviceAndAppManagementAssignmentFilter',
            displayName,
            description: String(description ?? '').slice(0, 1024),
            platform,
            rule,
            assignmentFilterManagementType: 'devices',
            roleScopeTags: ['0']
        });
    }

    async updateAssignmentFilterRule(filterId, rule) {
        await this.makeScopedRequest(
            CONFIG.msal.scopes,
            `/deviceManagement/assignmentFilters/${encodeURIComponent(filterId)}`,
            'PATCH',
            { '@odata.type': '#microsoft.graph.deviceAndAppManagementAssignmentFilter', rule }
        );
    }

    // Intune refuses to delete a filter it still sees as assigned (error 5001)
    // for a little while after its last assignment was removed, so retry.
    async deleteAssignmentFilter(filterId, { retries = 6, retryDelay = 5000 } = {}) {
        for (let attempt = 1; ; attempt++) {
            try {
                await this.makeScopedRequest(
                    CONFIG.msal.scopes,
                    `/deviceManagement/assignmentFilters/${encodeURIComponent(filterId)}`,
                    'DELETE'
                );
                return;
            } catch (error) {
                const stillAssigned = error.status === 400 && /5001|associated with existing assignments/i.test(error.message);
                if (!stillAssigned || attempt >= retries) throw error;
                await this.sleep(retryDelay);
            }
        }
    }

    /**
     * How Intune evaluated assignment filters for a device and policy at its
     * last check-in (beta). Returns [{ filterId, name, mode, result }] where
     * result is 'match' | 'notMatch' | other values Intune reports.
     */
    async getFilterEvaluation({ managedDeviceId, payloadId, userId = null, filterIds = [] }) {
        const result = await this.makeScopedRequest(
            CONFIG.msal.scopes,
            '/deviceManagement/getAssignmentFiltersStatusDetails',
            'POST',
            {
                managedDeviceId,
                payloadId,
                ...(userId ? { userId } : {}),
                assignmentFilterIds: filterIds,
                top: 50,
                skip: 0
            }
        );
        // The property is spelled "evalutionSummaries" in Graph.
        const details = result?.value ?? result ?? {};
        const summaries = details.evalutionSummaries ?? details.evaluationSummaries ?? [];
        return summaries.map(s => ({
            filterId: s.assignmentFilterId,
            name: s.assignmentFilterDisplayName,
            mode: String(s.assignmentFilterType ?? '').toLowerCase(),
            result: String(s.evaluationResult ?? '').toLowerCase(),
            evaluatedAt: s.evaluationDateTime ?? null
        }));
    }

    /**
     * Resolve Entra group IDs to display names via $batch. Requires
     * GroupMember.Read.All — throws ConsentRequiredError when not granted.
     * Deleted groups map to null.
     */
    async getGroupNames(groupIds, scopes = CONFIG.msal.groupScopes) {
        const ids = [...new Set(groupIds.map(id => id.toLowerCase()))];
        const names = new Map();
        for (let i = 0; i < ids.length; i += 20) {
            const chunk = ids.slice(i, i + 20);
            const response = await this.makeScopedRequest(
                scopes,
                'https://graph.microsoft.com/v1.0/$batch',
                'POST',
                { requests: chunk.map(id => ({ id, method: 'GET', url: `/groups/${id}?$select=id,displayName` })) }
            );
            for (const item of response?.responses ?? []) {
                if (item.status === 200) {
                    names.set(item.id, item.body?.displayName ?? null);
                } else if (item.status === 404) {
                    names.set(item.id, null);
                } else if (item.status === 403) {
                    const error = new Error('Insufficient privileges to read group names');
                    error.status = 403;
                    throw error;
                }
            }
        }
        return names;
    }

    // Report base type per normalised tenant policy type.
    static reportBaseTypeFor(type) {
        return {
            configurationPolicy: 'DeviceManagementConfigurationPolicy',
            endpointSecurity: 'Microsoft.Management.Services.Api.DeviceManagementIntent',
            deviceConfiguration: 'Microsoft.Management.Services.Api.DeviceConfiguration',
            adminTemplate: 'DeviceConfigurationAdmxPolicy'
        }[type] ?? null;
    }

    /**
     * Device status counts per policy — the numbers shown on a policy's
     * Overview page in Intune. Configuration-type policies use the summary
     * report (batched by base type); compliance uses deviceStatusOverview.
     *
     * @returns {Map<string, object>}  policy id → raw summary ({ source, row|overview } or { error })
     */
    async getPolicyStatusSummaries(tenantPolicies) {
        const scopes = CONFIG.msal.scopes;
        const result = new Map();

        // Configuration-type policies: one summary report call per base type per 20 IDs.
        const byBaseType = new Map();
        tenantPolicies.forEach(policy => {
            const baseType = GraphAPI.reportBaseTypeFor(policy.type);
            if (!baseType) return;
            if (!byBaseType.has(baseType)) byBaseType.set(baseType, []);
            byBaseType.get(baseType).push(policy);
        });

        for (const [baseType, policies] of byBaseType) {
            for (let i = 0; i < policies.length; i += 20) {
                const chunk = policies.slice(i, i + 20);
                const ids = chunk.map(p => `(PolicyId eq '${GraphAPI.escapeODataString(p.id)}')`).join(' or ');
                try {
                    const { rows } = await this.fetchReportRows(
                        scopes,
                        'getConfigurationPolicyDeviceSummaryReport',
                        { filter: `((PolicyBaseTypeName eq '${baseType}') and (${ids}))`, select: [] },
                        { paged: false }
                    );
                    const rowById = new Map(rows.map(row => [(row.PolicyId || '').toLowerCase(), row]));
                    chunk.forEach(policy => {
                        const row = rowById.get(policy.id.toLowerCase());
                        // No row = no device has reported on the policy yet.
                        result.set(policy.id, { source: 'configurationReport', row: row ?? null });
                    });
                } catch (error) {
                    if (error?.name === 'ConsentRequiredError') throw error;
                    chunk.forEach(policy => result.set(policy.id, { error: error.message }));
                }
            }
        }

        // Compliance policies: per-policy overview, small concurrency.
        const compliance = tenantPolicies.filter(p => p.type === 'compliance');
        const queue = [...compliance];
        const worker = async () => {
            while (queue.length) {
                const policy = queue.shift();
                try {
                    const overview = await this.makeScopedRequest(
                        scopes,
                        `/deviceManagement/deviceCompliancePolicies/${encodeURIComponent(policy.id)}/deviceStatusOverview`
                    );
                    result.set(policy.id, { source: 'complianceOverview', overview });
                } catch (error) {
                    if (error?.name === 'ConsentRequiredError') throw error;
                    result.set(policy.id, { error: error.message });
                }
            }
        };
        await Promise.all(Array.from({ length: Math.min(4, queue.length) }, worker));

        return result;
    }

    /**
     * Per-device status rows for one policy (the policy's "Device and user
     * check-in status" view). Returns { rows, total, truncated } with raw
     * report rows; normalise with policy-status.js.
     */
    async getPolicyDeviceStatuses(tenantPolicy, maxRows = 500) {
        const scopes = CONFIG.msal.scopes;
        const id = GraphAPI.escapeODataString(tenantPolicy.id);

        if (tenantPolicy.type === 'compliance') {
            const result = await this.fetchReportRows(
                scopes,
                'getDeviceStatusByCompliacePolicyReport', // sic — Graph's spelling
                { filter: `(PolicyId eq '${id}')`, select: [], orderBy: [] },
                { maxRows }
            );
            return { ...result, source: 'complianceReport' };
        }

        const baseType = GraphAPI.reportBaseTypeFor(tenantPolicy.type);
        if (!baseType) throw new Error(`No per-device report for policy type ${tenantPolicy.type}`);

        const result = await this.fetchReportRows(
            scopes,
            'getConfigurationPolicyDevicesReport',
            {
                filter: `((PolicyBaseTypeName eq '${baseType}') and (PolicyId eq '${id}'))`,
                select: ['DeviceName', 'UPN', 'PolicyStatus', 'IntuneDeviceId', 'PspdpuLastModifiedTimeUtc'],
                orderBy: []
            },
            { maxRows }
        );
        return { ...result, source: 'configurationReport' };
    }

    // ── Assignment groups ───────────────────────────────────────────────────
    // Creating groups and changing membership needs CONFIG.msal.groupManageScopes
    // (requested incrementally). Policy assignments themselves only need the
    // base DeviceManagementConfiguration scope.

    // Throws ConsentRequiredError when the group permissions aren't granted.
    async checkGroupManageAccess() {
        await authService.getAccessTokenForScopes(CONFIG.msal.groupManageScopes);
        return true;
    }

    async makeGroupRequest(endpoint, method = 'GET', body = null) {
        const url = endpoint.startsWith('https://') ? endpoint : `${this.config.v1Url}${endpoint}`;
        return this.makeScopedRequest(CONFIG.msal.groupManageScopes, url, method, body);
    }

    // v1.0 $batch with the group scopes, 20 requests per call.
    // Returns Map<request id, { status, body }>.
    async groupBatch(requests) {
        const responses = new Map();
        for (let i = 0; i < requests.length; i += 20) {
            const chunk = requests.slice(i, i + 20).map(request => (
                request.body ? { ...request, headers: { 'Content-Type': 'application/json' } } : request
            ));
            const result = await this.makeGroupRequest('/$batch', 'POST', { requests: chunk });
            (result?.responses ?? []).forEach(response => responses.set(response.id, response));
        }
        return responses;
    }

    async findGroupsByName(displayName) {
        const filter = encodeURIComponent(`displayName eq '${GraphAPI.escapeODataString(displayName)}'`);
        const result = await this.makeGroupRequest(
            `/groups?$filter=${filter}&$select=id,displayName,securityEnabled,groupTypes&$top=10`
        );
        return result?.value ?? [];
    }

    /**
     * Reuse the security group with this exact name, or create an assigned
     * (static) security group. Returns { id, displayName, created, dynamic, duplicates }.
     */
    async ensureGroup(displayName, description) {
        const matches = await this.findGroupsByName(displayName);
        const existing = matches.find(g => g.securityEnabled) || matches[0];
        if (existing) {
            return {
                id: existing.id,
                displayName: existing.displayName,
                created: false,
                dynamic: (existing.groupTypes || []).includes('DynamicMembership'),
                duplicates: matches.length > 1 ? matches.length : 0
            };
        }

        const group = await this.makeGroupRequest('/groups', 'POST', {
            displayName,
            description: description.slice(0, 1024),
            mailEnabled: false,
            securityEnabled: true,
            mailNickname: buildMailNickname(displayName)
        });
        return { id: group.id, displayName: group.displayName, created: true, dynamic: false, duplicates: 0 };
    }

    // Rename a group (and update its description). mailNickname is left as is.
    async renameGroup(groupId, displayName, description = null) {
        await this.makeGroupRequest(`/groups/${encodeURIComponent(groupId)}`, 'PATCH', {
            displayName,
            ...(description ? { description: description.slice(0, 1024) } : {})
        });
    }

    // Collection a policy lives in, for /assignments and /assign.
    static assignmentCollectionFor(policy) {
        if (policy?.endpoint) return policy.endpoint;
        return {
            configurationPolicy: '/deviceManagement/configurationPolicies',
            compliance: '/deviceManagement/deviceCompliancePolicies',
            deviceConfiguration: '/deviceManagement/deviceConfigurations',
            adminTemplate: '/deviceManagement/groupPolicyConfigurations',
            endpointSecurity: '/deviceManagement/intents'
        }[policy?.type] ?? null;
    }

    async getAssignmentsFor(policy) {
        const collection = GraphAPI.assignmentCollectionFor(policy);
        if (!collection) throw new Error(`Assignments not supported for policy type ${policy?.type}`);
        const result = await this.makeScopedRequest(
            CONFIG.msal.scopes,
            `${collection}/${encodeURIComponent(policy.id)}/assignments`
        );
        return result?.value ?? [];
    }

    // /assign replaces the policy's entire assignment list.
    async setAssignmentsFor(policy, assignments) {
        const collection = GraphAPI.assignmentCollectionFor(policy);
        if (!collection) throw new Error(`Assignments not supported for policy type ${policy?.type}`);
        return this.makeScopedRequest(
            CONFIG.msal.scopes,
            `${collection}/${encodeURIComponent(policy.id)}/assign`,
            'POST',
            { assignments }
        );
    }

    /**
     * Add targets to a policy while keeping its existing assignments.
     * A group created seconds ago may not have replicated to Intune yet, so
     * failed assigns are retried a few times.
     */
    async addAssignmentTargets(policy, targets, { retries = 3, retryDelay = 4000 } = {}) {
        const existing = await this.getAssignmentsFor(policy);
        const { assignments, added } = mergeAssignmentTargets(existing, targets);
        if (added === 0) return { added: 0 };

        for (let attempt = 1; ; attempt++) {
            try {
                await this.setAssignmentsFor(policy, assignments);
                return { added };
            } catch (error) {
                const retryable = error.status === 400 || error.status === 404;
                if (!retryable || attempt >= retries) throw error;
                logger.warn(`Assign failed (attempt ${attempt}), retrying — new groups may still be replicating:`, error.message);
                await this.sleep(retryDelay);
            }
        }
    }

    /**
     * Assign a freshly deployed (or existing) policy: to new include/exclude
     * groups, or to All devices / All users plus an optional exclude group.
     * Groups are created, or reused when one with the same name exists.
     *
     * @param policy  { id, endpoint?, type?, displayName }
     * @param options { naming, target, include, exclude }
     * @returns { builtIn, include?, exclude?, added, skipped?, error? }
     */
    async ensureAssignmentGroupsForPolicy(policy, options) {
        const plan = planPolicyAssignment(policy, options);
        const outcome = { added: 0, builtIn: plan.builtIn, filter: plan.filter };
        if (plan.skipped) {
            outcome.skipped = plan.skipped;
            return outcome;
        }
        try {
            for (const kind of plan.groups) {
                const name = buildGroupName(policy.displayName, kind, options.naming);
                outcome[kind] = await this.ensureGroup(
                    name,
                    `${kind === 'include' ? 'Include' : 'Exclude'} group for the Intune policy "${policy.displayName}". Created by OIB Deployer.`
                );
            }
            // The filter goes on the include target; exclusions can't have one.
            const targets = [
                ...(plan.builtIn ? [withFilter(builtInTarget(plan.builtIn), plan.filter)] : []),
                ...plan.groups.map(kind => (kind === 'include'
                    ? withFilter(groupTarget(outcome[kind].id, kind), plan.filter)
                    : groupTarget(outcome[kind].id, kind))),
            ];
            // Retrying only helps while new groups replicate to Intune.
            const { added } = await this.addAssignmentTargets(policy, targets, { retries: plan.groups.length > 0 ? 3 : 1 });
            outcome.added = added;
        } catch (error) {
            outcome.error = error;
        }
        return outcome;
    }

    // ── Group membership (version migration) ────────────────────────────────

    async getGroupMembers(groupId) {
        const items = [];
        let next = `/groups/${encodeURIComponent(groupId)}/members?$top=999`;
        while (next) {
            const page = await this.makeGroupRequest(next);
            items.push(...(page?.value ?? []));
            next = page?.['@odata.nextLink'] ?? null;
        }
        return items.map(normaliseDirectoryObject);
    }

    /**
     * Add directory objects to a group. Already-present members count as
     * success. Returns [{ id, ok, error? }].
     */
    async addGroupMembers(groupId, memberIds) {
        const responses = await this.groupBatch(memberIds.map((id, i) => ({
            id: String(i),
            method: 'POST',
            url: `/groups/${groupId}/members/$ref`,
            body: { '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${id}` }
        })));
        return memberIds.map((id, i) => {
            const response = responses.get(String(i));
            const message = response?.body?.error?.message || '';
            const ok = !!response && (response.status < 300 || (response.status === 400 && /already exist/i.test(message)));
            return { id, ok, error: ok ? null : (message || `HTTP ${response?.status ?? 'no response'}`) };
        });
    }

    // Remove directory objects from a group. Missing members count as success.
    async removeGroupMembers(groupId, memberIds) {
        const responses = await this.groupBatch(memberIds.map((id, i) => ({
            id: String(i),
            method: 'DELETE',
            url: `/groups/${groupId}/members/${id}/$ref`
        })));
        return memberIds.map((id, i) => {
            const response = responses.get(String(i));
            const ok = !!response && (response.status < 300 || response.status === 404);
            return { id, ok, error: ok ? null : (response?.body?.error?.message || `HTTP ${response?.status ?? 'no response'}`) };
        });
    }

    // Search Entra devices and/or users by display name (and UPN for users).
    async searchDirectory(query, types = ['device']) {
        const q = encodeURIComponent(GraphAPI.escapeODataString(query.trim()));
        if (!q) return [];
        const requests = [];
        if (types.includes('device')) {
            requests.push(this.makeGroupRequest(
                `/devices?$filter=startswith(displayName,'${q}')&$select=id,displayName,operatingSystem,operatingSystemVersion&$top=25`
            ).then(r => (r?.value ?? []).map(d => normaliseDirectoryObject({ '@odata.type': '#microsoft.graph.device', ...d }))));
        }
        if (types.includes('user')) {
            requests.push(this.makeGroupRequest(
                `/users?$filter=startswith(displayName,'${q}') or startswith(userPrincipalName,'${q}')&$select=id,displayName,userPrincipalName&$top=25`
            ).then(r => (r?.value ?? []).map(u => normaliseDirectoryObject({ '@odata.type': '#microsoft.graph.user', ...u }))));
        }
        const results = await Promise.all(requests);
        return results.flat();
    }

    async listAllGroupScoped(endpoint) {
        const items = [];
        let next = endpoint;
        while (next) {
            const page = await this.makeGroupRequest(next);
            items.push(...(page?.value ?? []));
            next = page?.['@odata.nextLink'] ?? null;
        }
        return items;
    }

    /**
     * Everyone a policy assigned to All devices / All users reaches, as
     * directory objects that can be added to a group: Entra devices of the
     * policy's platform, and/or member users (guests are skipped).
     *
     * @param types     ['device'] and/or ['user']
     * @param platform  policyPlatform() value; null lists every device
     */
    async listAudience(types, platform = null) {
        const members = [];
        if (types.includes('device')) {
            const devices = await this.listAllGroupScoped(
                '/devices?$select=id,displayName,operatingSystem,operatingSystemVersion&$top=999'
            );
            members.push(...devices
                .filter(d => deviceMatchesPlatform(d.operatingSystem, platform))
                .map(d => normaliseDirectoryObject({ '@odata.type': '#microsoft.graph.device', ...d })));
        }
        if (types.includes('user')) {
            // userType isn't in the basic profile User.ReadBasic.All allows,
            // so guests are recognised by their #EXT# UPN.
            const users = await this.listAllGroupScoped('/users?$select=id,displayName,userPrincipalName&$top=999');
            members.push(...users
                .filter(u => !String(u.userPrincipalName ?? '').includes('#EXT#'))
                .map(u => normaliseDirectoryObject({ '@odata.type': '#microsoft.graph.user', ...u })));
        }
        return members;
    }
}

// Create singleton instance
export const graphAPI = new GraphAPI();
export default graphAPI;
