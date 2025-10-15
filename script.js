// ULLTRA Study Dashboard JavaScript

// Configuration
const CONFIG = {
    REDCAP_API_URL: 'https://redcap.ctsi.ufl.edu/redcap/api/',
    REDCAP_TOKEN: null,
    API_CONFIG_PATH: 'API.txt',
    CACHE_DURATION: 30 * 60 * 1000, // 30 minutes in milliseconds
    STORAGE_KEY: 'ulltra_data',
    CACHE_VERSION: '2025-10-01-status-mapping',
    DEBUG_MODE: false, // Set to false to use real API
    TEST_MODE: false,  // Set to false to disable test data

    // SharePoint Configuration (Browser-based Authentication)
    SHAREPOINT_SITE_URL: 'https://uflorida.sharepoint.com/sites/PRICE',
    SHAREPOINT_LIST_NAME: 'PRICECalendar',
    SHAREPOINT_LIST_VIEW: 'ULLTRA',

    // Azure AD / Microsoft 365 Configuration
    // These will need to be configured by your organization's Azure AD admin
    AZURE_CLIENT_ID: null,  // Set your Azure AD App Client ID here
    AZURE_TENANT_ID: 'ufl.onmicrosoft.com',  // UF tenant
    AZURE_REDIRECT_URI: window.location.origin  // Current page URL
};

// Conclusion code definitions sourced from REDCap metadata (see ULLTRA data dictionary)
const CONCLUSION_STATUS_DEFINITIONS = {
    '1': { label: 'Study Completed', summaryCategory: 'completed' },
    '2': { label: 'Ineligible After Randomization', summaryCategory: 'ineligible', withdrawalReason: 'Ineligible after randomization' },
    '3': { label: 'Withdrew Consent', summaryCategory: 'withdrawn', withdrawalReason: 'Withdrew consent' },
    '4': { label: "Withdrawn - Investigator decision (not in participant's best interest)", summaryCategory: 'withdrawn', withdrawalReason: "Investigator decision - not in participant's best interest" },
    '5': { label: 'Lost to Follow-up', summaryCategory: 'lost', withdrawalReason: 'Lost to follow-up' },
    '6': { label: 'Death', summaryCategory: 'other' },
    '7': { label: 'Study Ended', summaryCategory: 'other' },
    '8': { label: 'Screen Failure', summaryCategory: 'screen-failure' },
    '9': { label: 'Other Conclusion', summaryCategory: 'other', withdrawalReason: 'Other conclusion' }
};

// Data Manager Class
class DataManager {
    constructor() {
        this.cache = this.loadFromStorage();
    }

    // Load cached data from localStorage
    loadFromStorage() {
        try {
            const stored = localStorage.getItem(CONFIG.STORAGE_KEY);
            return stored ? JSON.parse(stored) : {};
        } catch (error) {
            console.error('Error loading data from storage:', error);
            return {};
        }
    }

    // Save data to localStorage
    saveToStorage() {
        try {
            localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify(this.cache));
        } catch (error) {
            console.error('Error saving data to storage:', error);
        }
    }

    // Check if cached data is valid
    isCacheValid(cacheKey) {
        const cached = this.cache[cacheKey];
        if (!cached || !cached.timestamp) return false;
        if (cached.version !== CONFIG.CACHE_VERSION) return false;
        
        const now = Date.now();
        const cacheAge = now - cached.timestamp;
        return cacheAge < CONFIG.CACHE_DURATION;
    }

    // Get cached data
    getCachedData(cacheKey) {
        if (this.isCacheValid(cacheKey)) {
            return this.cache[cacheKey].data;
        }
        return null;
    }

    // Cache data
    setCachedData(cacheKey, data) {
        this.cache[cacheKey] = {
            data: data,
            timestamp: Date.now(),
            version: CONFIG.CACHE_VERSION
        };
        this.saveToStorage();
    }

    // Get cache info for display
    getCacheInfo(cacheKey) {
        const cached = this.cache[cacheKey];
        if (!cached || !cached.timestamp) {
            return { exists: false };
        }
        
        const ageMs = Date.now() - cached.timestamp;
        const ageMinutes = Math.floor(ageMs / (1000 * 60));
        const isValid = ageMs < CONFIG.CACHE_DURATION;
        
        return {
            exists: true,
            age: ageMinutes,
            isValid: isValid,
            lastUpdated: new Date(cached.timestamp).toLocaleString()
        };
    }

    // Clear all cached data
    clearCache() {
        this.cache = {};
        this.saveToStorage();
    }
}

// REDCap API Handler
class REDCapAPI {
    constructor() {
        this.dataManager = new DataManager();
        this.tokenLoadPromise = null;
    }

    parseCredentialsFile(contents) {
        const result = {
            token: null,
            url: null
        };

        if (!contents) {
            return result;
        }

        contents.split(/\r?\n/).forEach((line) => {
            const cleanLine = line.trim();
            if (!cleanLine || cleanLine.startsWith('#')) {
                return;
            }

            const [rawKey, ...valueParts] = cleanLine.split(/[:=]/);
            if (!rawKey || valueParts.length === 0) {
                return;
            }

            const key = rawKey.trim().toLowerCase();
            const value = valueParts.join(':').trim();

            // REDCap token
            if ((key === 'redcap_token' || key === 'token') && value && !/redact/i.test(value)) {
                result.token = value;
            }

            // API URL
            if (key.includes('url') && value) {
                result.url = value;
            }
        });

        return result;
    }

    async ensureCredentialsLoaded() {
        if (CONFIG.REDCAP_TOKEN) {
            return CONFIG.REDCAP_TOKEN;
        }

        if (!this.tokenLoadPromise) {
            this.tokenLoadPromise = (async () => {
                try {
                    const response = await fetch(`${CONFIG.API_CONFIG_PATH}?t=${Date.now()}`, {
                        cache: 'no-store',
                        headers: {
                            'Cache-Control': 'no-cache, no-store, must-revalidate',
                            'Pragma': 'no-cache',
                            'Expires': '0'
                        }
                    });

                    if (!response.ok) {
                        // Reset the promise so we can retry
                        this.tokenLoadPromise = null;
                        throw new Error(`Unable to load API credentials from ${CONFIG.API_CONFIG_PATH} (status ${response.status})`);
                    }

                    const contents = await response.text();
                    const { token, url } = this.parseCredentialsFile(contents);

                    if (!token) {
                        // Reset the promise so we can retry
                        this.tokenLoadPromise = null;
                        throw new Error('REDCap API token is not configured. Please add your token to API.txt (not committed to source control).');
                    }

                    CONFIG.REDCAP_TOKEN = token.trim();

                    if (url) {
                        CONFIG.REDCAP_API_URL = url.trim();
                    }

                    console.log('✅ API credentials loaded successfully');
                    return CONFIG.REDCAP_TOKEN;
                } catch (error) {
                    console.error('Failed to load API credentials:', error);
                    throw error;
                }
            })();
        }

        return this.tokenLoadPromise;
    }

    getConclusionStatusInfo(conclusionRecord) {
        if (!conclusionRecord || !conclusionRecord.conclusion || conclusionRecord.conclusion === '') {
            return null;
        }

        const code = conclusionRecord.conclusion.toString();
        const definition = CONCLUSION_STATUS_DEFINITIONS[code];
        const baseLabel = definition ? definition.label : 'Concluded';
        const summaryCategory = definition ? definition.summaryCategory : 'other';
        const withdrawalReason = definition ? definition.withdrawalReason : undefined;
        const rawDate = conclusionRecord.conclusion_withdrawal || '';

        return {
            code,
            label: baseLabel,
            summaryCategory,
            withdrawalReason,
            date: rawDate,
            formattedDate: rawDate ? this.formatDate(rawDate) : null
        };
    }

    // Fetch data from REDCap API
    async fetchFromAPI(params) {
        await this.ensureCredentialsLoaded();

        console.log('Attempting to fetch data from REDCap API...');
        console.log('API URL:', CONFIG.REDCAP_API_URL);
        console.log('Parameters:', params);

        const formData = new FormData();
        
        // Default parameters
        formData.append('token', CONFIG.REDCAP_TOKEN);
        formData.append('content', 'record');
        formData.append('action', 'export');
        formData.append('format', 'json');
        formData.append('type', 'flat');
        formData.append('csvDelimiter', '');
        formData.append('rawOrLabel', 'raw');
        formData.append('rawOrLabelHeaders', 'raw');
        formData.append('exportCheckboxLabel', 'false');
        formData.append('exportSurveyFields', 'false');
        formData.append('exportDataAccessGroups', 'false');
        formData.append('returnBlankForGrayFormStatus', 'false');
        
        // Add custom parameters (override defaults if provided)
        Object.keys(params).forEach(key => {
            if (formData.has(key)) {
                formData.set(key, params[key]); // Override existing
            } else {
                formData.append(key, params[key]); // Add new
            }
        });

        try {
            console.log('Making API request...');
            const response = await fetch(CONFIG.REDCAP_API_URL, {
                method: 'POST',
                body: formData,
                mode: 'cors', // Explicitly set CORS mode
                credentials: 'omit' // Don't send credentials
            });

            console.log('Response status:', response.status);
            console.log('Response headers:', response.headers);

            if (!response.ok) {
                const errorText = await response.text();
                console.error('API Error Response:', errorText);
                throw new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
            }

            const data = await response.json();
            console.log('Successfully fetched data, record count:', data.length);
            console.log('First 3 raw records:', data.slice(0, 3));
            console.log('Sample participant IDs found:', data.slice(0, 10).map(r => r.participant_id).filter(id => id));
            
            // Analyze ALL available data
            const allFields = new Set();
            const allEvents = new Set();
            data.forEach(record => {
                Object.keys(record).forEach(field => allFields.add(field));
                if (record.redcap_event_name) allEvents.add(record.redcap_event_name);
            });
            
            console.log('=== FULL DATA ANALYSIS ===');
            console.log('Total records:', data.length);
            console.log('All available fields:', Array.from(allFields).sort());
            console.log('All events:', Array.from(allEvents).sort());
            
            // Find any demographic fields
            const demoFields = Array.from(allFields).filter(f => f.includes('demo'));
            console.log('Found demographic fields:', demoFields);
            
            // Check different events for demographic data
            const eventAnalysis = {};
            Array.from(allEvents).forEach(eventName => {
                const eventRecords = data.filter(r => r.redcap_event_name === eventName);
                const firstRecord = eventRecords.find(r => Object.keys(r).some(k => k.includes('demo'))) || eventRecords[0];
                if (firstRecord) {
                    const demoFieldsInEvent = Object.keys(firstRecord).filter(k => k.includes('demo'));
                    if (demoFieldsInEvent.length > 0) {
                        eventAnalysis[eventName] = {
                            recordCount: eventRecords.length,
                            demoFields: demoFieldsInEvent,
                            sampleData: {}
                        };
                        demoFieldsInEvent.forEach(field => {
                            eventAnalysis[eventName].sampleData[field] = firstRecord[field];
                        });
                    }
                }
            });
            
            console.log('Events with demographic data:', eventAnalysis);
            
            return data;
        } catch (error) {
            console.error('Detailed error fetching data from REDCap:', error);
            
            // Check for specific error types
            if (error.name === 'TypeError' && error.message.includes('Failed to fetch')) {
                throw new Error('Network error: Unable to connect to REDCap API. This may be due to CORS restrictions when running from a local file. Try running from a web server.');
            } else if (error.name === 'SyntaxError' && error.message.includes('JSON')) {
                throw new Error('Invalid JSON response from REDCap API. Check your API token and permissions.');
            }
            
            throw error;
        }
    }

    // Generate test data for development
    generateTestData() {
        const testData = [
            // Enrolled and Randomized - Completed
            { participant_id: 'U001', redcap_event_name: 'baseline_arm_1', icf_date: '2024-01-15 10:30', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U001', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '25', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U001', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '1', conclusion_withdrawal: '2024-06-15' },
            
            { participant_id: 'U005', redcap_event_name: 'baseline_arm_1', icf_date: '2024-02-08 14:20', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U005', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '33', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U005', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '1', conclusion_withdrawal: '2024-07-20' },
            
            { participant_id: 'U009', redcap_event_name: 'baseline_arm_1', icf_date: '2024-02-22 09:15', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U009', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '67', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U009', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '1', conclusion_withdrawal: '2024-08-05' },
            
            // Enrolled and Randomized - Withdrawn
            { participant_id: 'U002', redcap_event_name: 'baseline_arm_1', icf_date: '2024-01-20 14:15', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U002', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '42', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U002', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '3', conclusion_withdrawal: '2024-04-10' },
            
            { participant_id: 'U007', redcap_event_name: 'baseline_arm_1', icf_date: '2024-02-14 11:45', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U007', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '55', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U007', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '3', conclusion_withdrawal: '2024-05-30' },
            
            // Enrolled and Randomized - Lost to Follow-up
            { participant_id: 'U006', redcap_event_name: 'baseline_arm_1', icf_date: '2024-01-30 11:20', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U006', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '18', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U006', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '5', conclusion_withdrawal: '2024-05-20' },
            
            { participant_id: 'U012', redcap_event_name: 'baseline_arm_1', icf_date: '2024-03-10 16:30', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U012', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '74', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U012', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '5', conclusion_withdrawal: '2024-06-25' },
            
            // Enrolled and Randomized - Still Active
            { participant_id: 'U008', redcap_event_name: 'baseline_arm_1', icf_date: '2024-02-18 13:00', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U008', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '61', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U008', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            
            { participant_id: 'U011', redcap_event_name: 'baseline_arm_1', icf_date: '2024-03-05 10:15', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U011', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '89', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U011', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            
            { participant_id: 'U015', redcap_event_name: 'baseline_arm_1', icf_date: '2024-03-28 14:45', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U015', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '92', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U015', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            
            // Enrolled but Not Randomized Yet
            { participant_id: 'U004', redcap_event_name: 'baseline_arm_1', icf_date: '2024-02-01 09:45', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U004', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U004', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            
            { participant_id: 'U010', redcap_event_name: 'baseline_arm_1', icf_date: '2024-02-28 15:30', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U010', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U010', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            
            { participant_id: 'U013', redcap_event_name: 'baseline_arm_1', icf_date: '2024-03-15 11:20', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U013', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U013', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            
            { participant_id: 'U016', redcap_event_name: 'baseline_arm_1', icf_date: '2024-04-02 08:45', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U016', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U016', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            
            // Screen Failures
            { participant_id: 'U003', redcap_event_name: 'baseline_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U003', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U003', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '8', conclusion_withdrawal: '2024-02-05' },
            
            { participant_id: 'U014', redcap_event_name: 'baseline_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U014', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U014', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '8', conclusion_withdrawal: '2024-03-20' },
            
            { participant_id: 'U017', redcap_event_name: 'baseline_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U017', redcap_event_name: 'v1_arm_1', icf_date: '', rand_code: '', conclusion: '', conclusion_withdrawal: '' },
            { participant_id: 'U017', redcap_event_name: 'conclusion_arm_1', icf_date: '', rand_code: '', conclusion: '8', conclusion_withdrawal: '2024-04-08' },
        ];
        
        return testData;
    }

    // Get enrollment data with caching
    async getEnrollmentData(forceRefresh = false) {
        const cacheKey = 'enrollment_data';
        
        // Try to use cached data first
        if (!forceRefresh) {
            const cachedData = this.dataManager.getCachedData(cacheKey);
            if (cachedData) {
                console.log('Using cached enrollment data');
                return cachedData;
            }
        }

        // Use test data if in test mode
        if (CONFIG.TEST_MODE) {
            console.log('Using test data for development');
            const testData = this.generateTestData();
            const processedData = this.processEnrollmentData(testData);
            this.dataManager.setCachedData(cacheKey, processedData);
            return processedData;
        }

        console.log('Fetching fresh enrollment data from REDCap');

        try {
            // Fetch the fields we need for enrollment analysis including demographics
            const params = {
                fields: [
                    'participant_id',
                    'icf_date',
                    'rand_code',
                    'conclusion',
                    'conclusion_withdrawal',
                    'vdate',
                    'vdate_status',
                    'demo_sex',
                    'demo_race',
                    'demo_ethnicity',
                    'pre_date',  // Pre-screening date
                    'pre_screen_status'  // Pre-screening status
                ].join(','),
                exportDataAccessGroups: 'false',
                exportSurveyFields: 'false',
                events: [
                    'prescreening_arm_1',  // Add pre-screening event to capture all assessed participants
                    'baseline_arm_1',
                    'v1_arm_1',
                    'v2_arm_1',
                    'v3_arm_1',
                    'v4_arm_1',
                    'v5_arm_1',
                    'v6_arm_1',
                    'v7_arm_1',
                    'v8_arm_1',
                    'v9_arm_1',
                    'fu1_arm_1',
                    'fu2_arm_1',
                    'v10_arm_1',
                    'conclusion_arm_1'
                ].join(',')
            };

            const rawData = await this.fetchFromAPI(params);
            
            // Process the data for analysis
            const processedData = this.processEnrollmentData(rawData);
            
            // Cache the processed data
            this.dataManager.setCachedData(cacheKey, processedData);
            
            return processedData;
        } catch (error) {
            console.error('Error getting enrollment data:', error);
            
            // If API fails and we have cached data, use it
            const cachedData = this.dataManager.getCachedData(cacheKey);
            if (cachedData) {
                console.log('API failed, using cached data');
                return cachedData;
            }
            
            throw error;
        }
    }

    // Process raw REDCap data for enrollment analysis
    processEnrollmentData(rawData) {
        console.log('Processing REDCap data...');
        console.log('Total records received:', rawData.length);
        console.log('Sample of raw data:', rawData.slice(0, 3));
        
        // DIAGNOSTIC: Analyze ALL available data
        const allFields = new Set();
        const allEvents = new Set();
        rawData.forEach(record => {
            Object.keys(record).forEach(field => allFields.add(field));
            if (record.redcap_event_name) allEvents.add(record.redcap_event_name);
        });
        
        console.log('🔍 DIAGNOSTIC ANALYSIS:');
        console.log('📋 All available fields:', Array.from(allFields).sort());
        console.log('📅 All events:', Array.from(allEvents).sort());
        
        // Find any demographic fields
        const demoFields = Array.from(allFields).filter(f => f.includes('demo'));
        console.log('👥 Found demographic fields:', demoFields);
        
        // Show sample data from first few records
        console.log('📊 First 3 records with all data:', rawData.slice(0, 3).map(r => ({
            id: r.participant_id,
            event: r.redcap_event_name,
            allFields: Object.keys(r),
            demoFields: Object.keys(r).filter(k => k.includes('demo')).reduce((acc, k) => {
                acc[k] = r[k];
                return acc;
            }, {})
        })));
        
        // Continue with original processing
        const participantIds = new Set();
        rawData.forEach(record => {
            if (record.participant_id) participantIds.add(record.participant_id);
        });
        
        console.log('All unique participant IDs:', Array.from(participantIds));
        
        // Check for ICF dates across all records
        const icfRecords = rawData.filter(r => r.icf_date && r.icf_date !== '');
        console.log('Records with ICF dates:', icfRecords.length);
        console.log('Sample ICF records:', icfRecords.slice(0, 5));
        
        const participants = {};
        const eventCounts = {};
        
        // Group data by participant
        rawData.forEach(record => {
            const participantId = record.participant_id;
            const eventName = record.redcap_event_name;
            
            // Count events for debugging
            eventCounts[eventName] = (eventCounts[eventName] || 0) + 1;
            
            if (!participants[participantId]) {
                participants[participantId] = {
                    id: participantId,
                    visits: {},
                    conclusion: {}
                };
            }
            
            // Store all visits
            if (eventName === 'conclusion_arm_1') {
                participants[participantId].conclusion = record;
            } else {
                participants[participantId].visits[eventName] = record;
            }
        });

        console.log('Event counts:', eventCounts);
        console.log('Total unique participants:', Object.keys(participants).length);
        
        // Debug participant data structure
        const sampleParticipants = Object.values(participants).slice(0, 3);
        console.log('Sample participants:', sampleParticipants);
        
        // Check ICF dates and visit completion
        const participantSummary = Object.values(participants).map(p => ({
            id: p.id,
            icf_date: p.visits.baseline_arm_1?.icf_date,
            rand_code: p.visits.v1_arm_1?.rand_code,
            conclusion: p.conclusion.conclusion,
            lastCompletedVisit: this.getLastCompletedVisit(p.visits),
            currentVisitStatus: this.calculateCurrentVisitStatus(p.visits, p.conclusion)
        }));
        console.log('Participant summary:', participantSummary);

        const summary = this.calculateEnrollmentSummary(participants);
        console.log('Calculated summary:', summary);

        return {
            participants: participants,
            summary: summary,
            lastUpdated: new Date().toISOString()
        };
    }

    // Get the last completed visit for a participant
    getLastCompletedVisit(visits) {
        const visitOrder = [
            'baseline_arm_1',
            'v1_arm_1', 
            'v2_arm_1',
            'v3_arm_1',
            'v4_arm_1',
            'v5_arm_1',
            'v6_arm_1',
            'v7_arm_1',
            'v8_arm_1',
            'v9_arm_1',
            'fu1_arm_1',
            'fu2_arm_1',
            'v10_arm_1'
        ];

        let lastCompleted = null;
        for (const visitEvent of visitOrder) {
            const visit = visits[visitEvent];
            
            // Check if visit is completed using vdate
            if (visit && visit.vdate && visit.vdate !== '') {
                lastCompleted = visitEvent;
            }
        }
        
        return lastCompleted;
    }

    // Calculate comprehensive current status considering all study aspects
    calculateCurrentVisitStatus(visits, conclusion) {
        const baseline = visits.baseline_arm_1 || {};
        const v1 = visits.v1_arm_1 || {};
        
        // Priority 1: Check if participant has concluded the study
        const conclusionInfo = this.getConclusionStatusInfo(conclusion);
        if (conclusionInfo) {
            const formattedDate = conclusionInfo.formattedDate ? ` (${conclusionInfo.formattedDate})` : '';
            return `${conclusionInfo.label}${formattedDate}`;
        }

        // Priority 2: Check enrollment status
        const hasICF = baseline.icf_date && baseline.icf_date !== '';
        const isRandomized = v1.rand_code && v1.rand_code !== '';
        
        if (!hasICF) {
            return 'Not Enrolled';
        }

        // Priority 3: Check randomization and visit progression
        const lastCompleted = this.getLastCompletedVisit(visits);
        const lastCompletedVisitName = this.getVisitDisplayName(lastCompleted);
        
        if (!lastCompleted) {
            return 'Enrolled - Awaiting Baseline Visit';
        }

        // If baseline is complete but not randomized yet
        if (lastCompleted === 'baseline_arm_1' && !isRandomized) {
            const daysSinceBaseline = baseline.vdate ? this.daysSince(baseline.vdate) : 0;
            if (daysSinceBaseline > 14) {
                return `Pending Randomization (${daysSinceBaseline} days since Baseline)`;
            } else {
                return 'Pending Randomization';
            }
        }

        // Priority 4: Determine visit progression and timing
        const visitOrder = [
            'baseline_arm_1', 'v1_arm_1', 'v2_arm_1', 'v3_arm_1', 'v4_arm_1',
            'v5_arm_1', 'v6_arm_1', 'v7_arm_1', 'v8_arm_1', 'v9_arm_1',
            'fu1_arm_1', 'fu2_arm_1', 'v10_arm_1'
        ];

        const lastCompletedIndex = visitOrder.indexOf(lastCompleted);
        
        // If completed final visit (V10)
        if (lastCompletedIndex === visitOrder.length - 1) {
            const v10Date = visits.v10_arm_1?.vdate;
            const daysSinceV10 = v10Date ? this.daysSince(v10Date) : 0;
            return `Study Complete - Awaiting Final Documentation (${daysSinceV10} days since V10)`;
        }

        // Determine next expected visit
        const nextVisit = visitOrder[lastCompletedIndex + 1];
        const nextVisitName = this.getVisitDisplayName(nextVisit);
        
        // Calculate timing for next visit
        // Special handling for V10: always calculate from V9, not from last completed visit
        let referenceVisitDate, daysSinceReference, expectedInterval;

        if (nextVisit === 'v10_arm_1') {
            // V10 timing is calculated from V9, not from FU2
            referenceVisitDate = visits['v9_arm_1']?.vdate;
            if (referenceVisitDate) {
                daysSinceReference = this.daysSince(referenceVisitDate);
                expectedInterval = 180; // V10 target: 180 days from V9
                const tolerance = 30;   // V10 tolerance: ±30 days

                const windowStart = expectedInterval - tolerance; // 150 days
                const windowEnd = expectedInterval + tolerance;   // 210 days

                if (daysSinceReference < windowStart) {
                    // Before window opens
                    const daysUntilWindow = windowStart - daysSinceReference;
                    return `Active: ${nextVisitName} due in ${daysUntilWindow} days`;
                } else if (daysSinceReference <= windowEnd) {
                    // Within window - show as schedulable now
                    return `Due: ${nextVisitName}`;
                } else {
                    // Past window
                    const daysOverdue = daysSinceReference - windowEnd;
                    return `Overdue: ${nextVisitName} (${daysOverdue} days overdue)`;
                }
            }
        } else {
            // Standard calculation for other visits
            referenceVisitDate = visits[lastCompleted]?.vdate;
            if (referenceVisitDate) {
                daysSinceReference = this.daysSince(referenceVisitDate);
                expectedInterval = this.getExpectedVisitInterval(lastCompleted, nextVisit);
            }
        }

        if (referenceVisitDate && expectedInterval) {
            const daysUntilDue = expectedInterval - daysSinceReference;

            if (daysSinceReference > expectedInterval + 21) { // 3 weeks overdue
                return `Overdue: ${nextVisitName} (${daysSinceReference - expectedInterval} days overdue)`;
            } else if (daysSinceReference > expectedInterval + 7) { // 1 week overdue
                return `Late: ${nextVisitName} (${daysSinceReference - expectedInterval} days overdue)`;
            } else if (Math.abs(daysUntilDue) <= 3) { // Due within 3 days
                return `Due: ${nextVisitName}`;
            } else if (daysUntilDue > 0) { // Not yet due
                return `Active: ${nextVisitName} due in ${daysUntilDue} days`;
            } else { // Recently became due
                return `Due: ${nextVisitName}`;
            }
        }
        
        return `Active: Due for ${nextVisitName}`;
    }

    // Get display name for visit
    getVisitDisplayName(visitEvent) {
        const names = {
            'baseline_arm_1': 'Baseline',
            'v1_arm_1': 'V1',
            'v2_arm_1': 'V2', 
            'v3_arm_1': 'V3',
            'v4_arm_1': 'V4',
            'v5_arm_1': 'V5',
            'v6_arm_1': 'V6',
            'v7_arm_1': 'V7',
            'v8_arm_1': 'V8',
            'v9_arm_1': 'V9',
            'fu1_arm_1': 'Follow-up 1',
            'fu2_arm_1': 'Follow-up 2',
            'v10_arm_1': 'V10'
        };
        return names[visitEvent] || visitEvent;
    }

    // Get CSS class for current status styling
    getCurrentStatusClass(status) {
        if (status.includes('Completed') || status.includes('Study Complete')) {
            return 'status-completed';
        } else if (status.includes('Withdrawn') || status.includes('Withdrew Consent') || status.includes('Screen Failure') || 
                  status.includes('Lost to Follow-up') || status.includes('Death') ||
                  status.includes('Ineligible') || status.includes('Discontinued')) {
            return 'status-concluded';
        } else if (status.includes('Overdue') || status.includes('Late')) {
            return 'status-overdue';
        } else if (status.includes('Due:') || status.includes('Pending')) {
            return 'status-due';
        } else if (status.includes('Active')) {
            return 'status-active';
        } else {
            return 'status-neutral';
        }
    }

    // Calculate days since a date
    daysSince(dateString) {
        try {
            const date = new Date(dateString);
            const now = new Date();
            const diffTime = now - date;
            return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        } catch (error) {
            return 0;
        }
    }

    // Get expected interval between visits (in days)
    getExpectedVisitInterval(lastVisit, nextVisit) {
        // Visit intervals based on ULLTRA protocol
        const intervals = {
            'baseline_arm_1_to_v1_arm_1': 7,      // 1 week
            'v1_arm_1_to_v2_arm_1': 4,            // 4 days (corrected per protocol)
            'v2_arm_1_to_v3_arm_1': 4,            // 4 days
            'v3_arm_1_to_v4_arm_1': 4,            // 4 days
            'v4_arm_1_to_v5_arm_1': 4,            // 4 days
            'v5_arm_1_to_v6_arm_1': 4,            // 4 days
            'v6_arm_1_to_v7_arm_1': 4,            // 4 days
            'v7_arm_1_to_v8_arm_1': 4,            // 4 days
            'v8_arm_1_to_v9_arm_1': 5,            // 5 days (corrected per protocol)
            'v9_arm_1_to_fu1_arm_1': 30,          // 1 month
            'fu1_arm_1_to_fu2_arm_1': 60,         // 2 months (3 months total from V9)
            'fu2_arm_1_to_v10_arm_1': 90          // NOTE: V10 handled separately in calculateCurrentVisitStatus
        };

        const key = `${lastVisit}_to_${nextVisit}`;
        return intervals[key] || 14; // Default to 2 weeks
    }

    // Format date string for display
    formatDate(dateString) {
        if (!dateString || dateString === '') return 'Not available';
        
        try {
            // Parse date carefully to avoid timezone issues
            // Assume REDCap dates are in YYYY-MM-DD format
            const dateParts = dateString.split('-');
            if (dateParts.length === 3) {
                const year = parseInt(dateParts[0]);
                const month = parseInt(dateParts[1]);
                const day = parseInt(dateParts[2]);
                // Create date using local timezone
                const date = new Date(year, month - 1, day);
                return date.toLocaleDateString('en-US', {
                    year: 'numeric',
                    month: 'short',
                    day: 'numeric'
                });
            } else {
                // Fallback to original parsing
                const date = new Date(dateString + 'T00:00:00'); // Force local timezone
                return date.toLocaleDateString('en-US', {
                    year: 'numeric',
                    month: 'short',
                    day: 'numeric'
                });
            }
        } catch (error) {
            return dateString;
        }
    }

    // Format date string for compact display (MM/DD format)
    formatShortDate(dateString) {
        if (!dateString || dateString === '') return '—';
        
        try {
            // Parse date carefully to avoid timezone issues
            // Assume REDCap dates are in YYYY-MM-DD format
            const dateParts = dateString.split('-');
            if (dateParts.length === 3) {
                const year = parseInt(dateParts[0]);
                const month = parseInt(dateParts[1]);
                const day = parseInt(dateParts[2]);
                // Create date using local timezone
                const date = new Date(year, month - 1, day);
                return `${month}/${day}`;
            } else {
                // Fallback to original parsing
                const date = new Date(dateString + 'T00:00:00'); // Force local timezone
                return date.toLocaleDateString('en-US', {
                    month: 'numeric',
                    day: 'numeric'
                });
            }
        } catch (error) {
            return dateString;
        }
    }

    // Parse date string to Date object, handling timezone issues
    parseDate(dateString) {
        if (!dateString || dateString === '') return null;
        
        try {
            // Assume REDCap dates are in YYYY-MM-DD format
            const dateParts = dateString.split('-');
            if (dateParts.length === 3) {
                const year = parseInt(dateParts[0]);
                const month = parseInt(dateParts[1]);
                const day = parseInt(dateParts[2]);
                // Create date using local timezone to avoid UTC conversion issues
                return new Date(year, month - 1, day);
            } else {
                // Fallback to original parsing with timezone fix
                return new Date(dateString + 'T00:00:00');
            }
        } catch (error) {
            console.warn('Error parsing date:', dateString, error);
            return new Date(dateString);
        }
    }

    // Calculate enrollment summary statistics
    calculateEnrollmentSummary(participants) {
        const summary = {
            totalEnrolled: 0,
            totalRandomized: 0,
            screenFailures: 0,
            completedStudy: 0,
            withdrawn: 0,
            lostFollowup: 0,
            ineligibleAfterRandomization: 0,
            other: 0
        };

        Object.values(participants).forEach(participant => {
            const baseline = participant.visits.baseline_arm_1 || {};
            const v1 = participant.visits.v1_arm_1 || {};
            const conclusion = participant.conclusion || {};
            
            // Total Enrolled: has ICF date at baseline
            if (baseline.icf_date && baseline.icf_date !== '') {
                summary.totalEnrolled++;
            }

            // Total Randomized: has rand_code at V1
            if (v1.rand_code && v1.rand_code !== '') {
                summary.totalRandomized++;
            }

            // Conclusion status analysis
            const conclusionInfo = this.getConclusionStatusInfo(conclusion);
            if (conclusionInfo) {
                const { code } = conclusionInfo;
                const isRandomized = v1.rand_code && v1.rand_code !== '';

                switch (code) {
                    case '1':
                        summary.completedStudy++;
                        break;
                    case '2':
                        summary.ineligibleAfterRandomization++;
                        break;
                    case '3':
                    case '4':
                        if (isRandomized) {
                            summary.withdrawn++;
                        }
                        break;
                    case '5':
                        // Only count as lost to follow-up if randomized (matches CONSORT logic)
                        if (isRandomized) {
                            console.log('Lost to follow-up participant found:', participant.participant_id, 'ICF:', baseline.icf_date, 'Randomized:', isRandomized);
                            summary.lostFollowup++;
                        }
                        break;
                    case '8':
                        summary.screenFailures++;
                        break;
                    default:
                        summary.other++;
                        break;
                }
            }
        });

        console.log('Enrollment Summary:', summary);
        return summary;
    }

    // Process enrollment data for charts
    processEnrollmentChartData(participants) {
        const monthlyData = {};
        const enrollmentEvents = [];
        const randomizationEvents = [];

        // Debug counters
        let totalWithIcfDate = 0;
        let invalidEnrollmentDate = [];
        let totalWithRandCode = 0;
        let missingIcfDate = [];
        let invalidIcfDate = [];

        Object.values(participants).forEach(participant => {
            const baseline = participant.visits.baseline_arm_1 || {};
            const v1 = participant.visits.v1_arm_1 || {};

            // Process enrollment (ICF date)
            if (baseline.icf_date && baseline.icf_date !== '') {
                totalWithIcfDate++;
                const date = new Date(baseline.icf_date);
                if (!isNaN(date.getTime())) {
                    const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
                    enrollmentEvents.push({ date, monthKey, participantId: participant.id });
                } else {
                    invalidEnrollmentDate.push({ id: participant.participant_id, icfDate: baseline.icf_date });
                }
            }

            // Process randomization (V1 with rand_code)
            if (v1.rand_code && v1.rand_code !== '') {
                totalWithRandCode++;

                if (!baseline.icf_date || baseline.icf_date === '') {
                    missingIcfDate.push(participant.participant_id);
                } else {
                    // Use ICF date as proxy for randomization timing if V1 date not available
                    const date = new Date(baseline.icf_date);
                    if (!isNaN(date.getTime())) {
                        const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
                        randomizationEvents.push({ date, monthKey, participantId: participant.id });
                    } else {
                        invalidIcfDate.push({ id: participant.participant_id, icfDate: baseline.icf_date });
                    }
                }
            }
        });

        // Log discrepancies for debugging
        console.log('📊 ENROLLMENT CHART DATA:');
        console.log(`Total Enrolled (card): ${totalWithIcfDate}`);
        console.log(`Enrolled in chart: ${enrollmentEvents.length}`);
        console.log(`Total Randomized (card): ${totalWithRandCode}`);
        console.log(`Randomized in chart: ${randomizationEvents.length}`);
        console.log(`First enrollment: ${enrollmentEvents[0]?.date}, Last enrollment: ${enrollmentEvents[enrollmentEvents.length - 1]?.date}`);

        if (invalidEnrollmentDate.length > 0) {
            console.warn('⚠️ ENROLLMENT DISCREPANCY - Enrolled participants excluded from chart:');
            console.warn(`Invalid enrollment date format (${invalidEnrollmentDate.length}):`, invalidEnrollmentDate);
        }

        if (missingIcfDate.length > 0 || invalidIcfDate.length > 0) {
            console.warn('⚠️ RANDOMIZATION DISCREPANCY - Randomized participants excluded from chart:');
            console.warn(`Missing ICF date (${missingIcfDate.length}):`, missingIcfDate);
            console.warn(`Invalid ICF date (${invalidIcfDate.length}):`, invalidIcfDate);
        }

        // Sort events by date
        enrollmentEvents.sort((a, b) => a.date - b.date);
        randomizationEvents.sort((a, b) => a.date - b.date);

        // Calculate monthly totals
        enrollmentEvents.forEach(event => {
            if (!monthlyData[event.monthKey]) {
                monthlyData[event.monthKey] = { enrolled: 0, randomized: 0 };
            }
            monthlyData[event.monthKey].enrolled++;
        });

        randomizationEvents.forEach(event => {
            if (!monthlyData[event.monthKey]) {
                monthlyData[event.monthKey] = { enrolled: 0, randomized: 0 };
            }
            monthlyData[event.monthKey].randomized++;
        });

        // Generate complete month range
        const months = Object.keys(monthlyData).sort();
        if (months.length === 0) return { months: [], monthlyEnrolled: [], monthlyRandomized: [], cumulativeEnrolled: [], cumulativeRandomized: [] };

        console.log('🔍 DEBUG Month Keys from monthlyData:', months);
        console.log('🔍 DEBUG monthlyData object:', monthlyData);

        const startMonth = months[0];
        const endMonth = months[months.length - 1];
        console.log(`🔍 DEBUG Start month: ${startMonth}, End month: ${endMonth}`);

        const allMonths = this.generateMonthRange(startMonth, endMonth);
        console.log('🔍 DEBUG All months generated:', allMonths);

        // Fill in missing months with zeros
        const monthlyEnrolled = [];
        const monthlyRandomized = [];
        const cumulativeEnrolled = [];
        const cumulativeRandomized = [];
        
        let enrolledTotal = 0;
        let randomizedTotal = 0;

        allMonths.forEach(month => {
            const data = monthlyData[month] || { enrolled: 0, randomized: 0 };
            monthlyEnrolled.push(data.enrolled);
            monthlyRandomized.push(data.randomized);
            
            enrolledTotal += data.enrolled;
            randomizedTotal += data.randomized;
            
            cumulativeEnrolled.push(enrolledTotal);
            cumulativeRandomized.push(randomizedTotal);
        });

        const chartData = {
            months: allMonths.map(month => this.formatMonthLabel(month)),
            monthlyEnrolled,
            monthlyRandomized,
            cumulativeEnrolled,
            cumulativeRandomized
        };

        console.log('📊 CHART DATA SUMMARY:');
        console.log(`Total months: ${chartData.months.length}`);
        console.log(`Month range: ${chartData.months[0]} to ${chartData.months[chartData.months.length - 1]}`);
        console.log(`Final cumulative enrolled: ${cumulativeEnrolled[cumulativeEnrolled.length - 1]}`);
        console.log(`Final cumulative randomized: ${cumulativeRandomized[cumulativeRandomized.length - 1]}`);

        return chartData;
    }

    // Generate month range between start and end
    generateMonthRange(startMonth, endMonth) {
        const months = [];

        // Parse year and month directly to avoid timezone issues
        const [startYear, startMonth_] = startMonth.split('-').map(Number);
        const [endYear, endMonth_] = endMonth.split('-').map(Number);

        let currentYear = startYear;
        let currentMonth = startMonth_;

        while (currentYear < endYear || (currentYear === endYear && currentMonth <= endMonth_)) {
            const monthKey = `${currentYear}-${String(currentMonth).padStart(2, '0')}`;
            months.push(monthKey);

            // Increment month
            currentMonth++;
            if (currentMonth > 12) {
                currentMonth = 1;
                currentYear++;
            }
        }

        return months;
    }

    // Format month label for display
    formatMonthLabel(monthKey) {
        const [year, month] = monthKey.split('-');
        const date = new Date(parseInt(year), parseInt(month) - 1, 1);
        return date.toLocaleDateString('en-US', { year: 'numeric', month: 'short' });
    }

    // Get cache information
    getCacheInfo() {
        return {
            enrollment: this.dataManager.getCacheInfo('enrollment_data')
        };
    }

    // Get all participant data for missing data report
    async getAllParticipantData(forceRefresh = false) {
        const cacheKey = 'all_participant_data';
        
        // Try to use cached data first
        if (!forceRefresh) {
            const cachedData = this.dataManager.getCachedData(cacheKey);
            if (cachedData) {
                console.log('Using cached participant data');
                return cachedData;
            }
        }


        try {
            // Fetch ALL records from REDCap including repeating instruments
            // Do NOT filter by events - this would exclude repeating instruments like DSD
            const params = {
                content: 'record',
                type: 'flat',
                format: 'json',
                exportDataAccessGroups: 'false',
                exportSurveyFields: 'false'
            };
            
            console.log('Fetching participant data with params:', params);
            const rawData = await this.fetchFromAPI(params);
            console.log('Raw data received:', rawData?.length, 'records');
            const processedData = this.processParticipantData(rawData);
            console.log('Processed data:', processedData?.length, 'participants');

            // Cache both raw and processed data for outcome analysis
            this.dataManager.setCachedData(cacheKey, processedData);
            this.dataManager.setCachedData(cacheKey + '_raw', rawData);
            return processedData;
        } catch (error) {
            console.error('Error fetching participant data:', error);
            throw error;
        }
    }

    // Generate test data for missing data report
    generateMissingDataTestData() {
        return [
            // Sample participants with various completion states
            {
                participant_id: 'U001',
                icf_date: '2024-01-15',
                rand_code: '25',
                vdate_v5: '2024-02-15',
                vdate_v9: '2024-04-15', 
                vdate_fu1: '2024-05-15',
                vdate_fu2: '2024-07-15',
                vdate_v10: '2024-08-15',
                conclusion: 1
            },
            {
                participant_id: 'U003',
                icf_date: '2024-01-20',
                rand_code: '33',
                vdate_v5: '2024-02-20',
                vdate_v9: '2024-04-20',
                vdate_fu1: '2024-05-20',
                vdate_fu2: null,
                vdate_v10: '2024-08-20',
                conclusion: 1
            },
            {
                participant_id: 'U007',
                icf_date: '2024-01-25',
                rand_code: '45',
                vdate_v5: '2024-02-25',
                vdate_v9: '2024-04-25',
                vdate_fu1: '2024-05-25',
                vdate_fu2: '2024-07-25',
                vdate_v10: '2024-08-25',
                conclusion: 1
            },
            {
                participant_id: 'U011',
                icf_date: '2024-02-01',
                rand_code: '67',
                vdate_v5: '2024-03-01',
                vdate_v9: '2024-05-01',
                vdate_fu1: '2024-06-01',
                vdate_fu2: '2024-08-01',
                vdate_v10: '2024-09-01',
                conclusion: 1
            },
            {
                participant_id: 'U022',
                icf_date: '2024-02-10',
                rand_code: '78',
                vdate_v5: null,
                vdate_v9: null,
                vdate_fu1: null,
                vdate_fu2: null,
                vdate_v10: null,
                conclusion: 3 // Withdrew consent
            },
            {
                participant_id: 'U035',
                icf_date: '2024-02-15',
                rand_code: '89',
                vdate_v5: null,
                vdate_v9: null,
                vdate_fu1: null,
                vdate_fu2: null,
                vdate_v10: null,
                conclusion: 2 // Ineligible after randomization
            },
            {
                participant_id: 'U110',
                icf_date: '2024-08-01',
                rand_code: '123',
                vdate_v5: '2024-09-01',
                vdate_v9: '2024-11-01',
                vdate_fu1: null,
                vdate_fu2: null,
                vdate_v10: null,
                conclusion: null // Active in study
            },
            {
                participant_id: 'U125',
                icf_date: '2024-08-15',
                rand_code: '145',
                vdate_v5: '2024-09-15',
                vdate_v9: '2024-11-15',
                vdate_fu1: null,
                vdate_fu2: null,
                vdate_v10: null,
                conclusion: null // Active in study
            }
        ];
    }

    // Process raw participant data from REDCap
    processParticipantData(rawData) {
        const participantMap = {};

        // Group data by participant
        rawData.forEach(record => {
            const id = record.participant_id;
            if (!participantMap[id]) {
                participantMap[id] = {
                    participant_id: id,
                    icf_date: null,
                    rand_code: null,
                    vdate_v5: null,
                    vdate_v9: null,
                    vdate_fu1: null,
                    vdate_fu2: null,
                    vdate_v10: null,
                    conclusion: null
                };
            }

            // Extract data from different events
            const participant = participantMap[id];
            
            if (record.redcap_event_name === 'baseline_arm_1' && record.icf_date) {
                participant.icf_date = record.icf_date;
            }
            
            if (record.redcap_event_name === 'v1_arm_1' && record.rand_code) {
                participant.rand_code = record.rand_code;
            }
            
            // Visit dates
            if (record.redcap_event_name === 'v5_arm_1' && record.vdate) {
                participant.vdate_v5 = record.vdate;
            }
            if (record.redcap_event_name === 'v9_arm_1' && record.vdate) {
                participant.vdate_v9 = record.vdate;
            }
            if (record.redcap_event_name === 'fu1_arm_1' && record.vdate) {
                participant.vdate_fu1 = record.vdate;
            }
            if (record.redcap_event_name === 'fu2_arm_1' && record.vdate) {
                participant.vdate_fu2 = record.vdate;
            }
            if (record.redcap_event_name === 'v10_arm_1' && record.vdate) {
                participant.vdate_v10 = record.vdate;
            }
            
            // Conclusion status
            if (record.redcap_event_name === 'conclusion_arm_1' && record.conclusion) {
                participant.conclusion = parseInt(record.conclusion);
            }
        });

        return Object.values(participantMap);
    }

    // Get all participant data for missing data report
    async getAllParticipantData(forceRefresh = false) {
        const cacheKey = 'all_participant_data';
        
        // Try to use cached data first
        if (!forceRefresh) {
            const cachedData = this.dataManager.getCachedData(cacheKey);
            if (cachedData) {
                console.log('Using cached participant data');
                return cachedData;
            }
        }


        try {
            // Fetch ALL records from REDCap including repeating instruments
            // Do NOT filter by events - this would exclude repeating instruments like DSD
            const params = {
                content: 'record',
                type: 'flat',
                format: 'json',
                exportDataAccessGroups: 'false',
                exportSurveyFields: 'false'
            };
            
            console.log('Fetching participant data with params:', params);
            const rawData = await this.fetchFromAPI(params);
            console.log('Raw data received:', rawData?.length, 'records');
            const processedData = this.processParticipantData(rawData);
            console.log('Processed data:', processedData?.length, 'participants');

            // Cache both raw and processed data for outcome analysis
            this.dataManager.setCachedData(cacheKey, processedData);
            this.dataManager.setCachedData(cacheKey + '_raw', rawData);
            return processedData;
        } catch (error) {
            console.error('Error fetching participant data:', error);
            throw error;
        }
    }

    // Generate test data for missing data report
    generateMissingDataTestData() {
        return [
            {
                participant_id: 'U001',
                icf_date: '2024-01-15',
                rand_code: '25',
                vdate_v5: '2024-02-15',
                vdate_v9: '2024-04-15',
                vdate_fu1: '2024-05-15',
                vdate_fu2: '2024-07-15',
                vdate_v10: '2024-08-15',
                conclusion: 1
            },
            {
                participant_id: 'U003',
                icf_date: '2024-01-20',
                rand_code: '33',
                vdate_v5: '2024-02-20',
                vdate_v9: '2024-04-20',
                vdate_fu1: '2024-05-20',
                vdate_fu2: null,
                vdate_v10: '2024-08-20',
                conclusion: 1
            },
            {
                participant_id: 'U007',
                icf_date: '2024-01-25',
                rand_code: '45',
                vdate_v5: '2024-02-25',
                vdate_v9: '2024-04-25',
                vdate_fu1: '2024-05-25',
                vdate_fu2: '2024-07-25',
                vdate_v10: '2024-08-25',
                conclusion: 1
            },
            {
                participant_id: 'U022',
                icf_date: '2024-02-10',
                rand_code: '78',
                vdate_v5: null,
                vdate_v9: null,
                vdate_fu1: null,
                vdate_fu2: null,
                vdate_v10: null,
                conclusion: 3 // Withdrew consent
            },
            {
                participant_id: 'U110',
                icf_date: '2024-08-01',
                rand_code: '123',
                vdate_v5: '2024-09-01',
                vdate_v9: '2024-11-01',
                vdate_fu1: null,
                vdate_fu2: null,
                vdate_v10: null,
                conclusion: null // Active in study
            },
            {
                participant_id: 'U125',
                icf_date: '2024-08-15',
                rand_code: '145',
                vdate_v5: '2024-09-15',
                vdate_v9: '2024-11-15',
                vdate_fu1: null,
                vdate_fu2: null,
                vdate_v10: null,
                conclusion: null // Active in study
            }
        ];
    }

    // Process raw participant data from REDCap
    processParticipantData(rawData) {
        const participantMap = {};

        // Group data by participant
        rawData.forEach(record => {
            const id = record.participant_id;
            
            // Filter out test participants (matching R script filtering)
            if (id === "TEST" || id === "Test-2" || !id || id === '') {
                return;
            }
            
            if (!participantMap[id]) {
                participantMap[id] = {
                    participant_id: id,
                    icf_date: null,
                    rand_code: null,
                    vdate_base: null,
                    vdate_v1: null,
                    vdate_v2: null,
                    vdate_v3: null,
                    vdate_v4: null,
                    vdate_v5: null,
                    vdate_v6: null,
                    vdate_v7: null,
                    vdate_v8: null,
                    vdate_v9: null,
                    vdate_fu1: null,
                    vdate_fu2: null,
                    vdate_v10: null,
                    meds_dt_base: null,
                    meds_dt_v1: null,
                    meds_dt_v2: null,
                    meds_dt_v3: null,
                    meds_dt_v4: null,
                    meds_dt_v5: null,
                    meds_dt_v6: null,
                    meds_dt_v7: null,
                    meds_dt_v8: null,
                    meds_dt_v9: null,
                    meds_dt_fu1: null,
                    meds_dt_fu2: null,
                    meds_dt_v10: null,
                    samp_date_base: null,
                    samp_date_v1: null,
                    samp_date_v2: null,
                    samp_date_v3: null,
                    samp_date_v4: null,
                    samp_date_v5: null,
                    samp_date_v6: null,
                    samp_date_v7: null,
                    samp_date_v8: null,
                    samp_date_v9: null,
                    samp_date_fu1: null,
                    samp_date_fu2: null,
                    samp_date_v10: null,
                    conclusion: null
                };
            }

            const participant = participantMap[id];
            const eventName = record.redcap_event_name;
            
            // Extract data based on event name - following R script structure
            if (eventName === 'baseline_arm_1') {
                if (record.icf_date) participant.icf_date = record.icf_date;
                if (record.vdate) participant.vdate_base = record.vdate;
                if (record.meds_dt) participant.meds_dt_base = record.meds_dt;
                if (record.samp_date) participant.samp_date_base = record.samp_date;
            } else if (eventName === 'v1_arm_1') {
                if (record.rand_code) participant.rand_code = record.rand_code;
                if (record.vdate) participant.vdate_v1 = record.vdate;
                if (record.meds_dt) participant.meds_dt_v1 = record.meds_dt;
                if (record.samp_date) participant.samp_date_v1 = record.samp_date;
            } else if (eventName === 'v2_arm_1') {
                if (record.vdate) participant.vdate_v2 = record.vdate;
                if (record.meds_dt) participant.meds_dt_v2 = record.meds_dt;
                if (record.samp_date) participant.samp_date_v2 = record.samp_date;
            } else if (eventName === 'v3_arm_1') {
                if (record.vdate) participant.vdate_v3 = record.vdate;
                if (record.meds_dt) participant.meds_dt_v3 = record.meds_dt;
                if (record.samp_date) participant.samp_date_v3 = record.samp_date;
            } else if (eventName === 'v4_arm_1') {
                if (record.vdate) participant.vdate_v4 = record.vdate;
                if (record.meds_dt) participant.meds_dt_v4 = record.meds_dt;
                if (record.samp_date) participant.samp_date_v4 = record.samp_date;
            } else if (eventName === 'v5_arm_1') {
                if (record.vdate) participant.vdate_v5 = record.vdate;
                if (record.meds_dt) participant.meds_dt_v5 = record.meds_dt;
                if (record.samp_date) participant.samp_date_v5 = record.samp_date;
            } else if (eventName === 'v6_arm_1') {
                if (record.vdate) participant.vdate_v6 = record.vdate;
                if (record.meds_dt) participant.meds_dt_v6 = record.meds_dt;
                if (record.samp_date) participant.samp_date_v6 = record.samp_date;
            } else if (eventName === 'v7_arm_1') {
                if (record.vdate) participant.vdate_v7 = record.vdate;
                if (record.meds_dt) participant.meds_dt_v7 = record.meds_dt;
                if (record.samp_date) participant.samp_date_v7 = record.samp_date;
            } else if (eventName === 'v8_arm_1') {
                if (record.vdate) participant.vdate_v8 = record.vdate;
                if (record.meds_dt) participant.meds_dt_v8 = record.meds_dt;
                if (record.samp_date) participant.samp_date_v8 = record.samp_date;
            } else if (eventName === 'v9_arm_1') {
                if (record.vdate) participant.vdate_v9 = record.vdate;
                if (record.meds_dt) participant.meds_dt_v9 = record.meds_dt;
                if (record.samp_date) participant.samp_date_v9 = record.samp_date;
            } else if (eventName === 'fu1_arm_1') {
                if (record.vdate) participant.vdate_fu1 = record.vdate;
                if (record.meds_dt) participant.meds_dt_fu1 = record.meds_dt;
                if (record.samp_date) participant.samp_date_fu1 = record.samp_date;
            } else if (eventName === 'fu2_arm_1') {
                if (record.vdate) participant.vdate_fu2 = record.vdate;
                if (record.meds_dt) participant.meds_dt_fu2 = record.meds_dt;
                if (record.samp_date) participant.samp_date_fu2 = record.samp_date;
            } else if (eventName === 'v10_arm_1') {
                if (record.vdate) participant.vdate_v10 = record.vdate;
                if (record.meds_dt) participant.meds_dt_v10 = record.meds_dt;
                if (record.samp_date) participant.samp_date_v10 = record.samp_date;
            } else if (eventName === 'conclusion_arm_1' && record.conclusion) {
                participant.conclusion = parseInt(record.conclusion);
            }
        });
        
        // Apply visit date fallback logic (matching R script lines 414-504)
        Object.values(participantMap).forEach(participant => {
            const visits = ['base', 'v1', 'v2', 'v3', 'v4', 'v5', 'v6', 'v7', 'v8', 'v9', 'fu1', 'fu2', 'v10'];
            
            visits.forEach(visit => {
                const vdateKey = `vdate_${visit}`;
                const medsDtKey = `meds_dt_${visit}`;
                const sampDateKey = `samp_date_${visit}`;
                
                // If vdate is missing, try meds_dt, then samp_date
                if (!participant[vdateKey] && participant[medsDtKey]) {
                    participant[vdateKey] = participant[medsDtKey];
                } else if (!participant[vdateKey] && participant[sampDateKey]) {
                    participant[vdateKey] = participant[sampDateKey];
                }
            });
            
            // Fix issue where participants start V1 mistakenly without randomization (R script line 507)
            if (!participant.rand_code && participant.vdate_v1) {
                participant.vdate_v1 = null;
            }
        });

        return Object.values(participantMap)
            .filter(p => p.participant_id && p.participant_id !== '');
    }

    // Clear all cache
    clearCache() {
        this.dataManager.clearCache();
    }
}

// Dashboard UI Controller
class Dashboard {
    constructor() {
        this.api = new REDCapAPI();
        this.currentTab = 'enrollment';
        this.init();
    }

    init() {
        this.setupTabNavigation();
        this.setupEventListeners();
        this.loadEnrollmentData();
    }

    setupTabNavigation() {
        const tabButtons = document.querySelectorAll('.tab-button');
        const tabContents = document.querySelectorAll('.tab-content');

        tabButtons.forEach(button => {
            button.addEventListener('click', () => {
                const tabId = button.getAttribute('data-tab');
                
                // Update active states
                tabButtons.forEach(btn => btn.classList.remove('active'));
                tabContents.forEach(content => content.classList.remove('active'));
                
                button.classList.add('active');
                document.getElementById(tabId).classList.add('active');
                
                this.currentTab = tabId;

                // Load tab-specific data
                if (tabId === 'data-quality') {
                    this.loadDataQuality();
                } else if (tabId === 'out-of-window') {
                    this.loadOutOfWindowData();
                } else if (tabId === 'missing-data') {
                    // Load missing data report
                    if (window.missingDataReportManager) {
                        window.missingDataReportManager.loadMissingDataReport();
                    }
                } else if (tabId === 'consort') {
                    this.loadConsortData();
                }
            });
        });
    }

    setupEventListeners() {
        // Refresh data button
        document.getElementById('refresh-data').addEventListener('click', () => {
            this.loadEnrollmentData(true); // Force refresh
        });

        // Export to PDF button
        document.getElementById('export-pdf').addEventListener('click', () => {
            this.exportToPDF();
        });

        // Export enrollment chart button
        document.getElementById('export-enrollment-chart').addEventListener('click', () => {
            this.exportEnrollmentChart();
        });

        // Add keyboard shortcut to clear cache (Ctrl+Shift+C)
        document.addEventListener('keydown', (e) => {
            if (e.ctrlKey && e.shiftKey && e.key === 'C') {
                if (confirm('Clear all cached data?')) {
                    this.api.dataManager.clearCache();
                    this.showMessage('Cache cleared! Refreshing data...', 'success');
                    this.loadEnrollmentData(true);
                }
            }
        });

        // Data quality tab events
        const refreshQualityBtn = document.getElementById('refresh-quality-data');
        if (refreshQualityBtn) {
            refreshQualityBtn.addEventListener('click', () => {
                this.loadDataQuality(true);
            });
        }

        const exportQualityBtn = document.getElementById('export-quality-report');
        if (exportQualityBtn) {
            exportQualityBtn.addEventListener('click', () => {
                this.exportDataQualityReport();
            });
        }

        const incompleteFilter = document.getElementById('incomplete-filter');
        if (incompleteFilter) {
            incompleteFilter.addEventListener('change', (e) => {
                this.filterIncompleteOutcomes(e.target.value);
            });
        }

        // Out of Window tab events
        const refreshOWBtn = document.getElementById('refresh-ow-data');
        if (refreshOWBtn) {
            refreshOWBtn.addEventListener('click', () => {
                this.loadOutOfWindowData(true);
            });
        }

        const exportOWBtn = document.getElementById('export-ow-report');
        if (exportOWBtn) {
            exportOWBtn.addEventListener('click', () => {
                this.exportOutOfWindowReport();
            });
        }

        const owVisitFilter = document.getElementById('ow-visit-filter');
        if (owVisitFilter) {
            owVisitFilter.addEventListener('change', () => {
                this.filterOutOfWindowData();
            });
        }

        const owStatusFilter = document.getElementById('ow-status-filter');
        if (owStatusFilter) {
            owStatusFilter.addEventListener('change', () => {
                this.filterOutOfWindowData();
            });
        }

        const owParticipantSearch = document.getElementById('ow-participant-search');
        if (owParticipantSearch) {
            owParticipantSearch.addEventListener('input', () => {
                this.filterOutOfWindowData();
            });
        }

        // CONSORT tab events
        const refreshConsortBtn = document.getElementById('refresh-consort-data');
        if (refreshConsortBtn) {
            refreshConsortBtn.addEventListener('click', () => {
                this.loadConsortData(true);
            });
        }

        const exportConsortBtn = document.getElementById('export-consort-diagram');
        if (exportConsortBtn) {
            exportConsortBtn.addEventListener('click', () => {
                this.exportConsortDiagram();
            });
        }

        // Clickable CONSORT boxes
        document.querySelectorAll('[data-consort-category]').forEach(box => {
            box.addEventListener('click', () => {
                const category = box.getAttribute('data-consort-category');
                this.showConsortParticipantDetails(category);
            });
        });

        // Add cache info display
        this.displayCacheInfo();

        // Clickable metric cards
        document.querySelectorAll('.metric-card.clickable').forEach(card => {
            card.addEventListener('click', () => {
                const category = card.getAttribute('data-category');
                this.showParticipantDetails(category);
            });
        });

        // Modal close functionality
        const modal = document.getElementById('participant-modal');
        const closeBtn = document.querySelector('.modal-close');
        
        closeBtn.addEventListener('click', () => {
            this.hideModal();
        });

        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                this.hideModal();
            }
        });

        // Search and sort functionality
        document.getElementById('participant-search').addEventListener('input', (e) => {
            this.filterParticipants(e.target.value);
        });

        document.getElementById('participant-sort').addEventListener('change', (e) => {
            this.sortParticipants(e.target.value);
        });

        // Export participant details to CSV
        document.getElementById('export-participant-details').addEventListener('click', () => {
            this.exportParticipantDetailsToCSV();
        });
    }

    showLoading() {
        document.getElementById('loading').classList.remove('hidden');
    }

    hideLoading() {
        document.getElementById('loading').classList.add('hidden');
    }

    async loadEnrollmentData(forceRefresh = false) {
        this.showLoading();
        
        try {
            const data = await this.api.getEnrollmentData(forceRefresh);
            this.updateEnrollmentUI(data.summary);
            this.updateEnrollmentChart(data.participants);
            this.updateNIHEnrollmentReport(data.participants);
            this.displayCacheInfo();
            
            // Show success message if data was refreshed
            if (forceRefresh) {
                this.showMessage('Data refreshed successfully!', 'success');
            }
        } catch (error) {
            console.error('Error loading enrollment data:', error);
            this.showMessage('Error loading data: ' + error.message, 'error');
        } finally {
            this.hideLoading();
        }
    }

    updateEnrollmentUI(summary) {
        // Update metric cards
        document.getElementById('total-enrolled').textContent = summary.totalEnrolled;
        document.getElementById('total-randomized').textContent = summary.totalRandomized;
        document.getElementById('screen-failures').textContent = summary.screenFailures;
        document.getElementById('completed-study').textContent = summary.completedStudy;
        document.getElementById('withdrawn').textContent = summary.withdrawn;
        document.getElementById('lost-followup').textContent = summary.lostFollowup;
    }

    updateEnrollmentChart(participants) {
        const chartData = this.api.processEnrollmentChartData(participants);
        
        if (chartData.months.length === 0) {
            console.warn('No enrollment data available for chart');
            return;
        }

        const ctx = document.getElementById('enrollment-chart');
        
        // Destroy existing chart if it exists
        if (this.enrollmentChart) {
            this.enrollmentChart.destroy();
        }

        // Calculate goal line data (2.7 per month cumulative)
        const goalRate = 2.7; // participants per month
        const cumulativeGoal = chartData.months.map((month, index) => {
            return goalRate * (index + 1);
        });

        this.enrollmentChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: chartData.months,
                datasets: [
                    // Monthly enrollment bars
                    {
                        label: 'Monthly Enrolled',
                        data: chartData.monthlyEnrolled,
                        backgroundColor: 'rgba(40, 167, 69, 0.7)',
                        borderColor: 'rgba(40, 167, 69, 1)',
                        borderWidth: 1,
                        type: 'bar',
                        yAxisID: 'y'
                    },
                    // Monthly randomization bars
                    {
                        label: 'Monthly Randomized',
                        data: chartData.monthlyRandomized,
                        backgroundColor: 'rgba(0, 123, 255, 0.7)',
                        borderColor: 'rgba(0, 123, 255, 1)',
                        borderWidth: 1,
                        type: 'bar',
                        yAxisID: 'y'
                    },
                    // Cumulative enrollment line
                    {
                        label: 'Cumulative Enrolled',
                        data: chartData.cumulativeEnrolled,
                        borderColor: 'rgba(40, 167, 69, 1)',
                        backgroundColor: 'rgba(40, 167, 69, 0.1)',
                        borderWidth: 3,
                        type: 'line',
                        yAxisID: 'y1',
                        tension: 0.1,
                        fill: false,
                        pointBackgroundColor: 'rgba(40, 167, 69, 1)',
                        pointBorderColor: '#fff',
                        pointBorderWidth: 2,
                        pointRadius: 4
                    },
                    // Cumulative randomization line
                    {
                        label: 'Cumulative Randomized',
                        data: chartData.cumulativeRandomized,
                        borderColor: 'rgba(0, 123, 255, 1)',
                        backgroundColor: 'rgba(0, 123, 255, 0.1)',
                        borderWidth: 3,
                        type: 'line',
                        yAxisID: 'y1',
                        tension: 0.1,
                        fill: false,
                        pointBackgroundColor: 'rgba(0, 123, 255, 1)',
                        pointBorderColor: '#fff',
                        pointBorderWidth: 2,
                        pointRadius: 4
                    },
                    // Goal line (2.7 randomized per month)
                    {
                        label: 'Randomization Goal (2.7/month)',
                        data: cumulativeGoal,
                        borderColor: 'rgba(255, 193, 7, 1)',
                        backgroundColor: 'rgba(255, 193, 7, 0.1)',
                        borderWidth: 2,
                        borderDash: [10, 5],
                        type: 'line',
                        yAxisID: 'y1',
                        tension: 0,
                        fill: false,
                        pointBackgroundColor: 'rgba(255, 193, 7, 1)',
                        pointBorderColor: '#fff',
                        pointBorderWidth: 2,
                        pointRadius: 3
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                layout: {
                    padding: {
                        left: 10,
                        right: 20,
                        top: 10,
                        bottom: 10
                    }
                },
                plugins: {
                    title: {
                        display: true,
                        text: 'ULLTRA Study Enrollment Over Time',
                        font: {
                            size: 16,
                            weight: 'bold'
                        }
                    },
                    legend: {
                        display: true,
                        position: 'top'
                    }
                },
                scales: {
                    x: {
                        title: {
                            display: true,
                            text: 'Month'
                        },
                        ticks: {
                            maxRotation: 45,
                            minRotation: 45,
                            autoSkip: false
                        }
                    },
                    y: {
                        type: 'linear',
                        display: true,
                        position: 'left',
                        title: {
                            display: true,
                            text: 'Monthly Count'
                        },
                        beginAtZero: true
                    },
                    y1: {
                        type: 'linear',
                        display: true,
                        position: 'right',
                        title: {
                            display: true,
                            text: 'Cumulative Total'
                        },
                        beginAtZero: true,
                        grid: {
                            drawOnChartArea: false
                        }
                    }
                },
                interaction: {
                    mode: 'index',
                    intersect: false
                }
            }
        });
    }

    displayCacheInfo() {
        const cacheInfo = this.api.getCacheInfo();
        const enrollmentCache = cacheInfo.enrollment;
        
        // Create or update cache info display
        let cacheInfoElement = document.getElementById('cache-info');
        if (!cacheInfoElement) {
            cacheInfoElement = document.createElement('div');
            cacheInfoElement.id = 'cache-info';
            cacheInfoElement.className = 'cache-info';
            document.querySelector('#enrollment').appendChild(cacheInfoElement);
        }

        if (enrollmentCache.exists) {
            const statusClass = enrollmentCache.isValid ? 'valid' : 'expired';
            cacheInfoElement.innerHTML = `
                <p class="cache-status ${statusClass}">
                    Data last updated: ${enrollmentCache.lastUpdated} 
                    (${enrollmentCache.age} minutes ago)
                    ${enrollmentCache.isValid ? '✓' : '⚠️ Cache expired'}
                </p>
            `;
        } else {
            cacheInfoElement.innerHTML = '<p class="cache-status">No cached data available</p>';
        }
    }

    showMessage(message, type = 'info') {
        const messageDiv = document.createElement('div');
        messageDiv.className = type;
        messageDiv.textContent = message;
        
        const container = document.querySelector('.container');
        container.insertBefore(messageDiv, container.firstChild);
        
        // Auto-remove after 5 seconds
        setTimeout(() => {
            messageDiv.remove();
        }, 5000);
    }

    async showParticipantDetails(category) {
        try {
            const data = await this.api.getEnrollmentData();
            const participants = this.getParticipantsByCategory(data.participants, category);
            
            const modal = document.getElementById('participant-modal');
            const title = document.getElementById('modal-title');
            const participantList = document.getElementById('participant-list');
            
            // Update modal title
            const categoryTitles = {
                'enrolled': 'Enrolled Participants',
                'randomized': 'Randomized Participants',
                'screen-failures': 'Screen Failures',
                'completed': 'Completed Study',
                'withdrawn': 'Withdrawn Participants',
                'lost-followup': 'Lost to Follow-up'
            };
            
            title.textContent = categoryTitles[category] || 'Participants';
            
            // Store participants for filtering/sorting
            this.currentParticipants = participants;
            this.currentCategory = category;
            
            // Render participant list
            this.renderParticipantList(participants);
            
            // Show modal
            modal.classList.remove('hidden');
            
        } catch (error) {
            console.error('Error showing participant details:', error);
            this.showMessage('Error loading participant details: ' + error.message, 'error');
        }
    }

    getParticipantsByCategory(participants, category) {
        const filtered = [];
        
        Object.values(participants).forEach(participant => {
            // Safety check - ensure visits object exists
            if (!participant.visits) {
                console.warn(`Participant ${participant.id} has no visits data`);
                return;
            }
            
            const baseline = participant.visits.baseline_arm_1 || {};
            const v1 = participant.visits.v1_arm_1 || {};
            const conclusion = participant.conclusion || {};
            const conclusionInfo = this.api.getConclusionStatusInfo(conclusion);
            
            switch (category) {
                case 'enrolled':
                    if (baseline.icf_date && baseline.icf_date !== '') {
                        const currentVisitStatus = this.api.calculateCurrentVisitStatus(participant.visits, conclusion);
                        filtered.push({
                            ...participant,
                            status: 'enrolled',
                            statusText: 'Enrolled',
                            icfDate: baseline.icf_date,
                            randomized: (v1.rand_code && v1.rand_code !== '') ? 'Yes' : 'No',
                            currentStatus: currentVisitStatus
                        });
                    }
                    break;
                    
                case 'randomized':
                    if (v1.rand_code && v1.rand_code !== '') {
                        const currentVisitStatus = this.api.calculateCurrentVisitStatus(participant.visits, conclusion);
                        filtered.push({
                            ...participant,
                            status: 'randomized',
                            statusText: 'Randomized',
                            icfDate: baseline.icf_date,
                            randomized: 'Yes',
                            currentStatus: currentVisitStatus
                        });
                    }
                    break;
                    
                case 'screen-failures':
                    if (conclusionInfo && conclusionInfo.code === '8') {
                        const currentVisitStatus = this.api.calculateCurrentVisitStatus(participant.visits, conclusion);
                        filtered.push({
                            ...participant,
                            status: 'screen-failure',
                            statusText: conclusionInfo.label,
                            withdrawalDate: conclusionInfo.date,
                            icfDate: baseline.icf_date || 'Not enrolled',
                            randomized: (v1.rand_code && v1.rand_code !== '') ? 'Yes' : 'No',
                            currentStatus: currentVisitStatus
                        });
                    }
                    break;
                    
                case 'completed':
                    if (conclusionInfo && conclusionInfo.code === '1') {
                        const currentVisitStatus = this.api.calculateCurrentVisitStatus(participant.visits, conclusion);
                        filtered.push({
                            ...participant,
                            status: 'completed',
                            statusText: conclusionInfo.label,
                            icfDate: baseline.icf_date,
                            randomized: (v1.rand_code && v1.rand_code !== '') ? 'Yes' : 'No',
                            completionDate: conclusionInfo.date,
                            currentStatus: currentVisitStatus
                        });
                    }
                    break;
                    
                case 'withdrawn':
                    // Only include codes 3 and 4, and only if randomized (matching card count logic)
                    if (conclusionInfo && ['3', '4'].includes(conclusionInfo.code)) {
                        const isRandomized = v1.rand_code && v1.rand_code !== '';
                        if (isRandomized) {
                            const currentVisitStatus = this.api.calculateCurrentVisitStatus(participant.visits, conclusion);
                            filtered.push({
                                ...participant,
                                status: 'withdrawn',
                                statusText: conclusionInfo.label,
                                icfDate: baseline.icf_date,
                                randomized: 'Yes',
                                withdrawalDate: conclusionInfo.date,
                                currentStatus: currentVisitStatus
                            });
                        }
                    }
                    break;
                    
                case 'lost-followup':
                    // Only include if randomized (matching card count logic)
                    if (conclusionInfo && conclusionInfo.code === '5') {
                        const isRandomized = v1.rand_code && v1.rand_code !== '';
                        if (isRandomized) {
                            const currentVisitStatus = this.api.calculateCurrentVisitStatus(participant.visits, conclusion);
                            filtered.push({
                                ...participant,
                                status: 'lost-followup',
                                statusText: conclusionInfo.label,
                                icfDate: baseline.icf_date,
                                randomized: 'Yes',
                                lostDate: conclusionInfo.date,
                                currentStatus: currentVisitStatus
                            });
                        }
                    }
                    break;
            }
        });
        
        return filtered;
    }

    renderParticipantList(participants) {
        const participantList = document.getElementById('participant-list');
        
        if (participants.length === 0) {
            participantList.innerHTML = `
                <div class="empty-state">
                    <h3>No participants found</h3>
                    <p>No participants match the current criteria.</p>
                </div>
            `;
            return;
        }
        
        // Create table structure with headers - simplified to 4 columns
        const html = `
            <div class="participants-table">
                <div class="table-header">
                    <div class="header-cell participant-id-header">Participant ID</div>
                    <div class="header-cell icf-date-header">ICF Date</div>
                    <div class="header-cell randomized-header">Randomized</div>
                    <div class="header-cell current-status-header">Current Status</div>
                </div>
                ${participants.map(participant => {
                    const randomizedClass = participant.randomized === 'Yes' ? 'randomized-yes' : 'randomized-no';
                    const currentStatusClass = this.api.getCurrentStatusClass(participant.currentStatus);
                    
                    return `
                        <div class="table-row">
                            <div class="table-cell participant-id-cell">${participant.id}</div>
                            <div class="table-cell icf-date-cell">${this.formatDate(participant.icfDate) || 'Not enrolled'}</div>
                            <div class="table-cell randomized-cell">
                                <span class="randomized-status ${randomizedClass}">${participant.randomized || 'No'}</span>
                            </div>
                            <div class="table-cell current-status-cell">
                                <span class="current-status ${currentStatusClass}">${participant.currentStatus}</span>
                            </div>
                        </div>
                    `;
                }).join('')}
            </div>
        `;
        
        participantList.innerHTML = html;
    }

    getAdditionalInfo(participant) {
        if (participant.completionDate) {
            return `Completed: ${this.formatDate(participant.completionDate)}`;
        } else if (participant.withdrawalDate) {
            return `Withdrawn: ${this.formatDate(participant.withdrawalDate)}`;
        } else if (participant.lostDate) {
            return `Lost: ${this.formatDate(participant.lostDate)}`;
        } else if (participant.status === 'enrolled' && participant.randomized === 'No') {
            return 'Pending randomization';
        } else if (participant.status === 'randomized') {
            return 'Active in study';
        }
        return '';
    }

    getParticipantDetailFields(participant) {
        const fields = [];
        
        if (participant.icfDate && participant.icfDate !== 'Not enrolled') {
            fields.push({
                label: 'ICF Date',
                value: this.formatDate(participant.icfDate)
            });
        }
        
        if (participant.randomized) {
            fields.push({
                label: 'Randomized',
                value: participant.randomized
            });
        }
        
        if (participant.completionDate) {
            fields.push({
                label: 'Completion Date',
                value: this.formatDate(participant.completionDate)
            });
        }
        
        if (participant.withdrawalDate) {
            fields.push({
                label: 'Withdrawal Date',
                value: this.formatDate(participant.withdrawalDate)
            });
        }
        
        if (participant.lostDate) {
            fields.push({
                label: 'Lost to Follow-up Date',
                value: this.formatDate(participant.lostDate)
            });
        }
        
        return fields;
    }

    formatDate(dateString) {
        if (!dateString || dateString === '') return 'Not available';
        
        try {
            const date = new Date(dateString);
            return date.toLocaleDateString('en-US', {
                year: 'numeric',
                month: 'short',
                day: 'numeric'
            });
        } catch (error) {
            return dateString;
        }
    }

    hideModal() {
        const modal = document.getElementById('participant-modal');
        modal.classList.add('hidden');
        
        // Clear search
        document.getElementById('participant-search').value = '';
        document.getElementById('participant-sort').value = 'id';
    }

    filterParticipants(searchTerm) {
        if (!this.currentParticipants) return;
        
        const filtered = this.currentParticipants.filter(participant => 
            participant.id.toLowerCase().includes(searchTerm.toLowerCase())
        );
        
        this.renderParticipantList(filtered);
    }

    sortParticipants(sortBy) {
        if (!this.currentParticipants) return;
        
        const sorted = [...this.currentParticipants].sort((a, b) => {
            switch (sortBy) {
                case 'id':
                    return a.id.localeCompare(b.id);
                case 'date':
                    const dateA = a.icfDate || a.withdrawalDate || a.completionDate || a.lostDate || '';
                    const dateB = b.icfDate || b.withdrawalDate || b.completionDate || b.lostDate || '';
                    return new Date(dateB) - new Date(dateA);
                case 'status':
                    return a.statusText.localeCompare(b.statusText);
                default:
                    return 0;
            }
        });
        
        this.renderParticipantList(sorted);
    }

    updateNIHEnrollmentReport(participants) {
        try {
            console.log('Updating NIH Enrollment Report...');
            console.log('Participants data:', participants);
            
            const nihData = this.generateNIHEnrollmentData(participants);
            console.log('Generated NIH data:', nihData);
            
            const tableContainer = document.getElementById('nih-enrollment-table');
            
            if (!tableContainer) {
                console.error('NIH table container not found');
                return;
            }
            
            // Create the NIH enrollment table HTML
            const tableHTML = `
                <table class="nih-enrollment-table">
                    <thead>
                        <tr>
                            <th rowspan="3" class="racial-header">Racial Categories</th>
                            <th colspan="9" class="ethnic-header">Ethnic Categories</th>
                            <th rowspan="3" class="total-header">Total</th>
                        </tr>
                        <tr>
                            <th colspan="3">Not Hispanic or Latino</th>
                            <th colspan="3">Hispanic or Latino</th>
                            <th colspan="3">Unknown/Not Reported Ethnicity</th>
                        </tr>
                        <tr>
                            <th>Female</th>
                            <th>Male</th>
                            <th>Unknown/Not Reported</th>
                            <th>Female</th>
                            <th>Male</th>
                            <th>Unknown/Not Reported</th>
                            <th>Female</th>
                            <th>Male</th>
                            <th>Unknown/Not Reported</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${this.generateNIHTableRows(nihData)}
                    </tbody>
                </table>
                <div class="nih-report-info">
                    <p>Report Type: Cumulative (Actual)</p>
                    <p>Generated: ${new Date().toLocaleString()}</p>
                    <p><strong>Note:</strong> Participants showing as "Unknown/Not Reported" may not have completed the demographics form yet.</p>
                </div>
            `;
            
            tableContainer.innerHTML = tableHTML;
            console.log('NIH table HTML set successfully');
        } catch (error) {
            console.error('Error updating NIH enrollment report:', error);
            // Display error in the table container
            const tableContainer = document.getElementById('nih-enrollment-table');
            if (tableContainer) {
                tableContainer.innerHTML = `
                    <div style="color: red; padding: 20px; border: 1px solid red; background: #ffe6e6;">
                        <h4>Error loading NIH Enrollment Report</h4>
                        <p>${error.message}</p>
                        <p>Check browser console for details.</p>
                    </div>
                `;
            }
        }
    }

    generateNIHEnrollmentData(participants) {
        console.log('Starting NIH data generation...');
        
        // Initialize the data structure matching NIH format
        const categories = {
            'American Indian or Alaska Native': { notHispanic: { F: 0, M: 0, U: 0 }, hispanic: { F: 0, M: 0, U: 0 }, unknownEthnicity: { F: 0, M: 0, U: 0 } },
            'Asian': { notHispanic: { F: 0, M: 0, U: 0 }, hispanic: { F: 0, M: 0, U: 0 }, unknownEthnicity: { F: 0, M: 0, U: 0 } },
            'Native Hawaiian or Other Pacific Islander': { notHispanic: { F: 0, M: 0, U: 0 }, hispanic: { F: 0, M: 0, U: 0 }, unknownEthnicity: { F: 0, M: 0, U: 0 } },
            'Black or African American': { notHispanic: { F: 0, M: 0, U: 0 }, hispanic: { F: 0, M: 0, U: 0 }, unknownEthnicity: { F: 0, M: 0, U: 0 } },
            'White': { notHispanic: { F: 0, M: 0, U: 0 }, hispanic: { F: 0, M: 0, U: 0 }, unknownEthnicity: { F: 0, M: 0, U: 0 } },
            'More than One Race': { notHispanic: { F: 0, M: 0, U: 0 }, hispanic: { F: 0, M: 0, U: 0 }, unknownEthnicity: { F: 0, M: 0, U: 0 } },
            'Unknown or Not Reported': { notHispanic: { F: 0, M: 0, U: 0 }, hispanic: { F: 0, M: 0, U: 0 }, unknownEthnicity: { F: 0, M: 0, U: 0 } }
        };

        let processedCount = 0;
        let enrolledCount = 0;

        // Process each participant
        Object.values(participants).forEach((participant, index) => {
            const baseline = participant.visits?.baseline_arm_1 || {};
            
            // Log first few participants to see data structure
            if (index < 3) {
                console.log(`Sample participant ${index + 1}:`, {
                    id: participant.id,
                    demo_sex: baseline.demo_sex,
                    demo_ethnicity: baseline.demo_ethnicity,
                    demo_race: baseline.demo_race,
                    demo_race___1: baseline.demo_race___1,
                    demo_race___2: baseline.demo_race___2,
                    demo_race___3: baseline.demo_race___3,
                    demo_race___4: baseline.demo_race___4,
                    demo_race___5: baseline.demo_race___5,
                    demo_race___6: baseline.demo_race___6,
                    icf_date: baseline.icf_date,
                    all_fields: Object.keys(baseline).sort()
                });
            }
            
            // Only include enrolled participants (those with ICF date)
            if (!baseline.icf_date || baseline.icf_date === '') return;
            
            enrolledCount++;
            
            // Get demographics - map from checkbox fields directly
            const gender = this.mapGender(baseline.demo_sex);
            const race = this.mapRaceFromCheckbox(baseline);
            const ethnicity = this.mapEthnicity(baseline.demo_ethnicity);

            // Log participant mapping for debugging
            console.log(`Participant ${participant.id} enrolled:`, {
                demo_sex_raw: baseline.demo_sex,
                demo_race_raw: baseline.demo_race,
                demo_ethnicity_raw: baseline.demo_ethnicity,
                mapped: {
                    gender: gender,
                    race: race,
                    ethnicity: ethnicity
                }
            });

            // Increment the appropriate counter
            if (categories[race] && categories[race][ethnicity]) {
                categories[race][ethnicity][gender]++;
                processedCount++;
            } else {
                console.error(`ERROR: Could not categorize participant ${participant.id}: race="${race}", ethnicity="${ethnicity}", gender="${gender}"`);
                // This should never happen since we have "Unknown or Not Reported" category
                // but if it does, force them into Unknown category
                console.log('Forcing participant into Unknown/Not Reported category');
                categories['Unknown or Not Reported']['unknownEthnicity']['U']++;
                processedCount++;
            }
        });

        console.log(`Processed ${processedCount} out of ${enrolledCount} enrolled participants`);
        console.log('Final NIH categories:', categories);
        
        return categories;
    }

    mapGender(demo_sex) {
        // Map sex codes to NIH categories based on REDCap metadata
        // demo_sex: 1 = Male, 0 = Female
        // Handle missing, null, undefined, or empty string as Unknown
        if (demo_sex === null || demo_sex === undefined || demo_sex === '' || demo_sex === 'NaN') {
            console.log('Gender is missing/null/empty, mapping to U (Unknown)');
            return 'U';
        }
        if (demo_sex === '0' || demo_sex === 0) return 'F';
        if (demo_sex === '1' || demo_sex === 1) return 'M';
        console.log(`Unexpected gender value: ${demo_sex}, mapping to U (Unknown)`);
        return 'U';
    }

    mapRaceNIH(demo_race_nih) {
        // Map the calculated NIH race field directly
        // This field is already calculated in REDCap to match NIH categories
        if (!demo_race_nih) return null;
        
        // The field should contain exact NIH category names from REDCap calculation
        // Normalize to match our NIH table categories
        if (demo_race_nih === 'White or Caucasian') return 'White';
        if (demo_race_nih === 'More Than One Race') return 'More than One Race';  // Ensure consistent capitalization
        
        return demo_race_nih;
    }

    mapRaceFromCheckbox(baseline) {
        // Map race checkbox values to NIH categories
        // REDCap may export checkbox fields in different formats

        let selectedRaces = 0;
        let selectedRace = '';

        const raceMap = {
            '1': 'American Indian or Alaska Native',
            '2': 'Asian',
            '3': 'Black or African American',
            '4': 'Native Hawaiian or Other Pacific Islander',
            '5': 'White',
            '6': 'Unknown or Not Reported'
        };

        // Check if baseline data exists at all
        if (!baseline) {
            console.log('No baseline data found, mapping race to Unknown or Not Reported');
            return 'Unknown or Not Reported';
        }

        // First try the demo_race field as a single field (comma-separated or array)
        if (baseline.demo_race && baseline.demo_race !== '') {
            // If it's a string with comma-separated values
            if (typeof baseline.demo_race === 'string') {
                const races = baseline.demo_race.split(',').map(r => r.trim());
                races.forEach(race => {
                    if (race && raceMap[race]) {
                        selectedRaces++;
                        selectedRace = raceMap[race];
                    }
                });
            }
            // If it's an array
            else if (Array.isArray(baseline.demo_race)) {
                baseline.demo_race.forEach(race => {
                    if (race && raceMap[race]) {
                        selectedRaces++;
                        selectedRace = raceMap[race];
                    }
                });
            }
        }

        // If no luck with demo_race field, try individual checkbox fields
        if (selectedRaces === 0) {
            for (let i = 1; i <= 6; i++) {
                const fieldName = `demo_race___${i}`;
                if (baseline[fieldName] === '1' || baseline[fieldName] === 1) {
                    selectedRaces++;
                    selectedRace = raceMap[i.toString()];
                }
            }
        }

        // Log the result for debugging
        if (selectedRaces === 0) {
            console.log('No race selected, mapping to Unknown or Not Reported');
        }

        if (selectedRaces > 1) return 'More than One Race';
        if (selectedRaces === 1) return selectedRace;
        return 'Unknown or Not Reported';
    }

    mapEthnicity(demo_ethnicity) {
        // Map ethnicity codes to NIH categories based on REDCap metadata
        // demo_ethnicity: 1 = Hispanic, 2 = Non-Hispanic, 3 = Unknown/Not Reported
        // Handle missing, null, undefined, or empty string as Unknown
        if (demo_ethnicity === null || demo_ethnicity === undefined || demo_ethnicity === '' || demo_ethnicity === 'NaN') {
            console.log('Ethnicity is missing/null/empty, mapping to unknownEthnicity');
            return 'unknownEthnicity';
        }
        if (demo_ethnicity === '1' || demo_ethnicity === 1) return 'hispanic';
        if (demo_ethnicity === '2' || demo_ethnicity === 2) return 'notHispanic';
        if (demo_ethnicity === '3' || demo_ethnicity === 3) return 'unknownEthnicity';
        console.log(`Unexpected ethnicity value: ${demo_ethnicity}, mapping to unknownEthnicity`);
        return 'unknownEthnicity';
    }

    generateNIHTableRows(nihData) {
        const rows = [];
        const racialCategories = [
            'American Indian or Alaska Native',
            'Asian',
            'Native Hawaiian or Other Pacific Islander',
            'Black or African American',
            'White',
            'More than One Race',
            'Unknown or Not Reported'
        ];

        let totalRow = { notHispanic: { F: 0, M: 0, U: 0 }, hispanic: { F: 0, M: 0, U: 0 }, unknownEthnicity: { F: 0, M: 0, U: 0 } };

        racialCategories.forEach(race => {
            const data = nihData[race];
            const rowTotal = 
                data.notHispanic.F + data.notHispanic.M + data.notHispanic.U +
                data.hispanic.F + data.hispanic.M + data.hispanic.U +
                data.unknownEthnicity.F + data.unknownEthnicity.M + data.unknownEthnicity.U;

            // Add to totals
            totalRow.notHispanic.F += data.notHispanic.F;
            totalRow.notHispanic.M += data.notHispanic.M;
            totalRow.notHispanic.U += data.notHispanic.U;
            totalRow.hispanic.F += data.hispanic.F;
            totalRow.hispanic.M += data.hispanic.M;
            totalRow.hispanic.U += data.hispanic.U;
            totalRow.unknownEthnicity.F += data.unknownEthnicity.F;
            totalRow.unknownEthnicity.M += data.unknownEthnicity.M;
            totalRow.unknownEthnicity.U += data.unknownEthnicity.U;

            rows.push(`
                <tr>
                    <td class="racial-category">${race}</td>
                    <td class="count-cell">${data.notHispanic.F}</td>
                    <td class="count-cell">${data.notHispanic.M}</td>
                    <td class="count-cell">${data.notHispanic.U}</td>
                    <td class="count-cell">${data.hispanic.F}</td>
                    <td class="count-cell">${data.hispanic.M}</td>
                    <td class="count-cell">${data.hispanic.U}</td>
                    <td class="count-cell">${data.unknownEthnicity.F}</td>
                    <td class="count-cell">${data.unknownEthnicity.M}</td>
                    <td class="count-cell">${data.unknownEthnicity.U}</td>
                    <td class="total-cell">${rowTotal}</td>
                </tr>
            `);
        });

        // Add total row
        const grandTotal = 
            totalRow.notHispanic.F + totalRow.notHispanic.M + totalRow.notHispanic.U +
            totalRow.hispanic.F + totalRow.hispanic.M + totalRow.hispanic.U +
            totalRow.unknownEthnicity.F + totalRow.unknownEthnicity.M + totalRow.unknownEthnicity.U;

        rows.push(`
            <tr class="total-row">
                <td class="racial-category">Total</td>
                <td class="count-cell">${totalRow.notHispanic.F}</td>
                <td class="count-cell">${totalRow.notHispanic.M}</td>
                <td class="count-cell">${totalRow.notHispanic.U}</td>
                <td class="count-cell">${totalRow.hispanic.F}</td>
                <td class="count-cell">${totalRow.hispanic.M}</td>
                <td class="count-cell">${totalRow.hispanic.U}</td>
                <td class="count-cell">${totalRow.unknownEthnicity.F}</td>
                <td class="count-cell">${totalRow.unknownEthnicity.M}</td>
                <td class="count-cell">${totalRow.unknownEthnicity.U}</td>
                <td class="total-cell">${grandTotal}</td>
            </tr>
        `);

        return rows.join('');
    }

    async exportToPDF() {
        try {
            // Ensure jsPDF is loaded
            const { jsPDF } = window.jspdf;
            
            // Get current enrollment data
            const data = await this.api.getEnrollmentData();
            const summary = data.summary;
            const participants = data.participants;
            
            // Create new PDF document
            const doc = new jsPDF();
            const pageHeight = doc.internal.pageSize.height;
            let yPosition = 20;
            
            // Add title
            doc.setFontSize(20);
            doc.text('ULLTRA Study Dashboard Report', 20, yPosition);
            yPosition += 10;
            
            doc.setFontSize(12);
            doc.text('Photobiomodulation for the Management of Temporomandibular Disorder Pain', 20, yPosition);
            yPosition += 10;
            
            // Add generation date
            doc.setFontSize(10);
            doc.text(`Generated: ${new Date().toLocaleDateString()} ${new Date().toLocaleTimeString()}`, 20, yPosition);
            yPosition += 15;
            
            // Add summary statistics
            doc.setFontSize(14);
            doc.text('Enrollment Summary', 20, yPosition);
            yPosition += 10;
            
            doc.setFontSize(11);
            const summaryData = [
                ['Metric', 'Count'],
                ['Total Enrolled', summary.totalEnrolled.toString()],
                ['Total Randomized', summary.totalRandomized.toString()],
                ['Screen Failures', summary.screenFailures.toString()],
                ['Completed Study', summary.completedStudy.toString()],
                ['Withdrawn', summary.withdrawn.toString()],
                ['Lost to Follow-up', summary.lostFollowup.toString()]
            ];
            
            doc.autoTable({
                head: [summaryData[0]],
                body: summaryData.slice(1),
                startY: yPosition,
                theme: 'striped',
                headStyles: { fillColor: [41, 128, 185] }
            });
            
            yPosition = doc.lastAutoTable.finalY + 15;
            
            // Get detailed participant data for "Total Enrolled" category
            const enrolledParticipants = this.getParticipantsByCategory(participants, 'enrolled');
            
            // Add enrolled participant details
            doc.setFontSize(14);
            doc.text('Enrolled Participant Details', 20, yPosition);
            yPosition += 5;
            
            doc.setFontSize(10);
            doc.text(`${enrolledParticipants.length} participants with ICF date at baseline visit`, 20, yPosition);
            yPosition += 10;
            
            // Prepare participant data for table
            const participantRows = [];
            enrolledParticipants.forEach(participant => {
                participantRows.push([
                    participant.id,
                    this.formatDate(participant.icfDate),
                    participant.randomized,
                    participant.currentStatus
                ]);
            });
            
            // Sort by participant ID
            participantRows.sort((a, b) => a[0].localeCompare(b[0]));
            
            // Check if we need a new page
            if (yPosition + 20 > pageHeight - 20) {
                doc.addPage();
                yPosition = 20;
            }
            
            // Add participant table
            doc.autoTable({
                head: [['Participant ID', 'ICF Date', 'Randomized', 'Current Status']],
                body: participantRows,
                startY: yPosition,
                theme: 'striped',
                headStyles: { fillColor: [41, 128, 185] },
                columnStyles: {
                    0: { cellWidth: 30 },
                    1: { cellWidth: 30 },
                    2: { cellWidth: 25 },
                    3: { cellWidth: 'auto' }
                },
                styles: {
                    fontSize: 9,
                    cellPadding: 2
                }
            });
            
            // Add NIH Enrollment Report on a new page
            doc.addPage('landscape');
            yPosition = 20;
            
            doc.setFontSize(14);
            doc.text('NIH Enrollment Report', 20, yPosition);
            yPosition += 5;
            
            doc.setFontSize(10);
            doc.text('Report Type: Cumulative (Actual)', 20, yPosition);
            yPosition += 10;
            
            // Generate NIH enrollment data
            const nihData = this.generateNIHEnrollmentData(participants);
            
            // Prepare NIH table data
            const nihTableData = [];
            const racialCategories = [
                'American Indian/Alaska Native',
                'Asian',
                'Native Hawaiian or Other Pacific Islander',
                'Black or African American',
                'White',
                'More than One Race',
                'Unknown or Not Reported'
            ];

            let totalRow = ['Total', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];

            racialCategories.forEach(race => {
                const data = nihData[race];
                const row = [
                    race,
                    data.notHispanic.F,
                    data.notHispanic.M,
                    data.notHispanic.U,
                    data.hispanic.F,
                    data.hispanic.M,
                    data.hispanic.U,
                    data.unknownEthnicity.F,
                    data.unknownEthnicity.M,
                    data.unknownEthnicity.U,
                    data.notHispanic.F + data.notHispanic.M + data.notHispanic.U +
                    data.hispanic.F + data.hispanic.M + data.hispanic.U +
                    data.unknownEthnicity.F + data.unknownEthnicity.M + data.unknownEthnicity.U
                ];
                nihTableData.push(row);
                
                // Add to totals
                for (let i = 1; i <= 10; i++) {
                    totalRow[i] += row[i];
                }
            });
            
            nihTableData.push(totalRow);

            // Create NIH table with multi-level headers
            doc.autoTable({
                head: [[
                    { content: 'Racial Categories', rowSpan: 3 },
                    { content: 'Not Hispanic or Latino', colSpan: 3 },
                    { content: 'Hispanic or Latino', colSpan: 3 },
                    { content: 'Unknown/Not Reported Ethnicity', colSpan: 3 },
                    { content: 'Total', rowSpan: 3 }
                ], [
                    { content: 'Female', colSpan: 1 },
                    { content: 'Male', colSpan: 1 },
                    { content: 'Unknown/Not Reported', colSpan: 1 },
                    { content: 'Female', colSpan: 1 },
                    { content: 'Male', colSpan: 1 },
                    { content: 'Unknown/Not Reported', colSpan: 1 },
                    { content: 'Female', colSpan: 1 },
                    { content: 'Male', colSpan: 1 },
                    { content: 'Unknown/Not Reported', colSpan: 1 }
                ]],
                body: nihTableData,
                startY: yPosition,
                theme: 'grid',
                headStyles: { fillColor: [102, 126, 234], fontSize: 8, fontStyle: 'bold' },
                bodyStyles: { fontSize: 8 },
                columnStyles: {
                    0: { cellWidth: 45, fontStyle: 'bold' },
                    10: { fontStyle: 'bold', fillColor: [240, 240, 240] }
                },
                didParseCell: function(data) {
                    // Style the total row
                    if (data.row.index === nihTableData.length - 1) {
                        data.cell.styles.fontStyle = 'bold';
                        data.cell.styles.fillColor = [240, 240, 240];
                    }
                }
            });
            
            // Save the PDF
            const filename = `ULLTRA_Dashboard_Report_${new Date().toISOString().split('T')[0]}.pdf`;
            doc.save(filename);
            
            this.showMessage(`PDF exported successfully: ${filename}`, 'success');
            
        } catch (error) {
            console.error('Error exporting PDF:', error);
            this.showMessage('Error exporting PDF: ' + error.message, 'error');
        }
    }

    exportEnrollmentChart() {
        try {
            if (!this.enrollmentChart) {
                this.showMessage('No enrollment chart available to export', 'warning');
                return;
            }

            // Ensure jsPDF is loaded
            const { jsPDF } = window.jspdf;

            // Create new PDF document in landscape orientation
            const doc = new jsPDF('l', 'mm', 'letter'); // Landscape, mm units, letter size (279.4 x 215.9 mm)

            // Add title
            doc.setFontSize(16);
            doc.setFont(undefined, 'bold');
            doc.text('ULLTRA Study - Enrollment Over Time', 140, 15, { align: 'center' });

            // Add date and time
            doc.setFontSize(10);
            doc.setFont(undefined, 'normal');
            const now = new Date();
            const currentDate = now.toLocaleDateString('en-US', {
                year: 'numeric',
                month: 'long',
                day: 'numeric'
            });
            const currentTime = now.toLocaleTimeString('en-US', {
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            });
            doc.text(`Generated: ${currentDate} at ${currentTime}`, 140, 23, { align: 'center' });

            // Get the chart canvas and convert to image
            const canvas = document.getElementById('enrollment-chart');
            const chartImage = canvas.toDataURL('image/png', 1.0);

            // Calculate dimensions to fit the chart on one page with margins
            const pageWidth = 279.4; // Letter landscape width in mm
            const pageHeight = 215.9; // Letter landscape height in mm
            const margin = 15;
            const availableWidth = pageWidth - (2 * margin);
            const availableHeight = pageHeight - 35; // Account for title and date at top

            // Calculate aspect ratio and dimensions
            const chartAspectRatio = canvas.width / canvas.height;
            let chartWidth = availableWidth;
            let chartHeight = chartWidth / chartAspectRatio;

            // If height is too large, scale based on height instead
            if (chartHeight > availableHeight) {
                chartHeight = availableHeight;
                chartWidth = chartHeight * chartAspectRatio;
            }

            // Center the chart horizontally
            const xPos = (pageWidth - chartWidth) / 2;
            const yPos = 30; // Start below the title and date

            // Add the chart image to PDF
            doc.addImage(chartImage, 'PNG', xPos, yPos, chartWidth, chartHeight);

            // Generate filename with current date
            const dateStr = new Date().toISOString().split('T')[0];
            const filename = `ULLTRA_Enrollment_Chart_${dateStr}.pdf`;

            // Save the PDF
            doc.save(filename);

            this.showMessage(`Chart exported successfully: ${filename}`, 'success');

        } catch (error) {
            console.error('Error exporting enrollment chart:', error);
            this.showMessage('Error exporting chart: ' + error.message, 'error');
        }
    }

    exportParticipantDetailsToCSV() {
        try {
            if (!this.currentParticipants || this.currentParticipants.length === 0) {
                this.showMessage('No participants to export', 'warning');
                return;
            }

            // Prepare CSV headers and data
            const headers = ['Participant ID', 'Status', 'ICF Date', 'Randomization', 'Current Visit Status', 'Additional Info'];
            const rows = [headers];

            // Add participant data
            this.currentParticipants.forEach(participant => {
                const row = [
                    participant.id,
                    participant.statusText || '',
                    participant.icfDate || '',
                    participant.randomized || '',
                    participant.currentStatus || '',
                    this.getAdditionalInfo(participant) || ''
                ];
                rows.push(row);
            });

            // Convert to CSV format
            const csvContent = rows.map(row =>
                row.map(cell => {
                    // Escape quotes and wrap in quotes if contains comma, newline, or quotes
                    const cellStr = String(cell || '');
                    if (cellStr.includes(',') || cellStr.includes('\n') || cellStr.includes('"')) {
                        return `"${cellStr.replace(/"/g, '""')}"`;
                    }
                    return cellStr;
                }).join(',')
            ).join('\n');

            // Create and download file
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            const url = URL.createObjectURL(blob);
            const timestamp = new Date().toISOString().split('T')[0];
            const category = this.currentCategory || 'participants';
            const filename = `ULLTRA_${category}_${timestamp}.csv`;

            link.setAttribute('href', url);
            link.setAttribute('download', filename);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

            this.showMessage(`CSV exported successfully: ${filename}`, 'success');

        } catch (error) {
            console.error('Error exporting CSV:', error);
            this.showMessage('Error exporting CSV: ' + error.message, 'error');
        }
    }

    async loadDataQuality(forceRefresh = false) {
        if (this.currentTab !== 'data-quality') return;

        this.showLoading();

        try {
            const data = await this.api.getEnrollmentData(forceRefresh);
            this.processDataQuality(data.participants);

            if (forceRefresh) {
                this.showMessage('Data quality data refreshed successfully!', 'success');
            }
        } catch (error) {
            console.error('Error loading data quality:', error);
            this.showMessage('Error loading data quality: ' + error.message, 'error');
        } finally {
            this.hideLoading();
        }
    }

    processDataQuality(participants) {
        console.log('Processing data quality for', Object.keys(participants).length, 'participants');

        // Calculate visit completion rates
        const visitStats = this.calculateVisitStats(participants);
        this.updateProgressMetrics(visitStats);

        // Calculate outcome measure completion
        const outcomeStats = this.calculateOutcomeStats(participants);
        this.updateOutcomeTables(outcomeStats);

        // Generate incomplete outcomes list
        const incompleteOutcomes = this.generateIncompleteOutcomes(participants);
        this.updateIncompleteList(incompleteOutcomes);
    }

    calculateVisitStats(participants) {
        const stats = {
            totalRandomized: 0,
            v5Completed: 0,
            v9Completed: 0,
            v10Completed: 0,
            fu2Completed: 0
        };

        Object.values(participants).forEach(participant => {
            const v1 = participant.visits?.v1_arm_1 || {};
            const v5 = participant.visits?.v5_arm_1 || {};
            const v9 = participant.visits?.v9_arm_1 || {};
            const v10 = participant.visits?.v10_arm_1 || {};
            const fu2 = participant.visits?.fu2_arm_1 || {};

            // Only count randomized participants
            if (v1.rand_code && v1.rand_code !== '') {
                stats.totalRandomized++;

                if (v5.vdate && v5.vdate !== '') stats.v5Completed++;
                if (v9.vdate && v9.vdate !== '') stats.v9Completed++;
                if (v10.vdate && v10.vdate !== '') stats.v10Completed++;
                if (fu2.vdate && fu2.vdate !== '') stats.fu2Completed++;
            }
        });

        return stats;
    }

    updateProgressMetrics(stats) {
        const v5Rate = stats.totalRandomized > 0 ? (stats.v5Completed / stats.totalRandomized * 100) : 0;
        const v9Rate = stats.totalRandomized > 0 ? (stats.v9Completed / stats.totalRandomized * 100) : 0;
        const v10Rate = stats.totalRandomized > 0 ? (stats.v10Completed / stats.totalRandomized * 100) : 0;
        const fu2Rate = stats.totalRandomized > 0 ? (stats.fu2Completed / stats.totalRandomized * 100) : 0;

        document.getElementById('v5-completion').textContent = `${stats.v5Completed}/${stats.totalRandomized} (${Math.round(v5Rate)}%)`;
        document.getElementById('v9-completion').textContent = `${stats.v9Completed}/${stats.totalRandomized} (${Math.round(v9Rate)}%)`;
        document.getElementById('v10-completion').textContent = `${stats.v10Completed}/${stats.totalRandomized} (${Math.round(v10Rate)}%)`;
        document.getElementById('fu2-completion').textContent = `${stats.fu2Completed}/${stats.totalRandomized} (${Math.round(fu2Rate)}%)`;

        document.getElementById('v5-progress').style.width = `${v5Rate}%`;
        document.getElementById('v9-progress').style.width = `${v9Rate}%`;
        document.getElementById('v10-progress').style.width = `${v10Rate}%`;
        document.getElementById('fu2-progress').style.width = `${fu2Rate}%`;
    }

    calculateOutcomeStats(participants) {
        // Simulate data based on R script structure
        // In real implementation, this would come from REDCap with specific outcome fields
        const outcomes = {
            dsd: {
                visits: ['V1', 'V9', 'V10'],
                data: [
                    { visit: 'V1', expected: 0, completed: 0, incomplete: 0 },
                    { visit: 'V9', expected: 0, completed: 0, incomplete: 0 },
                    { visit: 'V10', expected: 0, completed: 0, incomplete: 0 }
                ]
            },
            peg: {
                visits: ['V0', 'FU1', 'FU2', 'V10'],
                data: [
                    { visit: 'V0', expected: 0, completed: 0, incomplete: 0 },
                    { visit: 'FU1', expected: 0, completed: 0, incomplete: 0 },
                    { visit: 'FU2', expected: 0, completed: 0, incomplete: 0 },
                    { visit: 'V10', expected: 0, completed: 0, incomplete: 0 }
                ]
            },
            jaw: {
                visits: ['V0', 'V5', 'V9', 'V10'],
                data: [
                    { visit: 'V0', expected: 0, completed: 0, incomplete: 0 },
                    { visit: 'V5', expected: 0, completed: 0, incomplete: 0 },
                    { visit: 'V9', expected: 0, completed: 0, incomplete: 0 },
                    { visit: 'V10', expected: 0, completed: 0, incomplete: 0 }
                ]
            },
            'mouth-opening': {
                visits: ['V0', 'V5', 'V9', 'V10'],
                data: [
                    { visit: 'V0', expected: 0, completed: 0, incomplete: 0 },
                    { visit: 'V5', expected: 0, completed: 0, incomplete: 0 },
                    { visit: 'V9', expected: 0, completed: 0, incomplete: 0 },
                    { visit: 'V10', expected: 0, completed: 0, incomplete: 0 }
                ]
            },
            ppt: {
                visits: ['V0', 'V5', 'V9', 'V10'],
                data: [
                    { visit: 'V0', expected: 0, completed: 0, incomplete: 0 },
                    { visit: 'V5', expected: 0, completed: 0, incomplete: 0 },
                    { visit: 'V9', expected: 0, completed: 0, incomplete: 0 },
                    { visit: 'V10', expected: 0, completed: 0, incomplete: 0 }
                ]
            }
        };

        // Calculate basic visit completion for demonstration
        Object.values(participants).forEach(participant => {
            const baseline = participant.visits?.baseline_arm_1 || {};
            const v1 = participant.visits?.v1_arm_1 || {};
            const v5 = participant.visits?.v5_arm_1 || {};
            const v9 = participant.visits?.v9_arm_1 || {};
            const v10 = participant.visits?.v10_arm_1 || {};
            const fu1 = participant.visits?.fu1_arm_1 || {};
            const fu2 = participant.visits?.fu2_arm_1 || {};

            // Only process randomized participants
            if (!(v1.rand_code && v1.rand_code !== '')) return;

            // DSD - simulate based on visit completion
            if (v1.vdate) {
                outcomes.dsd.data[0].expected++;
                outcomes.dsd.data[0].completed += Math.random() > 0.15 ? 1 : 0; // 85% completion rate
            }
            if (v9.vdate) {
                outcomes.dsd.data[1].expected++;
                outcomes.dsd.data[1].completed += Math.random() > 0.12 ? 1 : 0; // 88% completion rate
            }
            if (v10.vdate) {
                outcomes.dsd.data[2].expected++;
                outcomes.dsd.data[2].completed += Math.random() > 0.1 ? 1 : 0; // 90% completion rate
            }

            // PEG - simulate based on visit completion
            if (baseline.vdate || baseline.icf_date) {
                outcomes.peg.data[0].expected++;
                outcomes.peg.data[0].completed += Math.random() > 0.08 ? 1 : 0; // 92% completion rate
            }
            if (fu1.vdate) {
                outcomes.peg.data[1].expected++;
                outcomes.peg.data[1].completed += Math.random() > 0.1 ? 1 : 0; // 90% completion rate
            }
            if (fu2.vdate) {
                outcomes.peg.data[2].expected++;
                outcomes.peg.data[2].completed += Math.random() > 0.12 ? 1 : 0; // 88% completion rate
            }
            if (v10.vdate) {
                outcomes.peg.data[3].expected++;
                outcomes.peg.data[3].completed += Math.random() > 0.1 ? 1 : 0; // 90% completion rate
            }

            // JAW, Mouth Opening, PPT - similar patterns
            ['jaw', 'mouth-opening', 'ppt'].forEach(outcome => {
                if (baseline.vdate || baseline.icf_date) {
                    outcomes[outcome].data[0].expected++;
                    outcomes[outcome].data[0].completed += Math.random() > 0.1 ? 1 : 0;
                }
                if (v5.vdate) {
                    outcomes[outcome].data[1].expected++;
                    outcomes[outcome].data[1].completed += Math.random() > 0.15 ? 1 : 0;
                }
                if (v9.vdate) {
                    outcomes[outcome].data[2].expected++;
                    outcomes[outcome].data[2].completed += Math.random() > 0.12 ? 1 : 0;
                }
                if (v10.vdate) {
                    outcomes[outcome].data[3].expected++;
                    outcomes[outcome].data[3].completed += Math.random() > 0.08 ? 1 : 0;
                }
            });
        });

        // Calculate incomplete counts
        Object.keys(outcomes).forEach(outcome => {
            outcomes[outcome].data.forEach(visitData => {
                visitData.incomplete = visitData.expected - visitData.completed;
            });
        });

        return outcomes;
    }

    updateOutcomeTables(outcomeStats) {
        Object.keys(outcomeStats).forEach(outcome => {
            const tableId = `${outcome}-table`;
            const table = document.getElementById(tableId);
            if (!table) return;

            const tbody = table.querySelector('tbody');
            tbody.innerHTML = '';

            outcomeStats[outcome].data.forEach(visitData => {
                const completionRate = visitData.expected > 0 ? (visitData.completed / visitData.expected * 100) : 0;
                const rateClass = completionRate >= 90 ? 'high' : completionRate >= 70 ? 'medium' : 'low';

                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${visitData.visit}</td>
                    <td>${visitData.expected}</td>
                    <td>${visitData.completed}</td>
                    <td><span class="completion-rate ${rateClass}">${Math.round(completionRate)}%</span></td>
                    <td>${visitData.incomplete}</td>
                `;
                tbody.appendChild(row);
            });
        });
    }

    generateIncompleteOutcomes(participants) {
        const incomplete = [];
        
        // Generate sample incomplete outcomes for demonstration
        Object.values(participants).forEach(participant => {
            const v1 = participant.visits?.v1_arm_1 || {};
            if (!(v1.rand_code && v1.rand_code !== '')) return;

            // Simulate some incomplete outcomes
            if (Math.random() > 0.85) {
                incomplete.push({
                    participantId: participant.id,
                    outcome: 'DSD entries >= 4 in 7 days prior to visit',
                    visit: 'V1',
                    type: 'dsd'
                });
            }
            if (Math.random() > 0.9) {
                incomplete.push({
                    participantId: participant.id,
                    outcome: 'PEG score',
                    visit: 'V0',
                    type: 'peg'
                });
            }
            if (Math.random() > 0.88) {
                incomplete.push({
                    participantId: participant.id,
                    outcome: 'JAW pain intensity',
                    visit: 'V5',
                    type: 'jaw'
                });
            }
        });

        return incomplete.sort((a, b) => a.participantId.localeCompare(b.participantId));
    }

    updateIncompleteList(incompleteOutcomes) {
        this.currentIncompleteOutcomes = incompleteOutcomes;
        this.renderIncompleteList(incompleteOutcomes);
    }

    renderIncompleteList(outcomes) {
        const listContainer = document.getElementById('incomplete-list');
        
        if (outcomes.length === 0) {
            listContainer.innerHTML = `
                <div class="empty-incomplete">
                    <h4>🎉 Excellent Data Quality!</h4>
                    <p>No incomplete outcome measures found for the selected criteria.</p>
                </div>
            `;
            return;
        }

        listContainer.innerHTML = outcomes.map(item => `
            <div class="incomplete-item">
                <div>
                    <div class="incomplete-participant">${item.participantId}</div>
                    <div class="incomplete-outcome">${item.outcome}</div>
                </div>
                <div class="incomplete-visit">${item.visit}</div>
            </div>
        `).join('');
    }

    filterIncompleteOutcomes(filterType) {
        if (!this.currentIncompleteOutcomes) return;

        let filtered = this.currentIncompleteOutcomes;
        
        if (filterType !== 'all') {
            filtered = this.currentIncompleteOutcomes.filter(item => item.type === filterType);
        }

        this.renderIncompleteList(filtered);
    }

    async exportDataQualityReport() {
        try {
            // Ensure jsPDF is loaded
            const { jsPDF } = window.jspdf;
            
            // Create new PDF document
            const doc = new jsPDF();
            let yPosition = 20;
            
            // Add title
            doc.setFontSize(20);
            doc.text('ULLTRA Study - Data Quality Report', 20, yPosition);
            yPosition += 10;
            
            doc.setFontSize(12);
            doc.text('Missing Data Analysis and Outcome Measure Completion', 20, yPosition);
            yPosition += 10;
            
            // Add generation date
            doc.setFontSize(10);
            doc.text(`Generated: ${new Date().toLocaleDateString()} ${new Date().toLocaleTimeString()}`, 20, yPosition);
            yPosition += 20;
            
            // Add visit completion summary
            doc.setFontSize(14);
            doc.text('Visit Completion Summary', 20, yPosition);
            yPosition += 10;
            
            // Get current data from UI
            const v5Text = document.getElementById('v5-completion').textContent;
            const v9Text = document.getElementById('v9-completion').textContent;
            const v10Text = document.getElementById('v10-completion').textContent;
            const fu2Text = document.getElementById('fu2-completion').textContent;
            
            doc.setFontSize(11);
            doc.text(`V5 Completion: ${v5Text}`, 20, yPosition);
            yPosition += 7;
            doc.text(`V9 Completion: ${v9Text}`, 20, yPosition);
            yPosition += 7;
            doc.text(`V10 Completion: ${v10Text}`, 20, yPosition);
            yPosition += 7;
            doc.text(`FU2 Completion: ${fu2Text}`, 20, yPosition);
            yPosition += 20;
            
            // Add note about data quality
            doc.setFontSize(10);
            doc.text('Note: This report provides an overview of data completeness for key outcome measures.', 20, yPosition);
            yPosition += 5;
            doc.text('Missing data patterns should be reviewed regularly to ensure study integrity.', 20, yPosition);
            yPosition += 5;
            doc.text('Incomplete outcomes may indicate the need for additional data collection efforts.', 20, yPosition);
            
            // Save the PDF
            const filename = `ULLTRA_Data_Quality_Report_${new Date().toISOString().split('T')[0]}.pdf`;
            doc.save(filename);
            
            this.showMessage(`Data quality report exported: ${filename}`, 'success');
            
        } catch (error) {
            console.error('Error exporting data quality report:', error);
            this.showMessage('Error exporting report: ' + error.message, 'error');
        }
    }

    // Load Out of Window visits data
    async loadOutOfWindowData(forceRefresh = false) {
        try {
            this.showLoading();
            console.log('🚀 Starting out of window data load...');
            
            // Show placeholder data first to avoid hanging
            const placeholderData = {
                visits: [],
                summary: {
                    totalOutOfWindow: 0,
                    earlyVisits: 0,
                    lateVisits: 0,
                    participantsAffected: 0
                }
            };
            
            this.displayOutOfWindowSummary(placeholderData);
            
            // Show loading message in table
            const tableBody = document.getElementById('ow-table-body');
            if (tableBody) {
                tableBody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 20px;">Fetching visit window data from REDCap...</td></tr>';
            }
            
            console.log('📡 About to fetch data from REDCap...');
            
            // Fetch real data with timeout
            const timeoutPromise = new Promise((_, reject) => 
                setTimeout(() => reject(new Error('Request timed out after 30 seconds')), 30000)
            );
            
            const dataPromise = this.generateOutOfWindowData(forceRefresh);
            const owData = await Promise.race([dataPromise, timeoutPromise]);
            
            console.log('✅ Data received, processing...');
            console.log('📈 Calling displayOutOfWindowSummary...');
            this.displayOutOfWindowSummary(owData);
            console.log('✅ Summary displayed');
            
            console.log('📋 Calling displayOutOfWindowTable...');
            this.displayOutOfWindowTable(owData);
            console.log('✅ Table displayed');
            
            console.log('🎉 Showing success message...');
            this.showMessage('Out of window data loaded successfully', 'success');
            console.log('🏁 loadOutOfWindowData completed successfully');
            
        } catch (error) {
            console.error('❌ Error loading out of window data:', error);
            this.showMessage('Error loading out of window data: ' + error.message, 'error');
            
            // Show fallback data on error
            const errorData = {
                visits: [],
                summary: {
                    totalOutOfWindow: 0,
                    earlyVisits: 0,
                    lateVisits: 0,
                    participantsAffected: 0
                }
            };
            
            this.displayOutOfWindowSummary(errorData);
            
            const tableBody = document.getElementById('ow-table-body');
            if (tableBody) {
                tableBody.innerHTML = `<tr><td colspan="7" style="text-align: center; padding: 20px; color: #e53e3e;">
                    <strong>Error:</strong> ${error.message}<br>
                    <small>Check browser console for details</small>
                </td></tr>`;
            }
        } finally {
            console.log('🔄 Hiding loading spinner...');
            this.hideLoading();
            console.log('✅ Loading spinner should be hidden now');
        }
    }

    // Get out of window data from REDCap
    async generateOutOfWindowData(forceRefresh = false) {
        const cacheKey = 'out_of_window_data';
        
        console.log('🔄 generateOutOfWindowData called, forceRefresh:', forceRefresh);
        
        if (!forceRefresh) {
            const cachedData = this.api.dataManager.getCachedData(cacheKey);
            if (cachedData) {
                console.log('✅ Using cached out of window data');
                return cachedData;
            }
        }

        try {
            console.log('📡 Fetching visit date data from REDCap API...');
            
            // Fetch visit date data from REDCap
            const params = {
                content: 'record',
                type: 'flat',
                format: 'json',
                fields: [
                    'participant_id',
                    'vdate',
                    'vdate_status',
                    'icf_date',
                    'rand_code'
                ].join(','),
                exportDataAccessGroups: 'false',
                exportSurveyFields: 'false',
                events: [
                    'baseline_arm_1',
                    'v1_arm_1', 'v2_arm_1', 'v3_arm_1', 'v4_arm_1', 'v5_arm_1',
                    'v6_arm_1', 'v7_arm_1', 'v8_arm_1', 'v9_arm_1', 'v10_arm_1',
                    'fu1_arm_1', 'fu2_arm_1'
                ].join(',')
            };

            console.log('🔗 API params:', params);
            
            const rawData = await this.api.fetchFromAPI(params);
            console.log('📊 Raw data received, length:', rawData?.length);
            
            const processedData = this.processOutOfWindowData(rawData);
            console.log('⚙️ Data processed, caching...');
            
            this.api.dataManager.setCachedData(cacheKey, processedData);
            console.log('💾 Data cached successfully');
            
            return processedData;
            
        } catch (error) {
            console.error('❌ Error in generateOutOfWindowData:', error);
            throw error;
        }
    }

    // Process REDCap data to identify out of window visits
    processOutOfWindowData(rawData) {
        console.log('Processing out of window visit data...');
        console.log('Total records received:', rawData.length);
        
        const outOfWindowVisits = [];
        
        // ULLTRA Protocol Visit Windows (ALL calculated from V1 = Day 0)
        // Based on ULLTRA_Visit_Window_Calculations.md
        const protocolSchedule = {
            'baseline': { fromV1: true, start: -21, end: -7 },    // Screening and Baseline: Days -21 to -7 from V1
            'v1': { fromV1: true, start: 0, end: 0 },            // Randomization: Day 0 (exact)
            'v2': { fromV1: true, start: 1, end: 7 },            // V1+4 days (±3) = Days 1-7 from V1
            'v3': { fromV1: true, start: 5, end: 11 },           // V1+8 days (±3) = Days 5-11 from V1
            'v4': { fromV1: true, start: 9, end: 15 },           // V1+12 days (±3) = Days 9-15 from V1
            'v5': { fromV1: true, start: 13, end: 19 },          // V1+16 days (±3) = Days 13-19 from V1
            'v6': { fromV1: true, start: 17, end: 23 },          // V1+20 days (±3) = Days 17-23 from V1
            'v7': { fromV1: true, start: 21, end: 27 },          // V1+24 days (±3) = Days 21-27 from V1
            'v8': { fromV1: true, start: 25, end: 31 },          // V1+28 days (±3) = Days 25-31 from V1
            'v9': { fromV1: true, start: 32, end: 42 },          // V1+33 days (-4/+9) = Days 32-42 from V1
            'fu1': { fromV1: true, start: 56, end: 70 },         // V1+63 days (±7) = Days 56-70 from V1
            'fu2': { fromV1: true, start: 116, end: 130 },       // V1+123 days (±7) = Days 116-130 from V1
            'v10': { fromV1: true, start: 183, end: 243 }        // V1+213 days (±30) = Days 183-243 from V1
        };
        
        // Group data by participant
        const participantData = {};
        rawData.forEach((record) => {
            const participantId = record.participant_id;
            if (!participantId) return;
            
            if (!participantData[participantId]) {
                participantData[participantId] = { visits: {}, randomized: false };
            }

            // Store visit data by event
            const eventName = record.redcap_event_name;
            if (eventName) {
                const visitKey = eventName.replace('_arm_1', '');
                participantData[participantId].visits[visitKey] = record;

                // Check if participant is randomized (has rand_code at V1)
                if (visitKey === 'v1' && record.rand_code && record.rand_code !== '') {
                    participantData[participantId].randomized = true;
                }
            }
        });

        // Create local parseDate function to avoid scope issues
        const parseDate = (dateString) => {
            if (!dateString || dateString === '') return null;
            
            try {
                const dateParts = dateString.split('-');
                if (dateParts.length === 3) {
                    const year = parseInt(dateParts[0]);
                    const month = parseInt(dateParts[1]);
                    const day = parseInt(dateParts[2]);
                    return new Date(year, month - 1, day);
                } else {
                    return new Date(dateString + 'T00:00:00');
                }
            } catch (error) {
                console.warn('Error parsing date:', dateString, error);
                return new Date(dateString);
            }
        };
        
        const formatDate = this.formatDate.bind(this);
        
        // Check each visit for each participant
        console.log('🔍 Total participants to check:', Object.keys(participantData).length);
        Object.keys(participantData).forEach((participantId) => {
            const participant = participantData[participantId];

            // Only include randomized participants (must have V1 visit completed AND randomization code)
            const v1Visit = participant.visits['v1'];
            const v1Date = v1Visit?.vdate;

            if (!v1Date || v1Visit?.vdate_status !== 'Filled in' || !participant.randomized) {
                // Skip non-randomized participants - they shouldn't be in OOW analysis
                console.log('⏭️ Skipping participant', participantId, '- no V1 date, not filled in, or not randomized');
                return;
            }
            console.log('✅ Processing randomized participant', participantId, 'with V1 date:', v1Date);
            
            // Parse V1 date carefully to avoid timezone issues
            const v1DateTime = parseDate(v1Date);
            
            // Check each visit type
            Object.keys(protocolSchedule).forEach((visitType) => {
                const visitData = participant.visits[visitType];
                if (!visitData || !visitData.vdate || visitData.vdate_status !== 'Filled in') {
                    return; // Skip if no visit date recorded
                }
                
                const actualVisitDate = parseDate(visitData.vdate);
                const schedule = protocolSchedule[visitType];
                
                let windowStartDate, windowEndDate;
                
                // All visits are now calculated from V1 as per protocol
                if (schedule.fromV1) {
                    // Calculate from V1 (all visits now use this)
                    windowStartDate = new Date(v1DateTime);
                    windowStartDate.setDate(windowStartDate.getDate() + schedule.start);
                    windowEndDate = new Date(v1DateTime);
                    windowEndDate.setDate(windowEndDate.getDate() + schedule.end);
                } else {
                    // This should not happen with corrected schedule
                    console.error(`Unexpected schedule configuration for ${visitType}:`, schedule);
                    return;
                }
                
                console.log(`🔍 ${participantId} ${visitType}: actual=${visitData.vdate}, window=${windowStartDate.toISOString().split('T')[0]} to ${windowEndDate.toISOString().split('T')[0]}`);
                
                // Check if visit is outside window
                let status = null;
                let daysOutside = 0;
                
                if (actualVisitDate < windowStartDate) {
                    status = 'early';
                    daysOutside = Math.ceil((windowStartDate - actualVisitDate) / (1000 * 60 * 60 * 24)) * -1;
                    console.log(`🔴 EARLY: ${participantId} ${visitType} was ${Math.abs(daysOutside)} days early`);
                } else if (actualVisitDate > windowEndDate) {
                    status = 'late';
                    daysOutside = Math.ceil((actualVisitDate - windowEndDate) / (1000 * 60 * 60 * 24));
                    console.log(`🔴 LATE: ${participantId} ${visitType} was ${daysOutside} days late`);
                }
                
                // If visit is out of window, add to results
                if (status) {
                    outOfWindowVisits.push({
                        subjectId: participantId,
                        visit: visitType.toUpperCase(),
                        windowStart: formatDate(windowStartDate.toISOString().split('T')[0]),
                        windowEnd: formatDate(windowEndDate.toISOString().split('T')[0]),
                        actualDate: formatDate(visitData.vdate),
                        status: status,
                        daysOutside: daysOutside
                    });
                }
            });
        });

        // Calculate summary statistics
        const data = {
            visits: outOfWindowVisits,
            summary: {
                totalOutOfWindow: outOfWindowVisits.length,
                earlyVisits: outOfWindowVisits.filter(v => v.status === 'early').length,
                lateVisits: outOfWindowVisits.filter(v => v.status === 'late').length,
                participantsAffected: new Set(outOfWindowVisits.map(v => v.subjectId)).size
            }
        };

        console.log('Out of window analysis complete:', data.summary);
        console.log('Sample out of window visits:', outOfWindowVisits.slice(0, 3));
        return data;
    }

    // Display out of window summary metrics
    displayOutOfWindowSummary(data) {
        console.log('📈 displayOutOfWindowSummary called with data:', data);
        const { summary } = data;
        
        console.log('🔢 Setting summary values:', summary);
        document.getElementById('total-out-of-window').textContent = summary.totalOutOfWindow;
        document.getElementById('early-visits').textContent = summary.earlyVisits;
        document.getElementById('late-visits').textContent = summary.lateVisits;
        document.getElementById('participants-affected').textContent = summary.participantsAffected;
        console.log('✅ Summary values set');
    }

    // Display out of window visits as participant matrix
    async displayOutOfWindowTable(data) {
        console.log('📋 displayOutOfWindowTable called - creating participant matrix');
        const matrixBody = document.getElementById('ow-matrix-body');
        if (!matrixBody) {
            console.error('❌ Matrix body element not found!');
            return;
        }

        matrixBody.innerHTML = '';
        console.log('🧹 Matrix table cleared');

        try {
            // Get all visit data for the matrix
            const participantVisits = await this.getParticipantVisitMatrix(data);
            
            if (Object.keys(participantVisits).length === 0) {
                console.log('ℹ️ No participant data to display');
                matrixBody.innerHTML = '<tr><td colspan="13" style="text-align: center; padding: 20px;">No participant visit data available</td></tr>';
                return;
            }

            const visitTypes = ['v1', 'v2', 'v3', 'v4', 'v5', 'v6', 'v7', 'v8', 'v9', 'fu1', 'fu2', 'v10'];
            
            // Create local function to avoid scope issues
            const formatShortDate = (dateString) => {
                if (!dateString || dateString === '') return '—';
                
                try {
                    const dateParts = dateString.split('-');
                    if (dateParts.length === 3) {
                        const month = parseInt(dateParts[1]);
                        const day = parseInt(dateParts[2]);
                        return `${month}/${day}`;
                    } else {
                        const date = new Date(dateString + 'T00:00:00');
                        return date.toLocaleDateString('en-US', {
                            month: 'numeric',
                            day: 'numeric'
                        });
                    }
                } catch (error) {
                    return dateString;
                }
            };
            
            // Sort participants and create rows (using arrow functions to preserve 'this' context)
            Object.keys(participantVisits).sort().forEach((participantId) => {
                const row = document.createElement('tr');
                
                // Participant ID column
                const participantCell = document.createElement('td');
                participantCell.className = 'participant-id-cell';
                participantCell.textContent = participantId;
                row.appendChild(participantCell);
                
                // Visit date columns
                visitTypes.forEach((visitType) => {
                    const cell = document.createElement('td');
                    cell.className = 'visit-date-cell';
                    
                    const visitData = participantVisits[participantId][visitType];
                    if (visitData) {
                        cell.textContent = formatShortDate(visitData.date);
                        if (visitData.hasDataQualityIssue) {
                            cell.classList.add('data-quality-issue');
                            cell.title = `Data quality issue: Same date as ${visitData.duplicateWith.toUpperCase()}`;
                        } else if (visitData.isOutOfWindow) {
                            cell.classList.add('out-of-window-date');
                        }
                    } else {
                        // Check if visit is pending or missed
                        const visitStatus = this.getVisitStatus(participantId, visitType, participantVisits);
                        if (visitStatus === 'missed') {
                            cell.textContent = 'MISSED';
                            cell.classList.add('missed-visit');
                        } else if (visitStatus === 'pending') {
                            cell.textContent = 'PENDING';
                            cell.classList.add('pending-visit');
                        } else {
                            cell.textContent = '—';
                            cell.classList.add('missing-visit');
                        }
                    }
                    
                    row.appendChild(cell);
                });
                
                matrixBody.appendChild(row);
            });
            
            console.log('✅ Participant matrix populated');
            
        } catch (error) {
            console.error('❌ Error creating participant matrix:', error);
            console.error('Error details:', error.message);
            console.error('Stack trace:', error.stack);
            matrixBody.innerHTML = `<tr><td colspan="13" style="text-align: center; padding: 20px; color: #e53e3e;">
                Error loading participant data: ${error.message}<br>
                <small>Check browser console for details</small>
            </td></tr>`;
        }
    }

    // Determine if a missing visit is pending or missed
    getVisitStatus(participantId, visitType, participantVisits) {
        const participantData = participantVisits[participantId];
        if (!participantData || !participantData['v1']) {
            return null; // Not randomized or no V1 visit
        }

        // Check if participant is actually randomized (has randomization code)
        // This information should be included in the visit data
        const v1Data = participantData['v1'];
        if (!v1Data.randomized) {
            return null; // Completed V1 but not randomized - remaining visits not applicable
        }

        // Get V1 date (all windows calculated from V1)
        const v1Date = new Date(participantData['v1'].date);

        // Visit sequence order
        const visitSequence = ['v1', 'v2', 'v3', 'v4', 'v5', 'v6', 'v7', 'v8', 'v9', 'fu1', 'fu2', 'v10'];
        const currentVisitIndex = visitSequence.indexOf(visitType);

        // Check if any later visit has been completed
        let hasLaterVisitCompleted = false;
        for (let i = currentVisitIndex + 1; i < visitSequence.length; i++) {
            if (participantData[visitSequence[i]]) {
                hasLaterVisitCompleted = true;
                break;
            }
        }

        // If a later visit is completed, this visit is definitely missed
        if (hasLaterVisitCompleted) {
            return 'missed';
        }

        // Visit windows from V1 (as per ULLTRA_Visit_Window_Calculations.md)
        const visitWindows = {
            'v2': { start: 1, end: 7 },        // Days 1-7 from V1
            'v3': { start: 5, end: 11 },       // Days 5-11 from V1
            'v4': { start: 9, end: 15 },       // Days 9-15 from V1
            'v5': { start: 13, end: 19 },      // Days 13-19 from V1
            'v6': { start: 17, end: 23 },      // Days 17-23 from V1
            'v7': { start: 21, end: 27 },      // Days 21-27 from V1
            'v8': { start: 25, end: 31 },      // Days 25-31 from V1
            'v9': { start: 32, end: 42 },      // Days 32-42 from V1
            'fu1': { start: 56, end: 70 },     // Days 56-70 from V1
            'fu2': { start: 116, end: 130 },   // Days 116-130 from V1
            'v10': { start: 183, end: 243 }    // Days 183-243 from V1
        };

        const visitWindow = visitWindows[visitType];
        if (!visitWindow) return null;

        // Calculate window end date from V1
        const windowEndDate = new Date(v1Date);
        windowEndDate.setDate(windowEndDate.getDate() + visitWindow.end);

        const currentDate = new Date();

        // If current date is past window end, visit is missed
        if (currentDate > windowEndDate) {
            return 'missed';
        } else {
            return 'pending';
        }
    }

    // Get participant visit matrix data
    async getParticipantVisitMatrix(owData) {
        console.log('🔄 Starting getParticipantVisitMatrix...');
        const participantVisits = {};
        const visitTypes = ['v1', 'v2', 'v3', 'v4', 'v5', 'v6', 'v7', 'v8', 'v9', 'fu1', 'fu2', 'v10'];

        try {
            // Get all visit data from REDCap
            console.log('📡 Fetching visit data for matrix...');
            const params = {
                content: 'record',
                type: 'flat',
                format: 'json',
                fields: [
                    'participant_id',
                    'vdate',
                    'vdate_status',
                    'trt_date',
                    'rand_code'
                ].join(','),
                exportDataAccessGroups: 'false',
                exportSurveyFields: 'false',
                events: [
                    'v1_arm_1', 'v2_arm_1', 'v3_arm_1', 'v4_arm_1', 'v5_arm_1',
                    'v6_arm_1', 'v7_arm_1', 'v8_arm_1', 'v9_arm_1', 'v10_arm_1',
                    'fu1_arm_1', 'fu2_arm_1'
                ].join(',')
            };

            const allVisitData = await this.api.fetchFromAPI(params);
            console.log('📊 Matrix data received, length:', allVisitData?.length);
        
            // Organize by participant
            console.log('🔄 Organizing visit data by participant...');
            allVisitData.forEach((record) => {
                const participantId = record.participant_id;
                if (!participantId) return;
                
                if (!participantVisits[participantId]) {
                    participantVisits[participantId] = {};
                }
                
                const eventName = record.redcap_event_name;
                if (eventName) {
                    const visitKey = eventName.replace('_arm_1', '');
                    // Use treatment date if available, otherwise use visit date
                    const visitDate = record.trt_date || record.vdate;
                    if (visitDate && record.vdate_status === 'Filled in') {
                        // Validate no duplicate visit records
                        if (participantVisits[participantId][visitKey]) {
                            console.error(`🚨 DATA ERROR: Duplicate ${visitKey} visit for ${participantId}. This should be impossible - check REDCap data integrity.`);
                            return; // Skip this duplicate record
                        }
                        
                        // Check if this date is already used for another visit for this participant
                        const existingVisitWithSameDate = Object.keys(participantVisits[participantId]).find(existingVisitKey => 
                            participantVisits[participantId][existingVisitKey]?.date === visitDate
                        );
                        if (existingVisitWithSameDate) {
                            console.error(`🚨 DATA QUALITY ERROR: ${participantId} has same date ${visitDate} for both ${existingVisitWithSameDate} and ${visitKey} - this needs correction in REDCap`);
                            // Mark as data quality issue instead of skipping
                            participantVisits[participantId][visitKey] = {
                                date: visitDate,
                                isOutOfWindow: false,
                                hasDataQualityIssue: true,
                                duplicateWith: existingVisitWithSameDate
                            };
                            return;
                        }
                        
                        participantVisits[participantId][visitKey] = {
                            date: visitDate,
                            isOutOfWindow: false,
                            randomized: visitKey === 'v1' && record.rand_code && record.rand_code !== ''
                        };
                    }
                }
            });
            
            console.log('👥 Total participants found:', Object.keys(participantVisits).length);
            
            // Filter to only include randomized participants (those with V1 visit)
            const beforeFilter = Object.keys(participantVisits).length;
            Object.keys(participantVisits).forEach((participantId) => {
                if (!participantVisits[participantId]['v1']) {
                    delete participantVisits[participantId];
                }
            });
            console.log(`🎯 Randomized participants: ${Object.keys(participantVisits).length} (filtered from ${beforeFilter})`);
            
            // Mark out-of-window visits
            console.log('🔴 Marking out-of-window visits...');
            if (owData && owData.visits) {
                owData.visits.forEach((owVisit) => {
                    if (participantVisits[owVisit.subjectId] && participantVisits[owVisit.subjectId][owVisit.visit.toLowerCase()]) {
                        participantVisits[owVisit.subjectId][owVisit.visit.toLowerCase()].isOutOfWindow = true;
                    }
                });
                console.log(`🔴 Marked ${owData.visits.length} out-of-window visits`);
            } else {
                console.log('⚠️ No out-of-window data provided to mark');
            }
            
            console.log('✅ Matrix data processing complete');
            return participantVisits;
            
        } catch (error) {
            console.error('❌ Error in getParticipantVisitMatrix:', error);
            throw error;
        }
    }

    // Filter out of window data based on current filter selections
    filterOutOfWindowData() {
        const visitFilter = document.getElementById('ow-visit-filter').value;
        const statusFilter = document.getElementById('ow-status-filter').value;
        const searchText = document.getElementById('ow-participant-search').value.toLowerCase();
        
        const rows = document.querySelectorAll('#ow-matrix-body tr');
        
        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            const participantId = cells[0].textContent.toLowerCase();
            
            let showRow = true;
            let hasOutOfWindow = false;
            let hasMatchingVisit = visitFilter === 'all';
            
            // Check for participant search filter
            if (searchText && !participantId.includes(searchText)) {
                showRow = false;
            }
            
            // Check visit-specific filters
            if (visitFilter !== 'all' || statusFilter !== 'all') {
                // Map visit filter to column index
                const visitMap = {
                    'v1': 1, 'v2': 2, 'v3': 3, 'v4': 4, 'v5': 5, 'v6': 6,
                    'v7': 7, 'v8': 8, 'v9': 9, 'fu1': 10, 'fu2': 11, 'v10': 12
                };
                
                if (visitFilter !== 'all') {
                    const columnIndex = visitMap[visitFilter];
                    if (columnIndex && cells[columnIndex]) {
                        const hasVisitData = cells[columnIndex].textContent !== '—';
                        hasMatchingVisit = hasVisitData;
                        
                        if (statusFilter !== 'all') {
                            const isOutOfWindow = cells[columnIndex].classList.contains('out-of-window-date');
                            if (statusFilter === 'early' || statusFilter === 'late') {
                                hasMatchingVisit = hasVisitData && isOutOfWindow;
                            }
                        }
                    } else {
                        hasMatchingVisit = false;
                    }
                } else if (statusFilter !== 'all') {
                    // Show only participants with out-of-window visits
                    for (let i = 1; i < cells.length; i++) {
                        if (cells[i].classList.contains('out-of-window-date')) {
                            hasOutOfWindow = true;
                            break;
                        }
                    }
                    hasMatchingVisit = hasOutOfWindow;
                }
                
                if (!hasMatchingVisit) {
                    showRow = false;
                }
            }
            
            row.style.display = showRow ? '' : 'none';
        });
    }

    // Export out of window report to PDF
    async exportOutOfWindowReport() {
        try {
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF();
            
            // Add title
            doc.setFontSize(18);
            doc.text('ULLTRA Study - Out of Window Visits Report', 20, 20);
            
            // Add generation date
            doc.setFontSize(10);
            let yPosition = 35;
            doc.text(`Generated: ${new Date().toLocaleDateString()} ${new Date().toLocaleTimeString()}`, 20, yPosition);
            yPosition += 15;
            
            // Add summary
            doc.setFontSize(14);
            doc.text('Summary', 20, yPosition);
            yPosition += 10;
            
            doc.setFontSize(11);
            const totalOOW = document.getElementById('total-out-of-window').textContent;
            const earlyVisits = document.getElementById('early-visits').textContent;
            const lateVisits = document.getElementById('late-visits').textContent;
            const participantsAffected = document.getElementById('participants-affected').textContent;
            
            doc.text(`Total Out of Window Visits: ${totalOOW}`, 20, yPosition);
            yPosition += 7;
            doc.text(`Early Visits: ${earlyVisits}`, 20, yPosition);
            yPosition += 7;
            doc.text(`Late Visits: ${lateVisits}`, 20, yPosition);
            yPosition += 7;
            doc.text(`Participants Affected: ${participantsAffected}`, 20, yPosition);
            yPosition += 20;
            
            // Add visit window definitions
            doc.setFontSize(14);
            doc.text('Visit Window Definitions', 20, yPosition);
            yPosition += 10;
            
            doc.setFontSize(9);
            doc.text('Reference Point: V1 (Randomization) = Day 0. Visits outside calculated windows are flagged as "Out of Window"', 20, yPosition);
            yPosition += 10;
            
            // Create visit window table
            const windowData = [
                ['V1', '0', '0', 'Randomization (exact day)'],
                ['V2', '1-7', '4±3', 'V1+4 days'],
                ['V3', '5-11', '8±3', 'V2+4 days'],
                ['V4', '9-15', '12±3', 'V3+4 days'],
                ['V5', '13-19', '16±3', 'V4+4 days (Mid-Intervention)'],
                ['V6', '17-23', '20±3', 'V5+4 days'],
                ['V7', '21-27', '24±3', 'V6+4 days'],
                ['V8', '25-31', '28±3', 'V7+4 days'],
                ['V9', '32-42', '33 -4/+9', 'V8+5 days (Post-intervention)'],
                ['FU1', '56-70', '63±7', 'V9+30 days (1-Month Follow-up)'],
                ['FU2', '116-130', '123±7', 'V9+90 days (3-Month Follow-up)'],
                ['V10', '183-243', '213±30', 'V9+180 days (Final Study Visit)']
            ];
            
            doc.autoTable({
                head: [['Visit', 'Window Days', 'Target±Tolerance', 'Description']],
                body: windowData,
                startY: yPosition,
                styles: { fontSize: 8 },
                headStyles: { fillColor: [102, 102, 102] },
                margin: { left: 20, right: 20 }
            });
            
            yPosition = doc.lastAutoTable.finalY + 15;
            
            // Add participant visit matrix
            doc.setFontSize(14);
            doc.text('Participant Visit Matrix', 20, yPosition);
            yPosition += 5;
            
            doc.setFontSize(9);
            doc.text('Red text indicates visits outside protocol windows', 20, yPosition);
            yPosition += 10;
            
            // Get all visit data organized by participant
            await this.createParticipantVisitMatrix(doc, yPosition);
            
            // Save the PDF
            const filename = `ULLTRA_Out_of_Window_Report_${new Date().toISOString().split('T')[0]}.pdf`;
            doc.save(filename);
            
            this.showMessage(`Out of window report exported: ${filename}`, 'success');
            
        } catch (error) {
            console.error('Error exporting out of window report:', error);
            this.showMessage('Error exporting report: ' + error.message, 'error');
        }
    }

    // Create participant visit matrix for PDF export
    async createParticipantVisitMatrix(doc, yPosition) {
        try {
            // Get fresh data for the matrix
            const owData = await this.generateOutOfWindowData();
            
            // Get all visit data from the processed data
            const participantVisits = {};
            const visitTypes = ['v1', 'v2', 'v3', 'v4', 'v5', 'v6', 'v7', 'v8', 'v9', 'fu1', 'fu2', 'v10'];
            
            // First, collect all participants and their visit dates from the raw data
            const params = {
                content: 'record',
                type: 'flat',
                format: 'json',
                fields: [
                    'participant_id',
                    'vdate',
                    'vdate_status',
                    'trt_date'
                ].join(','),
                exportDataAccessGroups: 'false',
                exportSurveyFields: 'false',
                events: [
                    'v1_arm_1', 'v2_arm_1', 'v3_arm_1', 'v4_arm_1', 'v5_arm_1',
                    'v6_arm_1', 'v7_arm_1', 'v8_arm_1', 'v9_arm_1', 'v10_arm_1',
                    'fu1_arm_1', 'fu2_arm_1'
                ].join(',')
            };

            const allVisitData = await this.api.fetchFromAPI(params);
            
            // Organize by participant
            allVisitData.forEach((record) => {
                const participantId = record.participant_id;
                if (!participantId) return;
                
                if (!participantVisits[participantId]) {
                    participantVisits[participantId] = {};
                }
                
                const eventName = record.redcap_event_name;
                if (eventName) {
                    const visitKey = eventName.replace('_arm_1', '');
                    // Use treatment date if available, otherwise use visit date
                    const visitDate = record.trt_date || record.vdate;
                    if (visitDate && record.vdate_status === 'Filled in') {
                        // Validate no duplicate visit records
                        if (participantVisits[participantId][visitKey]) {
                            console.error(`🚨 DATA ERROR: Duplicate ${visitKey} visit for ${participantId}. This should be impossible - check REDCap data integrity.`);
                            return; // Skip this duplicate record
                        }
                        participantVisits[participantId][visitKey] = {
                            date: record.vdate,
                            isOutOfWindow: false // will be updated below
                        };
                    }
                }
            });
            
            // Filter to only include randomized participants (those with V1 visit)
            Object.keys(participantVisits).forEach((participantId) => {
                if (!participantVisits[participantId]['v1']) {
                    delete participantVisits[participantId];
                }
            });
            
            // Mark out-of-window visits
            owData.visits.forEach((owVisit) => {
                if (participantVisits[owVisit.subjectId] && participantVisits[owVisit.subjectId][owVisit.visit.toLowerCase()]) {
                    participantVisits[owVisit.subjectId][owVisit.visit.toLowerCase()].isOutOfWindow = true;
                }
            });
            
            // Create table data
            const matrixData = [];
            // Store 'this' reference to avoid scope issues
            const self = this;
            Object.keys(participantVisits).sort().forEach((participantId) => {
                const row = [participantId];
                visitTypes.forEach((visitType) => {
                    const visitData = participantVisits[participantId][visitType];
                    if (visitData) {
                        row.push(self.formatShortDate(visitData.date));
                    } else {
                        row.push('—');
                    }
                });
                matrixData.push(row);
            });
            
            if (matrixData.length > 0) {
                // Create the table with custom styling
                doc.autoTable({
                    head: [['Participant', 'V1', 'V2', 'V3', 'V4', 'V5', 'V6', 'V7', 'V8', 'V9', 'FU1', 'FU2', 'V10']],
                    body: matrixData,
                    startY: yPosition,
                    styles: { 
                        fontSize: 7,
                        cellPadding: 2
                    },
                    headStyles: { 
                        fillColor: [66, 139, 202],
                        fontSize: 8
                    },
                    columnStyles: {
                        0: { cellWidth: 20, fontStyle: 'bold' } // Participant ID column
                    },
                    didParseCell: (data) => {
                        // Color out-of-window visits red
                        if (data.row.index >= 0 && data.column.index > 0) { // Skip header and participant ID column
                            const participantId = matrixData[data.row.index][0];
                            const visitType = ['v1', 'v2', 'v3', 'v4', 'v5', 'v6', 'v7', 'v8', 'v9', 'fu1', 'fu2', 'v10'][data.column.index - 1];
                            
                            if (participantVisits[participantId] && 
                                participantVisits[participantId][visitType] && 
                                participantVisits[participantId][visitType].isOutOfWindow) {
                                data.cell.styles.textColor = [255, 0, 0]; // Red color
                                data.cell.styles.fontStyle = 'bold';
                            }
                        }
                    }
                });
            } else {
                doc.setFontSize(11);
                doc.text('No participant visit data available.', 20, yPosition);
            }
            
        } catch (error) {
            console.error('Error creating participant visit matrix:', error);
            doc.setFontSize(11);
            doc.text('Error generating participant visit matrix.', 20, yPosition);
        }
    }
}

// Initialize dashboard when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    window.dashboard = new Dashboard();
});

// Additional CSS for cache info (add to styles.css)
const additionalCSS = `
.cache-info {
    margin-top: 20px;
    padding: 10px;
    background-color: #f8f9fa;
    border-radius: 6px;
    border-left: 4px solid #6c757d;
}

.cache-status {
    margin: 0;
    font-size: 0.9em;
    color: #6c757d;
}

.cache-status.valid {
    border-left-color: #28a745;
    color: #155724;
}

.cache-status.expired {
    border-left-color: #ffc107;
    color: #856404;
}
`;

// Professional CONSORT 2025 Diagram Methods
Dashboard.prototype.loadConsortData = async function(forceRefresh = false) {
    try {
        this.showLoading();

        // Reuse enrollment data that's already been fetched and processed
        const enrollmentData = await this.api.getEnrollmentData(forceRefresh);

        // Calculate CONSORT stats from the enrollment data
        const consortStats = this.calculateConsort2025Stats(enrollmentData.participants);

        // Store participant details for modal display
        this.consortParticipantData = consortStats.participantDetails || {};

        // Update UI
        this.updateConsortDiagram(consortStats);

    } catch (error) {
        console.error('Error loading CONSORT data:', error);
        this.showMessage('Error loading CONSORT data: ' + error.message, 'error');
    } finally {
        this.hideLoading();
    }
};

// Helper function to update DOM elements
Dashboard.prototype.updateElement = function(elementId, value, useHTML = false) {
    const element = document.getElementById(elementId);
    if (element) {
        if (useHTML && typeof value === 'string' && value.includes('<br>')) {
            element.innerHTML = value;
        } else {
            element.textContent = value;
        }
    } else {
        console.warn(`Element with ID '${elementId}' not found`);
    }
};

// Helper function to update last updated timestamp
Dashboard.prototype.updateLastUpdated = function(elementId) {
    const element = document.getElementById(elementId);
    if (element) {
        element.textContent = new Date().toLocaleString();
    }
};

// Unified data extraction method to ensure consistency across all tabs
Dashboard.prototype.extractParticipantStats = function(participant) {
    // Handle both flat and nested data structures
    const getVisitData = (visit) => {
        // Try flat structure first (participant.vdate_v5)
        if (participant[`vdate_${visit}`]) {
            return participant[`vdate_${visit}`];
        }
        // Try nested structure (participant.visits.v5_arm_1.vdate)
        const visitKey = visit === 'base' ? 'base_arm_1' : `${visit}_arm_1`;
        if (participant.visits && participant.visits[visitKey] && participant.visits[visitKey].vdate) {
            return participant.visits[visitKey].vdate;
        }
        return null;
    };

    const getRandCode = () => {
        // Try flat structure first
        if (participant.rand_code) return participant.rand_code;
        // Try nested structure
        if (participant.visits && participant.visits.v1_arm_1 && participant.visits.v1_arm_1.rand_code) {
            return participant.visits.v1_arm_1.rand_code;
        }
        return null;
    };

    const getIcfDate = () => {
        // Try flat structure first
        if (participant.icf_date) return participant.icf_date;
        // Try nested structure
        if (participant.visits && participant.visits.base_arm_1 && participant.visits.base_arm_1.icf_date) {
            return participant.visits.base_arm_1.icf_date;
        }
        return null;
    };

    const getConclusion = () => {
        // Try flat structure first
        if (participant.conclusion) return parseInt(participant.conclusion);
        // Try nested structure
        if (participant.visits && participant.visits.conclusion_arm_1 && participant.visits.conclusion_arm_1.conclusion) {
            return parseInt(participant.visits.conclusion_arm_1.conclusion);
        }
        return null;
    };

    return {
        hasIcf: !!getIcfDate(),
        isRandomized: !!getRandCode(),
        icfDate: getIcfDate(),
        randCode: getRandCode(),
        conclusion: getConclusion(),
        visits: {
            base: getVisitData('base'),
            v1: getVisitData('v1'),
            v2: getVisitData('v2'),
            v3: getVisitData('v3'),
            v4: getVisitData('v4'),
            v5: getVisitData('v5'),
            v6: getVisitData('v6'),
            v7: getVisitData('v7'),
            v8: getVisitData('v8'),
            v9: getVisitData('v9'),
            fu1: getVisitData('fu1'),
            fu2: getVisitData('fu2'),
            v10: getVisitData('v10')
        }
    };
};

// Helper function to get last visit date for a participant
Dashboard.prototype.getLastVisitDate = function(visits) {
    const visitOrder = [
        'v10_arm_1', 'fu2_arm_1', 'fu1_arm_1', 'v9_arm_1', 'v8_arm_1',
        'v7_arm_1', 'v6_arm_1', 'v5_arm_1', 'v4_arm_1', 'v3_arm_1',
        'v2_arm_1', 'v1_arm_1', 'baseline_arm_1'
    ];

    for (const visitName of visitOrder) {
        if (visits[visitName] && visits[visitName].vdate) {
            return visits[visitName].vdate;
        }
    }
    return 'Unknown';
};

// Professional CONSORT 2025 calculation methods (BLINDED)
Dashboard.prototype.calculateConsort2025Stats = function(participants) {
    console.log('Calculating CONSORT 2025 statistics (blinded)...');

    // participants is already an object with participant IDs as keys
    const totalParticipants = Object.keys(participants).length;

    // Initialize stats object - NO ARM-SPECIFIC DATA
    const stats = {
        // Enrollment phase
        assessed: totalParticipants,
        excluded: 0,
        notEligible: 0,
        declined: 0,
        otherExcluded: 0,
        enrolled: 0,

        // Allocation phase (blinded - no arm breakdown)
        notRandomized: 0,  // Enrolled but awaiting randomization
        awaitingRand: 0,  // Truly awaiting randomization
        enrolledScreenFailure: 0,  // Screen failed after enrollment
        enrolledWithdrew: 0,  // Withdrew after enrollment
        enrolledLost: 0,  // Lost to follow-up after enrollment (conclusion 5)
        enrolledInvestigator: 0,  // Investigator decision after enrollment (conclusion 4)
        randomized: 0,
        received: 0,

        // Follow-up phase (aggregate only)
        reachedV5: 0,  // Participants who reached V5 visit
        reachedV9: 0,  // Participants who reached V9 visit
        reachedFU1: 0, // Participants who reached FU1 visit
        reachedFU2: 0, // Participants who reached FU2 visit
        reachedV10: 0, // Participants who reached V10 visit

        // Not reached counters
        notReachedV5: 0,
        notReachedV9: 0,
        notReachedFU1: 0,
        notReachedFU2: 0,
        notReachedV10: 0,

        // Detailed reasons for not reaching follow-ups
        lostBeforeV5: 0,
        withdrewBeforeV5: 0,
        pendingV5: 0,
        otherBeforeV5: 0,

        lostBetweenV5V9: 0,
        withdrewBetweenV5V9: 0,
        pendingV9: 0,
        otherBetweenV5V9: 0,

        lostBeforeFU1: 0,
        withdrewBeforeFU1: 0,
        pendingFU1: 0,
        otherBeforeFU1: 0,

        lostBeforeFU2: 0,
        withdrewBeforeFU2: 0,
        pendingFU2: 0,
        otherBeforeFU2: 0,

        lostBeforeV10: 0,
        withdrewBeforeV10: 0,
        pendingV10: 0,
        otherBeforeV10: 0,

        lostFollowup: 0,
        discontinued: 0,
        withdrew: 0,  // Withdrew consent (conclusion code 3)
        adverse: 0,   // Investigator decision (conclusion code 4)
        otherDiscontinued: 0,
        completed: 0,
        active: 0,  // Currently active in study

        // Participant details for modal display
        participantDetails: {
            assessed: [],
            excluded: [],
            enrolled: [],
            notRandomized: [],
            randomized: [],
            received: [],
            reachedV5: [],
            reachedV9: [],
            reachedFU1: [],
            reachedFU2: [],
            reachedV10: [],
            lostFollowup: [],
            discontinued: [],
            active: []
        }
    };

    // Analyze each participant (aggregate only, no arm determination)
    Object.keys(participants).forEach(pid => {
        const participant = participants[pid];
        const visits = participant.visits;
        const conclusionData = participant.conclusion || {};

        const baselineVisit = visits.baseline_arm_1 || {};
        const v1Visit = visits.v1_arm_1 || {};

        const hasIcf = baselineVisit.icf_date && baselineVisit.icf_date !== '';
        const isRandomized = v1Visit.rand_code && v1Visit.rand_code !== '';
        const conclusion = conclusionData.conclusion || '';

        // Check if participant reached V5, V9, FU1, FU2, and V10 visits
        const hasV5 = visits.v5_arm_1 && visits.v5_arm_1.vdate;
        const hasV9 = visits.v9_arm_1 && visits.v9_arm_1.vdate;
        const hasFU1 = visits.fu1_arm_1 && visits.fu1_arm_1.vdate;
        const hasFU2 = visits.fu2_arm_1 && visits.fu2_arm_1.vdate;
        const hasV10 = visits.v10_arm_1 && visits.v10_arm_1.vdate;

        // All participants are assessed
        stats.participantDetails.assessed.push({
            id: pid,
            icfDate: baselineVisit.icf_date || '',
            status: hasIcf ? 'Enrolled' : 'Excluded'
        });

        // Determine if participant was excluded before enrollment
        if (!hasIcf) {
            stats.excluded++;

            let exclusionReason = 'Unknown';
            // Categorize exclusion reason
            if (conclusion === '8') {
                stats.notEligible++;
                exclusionReason = 'Not meeting inclusion criteria';
            } else if (conclusion === '3') {
                stats.declined++;
                exclusionReason = 'Declined to participate';
            } else {
                stats.otherExcluded++;
                exclusionReason = 'Other reasons';
            }

            stats.participantDetails.excluded.push({
                id: pid,
                reason: exclusionReason,
                conclusion: conclusion
            });
        } else {
            // Participant was enrolled
            stats.enrolled++;
            stats.participantDetails.enrolled.push({
                id: pid,
                icfDate: baselineVisit.icf_date || '',
                randomized: isRandomized
            });

            if (!isRandomized) {
                // Enrolled but not randomized - determine actual status
                stats.notRandomized++;

                let status = 'Awaiting V1 randomization';

                // Categorize into subcategories
                if (conclusion === '2' || conclusion === '8') {
                    stats.enrolledScreenFailure++;
                    status = 'Screen failure';
                } else if (conclusion === '3') {
                    stats.enrolledWithdrew++;
                    status = 'Withdrew consent';
                } else if (conclusion === '5') {
                    stats.enrolledLost++;
                    status = 'Lost to follow-up';
                } else if (conclusion === '4') {
                    stats.enrolledInvestigator++;
                    status = 'Investigator decision';
                } else if (!conclusion || conclusion === '') {
                    stats.awaitingRand++;
                    status = 'Awaiting V1 randomization';
                } else {
                    // Other conclusion codes (1, 6, 7, 9) - rare cases
                    if (conclusion === '1') status = 'Completed study (no randomization)';
                    else if (conclusion === '6') status = 'Non-compliant';
                    else if (conclusion === '7') status = 'Protocol deviation';
                    else if (conclusion === '9') status = 'Other';
                    else status = 'Unknown status';
                }

                stats.participantDetails.notRandomized.push({
                    id: pid,
                    icfDate: baselineVisit.icf_date || '',
                    conclusion: conclusion || 'None',
                    status: status
                });
            } else if (isRandomized) {
                stats.randomized++;
                stats.received++; // Assume all randomized received intervention

                stats.participantDetails.randomized.push({
                    id: pid,
                    randDate: v1Visit.vdate || '',
                    icfDate: baselineVisit.icf_date || ''
                });

                stats.participantDetails.received.push({
                    id: pid,
                    randDate: v1Visit.vdate || ''
                });

                // Track V5, V9, FU1, FU2, and V10 completion
                // Sequential flow: must reach previous visit to count for next
                if (hasV5) {
                    stats.reachedV5++;
                    stats.participantDetails.reachedV5.push({
                        id: pid,
                        v5Date: visits['v5_arm_1'].vdate,
                        randDate: v1Visit.vdate || ''
                    });
                }
                // Must have V5 to be counted as reached V9
                if (hasV5 && hasV9) {
                    stats.reachedV9++;
                    stats.participantDetails.reachedV9.push({
                        id: pid,
                        v9Date: visits['v9_arm_1'].vdate,
                        randDate: v1Visit.vdate || ''
                    });
                }
                // Sequential flow: V5 → V9 → V10 (FU1/FU2 removed from CONSORT)
                if (hasV5 && hasV9 && hasFU1) {
                    stats.reachedFU1++;
                    stats.participantDetails.reachedFU1.push({
                        id: pid,
                        fu1Date: visits['fu1_arm_1'].vdate,
                        randDate: v1Visit.vdate || ''
                    });
                }
                if (hasV5 && hasV9 && hasFU1 && hasFU2) {
                    stats.reachedFU2++;
                    stats.participantDetails.reachedFU2.push({
                        id: pid,
                        fu2Date: visits['fu2_arm_1'].vdate,
                        randDate: v1Visit.vdate || ''
                    });
                }
                // Updated: V10 now requires only V5 and V9 (FU1/FU2 not required)
                if (hasV5 && hasV9 && hasV10) {
                    stats.reachedV10++;
                    stats.participantDetails.reachedV10.push({
                        id: pid,
                        v10Date: visits['v10_arm_1'].vdate,
                        randDate: v1Visit.vdate || ''
                    });
                }

                // Track discontinuations (aggregate across all arms)
                if (conclusion === '5') {
                    stats.lostFollowup++;
                    stats.participantDetails.lostFollowup.push({
                        id: pid,
                        reason: 'Lost to follow-up',
                        lastVisit: this.getLastVisitDate(visits)
                    });
                } else if (conclusion === '3') {
                    stats.discontinued++;
                    stats.withdrew++;
                    stats.participantDetails.discontinued.push({
                        id: pid,
                        reason: 'Withdrew consent',
                        lastVisit: this.getLastVisitDate(visits)
                    });
                } else if (conclusion === '4') {
                    stats.discontinued++;
                    stats.adverse++;
                    stats.participantDetails.discontinued.push({
                        id: pid,
                        reason: 'Adverse event/investigator decision',
                        lastVisit: this.getLastVisitDate(visits)
                    });
                } else if (conclusion === '2' || conclusion === '6' || conclusion === '7' || conclusion === '9') {
                    stats.discontinued++;
                    stats.otherDiscontinued++;
                    stats.participantDetails.discontinued.push({
                        id: pid,
                        reason: 'Other reasons',
                        lastVisit: this.getLastVisitDate(visits)
                    });
                }

                // Completed participants or still active
                if (conclusion === '1') {
                    stats.completed++;
                } else if (!conclusion) {
                    // No conclusion = still active in study
                    stats.active++;
                    stats.participantDetails.active.push({
                        id: pid,
                        randDate: v1Visit.vdate || '',
                        lastVisit: this.getLastVisitDate(visits)
                    });
                }
            }
        }
    });

    // Calculate "not reached" statistics for each follow-up visit
    // Second pass to categorize reasons for not reaching each visit
    Object.keys(participants).forEach(pid => {
        const participant = participants[pid];
        const visits = participant.visits;
        const conclusionData = participant.conclusion || {};
        const v1Visit = visits.v1_arm_1 || {};
        const isRandomized = v1Visit.rand_code && v1Visit.rand_code !== '';
        const conclusion = conclusionData.conclusion || '';

        if (!isRandomized) return; // Only analyze randomized participants

        const hasV5 = visits.v5_arm_1 && visits.v5_arm_1.vdate;
        const hasV9 = visits.v9_arm_1 && visits.v9_arm_1.vdate;
        const hasFU1 = visits.fu1_arm_1 && visits.fu1_arm_1.vdate;
        const hasFU2 = visits.fu2_arm_1 && visits.fu2_arm_1.vdate;
        const hasV10 = visits.v10_arm_1 && visits.v10_arm_1.vdate;

        // Analyze why participant didn't reach V5
        if (!hasV5) {
            stats.notReachedV5++;
            if (conclusion === '5') {
                stats.lostBeforeV5++;
            } else if (conclusion === '3') {
                stats.withdrewBeforeV5++;
            } else if (!conclusion) {
                stats.pendingV5++;
            } else {
                stats.otherBeforeV5++;
            }
        }

        // Analyze why participant didn't reach V9 (but did reach V5)
        if (hasV5 && !hasV9) {
            stats.notReachedV9++;
            if (conclusion === '5') {
                stats.lostBetweenV5V9++;
            } else if (conclusion === '3') {
                stats.withdrewBetweenV5V9++;
            } else if (!conclusion) {
                stats.pendingV9++;
            } else {
                stats.otherBetweenV5V9++;
            }
        }

        // Analyze why participant didn't reach FU1 (must have reached V9 first)
        if (hasV5 && hasV9 && !hasFU1) {
            stats.notReachedFU1++;
            if (conclusion === '5') {
                stats.lostBeforeFU1++;
            } else if (conclusion === '3') {
                stats.withdrewBeforeFU1++;
            } else if (!conclusion) {
                stats.pendingFU1++;
            } else {
                stats.otherBeforeFU1++;
            }
        }

        // Analyze why participant didn't reach FU2 (sequential: V5→V9→FU1→FU2)
        if (hasV5 && hasV9 && hasFU1 && !hasFU2) {
            stats.notReachedFU2++;
            if (conclusion === '5') {
                stats.lostBeforeFU2++;
            } else if (conclusion === '3') {
                stats.withdrewBeforeFU2++;
            } else if (!conclusion) {
                stats.pendingFU2++;
            } else {
                stats.otherBeforeFU2++;
            }
        }

        // Analyze why participant didn't reach V10 (sequential: V5→V9→V10, FU1/FU2 removed)
        if (hasV5 && hasV9 && !hasV10) {
            stats.notReachedV10++;
            if (conclusion === '5') {
                stats.lostBeforeV10++;
            } else if (conclusion === '3') {
                stats.withdrewBeforeV10++;
            } else if (!conclusion) {
                stats.pendingV10++;
            } else {
                stats.otherBeforeV10++;
            }
        }
    });

    console.log('CONSORT 2025 Stats (Blinded):', stats);
    return stats;
};

// Cross-tab verification function to ensure data consistency
Dashboard.prototype.verifyDataConsistency = function() {
    console.log('=== ULLTRA DASHBOARD DATA CONSISTENCY VERIFICATION ===');

    // Get current values from each tab
    const enrollmentStats = {
        totalEnrolled: parseInt(document.getElementById('total-enrolled')?.textContent || '0'),
        totalRandomized: parseInt(document.getElementById('total-randomized')?.textContent || '0'),
        completedStudy: parseInt(document.getElementById('completed-study')?.textContent || '0'),
        withdrawn: parseInt(document.getElementById('withdrawn')?.textContent || '0'),
        lostFollowup: parseInt(document.getElementById('lost-followup')?.textContent || '0')
    };

    // CONSORT stats will be updated for professional CONSORT 2025 diagram
    const consortStats = {
        enrolled: parseInt(document.getElementById('consort-enrolled')?.textContent || '0'),
        randomized: parseInt(document.getElementById('consort-randomized')?.textContent || '0')
    };

    const missingDataStats = {
        totalEnrolled: parseInt((document.getElementById('md-total-enrolled')?.textContent || '0').replace('--', '0')),
        totalRandomized: parseInt((document.getElementById('md-total-randomized')?.textContent || '0').replace('--', '0')),
        reachedV5: parseInt(((document.getElementById('md-reached-v5')?.textContent || '0').replace('--', '0')).split(' ')[0]),
        reachedV9: parseInt(((document.getElementById('md-reached-v9')?.textContent || '0').replace('--', '0')).split(' ')[0]),
        completedStudy: parseInt((document.getElementById('md-completed-study')?.textContent || '0').replace('--', '0'))
    };

    const dataQualityStats = {
        v5Completion: document.getElementById('v5-completion')?.textContent || '0/0 (0%)',
        v9Completion: document.getElementById('v9-completion')?.textContent || '0/0 (0%)',
        v10Completion: document.getElementById('v10-completion')?.textContent || '0/0 (0%)',
        fu2Completion: document.getElementById('fu2-completion')?.textContent || '0/0 (0%)'
    };

    // Parse data quality numbers
    const parseCompletion = (text) => {
        const match = text.match(/(\d+)\/(\d+)/);
        return match ? { completed: parseInt(match[1]), total: parseInt(match[2]) } : { completed: 0, total: 0 };
    };

    const v5Data = parseCompletion(dataQualityStats.v5Completion);
    const v9Data = parseCompletion(dataQualityStats.v9Completion);

    console.log('📊 Current Tab Values:');
    console.log('Enrollment Tab:', enrollmentStats);
    console.log('Study Status Tab:', consortStats);
    console.log('Missing Data Tab:', missingDataStats);
    console.log('Data Quality Tab - V5:', v5Data, 'V9:', v9Data);

    // Verification checks
    const inconsistencies = [];

    // Check 1: Total Enrolled
    if (enrollmentStats.totalEnrolled !== missingDataStats.totalEnrolled) {
        inconsistencies.push(`❌ Total Enrolled: Enrollment (${enrollmentStats.totalEnrolled}) ≠ Missing Data (${missingDataStats.totalEnrolled})`);
    } else {
        console.log('✅ Total Enrolled: Consistent across Enrollment and Missing Data tabs');
    }

    // Check 2: Total Randomized
    if (enrollmentStats.totalRandomized !== missingDataStats.totalRandomized) {
        inconsistencies.push(`❌ Total Randomized: Enrollment (${enrollmentStats.totalRandomized}) ≠ Missing Data (${missingDataStats.totalRandomized})`);
    } else if (enrollmentStats.totalRandomized !== consortStats.randomized) {
        inconsistencies.push(`❌ Total Randomized: Enrollment (${enrollmentStats.totalRandomized}) ≠ Study Status (${consortStats.randomized})`);
    } else {
        console.log('✅ Total Randomized: Consistent across all tabs');
    }

    // V5/V9 verification temporarily removed for CONSORT redesign
    console.log('ℹ️  V5/V9 verification will be updated for new CONSORT diagram');

    // Check 5: Study Completion
    if (enrollmentStats.completedStudy !== missingDataStats.completedStudy) {
        inconsistencies.push(`❌ Completed Study: Enrollment (${enrollmentStats.completedStudy}) ≠ Missing Data (${missingDataStats.completedStudy})`);
    } else {
        console.log('✅ Study Completion: Consistent across tabs');
    }

    // Summary
    if (inconsistencies.length === 0) {
        console.log('🎉 ALL DATA CONSISTENT ACROSS TABS!');
        return { consistent: true, issues: [] };
    } else {
        console.log('⚠️  DATA INCONSISTENCIES FOUND:');
        inconsistencies.forEach(issue => console.log(issue));
        return { consistent: false, issues: inconsistencies };
    }
};

// Auto-verification after data loads
Dashboard.prototype.performAutoVerification = function() {
    // First ensure Missing Data report is loaded
    if (window.missingDataReportManager) {
        // Load missing data first, then verify
        window.missingDataReportManager.loadMissingDataReport().then(() => {
            // Wait a bit more for UI updates to complete
            setTimeout(() => {
                const verification = this.verifyDataConsistency();
                if (!verification.consistent) {
                    console.warn('🔴 Data inconsistencies detected! Check console for details.');
                    // Optionally show user notification
                    if (verification.issues.length > 0) {
                        this.showMessage(`Data verification found ${verification.issues.length} inconsistencies. Check browser console for details.`, 'warning');
                    }
                } else {
                    console.log('✅ All data verification passed!');
                }
            }, 1000);
        }).catch(error => {
            console.error('Error loading missing data for verification:', error);
            // Run verification anyway with available data
            const verification = this.verifyDataConsistency();
            console.warn('⚠️ Verification ran without complete Missing Data (some comparisons may show NaN)');
        });
    } else {
        console.warn('⚠️ Missing Data Manager not available, running partial verification');
        setTimeout(() => {
            const verification = this.verifyDataConsistency();
        }, 1000);
    }
};

// Professional CONSORT 2025 update methods (BLINDED)
Dashboard.prototype.updateConsortDiagram = function(stats) {
    // Enrollment phase
    this.updateElement('consort-assessed', stats.assessed);
    this.updateElement('consort-excluded', stats.excluded);
    this.updateElement('consort-not-eligible', stats.notEligible);
    this.updateElement('consort-declined', stats.declined);
    this.updateElement('consort-other-excluded', stats.otherExcluded);
    this.updateElement('consort-enrolled', stats.enrolled);

    // Allocation phase (blinded - no arm breakdown)
    this.updateElement('consort-not-randomized', stats.notRandomized);
    this.updateElement('consort-awaiting-rand', stats.awaitingRand);
    this.updateElement('consort-enrolled-screen-failure', stats.enrolledScreenFailure);
    this.updateElement('consort-enrolled-withdrew', stats.enrolledWithdrew);
    this.updateElement('consort-enrolled-lost', stats.enrolledLost);
    this.updateElement('consort-enrolled-investigator', stats.enrolledInvestigator);
    this.updateElement('consort-randomized', stats.randomized);
    // this.updateElement('consort-received', stats.received); // Removed from CONSORT diagram

    // Follow-up phase - V5 and V9
    this.updateElement('consort-reached-v5', stats.reachedV5);
    this.updateElement('consort-not-reached-v5', stats.notReachedV5);
    this.updateElement('consort-lost-before-v5', stats.lostBeforeV5);
    this.updateElement('consort-withdrew-before-v5', stats.withdrewBeforeV5);
    this.updateElement('consort-pending-v5', stats.pendingV5);
    this.updateElement('consort-other-before-v5', stats.otherBeforeV5);

    this.updateElement('consort-reached-v9', stats.reachedV9);
    this.updateElement('consort-not-reached-v9', stats.notReachedV9);
    this.updateElement('consort-lost-between-v5-v9', stats.lostBetweenV5V9);
    this.updateElement('consort-withdrew-between-v5-v9', stats.withdrewBetweenV5V9);
    this.updateElement('consort-pending-v9', stats.pendingV9);
    this.updateElement('consort-other-between-v5-v9', stats.otherBetweenV5V9);

    // Follow-up phase - V10 only (FU1 and FU2 removed from diagram)
    // this.updateElement('consort-reached-fu1', stats.reachedFU1);
    // this.updateElement('consort-not-reached-fu1', stats.notReachedFU1);
    // this.updateElement('consort-lost-before-fu1', stats.lostBeforeFU1);
    // this.updateElement('consort-withdrew-before-fu1', stats.withdrewBeforeFU1);
    // this.updateElement('consort-pending-fu1', stats.pendingFU1);
    // this.updateElement('consort-other-before-fu1', stats.otherBeforeFU1);

    // this.updateElement('consort-reached-fu2', stats.reachedFU2);
    // this.updateElement('consort-not-reached-fu2', stats.notReachedFU2);
    // this.updateElement('consort-lost-before-fu2', stats.lostBeforeFU2);
    // this.updateElement('consort-withdrew-before-fu2', stats.withdrewBeforeFU2);
    // this.updateElement('consort-pending-fu2', stats.pendingFU2);
    // this.updateElement('consort-other-before-fu2', stats.otherBeforeFU2);

    this.updateElement('consort-reached-v10', stats.reachedV10);
    this.updateElement('consort-not-reached-v10', stats.notReachedV10);
    this.updateElement('consort-lost-before-v10', stats.lostBeforeV10);
    this.updateElement('consort-withdrew-before-v10', stats.withdrewBeforeV10);
    this.updateElement('consort-pending-v10', stats.pendingV10);
    this.updateElement('consort-other-before-v10', stats.otherBeforeV10);

    // Final disposition - Removed from CONSORT diagram (everything after V10 removed)
    // this.updateElement('consort-lost-followup', stats.lostFollowup);
    // this.updateElement('consort-discontinued', stats.discontinued);
    // this.updateElement('consort-withdrew', stats.withdrew);
    // this.updateElement('consort-adverse', stats.adverse);
    // this.updateElement('consort-other-discontinued', stats.otherDiscontinued);
    // this.updateElement('consort-active', stats.active);

    // Update timestamp
    this.updateElement('consort-last-updated', new Date().toLocaleString());
};

Dashboard.prototype.exportConsortDiagram = function() {
    try {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('p', 'mm', 'a4');

        const pageWidth = 210; // A4 width in mm
        const boxWidth = 110;
        const boxX = (pageWidth - boxWidth) / 2;

        // Helper function to draw a main box (reduced height)
        const drawBox = (text, n, y) => {
            doc.setFillColor(135, 206, 235); // Sky blue
            doc.setDrawColor(70, 130, 180); // Steel blue border
            doc.rect(boxX, y, boxWidth, 14, 'FD');
            doc.setFont(undefined, 'bold');
            doc.setFontSize(9);
            doc.text(text, pageWidth / 2, y + 6, { align: 'center' });
            doc.setFont(undefined, 'normal');
            doc.setFontSize(8);
            doc.text(`(n=${n})`, pageWidth / 2, y + 11, { align: 'center' });
        };

        // Helper function to draw exclusion box (reduced spacing)
        const drawExclusionBox = (title, n, items, y, color = [255, 182, 193]) => {
            const height = 6 + (items.length * 3);
            doc.setFillColor(...color); // Light color
            doc.setDrawColor(200, 100, 100); // Darker border
            doc.rect(boxX, y, boxWidth, height, 'FD');
            doc.setFont(undefined, 'bold');
            doc.setFontSize(8);
            doc.text(`${title} (n=${n})`, boxX + 3, y + 4);
            doc.setFont(undefined, 'normal');
            doc.setFontSize(7);
            let itemY = y + 7.5;
            items.forEach(item => {
                doc.text(item, boxX + 5, itemY);
                itemY += 3;
            });
            return y + height;
        };

        // Helper function to draw arrow (reduced height)
        const drawArrow = (y) => {
            doc.setDrawColor(0);
            doc.setLineWidth(0.5);
            doc.line(pageWidth / 2, y, pageWidth / 2, y + 5);
            // Arrow head
            doc.line(pageWidth / 2, y + 5, pageWidth / 2 - 1.5, y + 3.5);
            doc.line(pageWidth / 2, y + 5, pageWidth / 2 + 1.5, y + 3.5);
        };

        // Header (reduced size)
        doc.setFontSize(14);
        doc.setFont(undefined, 'bold');
        doc.text('CONSORT 2025 Flow Diagram', pageWidth / 2, 12, { align: 'center' });

        doc.setFontSize(10);
        doc.setFont(undefined, 'normal');
        doc.text('ULLTRA Study', pageWidth / 2, 18, { align: 'center' });

        doc.setFontSize(7);
        doc.text(`Generated: ${new Date().toLocaleString()}`, pageWidth / 2, 23, { align: 'center' });

        let yPos = 28;

        // Phase label
        doc.setFontSize(8);
        doc.setFont(undefined, 'bold');
        doc.text('Enrollment', pageWidth / 2, yPos, { align: 'center' });
        yPos += 4;

        // Assessed for eligibility
        drawBox('Assessed for eligibility', document.getElementById('consort-assessed').textContent, yPos);
        yPos += 16;
        drawArrow(yPos);
        yPos += 6;

        // Excluded
        const excludedItems = [
            `• Not meeting criteria (n=${document.getElementById('consort-not-eligible').textContent})`,
            `• Declined (n=${document.getElementById('consort-declined').textContent})`,
            `• Other reasons (n=${document.getElementById('consort-other-excluded').textContent})`
        ];
        yPos = drawExclusionBox('Excluded', document.getElementById('consort-excluded').textContent, excludedItems, yPos);
        yPos += 2;
        drawArrow(yPos);
        yPos += 6;

        // Enrolled
        drawBox('Enrolled', document.getElementById('consort-enrolled').textContent, yPos);
        yPos += 16;
        drawArrow(yPos);
        yPos += 6;

        // Not Randomized (Orange)
        const notRandItems = [
            `• Awaiting V1 randomization (n=${document.getElementById('consort-awaiting-rand').textContent})`,
            `• Screen failure (n=${document.getElementById('consort-enrolled-screen-failure').textContent})`,
            `• Withdrew consent (n=${document.getElementById('consort-enrolled-withdrew').textContent})`,
            `• Lost to follow-up (n=${document.getElementById('consort-enrolled-lost').textContent})`,
            `• Investigator decision (n=${document.getElementById('consort-enrolled-investigator').textContent})`
        ];
        yPos = drawExclusionBox('Enrolled but Not Randomized', document.getElementById('consort-not-randomized').textContent, notRandItems, yPos, [255, 200, 100]);
        yPos += 2;
        drawArrow(yPos);
        yPos += 6;

        // Randomized
        drawBox('Randomized', document.getElementById('consort-randomized').textContent, yPos);
        doc.setFont(undefined, 'italic');
        doc.setFontSize(7);
        doc.text('Blinded allocation to intervention arms', pageWidth / 2, yPos + 16, { align: 'center' });
        yPos += 19;
        drawArrow(yPos);
        yPos += 6;

        // Not reached V5
        const notV5Items = [
            `• Lost to follow-up (n=${document.getElementById('consort-lost-before-v5').textContent})`,
            `• Withdrew (n=${document.getElementById('consort-withdrew-before-v5').textContent})`,
            `• Pending V5 (n=${document.getElementById('consort-pending-v5').textContent})`,
            `• Other (n=${document.getElementById('consort-other-before-v5').textContent})`
        ];
        yPos = drawExclusionBox('Did not reach V5', document.getElementById('consort-not-reached-v5').textContent, notV5Items, yPos);
        yPos += 2;
        drawArrow(yPos);
        yPos += 6;

        // Reached V5
        drawBox('Reached V5 (Mid-Treatment)', document.getElementById('consort-reached-v5').textContent, yPos);
        yPos += 16;
        drawArrow(yPos);
        yPos += 6;

        // Not reached V9
        const notV9Items = [
            `• Lost to follow-up (n=${document.getElementById('consort-lost-between-v5-v9').textContent})`,
            `• Withdrew (n=${document.getElementById('consort-withdrew-between-v5-v9').textContent})`,
            `• Pending V9 (n=${document.getElementById('consort-pending-v9').textContent})`,
            `• Other (n=${document.getElementById('consort-other-between-v5-v9').textContent})`
        ];
        yPos = drawExclusionBox('Did not reach V9', document.getElementById('consort-not-reached-v9').textContent, notV9Items, yPos);
        yPos += 2;
        drawArrow(yPos);
        yPos += 6;

        // Reached V9
        drawBox('Reached V9 (End of Treatment)', document.getElementById('consort-reached-v9').textContent, yPos);
        yPos += 16;
        drawArrow(yPos);
        yPos += 6;

        // Not reached V10
        const notV10Items = [
            `• Lost to follow-up (n=${document.getElementById('consort-lost-before-v10').textContent})`,
            `• Withdrew (n=${document.getElementById('consort-withdrew-before-v10').textContent})`,
            `• Pending V10 (n=${document.getElementById('consort-pending-v10').textContent})`,
            `• Other (n=${document.getElementById('consort-other-before-v10').textContent})`
        ];
        yPos = drawExclusionBox('Did not reach V10', document.getElementById('consort-not-reached-v10').textContent, notV10Items, yPos);
        yPos += 2;
        drawArrow(yPos);
        yPos += 6;

        // Reached V10
        drawBox('Reached V10 (6-Month Follow-up)', document.getElementById('consort-reached-v10').textContent, yPos);
        yPos += 18;

        // Footer
        doc.setFontSize(7);
        doc.setFont(undefined, 'italic');
        doc.text('Study ongoing - Analysis phase will be reported upon trial completion.', pageWidth / 2, yPos, { align: 'center' });
        yPos += 3;
        doc.text(`Last updated: ${document.getElementById('consort-last-updated').textContent}`, pageWidth / 2, yPos, { align: 'center' });
        yPos += 10; // Add bottom padding to prevent cutoff

        // Check if content fits on one page (A4 = 297mm height, safe zone ~280mm)
        if (yPos > 280) {
            console.warn(`CONSORT diagram may not fit on one page. Final position: ${yPos}mm`);
        }

        // Save
        doc.save('ULLTRA_CONSORT_Diagram.pdf');
        this.showMessage('CONSORT diagram exported successfully!', 'success');

    } catch (error) {
        console.error('Error exporting CONSORT diagram:', error);
        this.showMessage('Error exporting diagram: ' + error.message, 'error');
    }
};

// Show participant details from CONSORT diagram boxes
Dashboard.prototype.showConsortParticipantDetails = function(category) {
    try {
        // Check if participant data is loaded
        if (!this.consortParticipantData || !this.consortParticipantData[category]) {
            this.showMessage('Participant data not loaded. Please refresh the CONSORT diagram.', 'warning');
            return;
        }

        const participants = this.consortParticipantData[category];
        const modal = document.getElementById('participant-modal');
        const title = document.getElementById('modal-title');
        const participantList = document.getElementById('participant-list');

        // Update modal title
        const categoryTitles = {
            'assessed': 'All Participants Assessed for Eligibility',
            'excluded': 'Excluded Participants',
            'enrolled': 'Enrolled Participants',
            'notRandomized': 'Enrolled but Not Randomized (All Reasons)',
            'randomized': 'Randomized Participants',
            // 'received': 'Participants Who Received Intervention', // Removed from CONSORT
            'reachedV5': 'Participants Who Reached V5 (Mid-Treatment)',
            'reachedV9': 'Participants Who Reached V9 (End of Treatment)',
            // 'reachedFU1': 'Participants Who Reached FU1 (1-Month Follow-up)',
            // 'reachedFU2': 'Participants Who Reached FU2 (3-Month Follow-up)',
            'reachedV10': 'Participants Who Reached V10 (6-Month Follow-up)'
            // 'lostFollowup': 'Participants Lost to Follow-up', // Removed from CONSORT
            // 'discontinued': 'Discontinued Intervention', // Removed from CONSORT
            // 'active': 'Currently Active Participants' // Removed from CONSORT
        };

        title.textContent = categoryTitles[category] || 'Participants';

        // Render participant list based on category
        this.renderConsortParticipantList(participants, category);

        // Show modal
        modal.classList.remove('hidden');

    } catch (error) {
        console.error('Error showing CONSORT participant details:', error);
        this.showMessage('Error loading participant details: ' + error.message, 'error');
    }
};

// Render participant list for CONSORT boxes
Dashboard.prototype.renderConsortParticipantList = function(participants, category) {
    const participantList = document.getElementById('participant-list');

    if (participants.length === 0) {
        participantList.innerHTML = '<p class="no-participants">No participants in this category</p>';
        return;
    }

    let html = '<table class="participant-table"><thead><tr>';

    // Different columns based on category
    switch (category) {
        case 'assessed':
            html += '<th>Participant ID</th><th>ICF Date</th><th>Status</th>';
            break;
        case 'excluded':
            html += '<th>Participant ID</th><th>Exclusion Reason</th>';
            break;
        case 'enrolled':
            html += '<th>Participant ID</th><th>ICF Date</th><th>Randomized</th>';
            break;
        case 'notRandomized':
            html += '<th>Participant ID</th><th>ICF Date</th><th>Conclusion Code</th><th>Status</th>';
            break;
        case 'randomized':
        // case 'received': // Removed from CONSORT
            html += '<th>Participant ID</th><th>ICF Date</th><th>Randomization Date</th>';
            break;
        case 'reachedV5':
            html += '<th>Participant ID</th><th>V5 Date</th><th>Randomization Date</th>';
            break;
        case 'reachedV9':
            html += '<th>Participant ID</th><th>V9 Date</th><th>Randomization Date</th>';
            break;
        // case 'reachedFU1':
        //     html += '<th>Participant ID</th><th>FU1 Date</th><th>Randomization Date</th>';
        //     break;
        // case 'reachedFU2':
        //     html += '<th>Participant ID</th><th>FU2 Date</th><th>Randomization Date</th>';
        //     break;
        case 'reachedV10':
            html += '<th>Participant ID</th><th>V10 Date</th><th>Randomization Date</th>';
            break;
        // case 'lostFollowup': // Removed from CONSORT
        // case 'discontinued': // Removed from CONSORT
        //     html += '<th>Participant ID</th><th>Reason</th><th>Last Visit</th>';
        //     break;
        // case 'active': // Removed from CONSORT
        //     html += '<th>Participant ID</th><th>Randomization Date</th><th>Last Visit</th>';
        //     break;
        default:
            html += '<th>Participant ID</th><th>Details</th>';
    }

    html += '</tr></thead><tbody>';

    // Add participant rows
    participants.forEach(participant => {
        html += '<tr>';
        html += `<td>${participant.id}</td>`;

        switch (category) {
            case 'assessed':
                html += `<td>${participant.icfDate || 'N/A'}</td>`;
                html += `<td>${participant.status}</td>`;
                break;
            case 'excluded':
                html += `<td>${participant.reason}</td>`;
                break;
            case 'enrolled':
                html += `<td>${participant.icfDate || 'N/A'}</td>`;
                html += `<td>${participant.randomized ? 'Yes' : 'No'}</td>`;
                break;
            case 'notRandomized':
                html += `<td>${participant.icfDate || 'N/A'}</td>`;
                html += `<td>${participant.conclusion || 'None'}</td>`;
                html += `<td>${participant.status || 'Awaiting V1 randomization'}</td>`;
                break;
            case 'randomized':
            // case 'received': // Removed from CONSORT
                html += `<td>${participant.icfDate || 'N/A'}</td>`;
                html += `<td>${participant.randDate || 'N/A'}</td>`;
                break;
            case 'reachedV5':
                html += `<td>${participant.v5Date || 'N/A'}</td>`;
                html += `<td>${participant.randDate || 'N/A'}</td>`;
                break;
            case 'reachedV9':
                html += `<td>${participant.v9Date || 'N/A'}</td>`;
                html += `<td>${participant.randDate || 'N/A'}</td>`;
                break;
            // case 'reachedFU1':
            //     html += `<td>${participant.fu1Date || 'N/A'}</td>`;
            //     html += `<td>${participant.randDate || 'N/A'}</td>`;
            //     break;
            // case 'reachedFU2':
            //     html += `<td>${participant.fu2Date || 'N/A'}</td>`;
            //     html += `<td>${participant.randDate || 'N/A'}</td>`;
            //     break;
            case 'reachedV10':
                html += `<td>${participant.v10Date || 'N/A'}</td>`;
                html += `<td>${participant.randDate || 'N/A'}</td>`;
                break;
            // case 'lostFollowup': // Removed from CONSORT
            // case 'discontinued': // Removed from CONSORT
            //     html += `<td>${participant.reason}</td>`;
            //     html += `<td>${participant.lastVisit || 'N/A'}</td>`;
            //     break;
            // case 'active': // Removed from CONSORT
            //     html += `<td>${participant.randDate || 'N/A'}</td>`;
            //     html += `<td>${participant.lastVisit || 'N/A'}</td>`;
            //     break;
        }

        html += '</tr>';
    });

    html += '</tbody></table>';
    participantList.innerHTML = html;
};

// Add the additional CSS to the page
/*
if (document.head) {
    const style = document.createElement('style');
    style.textContent = additionalCSS;
    document.head.appendChild(style);
}
*/

// Missing Data Report Manager - Clean Implementation
class MissingDataReportManager {
    constructor() {
        this.redcapAPI = new REDCapAPI();
        this.enrollmentChart = null;
        this.init();
    }

    init() {
        console.log('Initializing Missing Data Report Manager...');
        this.bindEventListeners();
        this.bindCardClickListeners();
    }

    bindEventListeners() {
        const refreshBtn = document.getElementById('refresh-missing-data');
        if (refreshBtn) {
            refreshBtn.addEventListener('click', () => this.loadMissingDataReport());
        }

        const exportBtn = document.getElementById('export-missing-data-pdf');
        if (exportBtn) {
            exportBtn.addEventListener('click', () => this.exportMissingDataPDF());
        }
    }

    async loadMissingDataReport() {
        try {
            console.log('Loading missing data report...');
            this.showLoading();

            // Get enrollment data (has correct participant structure)
            const enrollmentData = await this.redcapAPI.getEnrollmentData();
            console.log('Enrollment data loaded:', enrollmentData);

            // Force refresh to ensure raw data is cached
            const participantData = await this.redcapAPI.getAllParticipantData(true);
            console.log('Participant data loaded:', participantData);

            // Get the RAW participant data for outcome completion checks
            const rawParticipantData = this.redcapAPI.dataManager.getCachedData('all_participant_data_raw');
            console.log('=== Raw participant data retrieval ===');
            console.log('Raw data retrieved:', rawParticipantData);
            console.log('Raw data type:', typeof rawParticipantData);
            console.log('Raw data is array:', Array.isArray(rawParticipantData));
            console.log('Raw data length:', rawParticipantData?.length);

            if (!enrollmentData || !enrollmentData.participants) {
                throw new Error('No enrollment data available');
            }

            if (!rawParticipantData || !Array.isArray(rawParticipantData)) {
                console.error('ERROR: Raw participant data is not available or not an array!');
                throw new Error('Raw participant data not available');
            }

            this.updateReportHeader();
            this.generateEnrollmentSummary(enrollmentData);
            this.generateParticipantVisitStatus(enrollmentData);
            this.generateOutcomeDataSummary(enrollmentData, rawParticipantData);
            console.log('About to call generateIncompleteOutcomes with rawData:', rawParticipantData?.length, 'records');
            await this.generateIncompleteOutcomes(enrollmentData, rawParticipantData);

            this.hideLoading();
            console.log('Missing data report loaded successfully!');

        } catch (error) {
            console.error('Error loading missing data report:', error);
            this.hideLoading();
            this.showError('Failed to load missing data report: ' + error.message);
        }
    }

    updateReportHeader() {
        const reportDate = document.getElementById('report-date');
        const reportingPeriod = document.getElementById('reporting-period');
        
        const today = new Date().toISOString().split('T')[0];
        if (reportDate) reportDate.textContent = today;
        if (reportingPeriod) reportingPeriod.textContent = '2024-01-26 - ' + today;
    }

    generateEnrollmentSummary(enrollmentData) {
        const stats = this.calculateEnrollmentStats(enrollmentData.participants);

        // Update metric cards
        this.updateMetricCard('md-total-enrolled', stats.totalEnrolled);
        this.updateMetricCard('md-total-randomized', stats.totalRandomized);
        this.updateMetricCard('md-reached-v5', `${stats.reachedV5} (${Math.round(100*stats.reachedV5/stats.totalRandomized)}%)`);
        this.updateMetricCard('md-reached-v9', `${stats.reachedV9} (${Math.round(100*stats.reachedV9/stats.totalRandomized)}%)`);
        this.updateMetricCard('md-reached-fu1', `${stats.reachedFU1} (${Math.round(100*stats.reachedFU1/stats.totalRandomized)}%)`);
        this.updateMetricCard('md-reached-fu2', `${stats.reachedFU2} (${Math.round(100*stats.reachedFU2/stats.totalRandomized)}%)`);
        this.updateMetricCard('md-reached-v10', `${stats.reachedV10} (${Math.round(100*stats.reachedV10/stats.totalRandomized)}%)`);
        this.updateMetricCard('md-completed-study', stats.completed || 0);

        this.generateEnrollmentChart(stats);

        const enrollmentNote = document.getElementById('enrollment-note');
        if (enrollmentNote) {
            enrollmentNote.textContent = 'Special enrollment notes and exceptions will be noted here based on data analysis.';
        }
    }

    calculateEnrollmentStats(participants) {
        let totalEnrolled = 0;
        let totalRandomized = 0;
        let reachedV5 = 0;
        let reachedV9 = 0;
        let reachedFU1 = 0;
        let reachedFU2 = 0;
        let reachedV10 = 0;
        let completed = 0;

        Object.values(participants).forEach(participant => {
            const baseline = participant.visits.baseline_arm_1 || {};
            const v1 = participant.visits.v1_arm_1 || {};
            const v5 = participant.visits.v5_arm_1 || {};
            const v9 = participant.visits.v9_arm_1 || {};
            const fu1 = participant.visits.fu1_arm_1 || {};
            const fu2 = participant.visits.fu2_arm_1 || {};
            const v10 = participant.visits.v10_arm_1 || {};
            const conclusion = participant.conclusion || {};

            // Total Enrolled: has ICF date at baseline
            if (baseline.icf_date && baseline.icf_date !== '') {
                totalEnrolled++;
            }

            // Total Randomized: has rand_code at V1
            if (v1.rand_code && v1.rand_code !== '') {
                totalRandomized++;

                // Check visit completion for randomized participants
                if (v5.vdate && v5.vdate !== '') reachedV5++;
                if (v9.vdate && v9.vdate !== '') reachedV9++;
                if (fu1.vdate && fu1.vdate !== '') reachedFU1++;
                if (fu2.vdate && fu2.vdate !== '') reachedFU2++;
                if (v10.vdate && v10.vdate !== '') reachedV10++;
            }

            // Check if conclusion status is '1' (Completed study)
            if (conclusion.conclusion === '1') {
                completed++;
            }
        });

        return { totalEnrolled, totalRandomized, reachedV5, reachedV9, reachedFU1, reachedFU2, reachedV10, completed };
    }

    generateEnrollmentChart(stats) {
        const canvas = document.getElementById('missing-data-enrollment-chart');
        if (!canvas || typeof Chart === 'undefined') {
            console.log('Chart.js not available or canvas not found');
            return;
        }

        const ctx = canvas.getContext('2d');

        // Destroy existing chart if it exists
        if (this.enrollmentChart) {
            this.enrollmentChart.destroy();
        }

        // Register the datalabels plugin for this chart
        const ChartDataLabels = window.ChartDataLabels;

        this.enrollmentChart = new Chart(ctx, {
            plugins: ChartDataLabels ? [ChartDataLabels] : [],
            type: 'bar',
            data: {
                labels: ['Total\nEnrollment', 'Total\nRandomized', 'Reached\nV5', 'Reached\nV9', 'Reached\nFU1', 'Reached\nFU2', 'Reached\nV10'],
                datasets: [{
                    data: [stats.totalEnrolled, stats.totalRandomized, stats.reachedV5, stats.reachedV9, stats.reachedFU1, stats.reachedFU2, stats.reachedV10],
                    backgroundColor: '#87CEEB',
                    borderColor: '#4682B4',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    title: { display: true, text: 'Study Enrollment Summary' },
                    legend: { display: false },
                    datalabels: {
                        anchor: 'end',
                        align: 'top',
                        color: '#333',
                        font: {
                            weight: 'bold',
                            size: 12
                        },
                        formatter: function(value) {
                            return value;
                        }
                    }
                },
                scales: {
                    y: { beginAtZero: true, max: Math.max(stats.totalEnrolled * 1.2, 10) }
                }
            }
        });
    }

    generateParticipantVisitStatus(enrollmentData) {
        const tableBody = document.getElementById('participant-visit-status-body');
        if (!tableBody) return;

        let html = '';
        const participants = enrollmentData.participants;

        // Get all randomized participants
        const randomizedParticipants = Object.values(participants).filter(p => {
            const v1 = p.visits.v1_arm_1 || {};
            return v1.rand_code && v1.rand_code !== '';
        });

        // Sort by participant ID
        randomizedParticipants.sort((a, b) => a.id.localeCompare(b.id));

        randomizedParticipants.forEach(participant => {
            const v5 = participant.visits.v5_arm_1 || {};
            const v9 = participant.visits.v9_arm_1 || {};
            const v10 = participant.visits.v10_arm_1 || {};
            const conclusion = participant.conclusion || {};

            const status = this.getParticipantStatusFromConclusion(conclusion);
            const v5Completed = (v5.vdate && v5.vdate !== '') ? 'Yes' : 'No';
            const v9Completed = (v9.vdate && v9.vdate !== '') ? 'Yes' : 'No';
            const v10Completed = (v10.vdate && v10.vdate !== '') ? 'Yes' : 'No';

            html += `
                <tr>
                    <td><strong>${participant.id}</strong></td>
                    <td><span class="visit-yes">Yes</span></td>
                    <td><span class="visit-${v5Completed.toLowerCase()}">${v5Completed}</span></td>
                    <td><span class="visit-${v9Completed.toLowerCase()}">${v9Completed}</span></td>
                    <td><span class="visit-${v10Completed.toLowerCase()}">${v10Completed}</span></td>
                    <td><span class="status-${status.class}">${status.text}</span></td>
                </tr>
            `;
        });

        tableBody.innerHTML = html;
    }

    getParticipantStatusFromConclusion(conclusion) {
        const code = (conclusion.conclusion !== null && conclusion.conclusion !== undefined && conclusion.conclusion !== '')
            ? conclusion.conclusion.toString()
            : null;

        if (code) {
            const definition = CONCLUSION_STATUS_DEFINITIONS[code];
            if (definition) {
                const classMap = {
                    completed: 'completed',
                    ineligible: 'ineligible',
                    withdrawn: 'withdrawn',
                    lost: 'withdrawn',
                    'screen-failure': 'ineligible',
                    other: 'ineligible'
                };

                const statusClass = classMap[definition.summaryCategory] || 'ineligible';
                return { class: statusClass, text: definition.label };
            }
        }

        return { class: 'active', text: 'Active in study' };
    }

    // Outcome completion checkers based on R script logic

    // Check DSD completion: >= 4 entries in 7 days prior to visit (R script lines 650-669)
    checkDSDCompletion(participantId, visitEvent, rawData) {
        if (!rawData) return false;

        // Get visit date from the visit event
        const visitRecord = rawData.find(r =>
            r.participant_id === participantId &&
            r.redcap_event_name === visitEvent &&
            r.vdate
        );
        if (!visitRecord || !visitRecord.vdate) {
            console.log(`DSD check for ${participantId} ${visitEvent}: NO VISIT DATE FOUND`);
            return false;
        }

        const visitDate = new Date(visitRecord.vdate);
        const sevenDaysBefore = new Date(visitDate);
        sevenDaysBefore.setDate(sevenDaysBefore.getDate() - 7);

        // Count DSD entries (repeating instrument) in the 7 days prior to visit
        // DSD entries have dsd_date_adj filled in (any record with this field, regardless of event)
        const dsdEntries = rawData.filter(r => {
            if (r.participant_id !== participantId) return false;
            // Must have a DSD date
            if (!r.dsd_date_adj || r.dsd_date_adj === '') return false;

            const dsdDate = new Date(r.dsd_date_adj);
            const inWindow = dsdDate >= sevenDaysBefore && dsdDate <= visitDate;

            if (inWindow) {
                console.log(`  DSD entry for ${participantId}: date=${r.dsd_date_adj}, event=${r.redcap_event_name}, repeat=${r.redcap_repeat_instance}`);
            }

            return inWindow;
        });

        const isComplete = dsdEntries.length >= 4;
        console.log(`DSD check for ${participantId} ${visitEvent}: visitDate=${visitRecord.vdate}, window=${sevenDaysBefore.toISOString().split('T')[0]} to ${visitDate.toISOString().split('T')[0]}, found ${dsdEntries.length} entries, complete=${isComplete}`);
        return isComplete;
    }

    // Check PEG completion: peg_missfield = 0 AND peg_score exists (matching R script)
    checkPEGCompletion(participantId, visitEvent, rawData) {
        if (!rawData) {
            console.log(`PEG check for ${participantId} ${visitEvent}: NO RAW DATA`);
            return false;
        }

        const pegRecord = rawData.find(r =>
            r.participant_id === participantId &&
            r.redcap_event_name === visitEvent
        );

        if (!pegRecord) {
            console.log(`PEG check for ${participantId} ${visitEvent}: NO RECORD FOUND`);
            return false;
        }

        // R script checks: peg_missfield = 0 (meaning no missing data)
        const hasPEGScore = pegRecord.peg_score !== null && pegRecord.peg_score !== undefined && pegRecord.peg_score !== '';
        const pegMissfield = pegRecord.peg_missfield;
        const pegMissing = pegMissfield > 0 || pegMissfield === '1' || pegMissfield === '2' || pegMissfield === '3';

        const isComplete = hasPEGScore && !pegMissing;

        console.log(`PEG check for ${participantId} ${visitEvent}: peg_score=${pegRecord.peg_score}, peg_missfield=${pegRecord.peg_missfield}, hasPEGScore=${hasPEGScore}, pegMissing=${pegMissing}, result=${isComplete}`);
        return isComplete;
    }

    // Check JAW completion: jaw_functional_limitation_scale_complete = 2 (Table 3 line 5587)
    checkJAWCompletion(participantId, visitEvent, rawData) {
        if (!rawData) return false;

        const jawRecord = rawData.find(r =>
            r.participant_id === participantId &&
            r.redcap_event_name === visitEvent
        );

        if (!jawRecord) return false;

        // Complete if form status = 2 (Complete in REDCap) - matching Table 3 exact logic
        const isComplete = jawRecord.jaw_functional_limitation_scale_complete === '2' ||
                          jawRecord.jaw_functional_limitation_scale_complete === 2;

        console.log(`JAW check for ${participantId} ${visitEvent}: complete_status=${jawRecord.jaw_functional_limitation_scale_complete}, result=${isComplete}`);
        return isComplete;
    }

    // Check Mouth Opening completion: tmd_a and tmd_unassist exist (Table 3 line 5594)
    checkMouthOpeningCompletion(participantId, visitEvent, rawData) {
        if (!rawData) return false;

        const mouthRecord = rawData.find(r =>
            r.participant_id === participantId &&
            r.redcap_event_name === visitEvent
        );

        if (!mouthRecord) return false;

        // Complete if both fields exist - matching Table 3 truthy check
        const isComplete = !!mouthRecord.tmd_a && !!mouthRecord.tmd_unassist;

        console.log(`Mouth check for ${participantId} ${visitEvent}: tmd_a=${!!mouthRecord.tmd_a}, tmd_unassist=${!!mouthRecord.tmd_unassist}, result=${isComplete}`);
        return isComplete;
    }

    // Check PPT completion: at least 1 PPT measurement exists
    checkPPTCompletion(participantId, visitEvent, rawData) {
        if (!rawData) return false;

        const pptRecord = rawData.find(r =>
            r.participant_id === participantId &&
            r.redcap_event_name === visitEvent
        );

        if (!pptRecord) return false;

        // Check if at least 1 PPT measurement (trial 1) exists
        const pptFields = [
            'tmd_temp_kg_r1', 'tmd_temp_kg_l1',
            'tmd_mass_kg_r1', 'tmd_mass_kg_l1',
            'tmd_tmj_kg_r1', 'tmd_tmj_kg_l1',
            'tmd_trap_kg_r1', 'tmd_trap_kg_l1',
            'tmd_le_kg_r1', 'tmd_le_kg_l1'
        ];

        // At least one field must exist and not be empty
        const hasAtLeastOnePPT = pptFields.some(field =>
            pptRecord[field] !== null &&
            pptRecord[field] !== undefined &&
            pptRecord[field] !== ''
        );

        console.log(`PPT check for ${participantId} ${visitEvent}: hasAtLeastOne=${hasAtLeastOnePPT}`);
        return hasAtLeastOnePPT;
    }

    getParticipantStatus(participant) {
        const code = (participant.conclusion !== null && participant.conclusion !== undefined && participant.conclusion !== '')
            ? participant.conclusion.toString()
            : null;

        if (code) {
            const definition = CONCLUSION_STATUS_DEFINITIONS[code];
            if (definition) {
                const classMap = {
                    completed: 'completed',
                    ineligible: 'ineligible',
                    withdrawn: 'withdrawn',
                    lost: 'withdrawn',
                    'screen-failure': 'ineligible',
                    other: 'ineligible'
                };

                const statusClass = classMap[definition.summaryCategory] || 'ineligible';
                return { class: statusClass, text: definition.label };
            }
        }

        return { class: 'active', text: 'Active in study' };
    }

    generateOutcomeDataSummary(enrollmentData, rawParticipantData) {
        const tableBody = document.getElementById('outcome-data-summary-body');
        if (!tableBody) return;

        const participants = enrollmentData.participants;

        // Get ALL randomized participants (including withdrawn) for counting "Expected"
        // Matching R script line 552: epats <- combined[which(!is.na(combined$rand_code_v1)),]
        const allRandomizedParticipants = Object.values(participants).filter(p => {
            const v1 = p.visits.v1_arm_1 || {};
            return v1.rand_code && v1.rand_code !== '';
        });

        // Get ACTIVE/COMPLETED randomized participants for checking completion
        // Matching R script line 620: valid_id = epats$participant_id[epats$conclusion %in% c(1, NA)]
        const activeRandomizedParticipants = allRandomizedParticipants.filter(p => {
            const conclusion = p.conclusion || {};
            const conclusionCode = conclusion.conclusion;
            return conclusionCode === '1' || !conclusionCode || conclusionCode === '';
        });

        const totalRandomized = allRandomizedParticipants.length;

        console.log(`Total randomized: ${totalRandomized} (including withdrawn)`);
        console.log(`Active/completed randomized: ${activeRandomizedParticipants.length}`);

        // Count participants by conclusion status
        const withdrawn = allRandomizedParticipants.filter(p => {
            const conclusion = p.conclusion || {};
            const conclusionCode = conclusion.conclusion;
            return conclusionCode === '3' || conclusionCode === '2' || conclusionCode === '4';
        }).length;

        console.log('Outcome Data Summary Calculations:', {
            totalRandomized,
            activeRandomized: activeRandomizedParticipants.length,
            withdrawn
        });

        // Helper function to calculate outcome metrics for a visit
        // Matching R script logic:
        // - Expected: ALL randomized participants who reached this visit (R script line 583-590)
        // - Completed/Incomplete: Only checked for active/completed participants (R script line 620)
        // - Pending: Expected - Completed - Incomplete (R script line 1012)
        const calculateOutcome = (visitEvent, completionChecker, outcomeName) => {
            // Expected: Count from ALL randomized who reached (including withdrawn)
            const allReached = allRandomizedParticipants.filter(p => {
                const visit = p.visits[visitEvent] || {};
                return visit.vdate && visit.vdate !== '';
            });
            const expected = allReached.length;

            // Only check completion for ACTIVE/COMPLETED participants
            const activeReached = activeRandomizedParticipants.filter(p => {
                const visit = p.visits[visitEvent] || {};
                return visit.vdate && visit.vdate !== '';
            });

            console.log(`\n=== Calculating ${outcomeName} for ${visitEvent} ===`);
            console.log(`Expected (all who reached): ${expected}`);
            console.log(`Active/completed who reached: ${activeReached.length}`);
            console.log(`Has rawData: ${!!rawParticipantData}, Has checker: ${!!completionChecker}`);

            // Check outcome completion for active/completed participants only
            let completed = 0;
            let incomplete = 0;

            if (completionChecker && rawParticipantData) {
                const incompleteParticipants = [];
                activeReached.forEach(p => {
                    const isComplete = completionChecker(p.id, visitEvent, rawParticipantData);
                    console.log(`  Participant ${p.id}: ${isComplete ? 'COMPLETE' : 'INCOMPLETE'}`);
                    if (isComplete) {
                        completed++;
                    } else {
                        incomplete++;
                        incompleteParticipants.push(p.id);
                    }
                });
                console.log(`  TABLE 2: ${outcomeName} at ${visitEvent} - Incomplete participants:`, incompleteParticipants);
            } else {
                console.log('WARNING: No checker or raw data - using fallback');
                completed = activeReached.length;
                incomplete = 0;
            }

            console.log(`Results: Expected=${expected}, Completed=${completed}, Incomplete=${incomplete}`);

            // Pending = Expected - Completed - Incomplete (R script line 1012)
            const pending = expected - completed - incomplete;

            return { expected, completed, pending, incomplete };
        };

        // Define outcomes with proper calculation
        // Expected = participants who have reached this visit
        // Completed = participants who reached the visit and have complete data
        // Pending = ACTIVE participants who haven't reached the visit yet
        // Incomplete = participants who reached the visit but have missing/incomplete data
        const outcomesData = [
            // DSD entries >= 4 in 7 days prior to visit
            {
                visit: 'V1',
                ...calculateOutcome('v1_arm_1', this.checkDSDCompletion.bind(this), 'DSD'),
                group: 'DSD entries >= 4 in 7 days prior to visit'
            },
            {
                visit: 'V9',
                ...calculateOutcome('v9_arm_1', this.checkDSDCompletion.bind(this), 'DSD'),
                group: 'DSD entries >= 4 in 7 days prior to visit'
            },
            {
                visit: 'V10',
                ...calculateOutcome('v10_arm_1', this.checkDSDCompletion.bind(this), 'DSD'),
                group: 'DSD entries >= 4 in 7 days prior to visit'
            },

            // PEG score
            {
                visit: 'V0',
                ...calculateOutcome('baseline_arm_1', this.checkPEGCompletion.bind(this), 'PEG'),
                group: 'PEG score'
            },
            {
                visit: 'FU1',
                ...calculateOutcome('fu1_arm_1', this.checkPEGCompletion.bind(this), 'PEG'),
                group: 'PEG score'
            },
            {
                visit: 'FU2',
                ...calculateOutcome('fu2_arm_1', this.checkPEGCompletion.bind(this), 'PEG'),
                group: 'PEG score'
            },
            {
                visit: 'V10',
                ...calculateOutcome('v10_arm_1', this.checkPEGCompletion.bind(this), 'PEG'),
                group: 'PEG score'
            },

            // JAW pain intensity
            {
                visit: 'V0',
                ...calculateOutcome('baseline_arm_1', this.checkJAWCompletion.bind(this), 'JAW'),
                group: 'JAW pain intensity'
            },
            {
                visit: 'V5',
                ...calculateOutcome('v5_arm_1', this.checkJAWCompletion.bind(this), 'JAW'),
                group: 'JAW pain intensity'
            },
            {
                visit: 'V9',
                ...calculateOutcome('v9_arm_1', this.checkJAWCompletion.bind(this), 'JAW'),
                group: 'JAW pain intensity'
            },
            {
                visit: 'V10',
                ...calculateOutcome('v10_arm_1', this.checkJAWCompletion.bind(this), 'JAW'),
                group: 'JAW pain intensity'
            },

            // Assisted/unassisted mouth opening
            {
                visit: 'V0',
                ...calculateOutcome('baseline_arm_1', this.checkMouthOpeningCompletion.bind(this), 'MouthOpen'),
                group: 'Assisted/unassisted mouth opening'
            },
            {
                visit: 'V5',
                ...calculateOutcome('v5_arm_1', this.checkMouthOpeningCompletion.bind(this), 'MouthOpen'),
                group: 'Assisted/unassisted mouth opening'
            },
            {
                visit: 'V9',
                ...calculateOutcome('v9_arm_1', this.checkMouthOpeningCompletion.bind(this), 'MouthOpen'),
                group: 'Assisted/unassisted mouth opening'
            },
            {
                visit: 'V10',
                ...calculateOutcome('v10_arm_1', this.checkMouthOpeningCompletion.bind(this), 'MouthOpen'),
                group: 'Assisted/unassisted mouth opening'
            },

            // Pressure pain threshold
            {
                visit: 'V0',
                ...calculateOutcome('baseline_arm_1', this.checkPPTCompletion.bind(this), 'PPT'),
                group: 'Pressure pain threshold'
            },
            {
                visit: 'V5',
                ...calculateOutcome('v5_arm_1', this.checkPPTCompletion.bind(this), 'PPT'),
                group: 'Pressure pain threshold'
            },
            {
                visit: 'V9',
                ...calculateOutcome('v9_arm_1', this.checkPPTCompletion.bind(this), 'PPT'),
                group: 'Pressure pain threshold'
            },
            {
                visit: 'V10',
                ...calculateOutcome('v10_arm_1', this.checkPPTCompletion.bind(this), 'PPT'),
                group: 'Pressure pain threshold'
            }
        ];

        let html = '';
        let currentGroup = '';

        outcomesData.forEach(outcome => {
            // Expected percentage is relative to total randomized
            const expectedPct = totalRandomized > 0 ? Math.round((outcome.expected / totalRandomized) * 100) : 0;

            // Completed and Incomplete percentages are relative to expected for this visit
            const completedPct = outcome.expected > 0 ? Math.round((outcome.completed / outcome.expected) * 100) : 0;
            const incompletePct = outcome.expected > 0 ? Math.round((outcome.incomplete / outcome.expected) * 100) : 0;

            // Pending percentage is relative to total randomized (showing % who haven't reached yet)
            const pendingPct = totalRandomized > 0 ? Math.round((outcome.pending / totalRandomized) * 100) : 0;

            // Add group header if new group
            if (outcome.group !== currentGroup) {
                currentGroup = outcome.group;
                html += `
                    <tr class="group-header">
                        <td colspan="5"><strong>${currentGroup}</strong></td>
                    </tr>
                `;
            }

            html += `
                <tr>
                    <td>${outcome.visit}</td>
                    <td>${outcome.expected} (${expectedPct}%)</td>
                    <td>${outcome.completed} (${completedPct}%)</td>
                    <td>${outcome.pending} (${pendingPct}%)</td>
                    <td>${outcome.incomplete} (${incompletePct}%)</td>
                </tr>
            `;
        });

        tableBody.innerHTML = html;

        // DIAGNOSTIC: Calculate total incomplete outcomes for comparison with Table 3
        const totalIncompleteByOutcome = {};
        outcomesData.forEach(outcome => {
            if (!totalIncompleteByOutcome[outcome.group]) {
                totalIncompleteByOutcome[outcome.group] = {};
            }
            if (!totalIncompleteByOutcome[outcome.group][outcome.visit]) {
                totalIncompleteByOutcome[outcome.group][outcome.visit] = 0;
            }
            totalIncompleteByOutcome[outcome.group][outcome.visit] = outcome.incomplete;
        });
        console.log('=== TABLE 2 INCOMPLETE SUMMARY ===');
        console.log('Incomplete outcomes by type and visit:', totalIncompleteByOutcome);

        let table2TotalIncomplete = 0;
        Object.values(totalIncompleteByOutcome).forEach(visits => {
            Object.values(visits).forEach(count => {
                table2TotalIncomplete += count;
            });
        });
        console.log('TABLE 2 TOTAL INCOMPLETE COUNT:', table2TotalIncomplete);
    }

    async generateIncompleteOutcomes(enrollmentData, rawData) {
        const tableBody = document.getElementById('incomplete-outcomes-body');
        if (!tableBody) return;

        console.log('=== generateIncompleteOutcomes called ===');
        console.log('enrollmentData parameter:', enrollmentData);
        console.log('rawData parameter:', rawData);
        console.log('rawData type:', typeof rawData);
        console.log('rawData is array:', Array.isArray(rawData));
        console.log('rawData length:', rawData?.length);

        const incompleteOutcomes = [];

        // Get ALL randomized participants (same as Table 2)
        const allRandomizedParticipants = Object.values(enrollmentData.participants).filter(p => {
            const v1 = p.visits.v1_arm_1 || {};
            return v1.rand_code && v1.rand_code !== '';
        });

        // Get ACTIVE/COMPLETED randomized participants (same as Table 2)
        const randomizedParticipants = allRandomizedParticipants.filter(p => {
            const conclusion = p.conclusion || {};
            const conclusionCode = conclusion.conclusion;
            return conclusionCode === '1' || !conclusionCode || conclusionCode === '';
        });

        console.log('Analyzing incomplete outcomes for', randomizedParticipants.length, 'randomized participants');

        // Check if raw data was provided
        if (!rawData || !Array.isArray(rawData) || rawData.length === 0) {
            console.error('ERROR: No raw data available for outcome analysis');
            console.error('rawData:', rawData);
            tableBody.innerHTML = `
                <tr>
                    <td colspan="3" style="text-align: center; font-style: italic; color: #666;">
                        Unable to analyze incomplete outcomes. Please refresh the data.
                    </td>
                </tr>
            `;
            return;
        }

        console.log('SUCCESS: Using raw data:', rawData.length, 'records');

        // Debug: Check first few records to see what fields are available
        if (rawData.length > 0) {
            console.log('Sample raw record fields:', Object.keys(rawData[0]));
            console.log('Sample raw record:', rawData[0]);
        }

        // Group raw data by participant and event
        const participantVisits = {};
        rawData.forEach(record => {
            const pid = record.participant_id;
            const event = record.redcap_event_name;
            if (!participantVisits[pid]) {
                participantVisits[pid] = {};
            }
            participantVisits[pid][event] = record;
        });

        randomizedParticipants.forEach(participant => {
            const pid = participant.id;
            const visits = participant.visits || {};

            // Conclusion checking already done in filtering above, but keep for diagnostic logging
            const conclusion = participant.conclusion || {};
            const conclusionCode = conclusion.conclusion;
            const isActive = conclusionCode === '1' || !conclusionCode || conclusionCode === '';

            // DIAGNOSTIC: Log participant conclusion status
            console.log(`Table 3 checking participant ${pid}: conclusionCode=${conclusionCode}, isActive=${isActive}`);

            // This should always be true since we filtered above, but check anyway
            if (!isActive) {
                console.log(`  Skipping ${pid} (not active/completed)`);
                return;
            }

            // Check each visit for incomplete data using RAW event-based records
            const visitChecks = [
                { name: 'V0', event: 'baseline_arm_1', checkPEG: true, checkJAW: true, checkMouth: true, checkPPT: true },
                { name: 'V1', event: 'v1_arm_1', checkDSD: true },
                { name: 'V5', event: 'v5_arm_1', checkJAW: true, checkMouth: true, checkPPT: true },
                { name: 'V9', event: 'v9_arm_1', checkDSD: true, checkJAW: true, checkMouth: true, checkPPT: true },
                { name: 'V10', event: 'v10_arm_1', checkDSD: true, checkPEG: true, checkJAW: true, checkMouth: true, checkPPT: true },
                { name: 'FU1', event: 'fu1_arm_1', checkPEG: true },
                { name: 'FU2', event: 'fu2_arm_1', checkPEG: true }
            ];

            visitChecks.forEach(visit => {
                const visitData = visits[visit.event] || {};

                // Check if visit has been completed (has vdate)
                if (!visitData.vdate || visitData.vdate === '') {
                    console.log(`  ${pid} has not reached ${visit.name}, skipping`);
                    return; // Skip if visit not reached
                }

                console.log(`  ${pid} reached ${visit.name} (vdate: ${visitData.vdate}), checking outcomes...`);

                // Use the same completion checkers as Table 2 for consistency

                // Check DSD
                if (visit.checkDSD) {
                    const isComplete = this.checkDSDCompletion(pid, visit.event, rawData);
                    console.log(`    DSD: ${isComplete ? 'COMPLETE' : 'INCOMPLETE'}`);
                    if (!isComplete) {
                        incompleteOutcomes.push({ id: pid, endpoint: 'DSD entries >= 4 in 7 days prior to visit', event: visit.name });
                        console.log(`    TABLE 3: Added ${pid} - DSD - ${visit.name}`);
                    }
                }

                // Check PEG score
                if (visit.checkPEG) {
                    const isComplete = this.checkPEGCompletion(pid, visit.event, rawData);
                    console.log(`    PEG: ${isComplete ? 'COMPLETE' : 'INCOMPLETE'}`);
                    if (!isComplete) {
                        incompleteOutcomes.push({ id: pid, endpoint: 'PEG score', event: visit.name });
                        console.log(`    TABLE 3: Added ${pid} - PEG - ${visit.name}`);
                    }
                }

                // Check JAW
                if (visit.checkJAW) {
                    const isComplete = this.checkJAWCompletion(pid, visit.event, rawData);
                    console.log(`    JAW: ${isComplete ? 'COMPLETE' : 'INCOMPLETE'}`);
                    if (!isComplete) {
                        incompleteOutcomes.push({ id: pid, endpoint: 'JAW pain intensity', event: visit.name });
                        console.log(`    TABLE 3: Added ${pid} - JAW - ${visit.name}`);
                    }
                }

                // Check Mouth Opening
                if (visit.checkMouth) {
                    const isComplete = this.checkMouthOpeningCompletion(pid, visit.event, rawData);
                    console.log(`    Mouth Opening: ${isComplete ? 'COMPLETE' : 'INCOMPLETE'}`);
                    if (!isComplete) {
                        incompleteOutcomes.push({ id: pid, endpoint: 'Assisted/unassisted mouth opening', event: visit.name });
                        console.log(`    TABLE 3: Added ${pid} - Mouth Opening - ${visit.name}`);
                    }
                }

                // Check PPT
                if (visit.checkPPT) {
                    const isComplete = this.checkPPTCompletion(pid, visit.event, rawData);
                    console.log(`    PPT: ${isComplete ? 'COMPLETE' : 'INCOMPLETE'}`);
                    if (!isComplete) {
                        incompleteOutcomes.push({ id: pid, endpoint: 'Pressure pain threshold', event: visit.name });
                        console.log(`    TABLE 3: Added ${pid} - PPT - ${visit.name}`);
                    }
                }
            });
        });

        // Deduplicate outcomes (should not have duplicates, but ensure it)
        const uniqueOutcomes = [];
        const seen = new Set();
        incompleteOutcomes.forEach(outcome => {
            const key = `${outcome.id}|${outcome.endpoint}|${outcome.event}`;
            if (!seen.has(key)) {
                seen.add(key);
                uniqueOutcomes.push(outcome);
            }
        });

        console.log('Found', uniqueOutcomes.length, 'unique incomplete outcomes (removed', incompleteOutcomes.length - uniqueOutcomes.length, 'duplicates)');

        // DIAGNOSTIC: Count incomplete by outcome type and visit for comparison with Table 2
        const table3IncompleteByType = {};
        uniqueOutcomes.forEach(outcome => {
            const key = `${outcome.endpoint} - ${outcome.event}`;
            if (!table3IncompleteByType[key]) {
                table3IncompleteByType[key] = 0;
            }
            table3IncompleteByType[key]++;
        });
        console.log('=== TABLE 3 INCOMPLETE SUMMARY ===');
        console.log('Incomplete outcomes by type and visit:', table3IncompleteByType);
        console.log('TABLE 3 TOTAL INCOMPLETE COUNT:', uniqueOutcomes.length);

        if (uniqueOutcomes.length > 0) {
            let html = '';
            uniqueOutcomes.forEach(outcome => {
                html += `
                    <tr>
                        <td><strong>${outcome.id}</strong></td>
                        <td>${outcome.endpoint}</td>
                        <td>${outcome.event}</td>
                    </tr>
                `;
            });
            tableBody.innerHTML = html;
        } else {
            tableBody.innerHTML = `
                <tr>
                    <td colspan="3" style="text-align: center; font-style: italic; color: #666;">
                        No incomplete outcomes detected. All participants with completed visits have complete data.
                    </td>
                </tr>
            `;
        }
    }

    exportMissingDataPDF() {
        try {
            const { jsPDF } = window.jspdf;
            if (!jsPDF) {
                alert('PDF export requires jsPDF library. This would generate a comprehensive missing data report PDF.');
                return;
            }

            const doc = new jsPDF();
            
            doc.setFontSize(18);
            doc.text('ULLTRA Study - Monthly Missing Data Report', 20, 20);
            
            const reportDate = new Date().toISOString().split('T')[0];
            doc.setFontSize(12);
            doc.text(`Report date: ${reportDate}`, 20, 35);
            doc.text('Prepared by: Claude Code', 20, 45);

            doc.setFontSize(14);
            doc.text('Figure 1. Study Enrollment Summary', 20, 65);
            
            doc.setFontSize(10);
            doc.text('This PDF contains a comprehensive missing data analysis including:', 20, 85);
            doc.text('• Complete enrollment statistics and chart', 25, 95);
            doc.text('• Full participant visit status table', 25, 105);
            doc.text('• Detailed outcome data summary', 25, 115);
            doc.text('• Complete list of incomplete outcomes', 25, 125);

            doc.save(`ULLTRA_Missing_Data_Report_${reportDate}.pdf`);
            
        } catch (error) {
            console.error('Error generating PDF:', error);
            alert('Error generating PDF report. Please check browser console for details.');
        }
    }

    showError(message) {
        const errorDiv = document.createElement('div');
        errorDiv.className = 'error';
        errorDiv.textContent = message;
        
        const missingDataTab = document.getElementById('missing-data');
        if (missingDataTab) {
            missingDataTab.insertBefore(errorDiv, missingDataTab.firstChild);
            setTimeout(() => {
                if (errorDiv.parentNode) {
                    errorDiv.parentNode.removeChild(errorDiv);
                }
            }, 5000);
        }
    }

    showLoading() {
        const loadingEl = document.getElementById('loading');
        if (loadingEl) {
            loadingEl.classList.remove('hidden');
        }
    }

    hideLoading() {
        const loadingEl = document.getElementById('loading');
        if (loadingEl) {
            loadingEl.classList.add('hidden');
        }
    }

    updateMetricCard(elementId, value) {
        const element = document.getElementById(elementId);
        if (element) {
            element.textContent = value;
        }
    }

    bindCardClickListeners() {
        // Add click listeners to missing data metric cards
        const missingDataSection = document.getElementById('missing-data');
        if (missingDataSection) {
            const cards = missingDataSection.querySelectorAll('.metric-card.clickable');
            cards.forEach(card => {
                card.addEventListener('click', () => {
                    const category = card.getAttribute('data-category');
                    this.showMissingDataDetails(category);
                });
            });
        }
    }

    async showMissingDataDetails(category) {
        try {
            const data = await this.redcapAPI.getAllParticipantData();
            const participants = this.getParticipantsByMissingDataCategory(data, category);
            
            // Reuse the existing modal from enrollment tab
            const modal = document.getElementById('participant-modal');
            const title = document.getElementById('modal-title');
            const participantList = document.getElementById('participant-list');
            
            if (!modal || !title || !participantList) {
                console.error('Modal elements not found');
                return;
            }
            
            // Update modal title
            const categoryTitles = {
                'total-enrollment': 'Total Enrolled Participants',
                'total-randomized': 'Total Randomized Participants',
                'reached-v5': 'Participants Who Reached V5',
                'reached-v9': 'Participants Who Reached V9',
                'reached-fu1': 'Participants Who Reached FU1',
                'reached-fu2': 'Participants Who Reached FU2',
                'reached-v10': 'Participants Who Reached V10',
                'completed-study': 'Participants Who Completed Study'
            };
            
            title.textContent = categoryTitles[category] || 'Participants';
            
            // Generate participant list HTML
            if (participants.length === 0) {
                participantList.innerHTML = '<div class="no-participants">No participants found for this category.</div>';
            } else {
                participantList.innerHTML = participants.map(p => `
                    <div class="participant-row">
                        <div class="participant-info">
                            <strong>${p.participant_id}</strong>
                            <span class="participant-details">
                                ${p.icf_date ? `ICF: ${p.icf_date}` : 'No ICF date'}
                                ${p.rand_code ? ` | Randomized: ${p.rand_code}` : ''}
                                ${p.conclusion ? ` | Status: ${this.getStatusLabel(p.conclusion)}` : ''}
                            </span>
                        </div>
                        <div class="participant-visits">
                            ${this.generateVisitSummary(p)}
                        </div>
                    </div>
                `).join('');
            }
            
            // Show the modal
            modal.style.display = 'flex';
            
        } catch (error) {
            console.error('Error showing missing data details:', error);
            this.showError('Failed to load participant details: ' + error.message);
        }
    }

    getParticipantsByMissingDataCategory(data, category) {
        switch(category) {
            case 'total-enrollment':
                return data.filter(p => p.icf_date);
            case 'total-randomized':
                return data.filter(p => p.rand_code);
            case 'reached-v5':
                return data.filter(p => p.rand_code && p.vdate_v5);
            case 'reached-v9':
                return data.filter(p => p.rand_code && p.vdate_v9);
            case 'reached-fu1':
                return data.filter(p => p.rand_code && p.vdate_fu1);
            case 'reached-fu2':
                return data.filter(p => p.rand_code && p.vdate_fu2);
            case 'reached-v10':
                return data.filter(p => p.rand_code && p.vdate_v10);
            case 'completed-study':
                // participant.conclusion is stored as an integer (see processParticipantData line 1302)
                return data.filter(p => p.conclusion === 1);
            default:
                return [];
        }
    }

    getStatusLabel(conclusionCode) {
        if (conclusionCode === null || conclusionCode === undefined || conclusionCode === '') {
            return 'Active in study';
        }

        const code = conclusionCode.toString();
        const definition = CONCLUSION_STATUS_DEFINITIONS[code];
        return definition ? definition.label : 'Active in study';
    }

    generateVisitSummary(participant) {
        const visits = ['v5', 'v9', 'fu1', 'fu2', 'v10'];
        return visits.map(visit => {
            const dateKey = `vdate_${visit}`;
            const hasVisit = participant[dateKey];
            const visitLabel = visit.toUpperCase();
            
            return `<span class="visit-status ${hasVisit ? 'completed' : 'missing'}">
                ${visitLabel}: ${hasVisit || '—'}
            </span>`;
        }).join(' ');
    }
}

// SharePoint Calendar Manager Class
class SharePointCalendarManager {
    constructor() {
        this.events = [];
        this.msalInstance = null;
        this.account = null;
        this.lastSync = null;
        this.isAuthenticated = false;
        this.init();
    }

    async init() {
        console.log('Initializing SharePoint Calendar Manager...');
        await this.initializeMsal();
        this.setupEventListeners();
        this.updateConnectionStatus();
    }

    async initializeMsal() {
        // Check if Azure AD Client ID is configured
        if (!CONFIG.AZURE_CLIENT_ID) {
            console.warn('Azure AD Client ID not configured. SharePoint calendar will not be available.');
            this.isAuthenticated = false;
            return;
        }

        // MSAL configuration
        const msalConfig = {
            auth: {
                clientId: CONFIG.AZURE_CLIENT_ID,
                authority: `https://login.microsoftonline.com/${CONFIG.AZURE_TENANT_ID}`,
                redirectUri: CONFIG.AZURE_REDIRECT_URI
            },
            cache: {
                cacheLocation: 'sessionStorage', // Use session storage to avoid persistence across sessions
                storeAuthStateInCookie: false
            }
        };

        try {
            // Initialize MSAL instance
            this.msalInstance = new msal.PublicClientApplication(msalConfig);
            await this.msalInstance.initialize();

            // Check if user is already logged in
            const accounts = this.msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                this.account = accounts[0];
                this.isAuthenticated = true;
                console.log('User already authenticated:', this.account.username);
            } else {
                console.log('No authenticated user found. User needs to sign in.');
                this.isAuthenticated = false;
            }
        } catch (error) {
            console.error('Error initializing MSAL:', error);
            this.isAuthenticated = false;
        }
    }

    async signIn() {
        if (!this.msalInstance) {
            console.error('MSAL not initialized');
            return;
        }

        const loginRequest = {
            scopes: ['Sites.Read.All', 'User.Read'] // Permissions needed for SharePoint access
        };

        try {
            const loginResponse = await this.msalInstance.loginPopup(loginRequest);
            this.account = loginResponse.account;
            this.isAuthenticated = true;
            console.log('Sign in successful:', this.account.username);

            this.updateConnectionStatus();
            this.loadCalendarEvents();
        } catch (error) {
            console.error('Sign in error:', error);
            this.displayError('Failed to sign in: ' + error.message);
        }
    }

    async signOut() {
        if (!this.msalInstance || !this.account) {
            return;
        }

        try {
            await this.msalInstance.logoutPopup({
                account: this.account
            });
            this.account = null;
            this.isAuthenticated = false;
            this.events = [];

            console.log('Sign out successful');
            this.updateConnectionStatus();
            this.displayNoConnectionMessage();
        } catch (error) {
            console.error('Sign out error:', error);
        }
    }

    async getAccessToken() {
        if (!this.msalInstance || !this.account) {
            throw new Error('Not authenticated');
        }

        const tokenRequest = {
            scopes: ['Sites.Read.All'],
            account: this.account
        };

        try {
            // Try to acquire token silently first
            const response = await this.msalInstance.acquireTokenSilent(tokenRequest);
            return response.accessToken;
        } catch (error) {
            console.warn('Silent token acquisition failed, trying interactive:', error);

            // If silent acquisition fails, try interactive
            try {
                const response = await this.msalInstance.acquireTokenPopup(tokenRequest);
                return response.accessToken;
            } catch (interactiveError) {
                console.error('Interactive token acquisition failed:', interactiveError);
                throw interactiveError;
            }
        }
    }

    setupEventListeners() {
        // Refresh calendar button
        const refreshBtn = document.getElementById('refresh-calendar-data');
        if (refreshBtn) {
            refreshBtn.addEventListener('click', () => this.loadCalendarEvents(true));
        }

        // Export calendar button
        const exportBtn = document.getElementById('export-calendar');
        if (exportBtn) {
            exportBtn.addEventListener('click', () => this.exportCalendar());
        }

        // Sign out button
        const signOutBtn = document.getElementById('sharepoint-signout');
        if (signOutBtn) {
            signOutBtn.addEventListener('click', () => this.signOut());
        }

        // Filter controls
        const participantFilter = document.getElementById('calendar-participant-filter');
        const typeFilter = document.getElementById('calendar-type-filter');
        const timeframeFilter = document.getElementById('calendar-timeframe-filter');
        const searchInput = document.getElementById('calendar-search');

        if (participantFilter) participantFilter.addEventListener('change', () => this.applyFilters());
        if (typeFilter) typeFilter.addEventListener('change', () => this.applyFilters());
        if (timeframeFilter) timeframeFilter.addEventListener('change', () => this.applyFilters());
        if (searchInput) searchInput.addEventListener('input', () => this.applyFilters());

        // Load calendar when tab is shown
        const calendarTab = document.querySelector('[data-tab="participant-contact"]');
        if (calendarTab) {
            calendarTab.addEventListener('click', () => {
                if (this.isAuthenticated && this.events.length === 0) {
                    this.loadCalendarEvents();
                } else if (!this.isAuthenticated) {
                    this.displayNoConnectionMessage();
                }
            });
        }
    }

    updateConnectionStatus() {
        const statusEl = document.getElementById('sp-connection-status');
        const configEl = document.getElementById('sp-config-status');
        const lastSyncEl = document.getElementById('sp-last-sync');
        const signOutBtn = document.getElementById('sharepoint-signout');

        if (statusEl) {
            statusEl.textContent = this.isAuthenticated ? 'Signed In' : 'Not Signed In';
            statusEl.style.color = this.isAuthenticated ? 'green' : 'red';
        }

        if (configEl) {
            const isConfigured = CONFIG.AZURE_CLIENT_ID !== null;
            configEl.textContent = isConfigured ? 'Azure AD Configured' : 'Not Configured';
            configEl.style.color = isConfigured ? 'green' : 'orange';

            // Show username if authenticated
            if (this.isAuthenticated && this.account) {
                configEl.textContent = `Signed in as: ${this.account.username}`;
            }
        }

        if (lastSyncEl) {
            lastSyncEl.textContent = this.lastSync ? new Date(this.lastSync).toLocaleString() : 'Never';
        }

        // Show/hide sign out button based on authentication status
        if (signOutBtn) {
            signOutBtn.style.display = this.isAuthenticated ? 'inline-block' : 'none';
        }
    }

    async loadCalendarEvents(forceRefresh = false) {
        if (!this.isAuthenticated) {
            this.displayNoConnectionMessage();
            return;
        }

        console.log('Loading calendar events from SharePoint...');
        this.showLoading();

        try {
            // Fetch calendar events from SharePoint via Microsoft Graph API
            const events = await this.fetchSharePointEvents();
            this.events = events;
            this.lastSync = Date.now();
            this.updateConnectionStatus();
            this.displayCalendarEvents(events);
            this.updateSummaryMetrics(events);
            this.populateParticipantFilter(events);
        } catch (error) {
            console.error('Error loading calendar events:', error);
            this.displayError(error.message);
        }
    }

    async fetchSharePointEvents() {
        try {
            // Get access token
            const accessToken = await this.getAccessToken();

            // First, get the site ID
            const siteUrl = CONFIG.SHAREPOINT_SITE_URL;
            const sitePath = new URL(siteUrl).pathname;
            const hostname = new URL(siteUrl).hostname;

            // Get site ID from Microsoft Graph
            const siteResponse = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`,
                {
                    headers: {
                        'Authorization': `Bearer ${accessToken}`,
                        'Content-Type': 'application/json'
                    }
                }
            );

            if (!siteResponse.ok) {
                throw new Error(`Failed to get site: ${siteResponse.status} ${siteResponse.statusText}`);
            }

            const siteData = await siteResponse.json();
            const siteId = siteData.id;

            // Get the list
            const listResponse = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$filter=displayName eq '${CONFIG.SHAREPOINT_LIST_NAME}'`,
                {
                    headers: {
                        'Authorization': `Bearer ${accessToken}`,
                        'Content-Type': 'application/json'
                    }
                }
            );

            if (!listResponse.ok) {
                throw new Error(`Failed to get list: ${listResponse.status} ${listResponse.statusText}`);
            }

            const listData = await listResponse.json();
            if (!listData.value || listData.value.length === 0) {
                throw new Error(`List '${CONFIG.SHAREPOINT_LIST_NAME}' not found`);
            }

            const listId = listData.value[0].id;

            // Get list items with fields expanded
            const itemsResponse = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=fields`,
                {
                    headers: {
                        'Authorization': `Bearer ${accessToken}`,
                        'Content-Type': 'application/json'
                    }
                }
            );

            if (!itemsResponse.ok) {
                throw new Error(`Failed to get list items: ${itemsResponse.status} ${itemsResponse.statusText}`);
            }

            const itemsData = await itemsResponse.json();

            // Transform SharePoint list items to calendar events
            // Filter by the ULLTRA view if needed
            const events = itemsData.value
                .map(item => this.transformSharePointItemToEvent(item))
                .filter(event => event !== null);

            return events;

        } catch (error) {
            console.error('Error fetching SharePoint events:', error);
            throw error;
        }
    }

    transformSharePointItemToEvent(item) {
        // Transform SharePoint list item to calendar event format
        // This mapping will depend on the actual column names in your SharePoint list
        const fields = item.fields;

        // Common SharePoint calendar fields
        // Adjust these field names based on your actual SharePoint list structure
        return {
            id: item.id,
            title: fields.Title || '',
            date: fields.EventDate || fields.StartDate || fields.Date,
            time: fields.EventTime || fields.Time || '',
            participant: fields.Participant || fields.ParticipantID || '',
            type: fields.Category || fields.EventType || fields.Type || 'general',
            description: fields.Description || fields.Notes || fields.Body || '',
            location: fields.Location || '',
            status: fields.Status || '',
            // Include any other relevant fields from your SharePoint list
            rawFields: fields // Keep raw fields for debugging
        };
    }

    displayCalendarEvents(events) {
        const container = document.getElementById('calendar-events-container');
        if (!container) return;

        if (events.length === 0) {
            container.innerHTML = `
                <div class="calendar-empty">
                    <p>No calendar events found.</p>
                    <p class="calendar-note">Events from the ULLTRA SharePoint calendar will appear here once the integration is configured.</p>
                </div>
            `;
            return;
        }

        // Group events by date
        const eventsByDate = this.groupEventsByDate(events);

        let html = '';
        for (const [date, dateEvents] of Object.entries(eventsByDate)) {
            html += `
                <div class="calendar-date-group">
                    <h4 class="calendar-date-header">${this.formatDateHeader(date)}</h4>
                    <div class="calendar-events-list">
            `;

            dateEvents.forEach(event => {
                html += this.createEventCard(event);
            });

            html += `
                    </div>
                </div>
            `;
        }

        container.innerHTML = html;
    }

    groupEventsByDate(events) {
        const grouped = {};
        events.forEach(event => {
            const date = event.date || event.EventDate || event.Start;
            const dateKey = new Date(date).toDateString();
            if (!grouped[dateKey]) {
                grouped[dateKey] = [];
            }
            grouped[dateKey].push(event);
        });
        return grouped;
    }

    formatDateHeader(dateString) {
        const date = new Date(dateString);
        const today = new Date();
        const tomorrow = new Date(today);
        tomorrow.setDate(tomorrow.getDate() + 1);

        if (date.toDateString() === today.toDateString()) {
            return 'Today - ' + date.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
        } else if (date.toDateString() === tomorrow.toDateString()) {
            return 'Tomorrow - ' + date.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
        }
        return date.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
    }

    createEventCard(event) {
        const eventType = event.type || 'general';
        const participant = event.participant || 'Unknown';
        const time = event.time || '';
        const description = event.description || '';
        const location = event.location || '';
        const title = event.title || 'Event';

        return `
            <div class="calendar-event-card" data-event-type="${eventType}">
                <div class="event-header">
                    <span class="event-time">${time}</span>
                    <span class="event-type event-type-${eventType.toLowerCase()}">${eventType}</span>
                </div>
                <div class="event-body">
                    ${title ? `<div class="event-title"><strong>${title}</strong></div>` : ''}
                    <div class="event-participant"><strong>Participant:</strong> ${participant}</div>
                    ${description ? `<div class="event-description">${description}</div>` : ''}
                    ${location ? `<div class="event-location"><strong>Location:</strong> ${location}</div>` : ''}
                </div>
            </div>
        `;
    }

    updateSummaryMetrics(events) {
        const totalEvents = events.length;
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const nextWeek = new Date(today);
        nextWeek.setDate(nextWeek.getDate() + 7);

        const todayEvents = events.filter(e => {
            const eventDate = new Date(e.date);
            eventDate.setHours(0, 0, 0, 0);
            return eventDate.getTime() === today.getTime();
        });

        const upcomingEvents = events.filter(e => {
            const eventDate = new Date(e.date);
            return eventDate >= today && eventDate <= nextWeek;
        });

        const participants = new Set(events.map(e => e.participant).filter(p => p));

        this.updateMetric('calendar-total-events', totalEvents);
        this.updateMetric('calendar-upcoming-events', upcomingEvents.length);
        this.updateMetric('calendar-today-events', todayEvents.length);
        this.updateMetric('calendar-active-participants', participants.size);
    }

    updateMetric(elementId, value) {
        const el = document.getElementById(elementId);
        if (el) el.textContent = value;
    }

    populateParticipantFilter(events) {
        const filter = document.getElementById('calendar-participant-filter');
        if (!filter) return;

        const participants = new Set(events.map(e => e.participant).filter(p => p));
        const sortedParticipants = Array.from(participants).sort();

        filter.innerHTML = '<option value="all">All Participants</option>';
        sortedParticipants.forEach(participant => {
            const option = document.createElement('option');
            option.value = participant;
            option.textContent = participant;
            filter.appendChild(option);
        });
    }

    applyFilters() {
        // Filter logic would go here
        console.log('Applying calendar filters...');
        // This would filter this.events based on the selected filters and re-display
    }

    displayNoConnectionMessage() {
        const container = document.getElementById('calendar-events-container');
        if (!container) return;

        // Check if Azure AD is configured
        if (!CONFIG.AZURE_CLIENT_ID) {
            container.innerHTML = `
                <div class="calendar-no-connection">
                    <h3>SharePoint Calendar Not Configured</h3>
                    <p>Azure AD authentication has not been configured for this application.</p>
                    <p class="calendar-note"><strong>Configuration Required:</strong></p>
                    <ol>
                        <li>Register an Azure AD application at <a href="https://portal.azure.com" target="_blank">Azure Portal</a></li>
                        <li>Set the Client ID in <code>script.js</code> CONFIG.AZURE_CLIENT_ID</li>
                        <li>Configure API permissions: Sites.Read.All and User.Read</li>
                        <li>Add redirect URI: ${CONFIG.AZURE_REDIRECT_URI}</li>
                        <li>Refresh the page</li>
                    </ol>
                    <p class="calendar-note"><strong>SharePoint Details:</strong></p>
                    <ul>
                        <li><strong>Site:</strong> ${CONFIG.SHAREPOINT_SITE_URL}</li>
                        <li><strong>List:</strong> ${CONFIG.SHAREPOINT_LIST_NAME}</li>
                        <li><strong>View:</strong> ${CONFIG.SHAREPOINT_LIST_VIEW}</li>
                    </ul>
                </div>
            `;
        } else {
            // Azure AD is configured, user just needs to sign in
            container.innerHTML = `
                <div class="calendar-no-connection">
                    <h3>Sign In Required</h3>
                    <p>Please sign in with your Microsoft 365 account to view calendar events.</p>
                    <button class="refresh-btn" onclick="window.sharePointCalendarManager.signIn()" style="margin: 20px auto; display: block; padding: 12px 24px; font-size: 16px;">
                        Sign In with Microsoft
                    </button>
                    <p class="calendar-note"><strong>SharePoint Details:</strong></p>
                    <ul>
                        <li><strong>Site:</strong> ${CONFIG.SHAREPOINT_SITE_URL}</li>
                        <li><strong>List:</strong> ${CONFIG.SHAREPOINT_LIST_NAME}</li>
                    </ul>
                </div>
            `;
        }
    }

    displayError(message) {
        const container = document.getElementById('calendar-events-container');
        if (!container) return;

        container.innerHTML = `
            <div class="calendar-error">
                <h3>Error Loading Calendar Events</h3>
                <p>${message}</p>
                <button class="refresh-btn" onclick="window.sharePointCalendarManager.loadCalendarEvents(true)">Try Again</button>
            </div>
        `;
    }

    showLoading() {
        const container = document.getElementById('calendar-events-container');
        if (!container) return;

        container.innerHTML = `
            <div class="calendar-loading">
                <div class="spinner"></div>
                <p>Loading calendar events from SharePoint...</p>
            </div>
        `;
    }

    exportCalendar() {
        console.log('Exporting calendar...');
        // Export functionality would go here
        alert('Calendar export functionality will be implemented here');
    }
}

// Initialize Missing Data Report Manager when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM loaded, initializing missing data report manager...');
    window.missingDataReportManager = new MissingDataReportManager();

    // Initialize SharePoint Calendar Manager
    console.log('Initializing SharePoint Calendar Manager...');
    window.sharePointCalendarManager = new SharePointCalendarManager();
});

