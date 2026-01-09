import React, { useState, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, Download, AlertTriangle, CheckCircle, XCircle, Settings, Users, ChevronDown, ChevronUp, Info, Filter } from 'lucide-react';

// Main Application Component
export default function ClinicianConsolidationTool() {
  // File state
  const [files, setFiles] = useState({
    geoSpatial: null,
    regionalAuthority: null,
    supportSite: null,
    ldgLedger: null
  });

  // Parsed data state
  const [data, setData] = useState({
    geoSpatial: [],
    regionalAuthority: [],
    supportSite: [],
    ldgLedger: []
  });

  // Validation warnings
  const [warnings, setWarnings] = useState({
    geoSpatial: [],
    regionalAuthority: [],
    supportSite: [],
    ldgLedger: []
  });

  // Processing state
  const [isProcessed, setIsProcessed] = useState(false);
  const [processedData, setProcessedData] = useState({ primaryCare: [], specialist: [] });
  const [dataQualityIssues, setDataQualityIssues] = useState([]);
  const [exclusions, setExclusions] = useState([]);

  // Configuration state
  const [specialtyMapping, setSpecialtyMapping] = useState({});
  const [deploymentTeam, setDeploymentTeam] = useState(['Ryland Steel', 'Art Dunham']);
  const [newTeamMember, setNewTeamMember] = useState('');
  const [showConfig, setShowConfig] = useState(false);
  const [activeTab, setActiveTab] = useState('upload');

  // Expected columns for each file
  const expectedColumns = {
    geoSpatial: ['CPSO', 'e-Referral', 'Hospital', 'Postal Code', 'Lead/Reach', 'Region', 'LDG', 'LDG Lead Org', 'Specialty', 'Type of Specialty'],
    regionalAuthority: ['clinicianProfessionalId', 'clinicianFirstName', 'clinicianSurname', 'siteNum', 'siteName', 'healthRegion', 'services', 'postalCode', 'eReferrals', 'eConsults'],
    supportSite: ['siteNum', 'siteName', 'clinicianType', 'professionalId', 'userFullName', 'username', 'referralLastSent'],
    ldgLedger: ['First Name', 'Last Name', 'CPSO #', 'eReferral Solution', 'Ocean Site Number (eReferral Ontario only)', 'Directory Listing Name (eReferral Ontario only)', 'Role (Sender, Receiver, Both)', 'Primary Specialty Pathway', 'Date Onboarding Completed']
  };

  // Parse uploaded file
  const parseFile = useCallback((file, fileType) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });

          // Get the first sheet (or specific sheet for LDG Ledger)
          let sheetName = workbook.SheetNames[0];
          if (fileType === 'ldgLedger') {
            const ledgerSheet = workbook.SheetNames.find(name =>
              name.toLowerCase().includes('onboarding') || name.toLowerCase().includes('ledger')
            );
            if (ledgerSheet) sheetName = ledgerSheet;
          }

          const sheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: '' });

          // Validate columns
          const fileWarnings = [];
          const actualColumns = jsonData.length > 0 ? Object.keys(jsonData[0]) : [];
          const expected = expectedColumns[fileType];

          const missingColumns = expected.filter(col =>
            !actualColumns.some(actual => actual.toLowerCase().trim() === col.toLowerCase().trim())
          );

          if (missingColumns.length > 0) {
            fileWarnings.push(`Missing expected columns: ${missingColumns.join(', ')}`);
          }

          resolve({ data: jsonData, warnings: fileWarnings });
        } catch (error) {
          reject(new Error(`Failed to parse file: ${error.message}`));
        }
      };
      reader.onerror = () => reject(new Error('Failed to read file'));
      reader.readAsArrayBuffer(file);
    });
  }, []);

  // Handle file upload
  const handleFileUpload = useCallback(async (fileType, file) => {
    if (!file) return;

    try {
      const result = await parseFile(file, fileType);

      // Check for duplicate CPSOs in Geo-Spatial LDG
      let fileWarnings = [...result.warnings];
      if (fileType === 'geoSpatial' && result.data.length > 0) {
        const cpsoCount = {};
        result.data.forEach(row => {
          const cpso = String(row['CPSO'] || '').replace(/\s/g, '').trim();
          if (cpso) {
            cpsoCount[cpso] = (cpsoCount[cpso] || 0) + 1;
          }
        });

        const duplicateCPSOs = Object.entries(cpsoCount).filter(([_, count]) => count > 1);
        const totalDuplicateRows = duplicateCPSOs.reduce((sum, [_, count]) => sum + count, 0);

        if (duplicateCPSOs.length > 0) {
          fileWarnings.push(`${totalDuplicateRows} rows contain duplicate CPSOs (${duplicateCPSOs.length} unique CPSOs appear multiple times). Last occurrence will be used for lookups.`);
        }
      }

      setFiles(prev => ({ ...prev, [fileType]: file }));
      setData(prev => ({ ...prev, [fileType]: result.data }));
      setWarnings(prev => ({ ...prev, [fileType]: fileWarnings }));
      setIsProcessed(false);

      // If Geo-Spatial LDG, extract unique specialties for mapping
      if (fileType === 'geoSpatial' && result.data.length > 0) {
        const specialties = [...new Set(result.data.map(row => row['Specialty']).filter(Boolean))];
        const defaultMapping = {};
        specialties.forEach(spec => {
          defaultMapping[spec] = 'specialist';
        });
        setSpecialtyMapping(defaultMapping);
      }
    } catch (error) {
      alert(`Error loading file: ${error.message}`);
    }
  }, [parseFile]);

  // Normalize CPSO (strip spaces)
  const normalizeCPSO = (cpso) => {
    if (!cpso) return '';
    return String(cpso).replace(/\s/g, '').trim();
  };

  // Parse full name into first and last name
  const parseFullName = (fullName) => {
    if (!fullName) return { firstName: '', lastName: '' };
    const parts = String(fullName).trim().split(' ');
    const firstName = parts[0] || '';
    const lastName = parts.slice(1).join(' ') || '';
    return { firstName, lastName };
  };

  // Check if record should be excluded
  const shouldExclude = (record, siteName) => {
    // Check deployment team
    const fullName = `${record.firstName || ''} ${record.lastName || ''}`.trim();
    for (const member of deploymentTeam) {
      if (fullName.toLowerCase() === member.toLowerCase()) {
        return { excluded: true, reason: `Deployment team member: ${member}` };
      }
    }

    // Check demo site
    if (siteName && siteName.toLowerCase().includes('demo')) {
      return { excluded: true, reason: `Demo site: ${siteName}` };
    }

    return { excluded: false, reason: null };
  };

  // Process all data
  const processData = useCallback(() => {
    const geoSpatialMap = new Map();
    const qualityIssues = [];
    const excludedRecords = [];
    const allRecords = [];

    // Build Geo-Spatial lookup map
    data.geoSpatial.forEach(row => {
      const cpso = normalizeCPSO(row['CPSO']);
      if (cpso) {
        geoSpatialMap.set(cpso, {
          cpso,
          eReferral: row['e-Referral'],
          hospital: row['Hospital'],
          postalCode: row['Postal Code'],
          leadReach: row['Lead/Reach'],
          region: row['Region'],
          ldg: row['LDG'],
          ldgLeadOrg: row['LDG Lead Org'],
          specialty: row['Specialty'],
          typeOfSpecialty: row['Type of Specialty']
        });
      }
    });

    // Track unique clinician-site combinations (CPSO + Site Number only)
    const recordKey = (cpso, siteNum) => `${cpso}|${siteNum}`;
    const processedKeys = new Set();

    // Process Regional Authority data (Specialists)
    data.regionalAuthority.forEach(row => {
      const cpso = normalizeCPSO(row['clinicianProfessionalId']);
      if (!cpso) return;

      const siteNum = row['siteNum'] || '';
      const siteName = row['siteName'] || '';
      const key = recordKey(cpso, siteNum);

      if (processedKeys.has(key)) return;
      processedKeys.add(key);

      const geoData = geoSpatialMap.get(cpso);
      const inGeoSpatial = !!geoData;

      if (!inGeoSpatial) {
        qualityIssues.push({
          type: 'NOT_IN_GEOSPATIAL',
          source: 'Regional Authority',
          cpso,
          name: `${row['clinicianFirstName'] || ''} ${row['clinicianSurname'] || ''}`.trim(),
          details: `CPSO ${cpso} found in Regional Authority but not in Geo-Spatial LDG`
        });
      }

      const record = {
        Professional_ID: cpso,
        First_Name: row['clinicianFirstName'] || (geoData ? '' : ''),
        Last_Name: row['clinicianSurname'] || (geoData ? '' : ''),
        Clinician_Type: '',
        CPSO_Specialty: geoData?.specialty || row['services'] || '',
        Type_of_Specialty: geoData?.typeOfSpecialty || '',
        Specialty_Pathway: 'Not Available - Regional CRM not provided',
        LDG_Name: geoData?.ldg || '',
        LDG_Lead_Org: geoData?.ldgLeadOrg || '',
        Region: geoData?.region || region,
        Network: 'OH',
        Solution_Type: '',
        Ocean_Site_Number: siteNum,
        Ocean_Site_Name: siteName,
        Hospital: geoData?.hospital || '',
        Address_Postal: geoData?.postalCode || row['postalCode'] || '',
        Lead_Reach: geoData?.leadReach || '',
        Role: '',
        Date_Onboarded: '',
        Training_Date: 'Not Available - Regional CRM not provided',
        Last_Referral_Sent: '',
        Regional_Dedupe_Flag: false,
        Provincial_Dedupe_Flag: false,
        Exclusion_Flag: false,
        Exclusion_Reason: '',
        In_GeoSpatial_LDG: inGeoSpatial,
        Data_Sources: inGeoSpatial ? 'Geo-Spatial,RA' : 'RA',
        _specialty: geoData?.specialty || 'Unknown'
      };

      const exclusionCheck = shouldExclude({ firstName: record.First_Name, lastName: record.Last_Name }, siteName);
      if (exclusionCheck.excluded) {
        record.Exclusion_Flag = true;
        record.Exclusion_Reason = exclusionCheck.reason;
        excludedRecords.push({ ...record, source: 'Regional Authority' });
      }

      allRecords.push(record);
    });

    // Process Support Site Analytics data (Primary Care)
    data.supportSite.forEach(row => {
      const cpso = normalizeCPSO(row['professionalId']);
      if (!cpso) return;

      const siteNum = row['siteNum'] || '';
      const siteName = row['siteName'] || '';
      const geoData = geoSpatialMap.get(cpso);
      const key = recordKey(cpso, siteNum);

      // Check if record already exists by CPSO + Site (consistent with LDG Ledger approach)
      const existingIdx = allRecords.findIndex(r =>
        r.Professional_ID === cpso && r.Ocean_Site_Number === siteNum
      );

      if (existingIdx >= 0) {
        // Update existing record with Support Site data
        allRecords[existingIdx].Last_Referral_Sent = row['referralLastSent'] || '';
        allRecords[existingIdx].Clinician_Type = row['clinicianType'] || allRecords[existingIdx].Clinician_Type;
        if (!allRecords[existingIdx].Data_Sources.includes('Support Site')) {
          allRecords[existingIdx].Data_Sources += ',Support Site';
        }
        return;
      }

      // Only create new record if no existing match
      if (processedKeys.has(key)) return;
      processedKeys.add(key);

      const inGeoSpatial = !!geoData;
      const { firstName, lastName } = parseFullName(row['userFullName']);

      if (!inGeoSpatial) {
        qualityIssues.push({
          type: 'NOT_IN_GEOSPATIAL',
          source: 'Support Site Analytics',
          cpso,
          name: row['userFullName'] || '',
          details: `CPSO ${cpso} found in Support Site Analytics but not in Geo-Spatial LDG`
        });
      }

      const record = {
        Professional_ID: cpso,
        First_Name: firstName,
        Last_Name: lastName,
        Clinician_Type: row['clinicianType'] || '',
        CPSO_Specialty: geoData?.specialty || '',
        Type_of_Specialty: geoData?.typeOfSpecialty || '',
        Specialty_Pathway: 'Not Available - Regional CRM not provided',
        LDG_Name: geoData?.ldg || '',
        LDG_Lead_Org: geoData?.ldgLeadOrg || '',
        Region: geoData?.region || '',
        Network: 'OH',
        Solution_Type: '',
        Ocean_Site_Number: siteNum,
        Ocean_Site_Name: siteName,
        Hospital: geoData?.hospital || '',
        Address_Postal: geoData?.postalCode || '',
        Lead_Reach: geoData?.leadReach || '',
        Role: '',
        Date_Onboarded: '',
        Training_Date: 'Not Available - Regional CRM not provided',
        Last_Referral_Sent: row['referralLastSent'] || '',
        Regional_Dedupe_Flag: false,
        Provincial_Dedupe_Flag: false,
        Exclusion_Flag: false,
        Exclusion_Reason: '',
        In_GeoSpatial_LDG: inGeoSpatial,
        Data_Sources: inGeoSpatial ? 'Geo-Spatial,Support Site' : 'Support Site',
        _specialty: geoData?.specialty || 'Unknown'
      };

      const exclusionCheck = shouldExclude({ firstName, lastName }, siteName);
      if (exclusionCheck.excluded) {
        record.Exclusion_Flag = true;
        record.Exclusion_Reason = exclusionCheck.reason;
        excludedRecords.push({ ...record, source: 'Support Site Analytics' });
      }

      allRecords.push(record);
    });

    // Process LDG Onboarding Ledger
    data.ldgLedger.forEach(row => {
      const cpso = normalizeCPSO(row['CPSO #']);
      if (!cpso) return;

      const siteNum = row['Ocean Site Number (eReferral Ontario only)'] || '';
      const siteName = row['Directory Listing Name (eReferral Ontario only)'] || '';
      const geoData = geoSpatialMap.get(cpso);
      const key = recordKey(cpso, siteNum);

      // Check if record already exists by CPSO + Site
      const existingIdx = allRecords.findIndex(r =>
        r.Professional_ID === cpso && r.Ocean_Site_Number === siteNum
      );

      if (existingIdx >= 0) {
        // Update existing record with LDG Ledger data
        allRecords[existingIdx].Solution_Type = row['eReferral Solution'] || allRecords[existingIdx].Solution_Type;
        allRecords[existingIdx].Role = row['Role (Sender, Receiver, Both)'] || allRecords[existingIdx].Role;
        allRecords[existingIdx].Specialty_Pathway = row['Primary Specialty Pathway'] || allRecords[existingIdx].Specialty_Pathway;
        if (row['Date Onboarding Completed']) {
          const existingDate = allRecords[existingIdx].Date_Onboarded;
          const newDate = row['Date Onboarding Completed'];
          if (!existingDate || newDate < existingDate) {
            allRecords[existingIdx].Date_Onboarded = newDate;
          }
        }
        if (!allRecords[existingIdx].Data_Sources.includes('LDG Ledger')) {
          allRecords[existingIdx].Data_Sources += ',LDG Ledger';
        }
        return;
      }

      if (processedKeys.has(key)) return;
      processedKeys.add(key);

      const inGeoSpatial = !!geoData;

      if (!inGeoSpatial) {
        qualityIssues.push({
          type: 'NOT_IN_GEOSPATIAL',
          source: 'LDG Onboarding Ledger',
          cpso,
          name: `${row['First Name'] || ''} ${row['Last Name'] || ''}`.trim(),
          details: `CPSO ${cpso} found in LDG Onboarding Ledger but not in Geo-Spatial LDG`
        });
      }

      const record = {
        Professional_ID: cpso,
        First_Name: row['First Name'] || '',
        Last_Name: row['Last Name'] || '',
        Clinician_Type: '',
        CPSO_Specialty: geoData?.specialty || '',
        Type_of_Specialty: geoData?.typeOfSpecialty || '',
        Specialty_Pathway: row['Primary Specialty Pathway'] || '',
        LDG_Name: geoData?.ldg || '',
        LDG_Lead_Org: geoData?.ldgLeadOrg || '',
        Region: geoData?.region || '',
        Network: '',
        Solution_Type: row['eReferral Solution'] || '',
        Ocean_Site_Number: siteNum,
        Ocean_Site_Name: siteName,
        Hospital: geoData?.hospital || '',
        Address_Postal: geoData?.postalCode || '',
        Lead_Reach: geoData?.leadReach || '',
        Role: row['Role (Sender, Receiver, Both)'] || '',
        Date_Onboarded: row['Date Onboarding Completed'] || '',
        Training_Date: 'Not Available - Regional CRM not provided',
        Last_Referral_Sent: '',
        Regional_Dedupe_Flag: false,
        Provincial_Dedupe_Flag: false,
        Exclusion_Flag: false,
        Exclusion_Reason: '',
        In_GeoSpatial_LDG: inGeoSpatial,
        Data_Sources: inGeoSpatial ? 'Geo-Spatial,LDG Ledger' : 'LDG Ledger',
        _specialty: geoData?.specialty || 'Unknown'
      };

      allRecords.push(record);
    });

    // Sort and set dedupe flags
    allRecords.sort((a, b) => {
      const pidA = String(a.Professional_ID || '');
      const pidB = String(b.Professional_ID || '');
      if (pidA !== pidB) return pidA.localeCompare(pidB);

      const regionA = String(a.Region || '');
      const regionB = String(b.Region || '');
      if (regionA !== regionB) return regionA.localeCompare(regionB);

      const siteA = String(a.Ocean_Site_Number || '');
      const siteB = String(b.Ocean_Site_Number || '');
      return siteA.localeCompare(siteB);
    });

    const seenRegional = new Map(); // cpso+region -> true
    const seenProvincial = new Set(); // cpso -> true

    allRecords.forEach(record => {
      const cpso = record.Professional_ID;
      const region = record.Region;
      const regionalKey = `${cpso}|${region}`;

      if (!seenRegional.has(regionalKey)) {
        record.Regional_Dedupe_Flag = true;
        seenRegional.set(regionalKey, true);
      }

      if (!seenProvincial.has(cpso)) {
        record.Provincial_Dedupe_Flag = true;
        seenProvincial.add(cpso);
      }
    });

    // Split into Primary Care and Specialist based on specialty mapping
    const primaryCare = [];
    const specialist = [];

    allRecords.forEach(record => {
      const specialty = record._specialty;
      const mapping = specialtyMapping[specialty];

      // Remove internal field before output
      const outputRecord = { ...record };
      delete outputRecord._specialty;

      if (mapping === 'primaryCare') {
        primaryCare.push(outputRecord);
      } else if (mapping === 'specialist') {
        specialist.push(outputRecord);
      } else {
        // Unknown specialty - default to specialist
        specialist.push(outputRecord);
      }
    });

    setProcessedData({ primaryCare, specialist });
    setDataQualityIssues(qualityIssues);
    setExclusions(excludedRecords);
    setIsProcessed(true);
    setActiveTab('summary');
  }, [data, specialtyMapping, deploymentTeam]);

  // Export to file
  const exportData = (dataArray, filename, format) => {
    const ws = XLSX.utils.json_to_sheet(dataArray);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');

    if (format === 'xlsx') {
      XLSX.writeFile(wb, `${filename}.xlsx`);
    } else {
      XLSX.writeFile(wb, `${filename}.csv`, { bookType: 'csv' });
    }
  };

  // Calculate summary statistics
  const summary = useMemo(() => {
    if (!isProcessed) return null;

    const allRecords = [...processedData.primaryCare, ...processedData.specialist];
    const nonExcluded = allRecords.filter(r => !r.Exclusion_Flag);

    const regionCounts = {};
    nonExcluded.filter(r => r.Regional_Dedupe_Flag).forEach(r => {
      const region = r.Region || 'Unknown';
      regionCounts[region] = (regionCounts[region] || 0) + 1;
    });

    const ldgCounts = {};
    nonExcluded.filter(r => r.Provincial_Dedupe_Flag).forEach(r => {
      const ldg = r.LDG_Name || 'Unknown';
      ldgCounts[ldg] = (ldgCounts[ldg] || 0) + 1;
    });

    // Count operational sources for merge statistics
    const countOperationalSources = (record) => {
      const sources = record.Data_Sources || '';
      let count = 0;
      if (sources.includes('RA')) count++;
      if (sources.includes('Support Site')) count++;
      if (sources.includes('LDG Ledger')) count++;
      return count;
    };

    // Calculate merge counts for Primary Care
    const primaryCareFullMerge = processedData.primaryCare.filter(r => countOperationalSources(r) === 3).length;
    const primaryCarePartialMerge = processedData.primaryCare.filter(r => countOperationalSources(r) === 2).length;
    const primaryCareSingleSource = processedData.primaryCare.filter(r => countOperationalSources(r) === 1).length;

    // Calculate merge counts for Specialist
    const specialistFullMerge = processedData.specialist.filter(r => countOperationalSources(r) === 3).length;
    const specialistPartialMerge = processedData.specialist.filter(r => countOperationalSources(r) === 2).length;
    const specialistSingleSource = processedData.specialist.filter(r => countOperationalSources(r) === 1).length;

    return {
      totalRecords: allRecords.length,
      totalUniqueClinicians: nonExcluded.filter(r => r.Provincial_Dedupe_Flag).length,
      primaryCareRecords: processedData.primaryCare.length,
      primaryCareUnique: processedData.primaryCare.filter(r => r.Provincial_Dedupe_Flag && !r.Exclusion_Flag).length,
      specialistRecords: processedData.specialist.length,
      specialistUnique: processedData.specialist.filter(r => r.Provincial_Dedupe_Flag && !r.Exclusion_Flag).length,
      excludedCount: allRecords.filter(r => r.Exclusion_Flag).length,
      notInGeoSpatial: allRecords.filter(r => !r.In_GeoSpatial_LDG).length,
      regionCounts,
      ldgCounts,
      qualityIssueCount: dataQualityIssues.length,
      primaryCareFullMerge,
      primaryCarePartialMerge,
      primaryCareSingleSource,
      specialistFullMerge,
      specialistPartialMerge,
      specialistSingleSource
    };
  }, [isProcessed, processedData, dataQualityIssues]);

  // Get unique specialties from loaded data
  const uniqueSpecialties = useMemo(() => {
    return Object.keys(specialtyMapping).sort();
  }, [specialtyMapping]);

  // File upload component
  const FileUploadBox = ({ fileType, label, description }) => {
    const file = files[fileType];
    const fileWarnings = warnings[fileType];
    const recordCount = data[fileType].length;
    const hasWarnings = fileWarnings && fileWarnings.length > 0;

    // Determine colors based on file state and warnings
    const cardStyle = !file
      ? 'border-gray-300 hover:border-blue-400'
      : hasWarnings
        ? 'border-yellow-400 bg-yellow-50'
        : 'border-green-400 bg-green-50';

    const iconColor = !file
      ? 'text-gray-400'
      : hasWarnings
        ? 'text-yellow-600'
        : 'text-green-600';

    const statusColor = hasWarnings ? 'text-yellow-700' : 'text-green-700';

    return (
      <div className={`border-2 border-dashed rounded-lg p-4 transition-all ${cardStyle}`}>
        <div className="flex items-start gap-3">
          <FileSpreadsheet className={`w-8 h-8 ${iconColor}`} />
          <div className="flex-1">
            <h3 className="font-semibold text-gray-800">{label}</h3>
            <p className="text-sm text-gray-500 mb-2">{description}</p>

            {file ? (
              <div className="space-y-1">
                <div className={`flex items-center gap-2 ${statusColor}`}>
                  {hasWarnings ? (
                    <AlertTriangle className="w-4 h-4" />
                  ) : (
                    <CheckCircle className="w-4 h-4" />
                  )}
                  <span className="text-sm font-medium">{file.name}</span>
                </div>
                <p className="text-sm text-gray-600">{recordCount.toLocaleString()} records loaded</p>
                {hasWarnings && (
                  <div className="mt-2 p-2 bg-yellow-100 rounded border border-yellow-300">
                    <div className="text-sm text-yellow-800">
                      {fileWarnings.map((w, i) => <p key={i}>{w}</p>)}
                    </div>
                  </div>
                )}
              </div>
            ) : (
              <label className="cursor-pointer">
                <div className="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition">
                  <Upload className="w-4 h-4" />
                  <span>Upload File</span>
                </div>
                <input
                  type="file"
                  accept=".csv,.xlsx,.xls,.xlsm"
                  className="hidden"
                  onChange={(e) => handleFileUpload(fileType, e.target.files[0])}
                />
              </label>
            )}
          </div>
          {file && (
            <button
              onClick={() => {
                setFiles(prev => ({ ...prev, [fileType]: null }));
                setData(prev => ({ ...prev, [fileType]: [] }));
                setWarnings(prev => ({ ...prev, [fileType]: [] }));
                setIsProcessed(false);
              }}
              className="text-gray-400 hover:text-red-500"
            >
              <XCircle className="w-5 h-5" />
            </button>
          )}
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <div className="bg-gradient-to-r from-blue-800 to-blue-600 text-white py-6 px-6 shadow-lg">
        <h1 className="text-2xl font-bold">Clinician Data Consolidation Tool</h1>
        <p className="text-blue-100 mt-1">Transform source files into Primary Care and Specialist outputs</p>
      </div>

      {/* Navigation Tabs */}
      <div className="bg-white border-b shadow-sm">
        <div className="max-w-6xl mx-auto px-6">
          <div className="flex gap-1">
            {['upload', 'config', 'summary', 'quality', 'export'].map(tab => (
              <button
                key={tab}
                onClick={() => setActiveTab(tab)}
                disabled={tab !== 'upload' && tab !== 'config' && !isProcessed}
                className={`px-4 py-3 font-medium capitalize transition ${
                  activeTab === tab
                    ? 'text-blue-600 border-b-2 border-blue-600'
                    : 'text-gray-500 hover:text-gray-700 disabled:text-gray-300 disabled:cursor-not-allowed'
                }`}
              >
                {tab === 'quality' ? 'Data Quality' : tab}
              </button>
            ))}
          </div>
        </div>
      </div>

      <div className="max-w-6xl mx-auto px-6 py-6">
        {/* Upload Tab */}
        {activeTab === 'upload' && (
          <div className="space-y-6">
            <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
              <div className="flex items-start gap-3">
                <Info className="w-5 h-5 text-blue-600 mt-0.5" />
                <div className="text-sm text-blue-800">
                  <p className="font-medium">Getting Started</p>
                  <p>Upload your source files below. Files can be in CSV or XLSX format. The Geo-Spatial LDG is required as the foundation dataset.</p>
                </div>
              </div>
            </div>

            <div className="grid md:grid-cols-2 gap-4">
              <FileUploadBox
                fileType="geoSpatial"
                label="Geo-Spatial LDG"
                description="Foundation dataset with CPSO, specialty, and LDG mappings"
              />
              <FileUploadBox
                fileType="regionalAuthority"
                label="Regional Authority Listings"
                description="Specialist directory listings from Ocean platform"
              />
              <FileUploadBox
                fileType="supportSite"
                label="Support Site Analytics"
                description="Primary care clinician activity data"
              />
              <FileUploadBox
                fileType="ldgLedger"
                label="LDG Onboarding Ledger"
                description="Onboarding completion and specialty pathways"
              />
            </div>

            {data.geoSpatial.length > 0 && (
              <div className="flex justify-center pt-4">
                <button
                  onClick={processData}
                  className="px-6 py-3 bg-green-600 text-white font-semibold rounded-lg hover:bg-green-700 transition flex items-center gap-2"
                >
                  <Settings className="w-5 h-5" />
                  Process Data
                </button>
              </div>
            )}
          </div>
        )}

        {/* Configuration Tab */}
        {activeTab === 'config' && (
          <div className="space-y-6">
            {/* Specialty Mapping */}
            <div className="bg-white rounded-lg shadow p-6">
              <h2 className="text-lg font-semibold text-gray-800 mb-4 flex items-center gap-2">
                <Filter className="w-5 h-5" />
                Specialty Type Mapping
              </h2>
              <p className="text-sm text-gray-600 mb-4">
                Configure which "Type of Specialty" values from Geo-Spatial LDG should go to Primary Care or Specialist output files.
              </p>

              {uniqueSpecialties.length > 0 ? (
                <div className="space-y-2">
                  {uniqueSpecialties.map(specialty => (
                    <div key={specialty} className="flex items-center gap-4 p-3 bg-gray-50 rounded">
                      <span className="flex-1 font-medium text-gray-700">{specialty}</span>
                      <select
                        value={specialtyMapping[specialty] || 'specialist'}
                        onChange={(e) => setSpecialtyMapping(prev => ({ ...prev, [specialty]: e.target.value }))}
                        className="px-3 py-1.5 border rounded-lg bg-white"
                      >
                        <option value="primaryCare">Primary Care</option>
                        <option value="specialist">Specialist</option>
                      </select>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-gray-500 italic">Load Geo-Spatial LDG file to configure specialty mapping</p>
              )}
            </div>

            {/* Deployment Team Exclusions */}
            <div className="bg-white rounded-lg shadow p-6">
              <h2 className="text-lg font-semibold text-gray-800 mb-4 flex items-center gap-2">
                <Users className="w-5 h-5" />
                Deployment Team Exclusion List
              </h2>
              <p className="text-sm text-gray-600 mb-4">
                These individuals will be excluded from the output files.
              </p>

              <div className="flex gap-2 mb-4">
                <input
                  type="text"
                  value={newTeamMember}
                  onChange={(e) => setNewTeamMember(e.target.value)}
                  placeholder="Enter name to exclude"
                  className="flex-1 px-3 py-2 border rounded-lg"
                  onKeyPress={(e) => {
                    if (e.key === 'Enter' && newTeamMember.trim()) {
                      setDeploymentTeam(prev => [...prev, newTeamMember.trim()]);
                      setNewTeamMember('');
                    }
                  }}
                />
                <button
                  onClick={() => {
                    if (newTeamMember.trim()) {
                      setDeploymentTeam(prev => [...prev, newTeamMember.trim()]);
                      setNewTeamMember('');
                    }
                  }}
                  className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700"
                >
                  Add
                </button>
              </div>

              <div className="flex flex-wrap gap-2">
                {deploymentTeam.map((member, idx) => (
                  <span key={idx} className="inline-flex items-center gap-1 px-3 py-1 bg-red-100 text-red-800 rounded-full">
                    {member}
                    <button
                      onClick={() => setDeploymentTeam(prev => prev.filter((_, i) => i !== idx))}
                      className="hover:text-red-600"
                    >
                      <XCircle className="w-4 h-4" />
                    </button>
                  </span>
                ))}
              </div>
            </div>

            {data.geoSpatial.length > 0 && (
              <div className="flex justify-center">
                <button
                  onClick={processData}
                  className="px-6 py-3 bg-green-600 text-white font-semibold rounded-lg hover:bg-green-700 transition flex items-center gap-2"
                >
                  <Settings className="w-5 h-5" />
                  {isProcessed ? 'Re-Process Data' : 'Process Data'}
                </button>
              </div>
            )}
          </div>
        )}

        {/* Summary Tab */}
        {activeTab === 'summary' && isProcessed && summary && (
          <div className="space-y-6">
            {/* Overview Cards */}
            <div className="grid md:grid-cols-4 gap-4">
              <div className="bg-white rounded-lg shadow p-4">
                <p className="text-sm text-gray-500">Total Records</p>
                <p className="text-2xl font-bold text-gray-800">{summary.totalRecords.toLocaleString()}</p>
              </div>
              <div className="bg-white rounded-lg shadow p-4">
                <p className="text-sm text-gray-500">Unique Clinicians</p>
                <p className="text-2xl font-bold text-blue-600">{summary.totalUniqueClinicians.toLocaleString()}</p>
              </div>
              <div className="bg-white rounded-lg shadow p-4">
                <p className="text-sm text-gray-500">Excluded Records</p>
                <p className="text-2xl font-bold text-red-600">{summary.excludedCount.toLocaleString()}</p>
              </div>
              <div className="bg-white rounded-lg shadow p-4">
                <p className="text-sm text-gray-500">Quality Issues</p>
                <p className="text-2xl font-bold text-yellow-600">{summary.qualityIssueCount.toLocaleString()}</p>
              </div>
            </div>

            {/* Output File Summary */}
            <div className="grid md:grid-cols-2 gap-4">
              <div className="bg-green-50 border border-green-200 rounded-lg p-4">
                <h3 className="font-semibold text-green-800 mb-2">Primary Care Output</h3>
                <p className="text-sm text-green-700">
                  {summary.primaryCareRecords.toLocaleString()} records ({summary.primaryCareUnique.toLocaleString()} unique clinicians)
                </p>
                <div className="mt-2 pt-2 border-t border-green-200">
                  <p className="text-xs text-green-600 font-medium mb-1">Data Source Merges:</p>
                  <p className="text-xs text-green-700">
                    Full (3 sources): {summary.primaryCareFullMerge.toLocaleString()} | Partial (2 sources): {summary.primaryCarePartialMerge.toLocaleString()} | Single: {summary.primaryCareSingleSource.toLocaleString()}
                  </p>
                </div>
              </div>
              <div className="bg-purple-50 border border-purple-200 rounded-lg p-4">
                <h3 className="font-semibold text-purple-800 mb-2">Specialist Output</h3>
                <p className="text-sm text-purple-700">
                  {summary.specialistRecords.toLocaleString()} records ({summary.specialistUnique.toLocaleString()} unique clinicians)
                </p>
                <div className="mt-2 pt-2 border-t border-purple-200">
                  <p className="text-xs text-purple-600 font-medium mb-1">Data Source Merges:</p>
                  <p className="text-xs text-purple-700">
                    Full (3 sources): {summary.specialistFullMerge.toLocaleString()} | Partial (2 sources): {summary.specialistPartialMerge.toLocaleString()} | Single: {summary.specialistSingleSource.toLocaleString()}
                  </p>
                </div>
              </div>
            </div>

            {/* Regional Breakdown */}
            <div className="bg-white rounded-lg shadow p-6">
              <h3 className="font-semibold text-gray-800 mb-4">Unique Clinicians by Region</h3>
              <div className="grid md:grid-cols-3 gap-2">
                {Object.entries(summary.regionCounts).sort((a, b) => b[1] - a[1]).map(([region, count]) => (
                  <div key={region} className="flex justify-between p-2 bg-gray-50 rounded">
                    <span className="text-gray-700">{region || 'Unknown'}</span>
                    <span className="font-semibold">{count.toLocaleString()}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Not in Geo-Spatial Warning */}
            {summary.notInGeoSpatial > 0 && (
              <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-4">
                <div className="flex items-start gap-3">
                  <AlertTriangle className="w-5 h-5 text-yellow-600 mt-0.5" />
                  <div>
                    <p className="font-medium text-yellow-800">
                      {summary.notInGeoSpatial.toLocaleString()} records not found in Geo-Spatial LDG
                    </p>
                    <p className="text-sm text-yellow-700 mt-1">
                      These records have In_GeoSpatial_LDG = FALSE. Review in Data Quality tab.
                    </p>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Data Quality Tab */}
        {activeTab === 'quality' && isProcessed && (
          <div className="space-y-6">
            {/* Quality Issues */}
            <div className="bg-white rounded-lg shadow">
              <div className="p-4 border-b">
                <h3 className="font-semibold text-gray-800">
                  Data Quality Issues ({dataQualityIssues.length})
                </h3>
              </div>
              {dataQualityIssues.length > 0 ? (
                <div className="max-h-96 overflow-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-gray-50 sticky top-0">
                      <tr>
                        <th className="px-4 py-2 text-left">Type</th>
                        <th className="px-4 py-2 text-left">Source</th>
                        <th className="px-4 py-2 text-left">CPSO</th>
                        <th className="px-4 py-2 text-left">Name</th>
                        <th className="px-4 py-2 text-left">Details</th>
                      </tr>
                    </thead>
                    <tbody>
                      {dataQualityIssues.map((issue, idx) => (
                        <tr key={idx} className="border-t hover:bg-gray-50">
                          <td className="px-4 py-2">
                            <span className="px-2 py-1 bg-yellow-100 text-yellow-800 rounded text-xs">
                              {issue.type.replace(/_/g, ' ')}
                            </span>
                          </td>
                          <td className="px-4 py-2">{issue.source}</td>
                          <td className="px-4 py-2 font-mono">{issue.cpso}</td>
                          <td className="px-4 py-2">{issue.name}</td>
                          <td className="px-4 py-2 text-gray-600">{issue.details}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div className="p-8 text-center text-gray-500">
                  <CheckCircle className="w-12 h-12 mx-auto mb-2 text-green-500" />
                  <p>No data quality issues found!</p>
                </div>
              )}
            </div>

            {/* Exclusions */}
            <div className="bg-white rounded-lg shadow">
              <div className="p-4 border-b">
                <h3 className="font-semibold text-gray-800">
                  Excluded Records ({exclusions.length})
                </h3>
              </div>
              {exclusions.length > 0 ? (
                <div className="max-h-96 overflow-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-gray-50 sticky top-0">
                      <tr>
                        <th className="px-4 py-2 text-left">CPSO</th>
                        <th className="px-4 py-2 text-left">Name</th>
                        <th className="px-4 py-2 text-left">Source</th>
                        <th className="px-4 py-2 text-left">Reason</th>
                      </tr>
                    </thead>
                    <tbody>
                      {exclusions.map((record, idx) => (
                        <tr key={idx} className="border-t hover:bg-gray-50">
                          <td className="px-4 py-2 font-mono">{record.Professional_ID}</td>
                          <td className="px-4 py-2">{record.First_Name} {record.Last_Name}</td>
                          <td className="px-4 py-2">{record.source}</td>
                          <td className="px-4 py-2">
                            <span className="px-2 py-1 bg-red-100 text-red-800 rounded text-xs">
                              {record.Exclusion_Reason}
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div className="p-8 text-center text-gray-500">
                  <p>No records excluded</p>
                </div>
              )}
            </div>

            {/* Export Quality Report */}
            <div className="flex justify-center gap-4">
              <button
                onClick={() => exportData(dataQualityIssues, 'Data_Quality_Report', 'xlsx')}
                className="px-4 py-2 bg-yellow-600 text-white rounded-lg hover:bg-yellow-700 flex items-center gap-2"
              >
                <Download className="w-4 h-4" />
                Export Quality Report (XLSX)
              </button>
              <button
                onClick={() => exportData(exclusions, 'Exclusions_Report', 'xlsx')}
                className="px-4 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 flex items-center gap-2"
              >
                <Download className="w-4 h-4" />
                Export Exclusions (XLSX)
              </button>
            </div>
          </div>
        )}

        {/* Export Tab */}
        {activeTab === 'export' && isProcessed && (
          <div className="space-y-6">
            <div className="bg-white rounded-lg shadow p-6">
              <h2 className="text-lg font-semibold text-gray-800 mb-6">Download Output Files</h2>

              <div className="grid md:grid-cols-2 gap-6">
                {/* Primary Care Export */}
                <div className="border rounded-lg p-6 bg-green-50">
                  <h3 className="font-semibold text-green-800 mb-2">Primary Care Output</h3>
                  <p className="text-sm text-green-700 mb-4">
                    {processedData.primaryCare.length.toLocaleString()} records
                  </p>
                  <div className="flex gap-2">
                    <button
                      onClick={() => exportData(processedData.primaryCare, 'Primary_Care_Output', 'xlsx')}
                      className="flex-1 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 flex items-center justify-center gap-2"
                    >
                      <Download className="w-4 h-4" />
                      XLSX
                    </button>
                    <button
                      onClick={() => exportData(processedData.primaryCare, 'Primary_Care_Output', 'csv')}
                      className="flex-1 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 flex items-center justify-center gap-2"
                    >
                      <Download className="w-4 h-4" />
                      CSV
                    </button>
                  </div>
                </div>

                {/* Specialist Export */}
                <div className="border rounded-lg p-6 bg-purple-50">
                  <h3 className="font-semibold text-purple-800 mb-2">Specialist Output</h3>
                  <p className="text-sm text-purple-700 mb-4">
                    {processedData.specialist.length.toLocaleString()} records
                  </p>
                  <div className="flex gap-2">
                    <button
                      onClick={() => exportData(processedData.specialist, 'Specialist_Output', 'xlsx')}
                      className="flex-1 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 flex items-center justify-center gap-2"
                    >
                      <Download className="w-4 h-4" />
                      XLSX
                    </button>
                    <button
                      onClick={() => exportData(processedData.specialist, 'Specialist_Output', 'csv')}
                      className="flex-1 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 flex items-center justify-center gap-2"
                    >
                      <Download className="w-4 h-4" />
                      CSV
                    </button>
                  </div>
                </div>
              </div>
            </div>

            {/* Additional Exports */}
            <div className="bg-white rounded-lg shadow p-6">
              <h2 className="text-lg font-semibold text-gray-800 mb-4">Additional Reports</h2>
              <div className="flex flex-wrap gap-4">
                <button
                  onClick={() => exportData(dataQualityIssues, 'Data_Quality_Report', 'xlsx')}
                  className="px-4 py-2 border border-gray-300 rounded-lg hover:bg-gray-50 flex items-center gap-2"
                >
                  <Download className="w-4 h-4" />
                  Data Quality Report
                </button>
                <button
                  onClick={() => exportData(exclusions, 'Excluded_Records', 'xlsx')}
                  className="px-4 py-2 border border-gray-300 rounded-lg hover:bg-gray-50 flex items-center gap-2"
                >
                  <Download className="w-4 h-4" />
                  Excluded Records
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
