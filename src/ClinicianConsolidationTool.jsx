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

  // Pre-processing quality checks state
  const [qualityChecks, setQualityChecks] = useState({
    geoSpatial: null,
    regionalAuthority: null,
    supportSite: null,
    ldgLedger: null
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

  // Quality check functions for each file type
  const runQualityChecks = useCallback((fileType, rows) => {
    const checks = { critical: [], warnings: [], summary: { totalRows: rows.length } };

    if (fileType === 'geoSpatial') {
      // Critical: Blank CPSO
      const blankCPSO = rows.filter(r => !String(r['CPSO'] || '').trim());
      if (blankCPSO.length > 0) {
        checks.critical.push({
          type: 'BLANK_CPSO',
          label: 'Blank CPSO values',
          count: blankCPSO.length,
          impact: 'These rows cannot be used for lookups and will be ignored during processing.',
          sampleRows: blankCPSO.slice(0, 3).map((r, i) => `Row ${rows.indexOf(r) + 2}`)
        });
      }

      // Critical: Non-numeric CPSO values
      const nonNumericCPSO = rows.filter(r => {
        const val = String(r['CPSO'] || '').trim();
        return val && !/^\d+$/.test(val);
      });
      if (nonNumericCPSO.length > 0) {
        checks.critical.push({
          type: 'NON_NUMERIC_CPSO',
          label: 'Non-numeric CPSO values',
          count: nonNumericCPSO.length,
          impact: 'CPSO values should contain only numbers. These records may fail to match with other files.',
          sampleRows: nonNumericCPSO.slice(0, 3).map(r => `"${r['CPSO']}" (Row ${rows.indexOf(r) + 2})`)
        });
      }

      // Warning: Duplicate CPSOs
      const cpsoCount = {};
      rows.forEach((row, idx) => {
        const cpso = String(row['CPSO'] || '').replace(/\s/g, '').trim();
        if (cpso) {
          if (!cpsoCount[cpso]) cpsoCount[cpso] = [];
          cpsoCount[cpso].push(idx + 2);
        }
      });
      const duplicates = Object.entries(cpsoCount).filter(([_, indices]) => indices.length > 1);
      if (duplicates.length > 0) {
        const totalDupeRows = duplicates.reduce((sum, [_, indices]) => sum + indices.length, 0);
        checks.warnings.push({
          type: 'DUPLICATE_CPSO',
          label: 'Duplicate CPSO values',
          count: totalDupeRows,
          uniqueCount: duplicates.length,
          impact: 'Last occurrence of each duplicate CPSO will be used for lookups. Earlier occurrences will be ignored.',
          sampleRows: duplicates.slice(0, 3).map(([cpso, indices]) => `CPSO ${cpso}: rows ${indices.join(', ')}`)
        });
      }

      // Warning: Blank Specialty
      const blankSpecialty = rows.filter(r => !String(r['Specialty'] || '').trim());
      if (blankSpecialty.length > 0) {
        checks.warnings.push({
          type: 'BLANK_SPECIALTY',
          label: 'Blank Specialty values',
          count: blankSpecialty.length,
          impact: 'Records will default to Specialist output file. Specialty mapping may be affected.',
          sampleRows: blankSpecialty.slice(0, 3).map(r => `CPSO ${r['CPSO'] || 'N/A'}`)
        });
      }

      // Warning: Blank Region
      const blankRegion = rows.filter(r => !String(r['Region'] || '').trim());
      if (blankRegion.length > 0) {
        checks.warnings.push({
          type: 'BLANK_REGION',
          label: 'Blank Region values',
          count: blankRegion.length,
          impact: 'Region field in output will be empty for these clinicians unless provided by other source files.',
          sampleRows: blankRegion.slice(0, 3).map(r => `CPSO ${r['CPSO'] || 'N/A'}`)
        });
      }

      // Warning: Blank LDG
      const blankLDG = rows.filter(r => !String(r['LDG'] || '').trim());
      if (blankLDG.length > 0) {
        checks.warnings.push({
          type: 'BLANK_LDG',
          label: 'Blank LDG values',
          count: blankLDG.length,
          impact: 'LDG Name field in output will be empty for these clinicians.',
          sampleRows: blankLDG.slice(0, 3).map(r => `CPSO ${r['CPSO'] || 'N/A'}`)
        });
      }

    } else if (fileType === 'regionalAuthority') {
      // Critical: Blank clinicianProfessionalId
      const blankCPSO = rows.filter(r => !String(r['clinicianProfessionalId'] || '').trim());
      if (blankCPSO.length > 0) {
        checks.critical.push({
          type: 'BLANK_PROFESSIONAL_ID',
          label: 'Blank Professional ID values',
          count: blankCPSO.length,
          impact: 'These rows will be skipped entirely - no records created.',
          sampleRows: blankCPSO.slice(0, 3).map((r, i) => `${r['clinicianFirstName'] || ''} ${r['clinicianSurname'] || ''} (Row ${rows.indexOf(r) + 2})`.trim())
        });
      }

      // Critical: Non-numeric Professional ID values
      const nonNumericProfId = rows.filter(r => {
        const val = String(r['clinicianProfessionalId'] || '').trim();
        return val && !/^\d+$/.test(val);
      });
      if (nonNumericProfId.length > 0) {
        checks.critical.push({
          type: 'NON_NUMERIC_PROFESSIONAL_ID',
          label: 'Non-numeric Professional ID values',
          count: nonNumericProfId.length,
          impact: 'Professional ID values should contain only numbers. These records may fail to match with other files.',
          sampleRows: nonNumericProfId.slice(0, 3).map(r => `"${r['clinicianProfessionalId']}" (Row ${rows.indexOf(r) + 2})`)
        });
      }

      // Warning: Blank siteNum
      const blankSite = rows.filter(r => !String(r['siteNum'] || '').trim());
      if (blankSite.length > 0) {
        checks.warnings.push({
          type: 'BLANK_SITE_NUM',
          label: 'Blank Site Number values',
          count: blankSite.length,
          impact: 'Record matching with other files may fail. Records may not merge properly with Support Site or LDG Ledger data.',
          sampleRows: blankSite.slice(0, 3).map(r => `CPSO ${r['clinicianProfessionalId'] || 'N/A'}`)
        });
      }

      // Warning: Blank healthRegion
      const blankRegion = rows.filter(r => !String(r['healthRegion'] || '').trim());
      if (blankRegion.length > 0) {
        checks.warnings.push({
          type: 'BLANK_HEALTH_REGION',
          label: 'Blank Health Region values',
          count: blankRegion.length,
          impact: 'Region field in output will rely on Geo-Spatial LDG lookup only.',
          sampleRows: blankRegion.slice(0, 3).map(r => `CPSO ${r['clinicianProfessionalId'] || 'N/A'}`)
        });
      }

      // Critical: Same Professional ID with different names
      const idToNames = {};
      rows.forEach((row, idx) => {
        const profId = String(row['clinicianProfessionalId'] || '').trim();
        if (profId) {
          const firstName = String(row['clinicianFirstName'] || '').trim().toLowerCase();
          const lastName = String(row['clinicianSurname'] || '').trim().toLowerCase();
          const fullName = `${firstName} ${lastName}`.trim();
          if (!idToNames[profId]) {
            idToNames[profId] = { names: new Set(), rows: [] };
          }
          if (fullName) {
            idToNames[profId].names.add(fullName);
          }
          idToNames[profId].rows.push(idx + 2);
        }
      });
      const nameMismatchIds = Object.entries(idToNames).filter(([_, data]) => data.names.size > 1);
      if (nameMismatchIds.length > 0) {
        const totalAffectedRows = nameMismatchIds.reduce((sum, [_, data]) => sum + data.rows.length, 0);
        checks.critical.push({
          type: 'PROFESSIONAL_ID_NAME_MISMATCH',
          label: 'Professional ID with multiple different names',
          count: totalAffectedRows,
          uniqueCount: nameMismatchIds.length,
          impact: 'Same Professional ID appears with different names, indicating potential data integrity issues.',
          sampleRows: nameMismatchIds.slice(0, 3).map(([id, data]) => `ID ${id}: ${Array.from(data.names).join(' vs ')}`)
        });
      }

    } else if (fileType === 'supportSite') {
      // Critical: Blank professionalId
      const blankCPSO = rows.filter(r => !String(r['professionalId'] || '').trim());
      if (blankCPSO.length > 0) {
        checks.critical.push({
          type: 'BLANK_PROFESSIONAL_ID',
          label: 'Blank Professional ID values',
          count: blankCPSO.length,
          impact: 'These rows will be skipped entirely - no records created.',
          sampleRows: blankCPSO.slice(0, 3).map(r => `${r['userFullName'] || 'Unknown'} (Row ${rows.indexOf(r) + 2})`)
        });
      }

      // Critical: Non-numeric Professional ID values
      const nonNumericProfId = rows.filter(r => {
        const val = String(r['professionalId'] || '').trim();
        return val && !/^\d+$/.test(val);
      });
      if (nonNumericProfId.length > 0) {
        checks.critical.push({
          type: 'NON_NUMERIC_PROFESSIONAL_ID',
          label: 'Non-numeric Professional ID values',
          count: nonNumericProfId.length,
          impact: 'Professional ID values should contain only numbers. These records may fail to match with other files.',
          sampleRows: nonNumericProfId.slice(0, 3).map(r => `"${r['professionalId']}" (Row ${rows.indexOf(r) + 2})`)
        });
      }

      // Warning: Blank siteNum
      const blankSite = rows.filter(r => !String(r['siteNum'] || '').trim());
      if (blankSite.length > 0) {
        checks.warnings.push({
          type: 'BLANK_SITE_NUM',
          label: 'Blank Site Number values',
          count: blankSite.length,
          impact: 'Record matching may fail. Records may not merge properly with Regional Authority or LDG Ledger data.',
          sampleRows: blankSite.slice(0, 3).map(r => `CPSO ${r['professionalId'] || 'N/A'}`)
        });
      }

      // Warning: Blank referralLastSent
      const blankLastRef = rows.filter(r => !String(r['referralLastSent'] || '').trim());
      if (blankLastRef.length > 0) {
        checks.warnings.push({
          type: 'BLANK_LAST_REFERRAL',
          label: 'Blank Last Referral Sent values',
          count: blankLastRef.length,
          impact: 'Last_Referral_Sent field in output will be empty for these clinicians.',
          sampleRows: blankLastRef.slice(0, 3).map(r => `CPSO ${r['professionalId'] || 'N/A'}`)
        });
      }

      // Critical: Same Professional ID with different names
      const idToNamesSS = {};
      rows.forEach((row, idx) => {
        const profId = String(row['professionalId'] || '').trim();
        if (profId) {
          const fullName = String(row['userFullName'] || '').trim().toLowerCase();
          if (!idToNamesSS[profId]) {
            idToNamesSS[profId] = { names: new Set(), rows: [] };
          }
          if (fullName) {
            idToNamesSS[profId].names.add(fullName);
          }
          idToNamesSS[profId].rows.push(idx + 2);
        }
      });
      const nameMismatchIdsSS = Object.entries(idToNamesSS).filter(([_, data]) => data.names.size > 1);
      if (nameMismatchIdsSS.length > 0) {
        const totalAffectedRows = nameMismatchIdsSS.reduce((sum, [_, data]) => sum + data.rows.length, 0);
        checks.critical.push({
          type: 'PROFESSIONAL_ID_NAME_MISMATCH',
          label: 'Professional ID with multiple different names',
          count: totalAffectedRows,
          uniqueCount: nameMismatchIdsSS.length,
          impact: 'Same Professional ID appears with different names, indicating potential data integrity issues.',
          sampleRows: nameMismatchIdsSS.slice(0, 3).map(([id, data]) => `ID ${id}: ${Array.from(data.names).join(' vs ')}`)
        });
      }

    } else if (fileType === 'ldgLedger') {
      // Critical: Blank CPSO #
      const blankCPSO = rows.filter(r => !String(r['CPSO #'] || '').trim());
      if (blankCPSO.length > 0) {
        checks.critical.push({
          type: 'BLANK_CPSO',
          label: 'Blank CPSO # values',
          count: blankCPSO.length,
          impact: 'These rows will be skipped entirely - no records created.',
          sampleRows: blankCPSO.slice(0, 3).map(r => `${r['First Name'] || ''} ${r['Last Name'] || ''} (Row ${rows.indexOf(r) + 2})`.trim())
        });
      }

      // Critical: Non-numeric CPSO # values
      const nonNumericCPSO = rows.filter(r => {
        const val = String(r['CPSO #'] || '').trim();
        return val && !/^\d+$/.test(val);
      });
      if (nonNumericCPSO.length > 0) {
        checks.critical.push({
          type: 'NON_NUMERIC_CPSO',
          label: 'Non-numeric CPSO # values',
          count: nonNumericCPSO.length,
          impact: 'CPSO # values should contain only numbers. These records may fail to match with other files.',
          sampleRows: nonNumericCPSO.slice(0, 3).map(r => `"${r['CPSO #']}" (Row ${rows.indexOf(r) + 2})`)
        });
      }

      // Warning: Blank Ocean Site Number
      const blankSite = rows.filter(r => !String(r['Ocean Site Number (eReferral Ontario only)'] || '').trim());
      if (blankSite.length > 0) {
        checks.warnings.push({
          type: 'BLANK_SITE_NUM',
          label: 'Blank Ocean Site Number values',
          count: blankSite.length,
          impact: 'Record matching may fail. Records may not merge properly with Regional Authority or Support Site data.',
          sampleRows: blankSite.slice(0, 3).map(r => `CPSO ${r['CPSO #'] || 'N/A'}`)
        });
      }

      // Warning: Blank Date Onboarding Completed
      const blankDate = rows.filter(r => !String(r['Date Onboarding Completed'] || '').trim());
      if (blankDate.length > 0) {
        checks.warnings.push({
          type: 'BLANK_ONBOARDING_DATE',
          label: 'Blank Onboarding Date values',
          count: blankDate.length,
          impact: 'Date_Onboarded field in output will be empty for these clinicians.',
          sampleRows: blankDate.slice(0, 3).map(r => `CPSO ${r['CPSO #'] || 'N/A'}`)
        });
      }

      // Warning: Blank Primary Specialty Pathway
      const blankPathway = rows.filter(r => !String(r['Primary Specialty Pathway'] || '').trim());
      if (blankPathway.length > 0) {
        checks.warnings.push({
          type: 'BLANK_SPECIALTY_PATHWAY',
          label: 'Blank Primary Specialty Pathway values',
          count: blankPathway.length,
          impact: 'Specialty_Pathway field in output will show default value for these clinicians.',
          sampleRows: blankPathway.slice(0, 3).map(r => `CPSO ${r['CPSO #'] || 'N/A'}`)
        });
      }

      // Critical: Same CPSO # with different names
      const idToNamesLDG = {};
      rows.forEach((row, idx) => {
        const cpso = String(row['CPSO #'] || '').trim();
        if (cpso) {
          const firstName = String(row['First Name'] || '').trim().toLowerCase();
          const lastName = String(row['Last Name'] || '').trim().toLowerCase();
          const fullName = `${firstName} ${lastName}`.trim();
          if (!idToNamesLDG[cpso]) {
            idToNamesLDG[cpso] = { names: new Set(), rows: [] };
          }
          if (fullName) {
            idToNamesLDG[cpso].names.add(fullName);
          }
          idToNamesLDG[cpso].rows.push(idx + 2);
        }
      });
      const nameMismatchIdsLDG = Object.entries(idToNamesLDG).filter(([_, data]) => data.names.size > 1);
      if (nameMismatchIdsLDG.length > 0) {
        const totalAffectedRows = nameMismatchIdsLDG.reduce((sum, [_, data]) => sum + data.rows.length, 0);
        checks.critical.push({
          type: 'CPSO_NAME_MISMATCH',
          label: 'CPSO # with multiple different names',
          count: totalAffectedRows,
          uniqueCount: nameMismatchIdsLDG.length,
          impact: 'Same CPSO # appears with different names, indicating potential data integrity issues.',
          sampleRows: nameMismatchIdsLDG.slice(0, 3).map(([id, data]) => `CPSO ${id}: ${Array.from(data.names).join(' vs ')}`)
        });
      }
    }

    // Calculate overall status
    checks.status = checks.critical.length > 0 ? 'critical' : checks.warnings.length > 0 ? 'warning' : 'good';
    return checks;
  }, []);

  // Get quality issues for a single row (used for flagged export)
  const getRowQualityIssues = useCallback((fileType, row, rowIndex, allRows) => {
    const issues = [];

    if (fileType === 'geoSpatial') {
      const cpso = String(row['CPSO'] || '').trim();

      // Check blank CPSO
      if (!cpso) {
        issues.push('Blank CPSO');
      }

      // Check non-numeric CPSO
      if (cpso && !/^\d+$/.test(cpso)) {
        issues.push('Non-numeric CPSO');
      }

      // Check duplicate CPSO (need to build count map)
      if (cpso) {
        const normalizedCpso = cpso.replace(/\s/g, '');
        const duplicateCount = allRows.filter(r =>
          String(r['CPSO'] || '').replace(/\s/g, '').trim() === normalizedCpso
        ).length;
        if (duplicateCount > 1) {
          issues.push('Duplicate CPSO');
        }
      }

      // Check blank Specialty
      if (!String(row['Specialty'] || '').trim()) {
        issues.push('Blank Specialty');
      }

      // Check blank Region
      if (!String(row['Region'] || '').trim()) {
        issues.push('Blank Region');
      }

      // Check blank LDG
      if (!String(row['LDG'] || '').trim()) {
        issues.push('Blank LDG');
      }

    } else if (fileType === 'regionalAuthority') {
      const profId = String(row['clinicianProfessionalId'] || '').trim();

      // Check blank Professional ID
      if (!profId) {
        issues.push('Blank Professional ID');
      }

      // Check non-numeric Professional ID
      if (profId && !/^\d+$/.test(profId)) {
        issues.push('Non-numeric Professional ID');
      }

      // Check blank siteNum
      if (!String(row['siteNum'] || '').trim()) {
        issues.push('Blank Site Number');
      }

      // Check blank healthRegion
      if (!String(row['healthRegion'] || '').trim()) {
        issues.push('Blank Health Region');
      }

      // Check if this Professional ID has multiple different names
      if (profId) {
        const idToNamesRA = {};
        allRows.forEach(r => {
          const id = String(r['clinicianProfessionalId'] || '').trim();
          if (id) {
            const firstName = String(r['clinicianFirstName'] || '').trim().toLowerCase();
            const lastName = String(r['clinicianSurname'] || '').trim().toLowerCase();
            const fullName = `${firstName} ${lastName}`.trim();
            if (!idToNamesRA[id]) {
              idToNamesRA[id] = new Set();
            }
            if (fullName) {
              idToNamesRA[id].add(fullName);
            }
          }
        });
        if (idToNamesRA[profId] && idToNamesRA[profId].size > 1) {
          issues.push('Professional ID has multiple different names');
        }
      }

    } else if (fileType === 'supportSite') {
      const profId = String(row['professionalId'] || '').trim();

      // Check blank Professional ID
      if (!profId) {
        issues.push('Blank Professional ID');
      }

      // Check non-numeric Professional ID
      if (profId && !/^\d+$/.test(profId)) {
        issues.push('Non-numeric Professional ID');
      }

      // Check blank siteNum
      if (!String(row['siteNum'] || '').trim()) {
        issues.push('Blank Site Number');
      }

      // Check blank referralLastSent
      if (!String(row['referralLastSent'] || '').trim()) {
        issues.push('Blank Last Referral Sent');
      }

      // Check if this Professional ID has multiple different names
      if (profId) {
        const idToNamesSS = {};
        allRows.forEach(r => {
          const id = String(r['professionalId'] || '').trim();
          if (id) {
            const fullName = String(r['userFullName'] || '').trim().toLowerCase();
            if (!idToNamesSS[id]) {
              idToNamesSS[id] = new Set();
            }
            if (fullName) {
              idToNamesSS[id].add(fullName);
            }
          }
        });
        if (idToNamesSS[profId] && idToNamesSS[profId].size > 1) {
          issues.push('Professional ID has multiple different names');
        }
      }

    } else if (fileType === 'ldgLedger') {
      const cpso = String(row['CPSO #'] || '').trim();

      // Check blank CPSO #
      if (!cpso) {
        issues.push('Blank CPSO #');
      }

      // Check non-numeric CPSO #
      if (cpso && !/^\d+$/.test(cpso)) {
        issues.push('Non-numeric CPSO #');
      }

      // Check blank Ocean Site Number
      if (!String(row['Ocean Site Number (eReferral Ontario only)'] || '').trim()) {
        issues.push('Blank Ocean Site Number');
      }

      // Check blank Date Onboarding Completed
      if (!String(row['Date Onboarding Completed'] || '').trim()) {
        issues.push('Blank Onboarding Date');
      }

      // Check blank Primary Specialty Pathway
      if (!String(row['Primary Specialty Pathway'] || '').trim()) {
        issues.push('Blank Specialty Pathway');
      }

      // Check if this CPSO # has multiple different names
      if (cpso) {
        const idToNamesLDG = {};
        allRows.forEach(r => {
          const id = String(r['CPSO #'] || '').trim();
          if (id) {
            const firstName = String(r['First Name'] || '').trim().toLowerCase();
            const lastName = String(r['Last Name'] || '').trim().toLowerCase();
            const fullName = `${firstName} ${lastName}`.trim();
            if (!idToNamesLDG[id]) {
              idToNamesLDG[id] = new Set();
            }
            if (fullName) {
              idToNamesLDG[id].add(fullName);
            }
          }
        });
        if (idToNamesLDG[cpso] && idToNamesLDG[cpso].size > 1) {
          issues.push('CPSO # has multiple different names');
        }
      }
    }

    return issues;
  }, []);

  // Export data with quality flags
  const exportWithQualityFlags = useCallback((fileType) => {
    const rows = data[fileType];
    const file = files[fileType];
    if (!rows || rows.length === 0 || !file) return;

    // Build enhanced data with quality flag columns
    const enhancedData = rows.map((row, idx) => {
      const issues = getRowQualityIssues(fileType, row, idx, rows);
      return {
        ...row,
        'Row_Contains_Data_Quality_Issue': issues.length > 0 ? 'TRUE' : 'FALSE',
        'Data_Quality_Issue_Description': issues.join('; ')
      };
    });

    // Generate filename
    const originalName = file.name.replace(/\.[^/.]+$/, ''); // Remove extension
    const filename = `${originalName}_QualityFlags`;

    // Export using XLSX
    const ws = XLSX.utils.json_to_sheet(enhancedData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    XLSX.writeFile(wb, `${filename}.xlsx`);
  }, [data, files, getRowQualityIssues]);

  // Handle file upload
  const handleFileUpload = useCallback(async (fileType, file) => {
    if (!file) return;

    try {
      const result = await parseFile(file, fileType);

      // Run comprehensive quality checks
      const checks = runQualityChecks(fileType, result.data);

      // Build warning messages from quality checks for backward compatibility
      let fileWarnings = [...result.warnings];
      checks.critical.forEach(issue => {
        fileWarnings.push(`Critical: ${issue.label} (${issue.count} rows) - ${issue.impact}`);
      });
      checks.warnings.forEach(issue => {
        fileWarnings.push(`Warning: ${issue.label} (${issue.count} rows) - ${issue.impact}`);
      });

      setFiles(prev => ({ ...prev, [fileType]: file }));
      setData(prev => ({ ...prev, [fileType]: result.data }));
      setWarnings(prev => ({ ...prev, [fileType]: fileWarnings }));
      setQualityChecks(prev => ({ ...prev, [fileType]: checks }));
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
  }, [parseFile, runQualityChecks]);

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
        Specialty_Pathway: 'Value not contained in Source files',
        LDG_Name: geoData?.ldg || '',
        LDG_Lead_Org: geoData?.ldgLeadOrg || '',
        Region: geoData?.region || row['healthRegion'] || '',
        Network: 'OH',
        Solution_Type: '',
        Ocean_Site_Number: siteNum,
        Ocean_Site_Name: siteName,
        Hospital: geoData?.hospital || '',
        Address_Postal: geoData?.postalCode || row['postalCode'] || '',
        Lead_Reach: geoData?.leadReach || '',
        Role: '',
        Date_Onboarded: '',
        Training_Date: 'Value not contained in Source files',
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
        // Only update Clinician_Type if not already set
        allRecords[existingIdx].Clinician_Type = allRecords[existingIdx].Clinician_Type || row['clinicianType'] || '';
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
        Specialty_Pathway: 'Value not contained in Source files',
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
        Training_Date: 'Value not contained in Source files',
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
        Training_Date: 'Value not contained in Source files',
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

  // Calculate pre-processing quality overview
  const preProcessingOverview = useMemo(() => {
    const fileTypes = ['geoSpatial', 'regionalAuthority', 'supportSite', 'ldgLedger'];
    const fileLabels = {
      geoSpatial: 'Geo-Spatial LDG',
      regionalAuthority: 'Regional Authority',
      supportSite: 'Support Site Analytics',
      ldgLedger: 'LDG Onboarding Ledger'
    };

    let totalCritical = 0;
    let totalWarnings = 0;
    let filesWithCritical = [];
    let filesWithWarnings = [];
    let allCriticalIssues = [];
    let allWarningIssues = [];

    fileTypes.forEach(ft => {
      const checks = qualityChecks[ft];
      if (checks) {
        if (checks.critical.length > 0) {
          totalCritical += checks.critical.length;
          filesWithCritical.push(fileLabels[ft]);
          checks.critical.forEach(issue => {
            allCriticalIssues.push({ ...issue, source: fileLabels[ft] });
          });
        }
        if (checks.warnings.length > 0) {
          totalWarnings += checks.warnings.length;
          if (!filesWithCritical.includes(fileLabels[ft])) {
            filesWithWarnings.push(fileLabels[ft]);
          }
          checks.warnings.forEach(issue => {
            allWarningIssues.push({ ...issue, source: fileLabels[ft] });
          });
        }
      }
    });

    const hasAnyFile = fileTypes.some(ft => files[ft]);
    const hasAnyIssues = totalCritical > 0 || totalWarnings > 0;

    return {
      totalCritical,
      totalWarnings,
      filesWithCritical,
      filesWithWarnings,
      allCriticalIssues,
      allWarningIssues,
      hasAnyFile,
      hasAnyIssues,
      overallStatus: totalCritical > 0 ? 'critical' : totalWarnings > 0 ? 'warning' : 'good'
    };
  }, [qualityChecks, files]);

  // Track expanded quality panels
  const [expandedQuality, setExpandedQuality] = useState({
    geoSpatial: false,
    regionalAuthority: false,
    supportSite: false,
    ldgLedger: false
  });

  // File upload component
  const FileUploadBox = ({ fileType, label, description }) => {
    const file = files[fileType];
    const recordCount = data[fileType].length;
    const checks = qualityChecks[fileType];
    const isExpanded = expandedQuality[fileType];

    // Determine status based on quality checks
    const hasCritical = checks?.critical?.length > 0;
    const hasWarnings = checks?.warnings?.length > 0;
    const hasIssues = hasCritical || hasWarnings;

    // Determine colors based on file state and quality status
    const cardStyle = !file
      ? 'border-gray-300 hover:border-blue-400'
      : hasCritical
        ? 'border-red-400 bg-red-50'
        : hasWarnings
          ? 'border-yellow-400 bg-yellow-50'
          : 'border-green-400 bg-green-50';

    const iconColor = !file
      ? 'text-gray-400'
      : hasCritical
        ? 'text-red-600'
        : hasWarnings
          ? 'text-yellow-600'
          : 'text-green-600';

    const statusColor = hasCritical
      ? 'text-red-700'
      : hasWarnings
        ? 'text-yellow-700'
        : 'text-green-700';

    const StatusIcon = hasCritical ? XCircle : hasWarnings ? AlertTriangle : CheckCircle;

    return (
      <div className={`border-2 border-dashed rounded-lg p-4 transition-all ${cardStyle}`}>
        <div className="flex items-start gap-3">
          <FileSpreadsheet className={`w-8 h-8 ${iconColor}`} />
          <div className="flex-1">
            <h3 className="font-semibold text-gray-800">{label}</h3>
            <p className="text-sm text-gray-500 mb-2">{description}</p>

            {file ? (
              <div className="space-y-2">
                <div className={`flex items-center gap-2 ${statusColor}`}>
                  <StatusIcon className="w-4 h-4" />
                  <span className="text-sm font-medium">{file.name}</span>
                </div>
                <p className="text-sm text-gray-600">{recordCount.toLocaleString()} records loaded</p>

                {/* Quality Status Summary */}
                {checks && (
                  <div className="flex items-center gap-2 text-xs">
                    {hasCritical && (
                      <span className="px-2 py-0.5 bg-red-200 text-red-800 rounded-full">
                        {checks.critical.length} critical
                      </span>
                    )}
                    {hasWarnings && (
                      <span className="px-2 py-0.5 bg-yellow-200 text-yellow-800 rounded-full">
                        {checks.warnings.length} warning{checks.warnings.length !== 1 ? 's' : ''}
                      </span>
                    )}
                    {!hasIssues && (
                      <span className="px-2 py-0.5 bg-green-200 text-green-800 rounded-full">
                        No issues found
                      </span>
                    )}
                  </div>
                )}

                {/* Expandable Quality Details */}
                {hasIssues && (
                  <div className="mt-2">
                    <button
                      onClick={() => setExpandedQuality(prev => ({ ...prev, [fileType]: !prev[fileType] }))}
                      className="flex items-center gap-1 text-sm text-gray-600 hover:text-gray-800"
                    >
                      {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                      <span>{isExpanded ? 'Hide' : 'Show'} quality details</span>
                    </button>

                    {isExpanded && (
                      <div className="mt-2 space-y-2">
                        {/* Critical Issues */}
                        {checks.critical.map((issue, idx) => (
                          <div key={`critical-${idx}`} className="p-2 bg-red-100 rounded border border-red-300">
                            <div className="flex items-start gap-2">
                              <XCircle className="w-4 h-4 text-red-600 mt-0.5 flex-shrink-0" />
                              <div className="text-sm">
                                <p className="font-medium text-red-800">{issue.label} ({issue.count} rows)</p>
                                <p className="text-red-700 text-xs mt-0.5">{issue.impact}</p>
                                {issue.sampleRows.length > 0 && (
                                  <p className="text-red-600 text-xs mt-1">
                                    Examples: {issue.sampleRows.join('; ')}
                                  </p>
                                )}
                              </div>
                            </div>
                          </div>
                        ))}

                        {/* Warnings */}
                        {checks.warnings.map((issue, idx) => (
                          <div key={`warning-${idx}`} className="p-2 bg-yellow-100 rounded border border-yellow-300">
                            <div className="flex items-start gap-2">
                              <AlertTriangle className="w-4 h-4 text-yellow-600 mt-0.5 flex-shrink-0" />
                              <div className="text-sm">
                                <p className="font-medium text-yellow-800">{issue.label} ({issue.count} rows)</p>
                                <p className="text-yellow-700 text-xs mt-0.5">{issue.impact}</p>
                                {issue.sampleRows.length > 0 && (
                                  <p className="text-yellow-600 text-xs mt-1">
                                    Examples: {issue.sampleRows.join('; ')}
                                  </p>
                                )}
                              </div>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}

                    {/* Download with Quality Flags Button */}
                    <button
                      onClick={() => exportWithQualityFlags(fileType)}
                      className="mt-2 flex items-center gap-1 text-xs px-3 py-1.5 bg-blue-100 text-blue-700 rounded border border-blue-300 hover:bg-blue-200 transition"
                    >
                      <Download className="w-3 h-3" />
                      <span>Download with Quality Flags</span>
                    </button>
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
                setQualityChecks(prev => ({ ...prev, [fileType]: null }));
                setExpandedQuality(prev => ({ ...prev, [fileType]: false }));
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

            {/* Pre-Processing Quality Overview */}
            {preProcessingOverview.hasAnyFile && preProcessingOverview.hasAnyIssues && (
              <div className={`rounded-lg p-4 ${
                preProcessingOverview.overallStatus === 'critical'
                  ? 'bg-red-50 border border-red-200'
                  : 'bg-yellow-50 border border-yellow-200'
              }`}>
                <div className="flex items-start gap-3">
                  {preProcessingOverview.overallStatus === 'critical' ? (
                    <XCircle className="w-5 h-5 text-red-600 mt-0.5" />
                  ) : (
                    <AlertTriangle className="w-5 h-5 text-yellow-600 mt-0.5" />
                  )}
                  <div className="flex-1">
                    <h3 className={`font-medium ${
                      preProcessingOverview.overallStatus === 'critical' ? 'text-red-800' : 'text-yellow-800'
                    }`}>
                      Pre-Processing Quality Overview
                    </h3>
                    <div className="mt-2 text-sm space-y-1">
                      {preProcessingOverview.totalCritical > 0 && (
                        <p className="text-red-700">
                          <span className="font-semibold">{preProcessingOverview.totalCritical} critical issue{preProcessingOverview.totalCritical !== 1 ? 's' : ''}</span>
                          {' '}in {preProcessingOverview.filesWithCritical.join(', ')}
                        </p>
                      )}
                      {preProcessingOverview.totalWarnings > 0 && (
                        <p className="text-yellow-700">
                          <span className="font-semibold">{preProcessingOverview.totalWarnings} warning{preProcessingOverview.totalWarnings !== 1 ? 's' : ''}</span>
                          {' '}across loaded files
                        </p>
                      )}
                    </div>

                    {preProcessingOverview.overallStatus === 'critical' && (
                      <div className="mt-3 p-3 bg-red-100 rounded border border-red-300">
                        <p className="text-sm font-medium text-red-800 mb-2">Impact on Processing:</p>
                        <ul className="text-xs text-red-700 space-y-1 list-disc list-inside">
                          {preProcessingOverview.allCriticalIssues.map((issue, idx) => (
                            <li key={idx}>
                              <span className="font-medium">{issue.source}</span>: {issue.label} ({issue.count} rows) - {issue.impact}
                            </li>
                          ))}
                        </ul>
                      </div>
                    )}

                    {preProcessingOverview.overallStatus === 'warning' && preProcessingOverview.allWarningIssues.length > 0 && (
                      <div className="mt-3 p-3 bg-yellow-100 rounded border border-yellow-300">
                        <p className="text-sm font-medium text-yellow-800 mb-2">Potential Impact:</p>
                        <ul className="text-xs text-yellow-700 space-y-1 list-disc list-inside">
                          {preProcessingOverview.allWarningIssues.slice(0, 5).map((issue, idx) => (
                            <li key={idx}>
                              <span className="font-medium">{issue.source}</span>: {issue.label} ({issue.count} rows)
                            </li>
                          ))}
                          {preProcessingOverview.allWarningIssues.length > 5 && (
                            <li className="italic">...and {preProcessingOverview.allWarningIssues.length - 5} more warnings</li>
                          )}
                        </ul>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}

            {data.geoSpatial.length > 0 && (
              <div className="flex flex-col items-center pt-4 gap-2">
                {preProcessingOverview.overallStatus === 'critical' && (
                  <p className="text-sm text-red-600 flex items-center gap-1">
                    <AlertTriangle className="w-4 h-4" />
                    Critical issues detected - some records will be skipped during processing
                  </p>
                )}
                <button
                  onClick={processData}
                  className={`px-6 py-3 font-semibold rounded-lg transition flex items-center gap-2 ${
                    preProcessingOverview.overallStatus === 'critical'
                      ? 'bg-orange-600 text-white hover:bg-orange-700'
                      : 'bg-green-600 text-white hover:bg-green-700'
                  }`}
                >
                  <Settings className="w-5 h-5" />
                  {preProcessingOverview.overallStatus === 'critical' ? 'Process Data (with issues)' : 'Process Data'}
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
