import React, { useState } from 'react';
import { Upload, Download, FileSpreadsheet, Home, Users, Eye, AlertCircle, CheckCircle, Loader, Table, Grid3X3 } from 'lucide-react';
import ExcelJS from 'exceljs';

const App = () => {
  const [currentPage, setCurrentPage] = useState('home');
  const [files, setFiles] = useState({
    year4: null,
    year3: null,
    year2: null
  });
  const [processing, setProcessing] = useState(false);
  const [combinedData, setCombinedData] = useState(null);
  const [previewData, setPreviewData] = useState([]);
  const [previewMode, setPreviewMode] = useState('table');
  const [googleSheetUrl, setGoogleSheetUrl] = useState('');
  const [sheetLoading, setSheetLoading] = useState(false);
  const [statistics, setStatistics] = useState(null);
  const [error, setError] = useState('');

  // Branch normalization mapping from Python script
  const normalizeBranchName = (branchName) => {
    const branchMapping = {
      'CSE(AIML)': 'Computer Science Engineering (AI & ML)',
      'AIML': 'Artificial Intelligence & Machine Learning',
      'CSE': 'Computer Science Engineering',
      'Data Science': 'CSE (Data Science)',
      'Cyber Security': 'CSE (Cyber Security)',
      'Aerospace Eng.': 'Aerospace Engineering',
      'Civil Eng.': 'Civil Engineering',
      'Chemical Eng.': 'Chemical Engineering',
      'Mechanical Eng.': 'Mechanical Engineering',
      'Information Science': 'Information Science & Engineering',
      'Biotechnology': 'Biotechnology',
      'EEE': 'Electrical & Electronics Engineering',
      'ECE': 'Electronics & Communication Engineering',
      'EIE': 'Electronics & Instrumentation Engineering',
      'ET': 'Electronics & Telecommunication Engineering',
      'IEM': 'Industrial Engineering & Management'
    };
    return branchMapping[branchName] || branchName;
  };

  // Extract batch info from filename
  const extractBatchInfo = (filename) => {
    if (filename.includes('2024') && filename.includes('2028')) {
      return { batch: '2024-2028', year: 'Year 2' };
    } else if (filename.includes('2023') && filename.includes('2027')) {
      return { batch: '2023-2027', year: 'Year 3' };
    } else if (filename.includes('2022') && filename.includes('2026')) {
      return { batch: '2022-2026', year: 'Year 4' };
    }
    return { batch: 'Unknown Batch', year: 'Unknown Year' };
  };

  // Clean cell values (remove NaN, null, 0 etc.)
  const cleanValue = (value) => {
    if (!value || value === '' || value === 0 || 
        String(value).toLowerCase() === 'nan' || 
        String(value).toLowerCase() === 'none' || 
        String(value).toLowerCase() === 'null') {
      return '';
    }
    return String(value).trim();
  };

  // Process individual worksheet
  const processSheet = async (worksheet, sheetName, batchYear) => {
    const data = [];
    
    // Convert worksheet to array format
    worksheet.eachRow((row, rowNumber) => {
      const rowData = [];
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        rowData[colNumber - 1] = cell.value || '';
      });
      data.push(rowData);
    });
    
    if (data.length <= 1) return [];

    // Find header row containing 'USN'
    let headerRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i].some(cell => String(cell).toUpperCase().includes('USN'))) {
        headerRow = i;
        break;
      }
    }

    if (headerRow === -1) return [];

    const headers = data[headerRow];
    const rows = data.slice(headerRow + 1).filter(row => 
      row.some(cell => cell && String(cell).trim() !== '')
    );

    // Column mapping from Python script
    const columnMapping = {
      'USN': 'USN',
      'FULL NAME': 'Full Name',
      'BRANCH ': 'Branch',
      'BRANCH': 'Branch',
      'SECTION': 'Section',
      'EMAIL': 'Email',
      'PHONE NUMBER': 'Phone Number',
      'COUNSELLOR': 'Counsellor',
      'E-Mail ID of the Counsellors': 'Counsellor Email',
      'COUNSELLOR DEPT.': 'Counsellor Department',
      'BATCH(20XX-20XX)': 'Batch',
      'BATCH': 'Batch'
    };

    // Process rows into standardized records
    const records = rows.map(row => {
      const record = {};
      headers.forEach((header, index) => {
        const standardHeader = columnMapping[String(header).toUpperCase()] || header;
        record[standardHeader] = cleanValue(row[index]);
      });

      // Ensure required fields exist
      const requiredFields = ['USN', 'Full Name', 'Branch', 'Section', 'Email', 
                            'Phone Number', 'Counsellor', 'Counsellor Email', 
                            'Counsellor Department', 'Batch'];
      
      requiredFields.forEach(field => {
        if (!record[field]) record[field] = '';
      });

      // Add normalized fields
      record['Student Branch'] = normalizeBranchName(record['Branch'] || '');
      record['Student Batch'] = record['Batch'] || batchYear;
      record['Student Email ID'] = record['Email'] || '';
      record['Counsellor Name'] = record['Counsellor'] || '';
      record['Counsellor Email ID'] = record['Counsellor Email'] || '';

      return record;
    }).filter(record => record['USN'] && record['USN'] !== '');

    return records;
  };

  // Main file processing function
  const combineFiles = async () => {
    if (!files.year2 && !files.year3 && !files.year4) {
      setError('Please upload at least one file');
      return;
    }

    setProcessing(true);
    setError('');

    try {
      const combinedRecords = [];
      const stats = {
        totalFiles: 0,
        totalSheets: 0,
        totalRecords: 0,
        branchesProcessed: new Set(),
        batchesProcessed: new Set()
      };

      // Process files in chronological order (Year 4, Year 3, Year 2)
      const fileOrder = [
        { key: 'year4', file: files.year4 },
        { key: 'year3', file: files.year3 },
        { key: 'year2', file: files.year2 }
      ];

      for (const { file } of fileOrder) {
        if (!file) continue;

        const batchInfo = extractBatchInfo(file.name);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(await file.arrayBuffer());
        
        // Add batch separator
        combinedRecords.push({
          type: 'batch_separator',
          text: `${batchInfo.batch} Batch (${batchInfo.year})`
        });

        // Get valid sheets (exclude utility sheets)
        const validSheets = [];
        workbook.eachSheet((worksheet, sheetId) => {
          const sheetName = worksheet.name;
          
          // Skip utility sheets
          if (sheetName.startsWith('Sheet') && sheetName !== 'Sheet1') return;
          if (['template', 'format', 'example', 'blank'].includes(sheetName.toLowerCase())) return;
          
          // Check if sheet has USN data
          let hasValidData = false;
          worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 10) return; // Only check first 10 rows
            row.eachCell((cell) => {
              if (String(cell.value).toUpperCase().includes('USN')) {
                hasValidData = true;
              }
            });
          });
          
          if (hasValidData) {
            validSheets.push({ name: sheetName, worksheet });
          }
        });

        // Sort sheets alphabetically by normalized branch names
        validSheets.sort((a, b) => 
          normalizeBranchName(a.name).toLowerCase().localeCompare(normalizeBranchName(b.name).toLowerCase())
        );

        // Process each valid sheet
        for (const { name: sheetName, worksheet } of validSheets) {
          const sheetData = await processSheet(worksheet, sheetName, batchInfo.batch);
          
          if (sheetData.length > 0) {
            // Add branch separator
            const normalizedBranch = normalizeBranchName(sheetName);
            combinedRecords.push({
              type: 'branch_separator',
              text: normalizedBranch
            });

            // Add student records
            combinedRecords.push(...sheetData);

            // Update statistics
            stats.branchesProcessed.add(normalizedBranch);
            stats.totalRecords += sheetData.length;
            stats.totalSheets++;
          }
        }

        stats.batchesProcessed.add(batchInfo.batch);
        stats.totalFiles++;
      }

      setCombinedData(combinedRecords);
      setPreviewData(combinedRecords.slice(0, 50));
      setStatistics({
        ...stats,
        branchesProcessed: Array.from(stats.branchesProcessed),
        batchesProcessed: Array.from(stats.batchesProcessed)
      });

    } catch (err) {
      setError(`Error processing files: ${err.message}`);
    } finally {
      setProcessing(false);
    }
  };

  // Create Google Sheet with data
  const createGoogleSheet = async (data) => {
    setSheetLoading(true);
    
    try {
      // Prepare data for Google Sheets
      const sheetData = [];
      
      // Add logo space
      sheetData.push(['LOGO SPACE - Insert RVCE Logo Here']);
      for (let i = 1; i < 8; i++) {
        sheetData.push(['']);
      }
      
      // Add headers
      const headers = [
        'USN', 'Full Name', 'Student Branch', 'Section', 'Student Email ID',
        'Phone Number', 'Counsellor Name', 'Counsellor Email ID',
        'Counsellor Department', 'Student Batch'
      ];
      sheetData.push(headers);
      
      // Add data
      data.forEach(item => {
        if (item.type === 'batch_separator') {
          sheetData.push([item.text]);
          sheetData.push(['']); // Empty row
        } else if (item.type === 'branch_separator') {
          sheetData.push([item.text]);
          sheetData.push(['']); // Empty row
        } else {
          // Student data
          sheetData.push([
            item.USN || '',
            item['Full Name'] || '',
            item['Student Branch'] || '',
            item.Section || '',
            item['Student Email ID'] || '',
            item['Phone Number'] || '',
            item['Counsellor Name'] || '',
            item['Counsellor Email ID'] || '',
            item['Counsellor Department'] || '',
            item['Student Batch'] || ''
          ]);
        }
      });

      // Create Google Sheet via API
      const response = await fetch('/api/create-sheet', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ 
          data: sheetData,
          title: `RVCE Counsellor Data - ${new Date().toISOString().split('T')[0]}`
        }),
      });

      if (!response.ok) {
        throw new Error('Failed to create Google Sheet');
      }

      const result = await response.json();
      setGoogleSheetUrl(result.url);
      
    } catch (error) {
      console.error('Error creating Google Sheet:', error);
      setError('Failed to create Google Sheet preview. Please try table view.');
    } finally {
      setSheetLoading(false);
    }
  };

  // Handle Excel view mode
  const handleExcelViewClick = async () => {
    setPreviewMode('excel');
    if (combinedData && !googleSheetUrl) {
      await createGoogleSheet(combinedData);
    }
  };

  // Download Excel file with professional formatting
  const downloadExcel = async () => {
    if (!combinedData) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Combined Data');

    const headers = [
      'USN', 'Full Name', 'Student Branch', 'Section', 'Student Email ID',
      'Phone Number', 'Counsellor Name', 'Counsellor Email ID',
      'Counsellor Department', 'Student Batch'
    ];

    let rowIndex = 1;

    // Logo space (rows 1-8) - merged cells
    worksheet.mergeCells('A1:J8');
    const logoCell = worksheet.getCell('A1');
    logoCell.value = 'LOGO SPACE - Insert RVCE Logo Here';
    logoCell.font = { size: 14, color: { argb: '666666' }, italic: true };
    logoCell.alignment = { horizontal: 'center', vertical: 'middle' };
    logoCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'F8F9FA' }
    };
    
    rowIndex = 9;

    // Professional headers
    worksheet.addRow(headers);
    const headerRow = worksheet.getRow(rowIndex);
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFF' }, size: 12 };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '366092' }
      };
      cell.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFFFFF' } },
        left: { style: 'thin', color: { argb: 'FFFFFF' } },
        bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
        right: { style: 'thin', color: { argb: 'FFFFFF' } }
      };
    });
    
    rowIndex++;

    // Add data with formatting
    combinedData.forEach(item => {
      if (item.type === 'batch_separator') {
        // Batch separator with merged cells
        worksheet.mergeCells(`A${rowIndex}:J${rowIndex}`);
        const batchCell = worksheet.getCell(`A${rowIndex}`);
        batchCell.value = item.text;
        batchCell.font = { bold: true, color: { argb: 'FFFFFF' }, size: 16 };
        batchCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '0B5394' }
        };
        batchCell.alignment = { horizontal: 'center', vertical: 'center' };
        worksheet.getRow(rowIndex).height = 30;
        rowIndex++;
        rowIndex++; // Empty row
      } else if (item.type === 'branch_separator') {
        // Branch separator with merged cells
        worksheet.mergeCells(`A${rowIndex}:J${rowIndex}`);
        const branchCell = worksheet.getCell(`A${rowIndex}`);
        branchCell.value = item.text;
        branchCell.font = { bold: true, color: { argb: 'FFFFFF' }, size: 14 };
        branchCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '6AA84F' }
        };
        branchCell.alignment = { horizontal: 'center', vertical: 'center' };
        worksheet.getRow(rowIndex).height = 25;
        rowIndex++;
        rowIndex++; // Empty row
      } else {
        // Student data
        const rowData = [
          item.USN || '',
          item['Full Name'] || '',
          item['Student Branch'] || '',
          item.Section || '',
          item['Student Email ID'] || '',
          item['Phone Number'] || '',
          item['Counsellor Name'] || '',
          item['Counsellor Email ID'] || '',
          item['Counsellor Department'] || '',
          item['Student Batch'] || ''
        ];
        
        const dataRow = worksheet.addRow(rowData);
        
        // Apply alternating row colors and borders
        dataRow.eachCell((cell, colNumber) => {
          if (rowIndex % 2 === 0) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'F8F9FA' }
            };
          }
          
          cell.border = {
            top: { style: 'thin', color: { argb: 'E1E1E1' } },
            left: { style: 'thin', color: { argb: 'E1E1E1' } },
            bottom: { style: 'thin', color: { argb: 'E1E1E1' } },
            right: { style: 'thin', color: { argb: 'E1E1E1' } }
          };
          
          // Alignment based on column
          if (colNumber === 1) {
            cell.alignment = { horizontal: 'center', vertical: 'center' };
          } else {
            cell.alignment = { horizontal: 'left', vertical: 'center' };
          }
        });
        
        rowIndex++;
      }
    });

    // Set professional column widths
    const columnWidths = [15, 30, 35, 10, 40, 15, 25, 40, 20, 15];
    columnWidths.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width;
    });

    // Download the file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'Combined_Student_Counsellor_Data_Professional.xlsx';
    link.click();
    window.URL.revokeObjectURL(url);
  };

  // File upload handler
  const handleFileUpload = (yearKey, file) => {
    if (file && file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
      setFiles(prev => ({ ...prev, [yearKey]: file }));
      setError('');
    } else {
      setError('Please upload only .xlsx files');
    }
  };

  // Home page component
  const HomePage = () => (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100">
      <div className="container mx-auto px-4 py-16">
        <div className="text-center mb-16">
          <div className="mb-8">
            <div className="w-32 h-32 mx-auto bg-blue-600 rounded-full flex items-center justify-center mb-6">
              <Users className="w-16 h-16 text-white" />
            </div>
          </div>
          <h1 className="text-5xl font-bold text-gray-800 mb-4">
            Combined Counsellors Data
          </h1>
          <h2 className="text-3xl font-semibold text-blue-600 mb-6">RVCE</h2>
          <p className="text-xl text-gray-600 mb-12 max-w-2xl mx-auto">
            Professional tool for combining student-counsellor data from multiple Excel files 
            with advanced formatting, preview, and download capabilities.
          </p>
          <button
            onClick={() => setCurrentPage('combine')}
            className="bg-blue-600 hover:bg-blue-700 text-white px-8 py-4 rounded-lg text-lg font-semibold transition-colors duration-200 shadow-lg"
          >
            Start Combining Data
          </button>
        </div>

        <div className="grid md:grid-cols-3 gap-8 max-w-4xl mx-auto">
          <div className="bg-white p-6 rounded-lg shadow-lg text-center">
            <Upload className="w-12 h-12 text-blue-600 mx-auto mb-4" />
            <h3 className="text-xl font-semibold mb-2">Easy Upload</h3>
            <p className="text-gray-600">Upload Excel files for each batch year with drag & drop support</p>
          </div>
          <div className="bg-white p-6 rounded-lg shadow-lg text-center">
            <Eye className="w-12 h-12 text-green-600 mx-auto mb-4" />
            <h3 className="text-xl font-semibold mb-2">Live Preview</h3>
            <p className="text-gray-600">Preview combined data before downloading with full formatting</p>
          </div>
          <div className="bg-white p-6 rounded-lg shadow-lg text-center">
            <Download className="w-12 h-12 text-purple-600 mx-auto mb-4" />
            <h3 className="text-xl font-semibold mb-2">Professional Export</h3>
            <p className="text-gray-600">Download professionally formatted Excel with logo space</p>
          </div>
        </div>
      </div>
    </div>
  );

  // Main combine page component
  const CombinePage = () => (
    <div className="min-h-screen bg-gray-50">
      <div className="bg-white shadow-sm border-b">
        <div className="container mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center space-x-4">
            <button
              onClick={() => setCurrentPage('home')}
              className="flex items-center space-x-2 text-blue-600 hover:text-blue-800"
            >
              <Home className="w-5 h-5" />
              <span>Home</span>
            </button>
            <span className="text-gray-400">/</span>
            <span className="text-gray-700 font-semibold">Combine Counsellor Data</span>
          </div>
        </div>
      </div>

      <div className="container mx-auto px-4 py-8">
        <div className="max-w-6xl mx-auto">
          <h1 className="text-3xl font-bold text-gray-800 mb-8">Combine Counsellor Data</h1>

          {/* File Upload Section */}
          <div className="bg-white rounded-lg shadow-lg p-6 mb-8">
            <h2 className="text-xl font-semibold mb-6">Upload Excel Files by Batch</h2>
            
            <div className="grid md:grid-cols-3 gap-6">
              {[
                { key: 'year4', label: 'Year 4 (2022-2026)', color: 'bg-red-50 border-red-200' },
                { key: 'year3', label: 'Year 3 (2023-2027)', color: 'bg-yellow-50 border-yellow-200' },
                { key: 'year2', label: 'Year 2 (2024-2028)', color: 'bg-green-50 border-green-200' }
              ].map(({ key, label, color }) => (
                <div key={key} className={`border-2 border-dashed rounded-lg p-6 ${color}`}>
                  <div className="text-center">
                    <FileSpreadsheet className="w-12 h-12 mx-auto mb-4 text-gray-400" />
                    <h3 className="font-semibold mb-2">{label}</h3>
                    
                    <input
                      type="file"
                      accept=".xlsx"
                      onChange={(e) => handleFileUpload(key, e.target.files[0])}
                      className="hidden"
                      id={`file-${key}`}
                    />
                    <label
                      htmlFor={`file-${key}`}
                      className="cursor-pointer bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded text-sm transition-colors duration-200"
                    >
                      Choose File
                    </label>
                    
                    {files[key] && (
                      <div className="mt-3">
                        <div className="flex items-center justify-center text-sm text-green-600">
                          <CheckCircle className="w-4 h-4 mr-1" />
                          {files[key].name}
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              ))}
            </div>

            {error && (
              <div className="mt-4 bg-red-50 border border-red-200 rounded-lg p-4">
                <div className="flex items-center text-red-700">
                  <AlertCircle className="w-5 h-5 mr-2" />
                  {error}
                </div>
              </div>
            )}

            <div className="mt-6 text-center">
              <button
                onClick={combineFiles}
                disabled={processing || (!files.year2 && !files.year3 && !files.year4)}
                className="bg-green-600 hover:bg-green-700 disabled:bg-gray-400 text-white px-8 py-3 rounded-lg font-semibold transition-colors duration-200 shadow-lg"
              >
                {processing ? (
                  <div className="flex items-center">
                    <Loader className="w-5 h-5 mr-2 animate-spin" />
                    Processing...
                  </div>
                ) : (
                  'Combine Data'
                )}
              </button>
            </div>
          </div>

          {/* Statistics */}
          {statistics && (
            <div className="bg-white rounded-lg shadow-lg p-6 mb-8">
              <h2 className="text-xl font-semibold mb-4">Processing Statistics</h2>
              <div className="grid md:grid-cols-4 gap-4">
                <div className="text-center">
                  <div className="text-2xl font-bold text-blue-600">{statistics.totalFiles}</div>
                  <div className="text-sm text-gray-600">Files Processed</div>
                </div>
                <div className="text-center">
                  <div className="text-2xl font-bold text-green-600">{statistics.totalSheets}</div>
                  <div className="text-sm text-gray-600">Sheets Processed</div>
                </div>
                <div className="text-center">
                  <div className="text-2xl font-bold text-purple-600">{statistics.totalRecords}</div>
                  <div className="text-sm text-gray-600">Total Records</div>
                </div>
                <div className="text-center">
                  <div className="text-2xl font-bold text-orange-600">{statistics.branchesProcessed.length}</div>
                  <div className="text-sm text-gray-600">Branches Found</div>
                </div>
              </div>
              
              <div className="mt-4">
                <div className="text-sm text-gray-600 mb-2">Batches: {statistics.batchesProcessed.join(', ')}</div>
                <div className="text-sm text-gray-600">Branches: {statistics.branchesProcessed.join(', ')}</div>
              </div>
            </div>
          )}

          {/* Preview and Download */}
          {combinedData && (
            <div className="bg-white rounded-lg shadow-lg p-6">
              <div className="flex justify-between items-center mb-6">
                <div className="flex items-center space-x-4">
                  <h2 className="text-xl font-semibold">Data Preview</h2>
                  <div className="flex bg-gray-100 rounded-lg p-1">
                    <button
                      onClick={() => setPreviewMode('table')}
                      className={`flex items-center px-3 py-1 rounded text-sm transition-colors ${
                        previewMode === 'table' 
                          ? 'bg-white shadow text-blue-600' 
                          : 'text-gray-600 hover:text-gray-800'
                      }`}
                    >
                      <Table className="w-4 h-4 mr-1" />
                      Table View
                    </button>
                    <button
                      onClick={handleExcelViewClick}
                      className={`flex items-center px-3 py-1 rounded text-sm transition-colors ${
                        previewMode === 'excel' 
                          ? 'bg-white shadow text-green-600' 
                          : 'text-gray-600 hover:text-gray-800'
                      }`}
                    >
                      <Grid3X3 className="w-4 h-4 mr-1" />
                      Google Sheets View
                    </button>
                  </div>
                </div>
                <button
                  onClick={downloadExcel}
                  className="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded-lg font-semibold transition-colors duration-200 shadow-lg flex items-center"
                >
                  <Download className="w-5 h-5 mr-2" />
                  Download Excel
                </button>
              </div>

              <div className="preview-container">
                {previewMode === 'table' ? (
                  <div className="overflow-x-auto">
                    <div className="max-h-96 overflow-y-auto border rounded-lg">
                      <table className="w-full text-sm">
                        <thead className="bg-blue-600 text-white sticky top-0">
                          <tr>
                            <th className="px-4 py-2 text-left">USN</th>
                            <th className="px-4 py-2 text-left">Full Name</th>
                            <th className="px-4 py-2 text-left">Branch</th>
                            <th className="px-4 py-2 text-left">Section</th>
                            <th className="px-4 py-2 text-left">Email</th>
                            <th className="px-4 py-2 text-left">Counsellor</th>
                            <th className="px-4 py-2 text-left">Batch</th>
                          </tr>
                        </thead>
                        <tbody>
                          {previewData.map((item, index) => {
                            if (item.type === 'batch_separator') {
                              return (
                                <tr key={index} className="bg-blue-100">
                                  <td colSpan="7" className="px-4 py-3 font-bold text-center text-blue-800">
                                    {item.text}
                                  </td>
                                </tr>
                              );
                            } else if (item.type === 'branch_separator') {
                              return (
                                <tr key={index} className="bg-green-100">
                                  <td colSpan="7" className="px-4 py-2 font-semibold text-center text-green-800">
                                    {item.text}
                                  </td>
                                </tr>
                              );
                            } else {
                              return (
                                <tr key={index} className={index % 2 === 0 ? 'bg-gray-50' : 'bg-white'}>
                                  <td className="px-4 py-2 border-b">{item.USN}</td>
                                  <td className="px-4 py-2 border-b">{item['Full Name']}</td>
                                  <td className="px-4 py-2 border-b">{item['Student Branch']}</td>
                                  <td className="px-4 py-2 border-b">{item.Section}</td>
                                  <td className="px-4 py-2 border-b">{item['Student Email ID']}</td>
                                  <td className="px-4 py-2 border-b">{item['Counsellor Name']}</td>
                                  <td className="px-4 py-2 border-b">{item['Student Batch']}</td>
                                </tr>
                              );
                            }
                          })}
                        </tbody>
                      </table>
                    </div>
                    {combinedData.length > 50 && (
                      <p className="text-sm text-gray-600 mt-2 text-center">
                        Showing first 50 items. Full data will be included in download.
                      </p>
                    )}
                  </div>
                ) : (
                  <div className="border rounded-lg overflow-hidden">
                    <div className="bg-gray-50 p-3 border-b">
                      <div className="flex items-center justify-between">
                        <p className="text-sm text-gray-600">
                          Google Sheets preview with full Excel functionality
                        </p>
                        <div className="text-xs text-gray-500">
                          Interactive spreadsheet view
                        </div>
                      </div>
                    </div>
                    
                    {sheetLoading ? (
                      <div className="flex items-center justify-center h-96 bg-gray-50">
                        <div className="text-center">
                          <Loader className="w-8 h-8 animate-spin mx-auto mb-2 text-green-600" />
                          <p className="text-gray-600">Creating Google Sheet...</p>
                        </div>
                      </div>
                    ) : googleSheetUrl ? (
                      <div className="h-96">
                        <iframe
                          src={`${googleSheetUrl}&rm=embedded`}
                          width="100%"
                          height="100%"
                          frameBorder="0"
                          className="rounded-b-lg"
                          title="Google Sheets Preview"
                        />
                      </div>
                    ) : (
                      <div className="flex items-center justify-center h-96 bg-gray-50">
                        <div className="text-center">
                          <p className="text-gray-600 mb-4">Click "Google Sheets View" to create interactive preview</p>
                          <button
                            onClick={handleExcelViewClick}
                            className="bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-lg"
                          >
                            Create Google Sheet
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );

  return currentPage === 'home' ? <HomePage /> : <CombinePage />;
};

export default App;