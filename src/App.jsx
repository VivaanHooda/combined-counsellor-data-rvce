import React, { useState, useEffect } from 'react';
import { Upload, Download, FileText, CheckCircle, AlertCircle, Loader, Menu, X } from 'lucide-react';
import ExcelJS from 'exceljs';

const App = () => {
  const [files, setFiles] = useState({
    year4: null,
    year3: null,
    year2: null
  });
  const [processing, setProcessing] = useState(false);
  const [combinedData, setCombinedData] = useState(null);
  const [statistics, setStatistics] = useState(null);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

  // Auto-clear messages
  useEffect(() => {
    if (error || success) {
      const timer = setTimeout(() => {
        setError('');
        setSuccess('');
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [error, success]);

  // Enhanced Excel cell value extraction specifically for emails and complex objects
  const extractCellValue = (cellValue) => {
    if (cellValue === null || cellValue === undefined) {
      return '';
    }
    
    // If it's already a simple value
    if (typeof cellValue !== 'object') {
      return String(cellValue).trim();
    }
    
    // Handle Excel Date objects
    if (cellValue instanceof Date) {
      return cellValue.toLocaleDateString();
    }
    
    // Handle Excel hyperlink objects (common for emails)
    if (cellValue.hyperlink) {
      // Email hyperlinks often have the email in the hyperlink property
      if (cellValue.text) {
        return String(cellValue.text).trim();
      }
      if (typeof cellValue.hyperlink === 'string') {
        // Extract email from mailto: links
        return cellValue.hyperlink.replace('mailto:', '').trim();
      }
      if (cellValue.hyperlink.text) {
        return String(cellValue.hyperlink.text).trim();
      }
      return String(cellValue.hyperlink).trim();
    }
    
    // Handle Excel rich text objects
    if (cellValue.richText && Array.isArray(cellValue.richText)) {
      return cellValue.richText.map(rt => rt.text || '').join('').trim();
    }
    
    // Handle Excel formula objects
    if (cellValue.formula !== undefined) {
      return String(cellValue.result || cellValue.formula || '').trim();
    }
    
    // Handle objects with result property
    if (cellValue.result !== undefined) {
      return String(cellValue.result).trim();
    }
    
    // Handle objects with text property
    if (cellValue.text !== undefined) {
      return String(cellValue.text).trim();
    }
    
    // Handle objects with value property
    if (cellValue.value !== undefined) {
      return String(cellValue.value).trim();
    }
    
    // For any other object, try to extract meaningful content
    if (cellValue.toString && cellValue.toString() !== '[object Object]') {
      return String(cellValue).trim();
    }
    
    // Last resort: try to find any string property in the object
    if (typeof cellValue === 'object') {
      for (const key in cellValue) {
        if (typeof cellValue[key] === 'string' && cellValue[key].includes('@')) {
          return cellValue[key].trim(); // Likely an email
        }
        if (typeof cellValue[key] === 'string' && cellValue[key].length > 0) {
          return cellValue[key].trim();
        }
      }
    }
    
    return '';
  };

  // EXACT replica of Python script's clean_value function with enhanced Excel object handling
  const cleanValue = (value) => {
    const extractedValue = extractCellValue(value);
    
    if (!extractedValue || extractedValue === '' || extractedValue === '0' || 
        String(extractedValue).toLowerCase() === 'nan' || 
        String(extractedValue).toLowerCase() === 'none' || 
        String(extractedValue).toLowerCase() === 'null') {
      return '';
    }
    return String(extractedValue).trim();
  };

  // EXACT replica of Python script's normalize_branch_name function
  const normalizeBranchName = (branchName) => {
    const branchMapping = {
      // AI/ML branches - distinguish between the two different programs
      'CSE(AIML)': 'Computer Science Engineering (AI & ML)',  // Older program
      'AIML': 'Artificial Intelligence & Machine Learning',    // New dedicated program (2024-2028)
      
      // Regular branches with correct names
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

  // EXACT replica of Python script's extract_batch_year function
  const extractBatchYear = (filename) => {
    if (filename.includes('2024') && filename.includes('2028')) {
      return { batch: '2024-2028', year: 'Year 2' };
    } else if (filename.includes('2023') && filename.includes('2027')) {
      return { batch: '2023-2027', year: 'Year 3' };
    } else if (filename.includes('2022') && filename.includes('2026')) {
      return { batch: '2022-2026', year: 'Year 4' };
    } else {
      console.warn(`Could not extract batch year from filename: ${filename}`);
      return { batch: 'Unknown Batch', year: 'Unknown Year' };
    }
  };

  // EXACT replica of Python script's is_valid_sheet function
  const isValidSheet = (data, sheetName) => {
    if (!data || data.length === 0) {
      console.warn(`Sheet '${sheetName}' is empty`);
      return false;
    }
    
    if (data.length < 2) {
      console.warn(`Sheet '${sheetName}' has insufficient rows (${data.length})`);
      return false;
    }
    
    // Look for USN column indicator
    let hasUsn = false;
    for (const row of data) {
      if (row.some(cell => String(cell).toUpperCase().includes('USN'))) {
        hasUsn = true;
        break;
      }
    }
    
    if (!hasUsn) {
      console.warn(`Sheet '${sheetName}' does not contain USN column`);
      return false;
    }
    
    // Check if there's actual student data
    let usnPatternFound = false;
    for (const row of data) {
      for (const cell of row) {
        if (cell && typeof cell === 'string') {
          if (cell.length > 5 && /\d/.test(cell) && /[a-zA-Z]/.test(cell)) {
            usnPatternFound = true;
            break;
          }
        }
      }
      if (usnPatternFound) break;
    }
    
    if (!usnPatternFound) {
      console.warn(`Sheet '${sheetName}' does not contain valid student data`);
      return false;
    }
    
    return true;
  };

  // EXACT replica of Python script's process_sheet function
  const processSheet = (worksheet, sheetName, batchYear, fileName) => {
    const data = [];
    
    // Convert worksheet to array format with direct cell hyperlink access
    worksheet.eachRow((row, rowNumber) => {
      const rowData = [];
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        let cellValue = '';
        
        // DIRECT HYPERLINK CHECK - ExcelJS stores hyperlinks separately
        if (cell.hyperlink) {
          // Get the display text from hyperlink
          if (cell.value && typeof cell.value === 'string') {
            cellValue = cell.value;
          } else if (cell.text) {
            cellValue = cell.text;
          } else if (typeof cell.hyperlink === 'string') {
            cellValue = cell.hyperlink.replace('mailto:', '');
          } else if (cell.hyperlink.text) {
            cellValue = cell.hyperlink.text;
          } else {
            cellValue = String(cell.hyperlink).replace('mailto:', '');
          }
        } else {
          // Regular cell value extraction
          cellValue = extractCellValue(cell.value);
        }
        
        rowData[colNumber - 1] = cellValue || '';
      });
      data.push(rowData);
    });
    
    if (data.length === 0) {
      console.warn(`Empty sheet: ${sheetName} in ${fileName}`);
      return [];
    }
    
    if (data.length <= 1) {
      console.warn(`Sheet ${sheetName} has insufficient data`);
      return [];
    }
    
    // Find the actual header row
    let headerRow = null;
    for (let idx = 0; idx < data.length; idx++) {
      if (data[idx].some(cell => String(cell).toUpperCase().includes('USN'))) {
        headerRow = idx;
        break;
      }
    }
    
    if (headerRow === null) {
      console.warn(`No header row found in sheet: ${sheetName}`);
      return [];
    }
    
    // Set proper headers and get data
    const headers = data[headerRow];
    let rows = data.slice(headerRow + 1);
    
    // Clean and filter data - remove empty rows
    rows = rows.filter(row => row.some(cell => cell && String(cell).trim() !== ''));
    
    if (rows.length === 0) {
      console.warn(`No valid data rows in sheet: ${sheetName}`);
      return [];
    }
    
    // EXACT replica of Python script's column mapping
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
    
    // Process rows into records
    const records = rows.map(row => {
      const record = {};
      
      // Map columns based on headers
      headers.forEach((header, index) => {
        const standardHeader = columnMapping[String(header).toUpperCase()] || header;
        record[standardHeader] = cleanValue(row[index]);
      });
      
      // Ensure required columns exist
      const requiredColumns = ['USN', 'Full Name', 'Branch', 'Email', 'Phone Number', 'Counsellor', 'Batch'];
      requiredColumns.forEach(col => {
        if (!record[col]) record[col] = '';
      });
      
      // Add missing columns with empty values
      const allColumns = ['USN', 'Full Name', 'Branch', 'Section', 'Email', 
                         'Phone Number', 'Counsellor', 'Counsellor Email', 'Counsellor Department', 'Batch'];
      allColumns.forEach(col => {
        if (!record[col]) record[col] = '';
      });
      
      // Add normalized branch name
      record['Normalized Branch'] = normalizeBranchName(record['Branch'] || '');
      
      // Update batch info if missing
      if (!record['Batch'] || record['Batch'] === '') {
        record['Batch'] = batchYear;
      }
      
      return record;
    }).filter(record => record['USN'] && record['USN'] !== '');
    
    console.log(`Processed sheet '${sheetName}': ${records.length} records`);
    return records;
  };

  // EXACT replica of Python script's main processing logic
  const combineFiles = async () => {
    if (!files.year2 && !files.year3 && !files.year4) {
      setError('Please upload at least one file');
      return;
    }

    setProcessing(true);
    setError('');
    setSuccess('');

    try {
      const combinedRecords = [];
      const stats = {
        totalFiles: 0,
        totalSheets: 0,
        totalRecords: 0,
        branchesProcessed: new Set(),
        batchesProcessed: new Set()
      };

      // EXACT replica: Process files in chronological order (Year 4, Year 3, Year 2)
      const fileOrder = [
        { key: 'year4', file: files.year4 },
        { key: 'year3', file: files.year3 },
        { key: 'year2', file: files.year2 }
      ];

      for (const { file } of fileOrder) {
        if (!file) continue;

        const { batch, year } = extractBatchYear(file.name);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(await file.arrayBuffer());
        
        console.log(`Processing file: ${file.name} (Batch: ${batch}, ${year})`);
        
        // Add batch separator
        combinedRecords.push({
          type: 'batch_separator',
          text: `${batch} Batch (${year})`
        });

        // EXACT replica: Filter and validate sheets
        const validSheets = [];
        workbook.eachSheet((worksheet, sheetId) => {
          const sheetName = worksheet.name;
          
          // EXACT replica: Skip utility sheets
          if ((sheetName.startsWith('Sheet') && sheetName !== 'Sheet1') || 
              ['template', 'format', 'example', 'blank'].includes(sheetName.toLowerCase())) {
            console.log(`Skipping utility sheet: ${sheetName}`);
            return;
          }
          
          // Convert worksheet to data for validation with direct hyperlink access
          const data = [];
          worksheet.eachRow((row, rowNumber) => {
            const rowData = [];
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
              let cellValue = '';
              
              // DIRECT HYPERLINK CHECK first
              if (cell.hyperlink) {
                if (cell.value && typeof cell.value === 'string') {
                  cellValue = cell.value;
                } else if (cell.text) {
                  cellValue = cell.text;
                } else if (typeof cell.hyperlink === 'string') {
                  cellValue = cell.hyperlink.replace('mailto:', '');
                } else if (cell.hyperlink.text) {
                  cellValue = cell.hyperlink.text;
                } else {
                  cellValue = String(cell.hyperlink).replace('mailto:', '');
                }
              } else {
                cellValue = extractCellValue(cell.value);
              }
              
              rowData[colNumber - 1] = cellValue || '';
            });
            data.push(rowData);
          });
          
          if (isValidSheet(data, sheetName)) {
            validSheets.push({ name: sheetName, worksheet });
          } else {
            console.log(`Skipping invalid/empty sheet: ${sheetName}`);
          }
        });

        // EXACT replica: Sort sheets alphabetically by normalized branch names
        validSheets.sort((a, b) => {
          const normalizedA = normalizeBranchName(a.name).toLowerCase();
          const normalizedB = normalizeBranchName(b.name).toLowerCase();
          return normalizedA.localeCompare(normalizedB);
        });

        let fileRecords = 0;
        for (const { name: sheetName, worksheet } of validSheets) {
          try {
            // Add branch separator
            const normalizedBranch = normalizeBranchName(sheetName);
            combinedRecords.push({
              type: 'branch_separator',
              text: normalizedBranch
            });

            // Process the sheet
            const sheetData = processSheet(worksheet, sheetName, batch, file.name);
            if (sheetData.length > 0) {
              combinedRecords.push(...sheetData);
              fileRecords += sheetData.length;
              
              // Update statistics
              stats.branchesProcessed.add(sheetName);
              stats.totalRecords += sheetData.length;
              stats.totalSheets += 1;
            }
          } catch (err) {
            console.error(`Error processing sheet '${sheetName}' in ${file.name}: ${err.message}`);
            continue;
          }
        }

        stats.batchesProcessed.add(batch);
        stats.totalFiles += 1;
        console.log(`Completed file ${file.name}: ${fileRecords} total records`);
      }

      setCombinedData(combinedRecords);
      setStatistics({
        ...stats,
        branchesProcessed: Array.from(stats.branchesProcessed),
        batchesProcessed: Array.from(stats.batchesProcessed)
      });

      setSuccess('Files processed successfully! Ready to download.');

    } catch (err) {
      setError(`Error processing files: ${err.message}`);
    } finally {
      setProcessing(false);
    }
  };

  // EXACT replica of Python script's Excel formatting
  const downloadExcel = async () => {
    if (!combinedData) return;

    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Combined Data');

      // EXACT replica: Professional column headers
      const headers = [
        'USN', 'Full Name', 'Student Branch', 'Section', 'Student Email ID', 
        'Phone Number', 'Counsellor Name', 'Counsellor Email ID', 
        'Counsellor Department', 'Student Batch'
      ];

      let currentRow = 1;

      // EXACT replica: Row 1 - Column headers with professional styling
      worksheet.addRow(headers);
      const headerRow = worksheet.getRow(currentRow);
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
      currentRow++;

      // EXACT replica: Row 2 - Empty row for spacing
      worksheet.mergeCells(`A${currentRow}:J${currentRow}`);
      currentRow++;

      // EXACT replica: Rows 3-8 - Logo space with original size and centered positioning
      worksheet.mergeCells(`A${currentRow}:J${currentRow + 5}`);
      const logoCell = worksheet.getCell(`A${currentRow}`);
      
      // Fetch and add RVCE logo with original size - SAFE METHOD
      try {
        const logoResponse = await fetch('https://csitss.ieee-rvce.org/Logo3.png');
        
        if (logoResponse.ok) {
          const logoBuffer = await logoResponse.arrayBuffer();
          
          if (logoBuffer && logoBuffer.byteLength > 0) {
            const logoId = workbook.addImage({
              buffer: logoBuffer,
              extension: 'png',
            });

            // Insert logo at the extreme far left
            worksheet.addImage(logoId, {
              tl: { col: 0.5, row: currentRow - 0.5 }, // Extreme left (column A start)
              ext: { width: 250, height: 100 }, // Keep the size you liked
              editAs: 'oneCell'
            });
          } else {
            throw new Error('Empty logo buffer');
          }
        } else {
          throw new Error('Failed to fetch logo');
        }
        
        // Set the merged cell background to complement the logo
        logoCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFF' }
        };
        logoCell.alignment = { horizontal: 'center', vertical: 'middle' };
        
      } catch (logoError) {
        console.warn('Logo insertion failed:', logoError.message);
        // Fallback to text if logo fails
        logoCell.value = 'RVCE LOGO SPACE';
        logoCell.font = { size: 16, color: { argb: '2563eb' }, bold: true };
        logoCell.alignment = { horizontal: 'center', vertical: 'middle' };
        logoCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'F8F9FA' }
        };
      }
      
      currentRow += 6; // Skip logo space (rows 3-8)

      // EXACT replica: Add data with formatting
      combinedData.forEach(item => {
        if (item.type === 'batch_separator') {
          // EXACT replica: Batch separator with merged cells
          worksheet.mergeCells(`A${currentRow}:J${currentRow}`);
          const batchCell = worksheet.getCell(`A${currentRow}`);
          batchCell.value = item.text;
          batchCell.font = { bold: true, color: { argb: 'FFFFFF' }, size: 16 };
          batchCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '0B5394' }
          };
          batchCell.alignment = { horizontal: 'center', vertical: 'center' };
          batchCell.border = {
            top: { style: 'thick', color: { argb: 'FFFFFF' } },
            left: { style: 'thick', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thick', color: { argb: 'FFFFFF' } },
            right: { style: 'thick', color: { argb: 'FFFFFF' } }
          };
          worksheet.getRow(currentRow).height = 30;
          currentRow++;
          currentRow++; // Empty row after batch separator
        } else if (item.type === 'branch_separator') {
          // EXACT replica: Branch separator with merged cells
          worksheet.mergeCells(`A${currentRow}:J${currentRow}`);
          const branchCell = worksheet.getCell(`A${currentRow}`);
          branchCell.value = item.text;
          branchCell.font = { bold: true, color: { argb: 'FFFFFF' }, size: 14 };
          branchCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '6AA84F' }
          };
          branchCell.alignment = { horizontal: 'center', vertical: 'center' };
          branchCell.border = {
            top: { style: 'medium', color: { argb: 'FFFFFF' } },
            left: { style: 'medium', color: { argb: 'FFFFFF' } },
            bottom: { style: 'medium', color: { argb: 'FFFFFF' } },
            right: { style: 'medium', color: { argb: 'FFFFFF' } }
          };
          worksheet.getRow(currentRow).height = 25;
          currentRow++;
          currentRow++; // Empty row after branch separator
        } else {
          // EXACT replica: Student data with clean formatting
          const rowData = [
            cleanValue(item['USN']),
            cleanValue(item['Full Name']),
            cleanValue(item['Normalized Branch']),  // Student Branch
            cleanValue(item['Section']),
            cleanValue(item['Email']),  // Student Email ID
            cleanValue(item['Phone Number']),
            cleanValue(item['Counsellor']),  // Counsellor Name
            cleanValue(item['Counsellor Email']),  // Counsellor Email ID
            cleanValue(item['Counsellor Department']),
            cleanValue(item['Batch'])  // Student Batch
          ];
          
          const dataRow = worksheet.addRow(rowData);
          
          // EXACT replica: Apply formatting
          dataRow.eachCell((cell, colNumber) => {
            // Alternating row colors
            if (currentRow % 2 === 0) {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'F8F9FA' }
              };
            }
            
            // Borders
            cell.border = {
              top: { style: 'thin', color: { argb: 'E1E1E1' } },
              left: { style: 'thin', color: { argb: 'E1E1E1' } },
              bottom: { style: 'thin', color: { argb: 'E1E1E1' } },
              right: { style: 'thin', color: { argb: 'E1E1E1' } }
            };
            
            // EXACT replica: Alignment
            if (colNumber === 1) { // USN column - center align
              cell.alignment = { horizontal: 'center', vertical: 'center' };
            } else if (colNumber === 5 || colNumber === 8) { // Email columns - left align
              cell.alignment = { horizontal: 'left', vertical: 'center' };
            } else {
              cell.alignment = { horizontal: 'left', vertical: 'center' };
            }
          });
          
          currentRow++;
        }
      });

      // EXACT replica: Set column widths
      const columnWidths = [15, 30, 35, 10, 40, 15, 25, 40, 20, 15];
      columnWidths.forEach((width, index) => {
        const colLetter = String.fromCharCode(65 + index); // A, B, C, etc.
        worksheet.getColumn(colLetter).width = width;
      });

      // EXACT replica: Freeze header row and set height
      worksheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];
      worksheet.getRow(1).height = 40;

      // Download the file
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'Combined Student-Counsellor Info (Year 2,3,4).xlsx';
      link.click();
      window.URL.revokeObjectURL(url);

      setSuccess('Excel file downloaded successfully.');
    } catch (err) {
      setError('Error creating Excel file: ' + err.message);
    }
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

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col" style={{ fontFamily: '-apple-system, BlinkMacSystemFont, "SF Pro Display", "Helvetica Neue", Helvetica, Arial, sans-serif' }}>
      {/* Header with RVCE Logo - Fully Responsive */}
      <div className="bg-white border-b border-gray-100 shadow-sm">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 py-4 sm:py-6">
          <div className="flex flex-col sm:flex-row items-center justify-center space-y-4 sm:space-y-0 sm:space-x-6">
            <a 
              href="https://rvce.edu.in/" 
              target="_blank" 
              rel="noopener noreferrer"
              className="transition-transform duration-200 hover:scale-105"
            >
              <img 
                src="https://csitss.ieee-rvce.org/Logo3.png" 
                alt="RVCE Logo" 
                className="h-12 sm:h-16 w-auto"
              />
            </a>
            <div className="text-center sm:text-left">
              <h1 className="text-xl sm:text-2xl font-semibold text-gray-900">
                Counsellor Data Combiner
              </h1>
              <p className="text-gray-500 mt-1 text-sm sm:text-base">
                R.V. College of Engineering
              </p>
            </div>
          </div>
        </div>
      </div>

      {/* Main Content - Responsive */}
      <div className="flex-1 max-w-6xl mx-auto px-4 sm:px-6 py-6 sm:py-12 w-full">
        
        {/* File Upload Section - Mobile Optimized */}
        <div className="bg-white rounded-xl sm:rounded-2xl shadow-sm border border-gray-100 p-4 sm:p-8 mb-6 sm:mb-8">
          <h2 className="text-lg sm:text-xl font-semibold text-gray-900 mb-6 sm:mb-8 text-center">
            Upload Excel Files
          </h2>
          
          {/* Mobile: Stack cards vertically, Desktop: 3-column grid */}
          <div className="flex flex-col sm:grid sm:grid-cols-3 gap-4 sm:gap-6">
            {[
              { key: 'year4', label: 'Year 4', sublabel: '2022-2026', color: 'from-red-500 to-pink-500' },
              { key: 'year3', label: 'Year 3', sublabel: '2023-2027', color: 'from-yellow-500 to-orange-500' },
              { key: 'year2', label: 'Year 2', sublabel: '2024-2028', color: 'from-green-500 to-emerald-500' }
            ].map(({ key, label, sublabel, color }) => (
              <div key={key} className="group w-full">
                <div className="relative">
                  <input
                    type="file"
                    accept=".xlsx"
                    onChange={(e) => handleFileUpload(key, e.target.files[0])}
                    className="hidden"
                    id={`file-${key}`}
                  />
                  <label
                    htmlFor={`file-${key}`}
                    className="block cursor-pointer w-full"
                  >
                    <div className="bg-white border border-gray-200 rounded-xl sm:rounded-2xl p-4 sm:p-6 hover:border-gray-300 transition-all duration-200 hover:shadow-md group-hover:scale-105 min-h-[120px] sm:min-h-[140px] flex flex-col justify-center">
                      <div className="text-center">
                        <div className={`w-12 h-12 sm:w-16 sm:h-16 mx-auto mb-3 sm:mb-4 rounded-xl sm:rounded-2xl bg-gradient-to-r ${color} flex items-center justify-center`}>
                          {files[key] ? (
                            <CheckCircle className="w-6 h-6 sm:w-8 sm:h-8 text-white" />
                          ) : (
                            <FileText className="w-6 h-6 sm:w-8 sm:h-8 text-white" />
                          )}
                        </div>
                        <h3 className="font-semibold text-gray-900 mb-1 text-sm sm:text-base">{label}</h3>
                        <p className="text-xs sm:text-sm text-gray-500 mb-3 sm:mb-4">{sublabel}</p>
                        
                        {files[key] ? (
                          <div className="space-y-2 sm:space-y-3">
                            <div 
                              className="flex items-center justify-center text-green-600 mb-2"
                              title={files[key].name} // Tooltip with filename on hover
                            >
                              <CheckCircle className="w-5 h-5 sm:w-6 sm:h-6" />
                            </div>
                            <label
                              htmlFor={`file-${key}`}
                              className="block bg-blue-50 hover:bg-blue-100 text-blue-600 px-3 py-2 sm:px-4 sm:py-2 rounded-lg text-xs sm:text-sm font-medium transition-all duration-200 cursor-pointer text-center"
                            >
                              Change File
                            </label>
                          </div>
                        ) : (
                          <div className="text-xs sm:text-sm text-blue-600 font-medium">
                            Choose File
                          </div>
                        )}
                      </div>
                    </div>
                  </label>
                </div>
              </div>
            ))}
          </div>

          {/* Action Buttons - Responsive */}
          <div className="flex flex-col sm:flex-row justify-center mt-6 sm:mt-8 space-y-3 sm:space-y-0 sm:space-x-4">
            <button
              onClick={combineFiles}
              disabled={processing || (!files.year2 && !files.year3 && !files.year4)}
              className="w-full sm:w-auto bg-blue-600 hover:bg-blue-700 disabled:bg-gray-300 text-white px-6 sm:px-8 py-3 rounded-full font-semibold transition-all duration-200 hover:scale-105 disabled:hover:scale-100 shadow-sm hover:shadow-md flex items-center justify-center text-sm sm:text-base"
            >
              {processing ? (
                <>
                  <Loader className="w-4 h-4 sm:w-5 sm:h-5 mr-2 animate-spin" />
                  Processing...
                </>
              ) : (
                'Combine Data'
              )}
            </button>

            {combinedData && (
              <button
                onClick={downloadExcel}
                className="w-full sm:w-auto bg-green-600 hover:bg-green-700 text-white px-6 sm:px-8 py-3 rounded-full font-semibold transition-all duration-200 hover:scale-105 shadow-sm hover:shadow-md flex items-center justify-center text-sm sm:text-base"
              >
                <Download className="w-4 h-4 sm:w-5 sm:h-5 mr-2" />
                Download Excel
              </button>
            )}
          </div>
        </div>

        {/* Statistics - Mobile Optimized */}
        {statistics && (
          <div className="bg-white rounded-xl sm:rounded-2xl shadow-sm border border-gray-100 p-4 sm:p-8 mb-6 sm:mb-8">
            <h3 className="text-base sm:text-lg font-semibold text-gray-900 mb-4 sm:mb-6 text-center">
              Processing Summary
            </h3>
            {/* Mobile: 2x2 grid, Desktop: 4-column grid */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 sm:gap-6">
              <div className="text-center">
                <div className="text-2xl sm:text-3xl font-bold text-blue-600 mb-1">{statistics.totalFiles}</div>
                <div className="text-xs sm:text-sm text-gray-500">Files</div>
              </div>
              <div className="text-center">
                <div className="text-2xl sm:text-3xl font-bold text-green-600 mb-1">{statistics.totalSheets}</div>
                <div className="text-xs sm:text-sm text-gray-500">Sheets</div>
              </div>
              <div className="text-center">
                <div className="text-2xl sm:text-3xl font-bold text-purple-600 mb-1">{statistics.totalRecords}</div>
                <div className="text-xs sm:text-sm text-gray-500">Records</div>
              </div>
              <div className="text-center">
                <div className="text-2xl sm:text-3xl font-bold text-orange-600 mb-1">{statistics.branchesProcessed.length}</div>
                <div className="text-xs sm:text-sm text-gray-500">Branches</div>
              </div>
            </div>
          </div>
        )}

        {/* Messages - Mobile Optimized */}
        {error && (
          <div className="bg-red-50 border border-red-200 rounded-xl sm:rounded-2xl p-3 sm:p-4 mb-4 sm:mb-6 flex items-start sm:items-center">
            <AlertCircle className="w-4 h-4 sm:w-5 sm:h-5 text-red-500 mr-2 sm:mr-3 flex-shrink-0 mt-0.5 sm:mt-0" />
            <span className="text-red-700 text-sm sm:text-base break-words">{error}</span>
          </div>
        )}

        {success && (
          <div className="bg-green-50 border border-green-200 rounded-xl sm:rounded-2xl p-3 sm:p-4 mb-4 sm:mb-6 flex items-start sm:items-center">
            <CheckCircle className="w-4 h-4 sm:w-5 sm:h-5 text-green-500 mr-2 sm:mr-3 flex-shrink-0 mt-0.5 sm:mt-0" />
            <span className="text-green-700 text-sm sm:text-base break-words">{success}</span>
          </div>
        )}
      </div>

      {/* Footer with Coding Club Logo - Mobile Optimized */}
      <div className="bg-white border-t border-gray-100 mt-6 sm:mt-12">
        <div className="max-w-4xl mx-auto px-4 sm:px-6 py-6 sm:py-8">
          <div className="flex flex-col items-center space-y-3 sm:space-y-4">
            <a 
              href="https://www.linkedin.com/company/coding-club-rvce" 
              target="_blank" 
              rel="noopener noreferrer"
              className="transition-transform duration-200 hover:scale-105"
            >
              <img 
                src="https://avatars.githubusercontent.com/u/54234255?v=4" 
                alt="Coding Club Logo" 
                className="h-16 w-16 sm:h-24 sm:w-24 rounded-xl sm:rounded-2xl shadow-lg"
              />
            </a>
            <div className="text-center">
              <p className="text-base sm:text-lg font-semibold text-gray-900">Coding Club RVCE</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default App;