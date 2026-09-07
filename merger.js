/**
 * merger.js - 엑셀 데이터 병합, 중복 제거, 엑셀/CSV 및 ZIP 내보내기 엔진
 */

const ExcelMerger = (() => {

  /**
   * Merges all files in a schema group according to the current settings
   * @param {Object} group - schema group object from ExcelParser
   * @param {Object} options - merge options
   * @returns {Object} { headers: Array<string>, rows: Array<Object>, deduplicatedCount: number }
   */
  function mergeGroupData(group, options = {}) {
    const finalHeaders = [];
    
    // Add Source Tracking Columns if requested
    if (options.addFileName) {
      finalHeaders.push('_원본파일명');
    }
    if (options.addSheetName) {
      finalHeaders.push('_시트명');
    }

    // Append master headers
    finalHeaders.push(...group.masterHeaders);

    const mergedRows = [];
    const seenRowSignatures = new Set();
    let deduplicatedCount = 0;

    group.files.forEach(fileItem => {
      fileItem.dataRows.forEach(rawRow => {
        const rowObj = {};

        // Fill source tracking
        if (options.addFileName) {
          rowObj['_원본파일명'] = fileItem.fileName;
        }
        if (options.addSheetName) {
          rowObj['_시트명'] = fileItem.sheetName;
        }

        // Fill canonical headers (mapping keys accurately even if order differs)
        group.masterHeaders.forEach(header => {
          let val = rawRow[header];
          // If not found directly, check case/whitespace insensitive match
          if (val === undefined || val === null) {
            const matchedKey = Object.keys(rawRow).find(k => k.trim() === header.trim());
            val = matchedKey ? rawRow[matchedKey] : '';
          }
          rowObj[header] = val !== undefined && val !== null ? val : '';
        });

        // Deduplication Check
        if (options.removeDuplicates) {
          // Compute signature from business data columns (excluding source filenames)
          const dataSignature = group.masterHeaders.map(h => String(rowObj[h] || '')).join('|||');
          if (seenRowSignatures.has(dataSignature)) {
            deduplicatedCount++;
            return; // Skip duplicate row
          }
          seenRowSignatures.add(dataSignature);
        }

        mergedRows.push(rowObj);
      });
    });

    return {
      headers: finalHeaders,
      rows: mergedRows,
      totalRows: mergedRows.length,
      deduplicatedCount: deduplicatedCount
    };
  }

  /**
   * Generates a SheetJS worksheet from headers and rows
   */
  function createWorksheet(headers, rows) {
    // Array of arrays format for maximum fidelity
    const dataAOA = [];
    
    // Row 1: Headers
    dataAOA.push(headers);

    // Following Rows: Data
    rows.forEach(rowObj => {
      const rowArr = headers.map(h => rowObj[h] !== undefined && rowObj[h] !== null ? rowObj[h] : '');
      dataAOA.push(rowArr);
    });

    const ws = XLSX.utils.aoa_to_sheet(dataAOA);

    // Auto calculate column widths
    const colWidths = headers.map((h, colIdx) => {
      let maxLen = String(h).length;
      // Sample first 50 rows for width
      const sampleLimit = Math.min(rows.length, 50);
      for (let r = 0; r < sampleLimit; r++) {
        const cellVal = String(rows[r][h] || '');
        maxLen = Math.max(maxLen, cellVal.length);
      }
      return { wch: Math.min(Math.max(maxLen + 4, 10), 50) };
    });
    ws['!cols'] = colWidths;

    return ws;
  }

  /**
   * Exports a single group to XLSX file
   */
  function exportGroupToXLSX(group, options = {}, customFileName = null) {
    const merged = mergeGroupData(group, options);
    const ws = createWorksheet(merged.headers, merged.rows);
    const wb = XLSX.utils.book_new();

    const sheetTitle = sanitizeSheetName(group.groupName.split('(')[0].trim() || '병합데이터');
    XLSX.utils.book_append_sheet(wb, ws, sheetTitle);

    const fileName = customFileName || `${cleanFileName(group.groupName)}_병합결과.xlsx`;
    XLSX.writeFile(wb, fileName);
  }

  /**
   * Exports a single group to CSV file (with UTF-8 BOM for Korean Excel compatibility)
   */
  function exportGroupToCSV(group, options = {}, customFileName = null) {
    const merged = mergeGroupData(group, options);
    const ws = createWorksheet(merged.headers, merged.rows);
    const csvContent = XLSX.utils.sheet_to_csv(ws);

    // Add UTF-8 BOM (\uFEFF) so Excel opens Korean characters correctly
    const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const fileName = customFileName || `${cleanFileName(group.groupName)}_병합결과.csv`;
    triggerBlobDownload(blob, fileName);
  }

  /**
   * Exports all schema groups into a single Multi-Sheet Excel Workbook
   */
  function exportAllToMultiSheetWorkbook(groups, options = {}, customFileName = '전체_규격별_통합병합.xlsx') {
    const wb = XLSX.utils.book_new();
    const usedSheetNames = new Set();

    groups.forEach((group, index) => {
      const merged = mergeGroupData(group, options);
      const ws = createWorksheet(merged.headers, merged.rows);

      // Create unique, sanitized sheet name (max 31 chars for Excel)
      let rawName = `규격${index + 1}_${group.masterHeaders.slice(0, 2).join('_')}`;
      let sheetName = sanitizeSheetName(rawName).substring(0, 28);
      if (!sheetName) sheetName = `규격그룹_${index + 1}`;

      if (usedSheetNames.has(sheetName)) {
        sheetName = `${sheetName.substring(0, 25)}_${index + 1}`;
      }
      usedSheetNames.add(sheetName);

      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    XLSX.writeFile(wb, customFileName);
  }

  /**
   * Exports all groups as separate XLSX files bundled inside a single ZIP archive
   */
  async function exportAllToZip(groups, options = {}, zipFileName = '규격별_엑셀_병합_일괄.zip') {
    if (typeof JSZip === 'undefined') {
      throw new Error('JSZip 라이브러리가 로드되지 않았습니다.');
    }

    const zip = new JSZip();

    groups.forEach((group, index) => {
      const merged = mergeGroupData(group, options);
      const ws = createWorksheet(merged.headers, merged.rows);
      const wb = XLSX.utils.book_new();
      
      const sheetTitle = sanitizeSheetName(`규격_${index + 1}`);
      XLSX.utils.book_append_sheet(wb, ws, sheetTitle);

      // Convert workbook to binary array
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      
      const singleFileName = `규격그룹_${index + 1}_${cleanFileName(group.masterHeaders.slice(0, 3).join('_'))}.xlsx`;
      zip.file(singleFileName, wbout);
    });

    const contentBlob = await zip.generateAsync({ type: 'blob' });
    triggerBlobDownload(contentBlob, zipFileName);
  }

  /**
   * Helper: Triggers browser file download from Blob
   */
  function triggerBlobDownload(blob, fileName) {
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    setTimeout(() => {
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }, 100);
  }

  /**
   * Helper: Sanitizes sheet name for Excel rules (max 31 chars, no brackets, slash, question mark, etc.)
   */
  function sanitizeSheetName(name) {
    return name.replace(/[\[\]\*\/\\\?\:]/g, '_').trim().substring(0, 31) || 'Sheet';
  }

  /**
   * Helper: Cleans file name for filesystem safety
   */
  function cleanFileName(name) {
    return name.replace(/[\/\\:\*\?"<>\|]/g, '_').replace(/\s+/g, '_').trim();
  }

  return {
    mergeGroupData,
    createWorksheet,
    exportGroupToXLSX,
    exportGroupToCSV,
    exportAllToMultiSheetWorkbook,
    exportAllToZip
  };

})();
