/**
 * parser.js - 스마트 엑셀/CSV 파서 및 스키마 분석 엔진
 * Handles multi-file ingestion, header extraction, and automatic schema clustering.
 */

const ExcelParser = (() => {

  /**
   * Reads a File object using SheetJS
   * @param {File} file 
   * @param {Object} options - parsing settings
   * @returns {Promise<Array<Object>>} list of parsed sheet data objects
   */
  async function parseFile(file, options = {}) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, {
            type: 'array',
            cellDates: true,
            cellText: false,
            raw: true
          });

          const parsedSheets = [];
          const sheetNames = options.sheetMode === 'all' ? workbook.SheetNames : [workbook.SheetNames[0]];

          sheetNames.forEach((sheetName) => {
            const worksheet = workbook.Sheets[sheetName];
            if (!worksheet) return;

            // Convert worksheet to 2D array (array of arrays)
            const rawRows = XLSX.utils.sheet_to_json(worksheet, {
              header: 1,
              defval: '',
              blankrows: false
            });

            if (!rawRows || rawRows.length === 0) {
              return; // Skip completely empty sheet
            }

            // Determine Header Row
            const { headerRowIndex, headers } = extractHeaders(rawRows, options);
            if (headers.length === 0) return;

            // Extract Data Rows
            const dataRows = [];
            for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
              const row = rawRows[i];
              if (!row || row.length === 0) continue;

              // Check if row is completely blank
              if (options.skipEmptyRows) {
                const isAllBlank = row.every(cell => cell === null || cell === undefined || String(cell).trim() === '');
                if (isAllBlank) continue;
              }

              // Map row cells to headers
              const rowObj = {};
              headers.forEach((header, colIdx) => {
                let cellVal = row[colIdx];
                if (cellVal === undefined || cellVal === null) cellVal = '';
                // If date object, format cleanly
                if (cellVal instanceof Date) {
                  cellVal = formatDate(cellVal);
                }
                rowObj[header] = cellVal;
              });

              dataRows.push(rowObj);
            }

            parsedSheets.push({
              fileId: generateUniqueId(),
              fileName: file.name,
              fileSize: file.size,
              lastModified: file.lastModified,
              sheetName: sheetName,
              headers: headers,
              dataRows: dataRows,
              rowCount: dataRows.length,
              columnCount: headers.length
            });
          });

          resolve(parsedSheets);
        } catch (err) {
          console.error("Error reading file:", file.name, err);
          reject(new Error(`파일 "${file.name}" 파싱 중 오류가 발생했습니다: ${err.message}`));
        }
      };

      reader.onerror = () => {
        reject(new Error(`파일 "${file.name}"을(를) 읽을 수 없습니다.`));
      };

      reader.readAsArrayBuffer(file);
    });
  }

  /**
   * Extracts headers based on user options
   */
  function extractHeaders(rawRows, options) {
    let headerRowIndex = 0;

    if (options.headerRow === 'auto') {
      // Find the first row that has at least 1 non-empty cell
      for (let i = 0; i < rawRows.length; i++) {
        if (rawRows[i] && rawRows[i].some(cell => cell !== null && cell !== undefined && String(cell).trim() !== '')) {
          headerRowIndex = i;
          break;
        }
      }
    } else {
      const requestedIndex = parseInt(options.headerRow, 10) - 1;
      headerRowIndex = Math.max(0, Math.min(requestedIndex, rawRows.length - 1));
    }

    const rawHeaderRow = rawRows[headerRowIndex] || [];
    const headers = [];
    const seenHeaders = new Map();

    rawHeaderRow.forEach((cell, idx) => {
      let headerName = cell !== null && cell !== undefined ? String(cell) : '';
      if (options.trimHeaders) {
        headerName = headerName.trim();
      }

      if (!headerName) {
        headerName = `열_${idx + 1}`;
      }

      // Handle duplicate header names in the same file
      if (seenHeaders.has(headerName)) {
        const count = seenHeaders.get(headerName) + 1;
        seenHeaders.set(headerName, count);
        headerName = `${headerName}_${count}`;
      } else {
        seenHeaders.set(headerName, 1);
      }

      headers.push(headerName);
    });

    return { headerRowIndex, headers };
  }

  /**
   * Generates a unique schema signature for clustering
   */
  function computeSchemaSignature(headers, options) {
    let normalizedHeaders = headers.map(h => {
      let val = String(h);
      if (options.trimHeaders) val = val.trim();
      return val;
    });

    if (options.matchMode === 'flexible') {
      // Order-independent match: sort headers
      normalizedHeaders = [...normalizedHeaders].sort((a, b) => a.localeCompare(b));
    }

    return normalizedHeaders.join('###');
  }

  /**
   * Clusters parsed sheet items by schema signature
   * @param {Array<Object>} parsedItems - list of parsed sheet objects
   * @param {Object} options - clustering options
   * @returns {Array<Object>} schema groups
   */
  function clusterBySchema(parsedItems, options = {}) {
    const groupsMap = new Map();

    parsedItems.forEach(item => {
      const signature = computeSchemaSignature(item.headers, options);

      if (!groupsMap.has(signature)) {
        groupsMap.set(signature, {
          groupId: generateUniqueId(),
          groupName: '', // Will be assigned later
          signature: signature,
          masterHeaders: [...item.headers], // Canonical header list
          files: [],
          totalRowCount: 0,
          uniqueRowCount: 0
        });
      }

      const group = groupsMap.get(signature);
      group.files.push(item);
      group.totalRowCount += item.rowCount;
    });

    // Convert map to sorted array of groups (largest groups first)
    const groups = Array.from(groupsMap.values()).sort((a, b) => b.files.length - a.files.length);

    // Assign friendly names & order canonical headers
    groups.forEach((group, index) => {
      const mainHeadersPreview = group.masterHeaders.slice(0, 3).join(', ');
      const extraCount = group.masterHeaders.length > 3 ? ` 외 ${group.masterHeaders.length - 3}개` : '';
      group.groupName = `규격 그룹 ${index + 1} (${group.masterHeaders.length}개 열: [${mainHeadersPreview}${extraCount}])`;
    });

    return groups;
  }

  /**
   * Helper: Date formatting
   */
  function formatDate(d) {
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  /**
   * Helper: Random unique ID
   */
  function generateUniqueId() {
    return 'id_' + Math.random().toString(36).substring(2, 9) + Date.now().toString(36);
  }

  return {
    parseFile,
    computeSchemaSignature,
    clusterBySchema,
    generateUniqueId
  };

})();
