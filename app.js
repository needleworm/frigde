/**
 * app.js - 스마트 엑셀 규격별 자동 병합기 메인 컨트롤러 (컴팩트 고밀도 UI)
 */

document.addEventListener('DOMContentLoaded', () => {
  // Initialize Lucide Icons
  if (window.lucide) {
    lucide.createIcons();
  }

  // State
  const state = {
    rawFiles: [],        // List of uploaded File objects
    parsedItems: [],     // All parsed sheets/files from ExcelParser
    schemaGroups: [],    // Clustered groups
    options: {
      matchMode: 'flexible',
      headerRow: '1',
      sheetMode: 'first',
      addFileName: true,
      addSheetName: false,
      removeDuplicates: false,
      skipEmptyRows: true,
      trimHeaders: true
    },
    modalGroup: null,    // Currently viewed group in modal
    modalFilterText: '',
    theme: localStorage.getItem('excel_merger_theme') || 'dark'
  };

  // DOM Elements
  const elements = {
    // Theme & Header
    themeToggle: document.getElementById('btn-theme-toggle'),
    resetAllBtn: document.getElementById('btn-reset-all'),
    sampleDemoBtn: document.getElementById('btn-sample-demo'),

    // Upload & Dropzone
    dropZone: document.getElementById('drop-zone'),
    fileInput: document.getElementById('file-input'),
    fileInputAppend: document.getElementById('file-input-append'),
    parsingOverlay: document.getElementById('parsing-overlay'),
    parsingTitle: document.getElementById('parsing-status-title'),
    parsingDesc: document.getElementById('parsing-status-desc'),
    parsingBar: document.getElementById('parsing-progress-bar'),

    // Settings
    settingsToggleTrigger: document.getElementById('settings-toggle-trigger'),
    settingsToggleText: document.getElementById('settings-toggle-text'),
    settingsChevron: document.getElementById('settings-chevron'),
    settingsContent: document.getElementById('settings-content'),
    optMatchMode: document.getElementById('opt-match-mode'),
    optHeaderRow: document.getElementById('opt-header-row'),
    optSheetMode: document.getElementById('opt-sheet-mode'),
    optAddFileName: document.getElementById('opt-add-filename'),
    optAddSheetName: document.getElementById('opt-add-sheetname'),
    optRemoveDuplicates: document.getElementById('opt-remove-duplicates'),
    optSkipEmptyRows: document.getElementById('opt-skip-empty-rows'),
    optTrimHeaders: document.getElementById('opt-trim-headers'),

    // Sections
    uploadedFilesSection: document.getElementById('uploaded-files-section'),
    uploadedFilesCount: document.getElementById('uploaded-files-count'),
    fileCardsList: document.getElementById('file-cards-list'),
    btnClearFiles: document.getElementById('btn-clear-files'),

    resultsSection: document.getElementById('results-section'),
    statTotalFiles: document.getElementById('stat-total-files'),
    statTotalGroups: document.getElementById('stat-total-groups'),
    statTotalRows: document.getElementById('stat-total-rows'),
    statDedupRows: document.getElementById('stat-dedup-rows'),
    statDedupWrap: document.getElementById('stat-dedup-wrap'),

    btnDownloadAllSheets: document.getElementById('btn-download-all-sheets'),
    btnDownloadAllZip: document.getElementById('btn-download-all-zip'),
    schemaGroupsList: document.getElementById('schema-groups-list'),

    // Modal
    previewModal: document.getElementById('preview-modal'),
    modalGroupTitle: document.getElementById('modal-group-title'),
    modalRowBadge: document.getElementById('modal-row-badge'),
    modalSearchInput: document.getElementById('modal-search-input'),
    modalTableContainer: document.getElementById('modal-table-container'),
    modalPreviewInfo: document.getElementById('modal-preview-info'),
    btnCloseModal: document.getElementById('btn-close-modal'),
    btnModalCloseFooter: document.getElementById('btn-modal-close-footer'),
    btnModalCopy: document.getElementById('btn-modal-copy'),
    btnModalDownloadXlsx: document.getElementById('btn-modal-download-xlsx'),

    // Toast
    toastContainer: document.getElementById('toast-container')
  };

  // Set Initial Theme
  document.documentElement.setAttribute('data-theme', state.theme);

  // Bind Events
  initEventListeners();

  /* ==========================================================================
     Event Listeners
     ========================================================================== */
  function initEventListeners() {
    // Theme Toggle
    elements.themeToggle.addEventListener('click', toggleTheme);

    // Reset All
    elements.resetAllBtn.addEventListener('click', resetAll);

    // Sample Demo Loader
    elements.sampleDemoBtn.addEventListener('click', loadSampleDemoData);

    // Drop Zone Drag & Drop
    ['dragenter', 'dragover'].forEach(eventName => {
      elements.dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        elements.dropZone.classList.add('drag-over');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      elements.dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        elements.dropZone.classList.remove('drag-over');
      });
    });

    elements.dropZone.addEventListener('drop', (e) => {
      const files = Array.from(e.dataTransfer.files || []);
      if (files.length > 0) {
        handleFilesSelected(files);
      }
    });

    // Clicking Dropzone triggers file picker
    elements.dropZone.addEventListener('click', (e) => {
      if (e.target.closest('.file-input-hidden')) return;
      elements.fileInput.click();
    });

    // File Input Select
    elements.fileInput.addEventListener('change', (e) => {
      const files = Array.from(e.target.files || []);
      if (files.length > 0) {
        handleFilesSelected(files);
      }
      e.target.value = '';
    });

    elements.fileInputAppend.addEventListener('change', (e) => {
      const files = Array.from(e.target.files || []);
      if (files.length > 0) {
        handleFilesSelected(files, true);
      }
      e.target.value = '';
    });

    // Clear All Files
    elements.btnClearFiles.addEventListener('click', resetAll);

    // Settings Toggle Accordion
    elements.settingsToggleTrigger.addEventListener('click', toggleSettingsAccordion);

    // Settings Change Listeners
    const settingInputs = [
      elements.optMatchMode,
      elements.optHeaderRow,
      elements.optSheetMode,
      elements.optAddFileName,
      elements.optAddSheetName,
      elements.optRemoveDuplicates,
      elements.optSkipEmptyRows,
      elements.optTrimHeaders
    ];

    settingInputs.forEach(input => {
      input.addEventListener('change', () => {
        readOptionsFromUI();
        if (state.rawFiles.length > 0) {
          // Re-parse and re-cluster
          reprocessAllFiles();
        }
      });
    });

    // Batch Downloads
    elements.btnDownloadAllSheets.addEventListener('click', () => {
      if (state.schemaGroups.length === 0) return;
      try {
        ExcelMerger.exportAllToMultiSheetWorkbook(state.schemaGroups, state.options);
        showToast('통합 엑셀 파일(시트별 분리)이 다운로드되었습니다.', 'success');
      } catch (err) {
        showToast(`다운로드 실패: ${err.message}`, 'error');
      }
    });

    elements.btnDownloadAllZip.addEventListener('click', async () => {
      if (state.schemaGroups.length === 0) return;
      try {
        showToast('ZIP 압축 파일을 생성하고 있습니다...', 'info');
        await ExcelMerger.exportAllToZip(state.schemaGroups, state.options);
        showToast('ZIP 압축 파일이 다운로드되었습니다.', 'success');
      } catch (err) {
        showToast(`ZIP 생성 실패: ${err.message}`, 'error');
      }
    });

    // Modal Events
    elements.btnCloseModal.addEventListener('click', closeModal);
    elements.btnModalCloseFooter.addEventListener('click', closeModal);
    elements.previewModal.addEventListener('click', (e) => {
      if (e.target === elements.previewModal) closeModal();
    });

    elements.modalSearchInput.addEventListener('input', (e) => {
      state.modalFilterText = e.target.value.trim().toLowerCase();
      renderModalTable();
    });

    elements.btnModalCopy.addEventListener('click', copyModalTableToClipboard);
    
    elements.btnModalDownloadXlsx.addEventListener('click', () => {
      if (state.modalGroup) {
        ExcelMerger.exportGroupToXLSX(state.modalGroup, state.options);
        showToast('엑셀 파일이 다운로드되었습니다.', 'success');
      }
    });
  }

  /* ==========================================================================
     File Processing Pipeline
     ========================================================================== */
  async function handleFilesSelected(newFiles, append = false) {
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const validFiles = newFiles.filter(file => {
      const ext = '.' + file.name.split('.').pop().toLowerCase();
      return validExtensions.includes(ext);
    });

    if (validFiles.length === 0) {
      showToast('지원되는 파일 형식(.xlsx, .xls, .csv)이 없습니다.', 'error');
      return;
    }

    if (append) {
      const existingNames = new Set(state.rawFiles.map(f => f.name));
      const filtered = validFiles.filter(f => !existingNames.has(f.name));
      if (filtered.length < validFiles.length) {
        showToast('중복 파일은 제외되었습니다.', 'info');
      }
      state.rawFiles.push(...filtered);
    } else {
      state.rawFiles = validFiles;
    }

    readOptionsFromUI();
    await reprocessAllFiles();
  }

  async function reprocessAllFiles() {
    if (state.rawFiles.length === 0) {
      resetAll();
      return;
    }

    showOverlay(true, '파일 분석 중...', `총 ${state.rawFiles.length}개 파일 판독 중`);
    
    const parsedSheets = [];
    const total = state.rawFiles.length;

    for (let i = 0; i < total; i++) {
      const file = state.rawFiles[i];
      updateProgress((i / total) * 90, `${file.name} (${i + 1}/${total})`);
      
      try {
        const sheets = await ExcelParser.parseFile(file, state.options);
        parsedSheets.push(...sheets);
      } catch (err) {
        console.error(err);
        showToast(`"${file.name}" 파일 처리 중 오류: ${err.message}`, 'error');
      }
    }

    updateProgress(95, '규격 자동 분류 중...');
    state.parsedItems = parsedSheets;

    // Cluster into schema groups
    state.schemaGroups = ExcelParser.clusterBySchema(parsedSheets, state.options);

    updateProgress(100, '완료');
    setTimeout(() => {
      showOverlay(false);
      renderUI();
      showToast(`총 ${state.parsedItems.length}개 파일이 ${state.schemaGroups.length}개 규격 그룹으로 분류되었습니다.`, 'success');
    }, 200);
  }

  function readOptionsFromUI() {
    state.options = {
      matchMode: elements.optMatchMode.value,
      headerRow: elements.optHeaderRow.value,
      sheetMode: elements.optSheetMode.value,
      addFileName: elements.optAddFileName.checked,
      addSheetName: elements.optAddSheetName.checked,
      removeDuplicates: elements.optRemoveDuplicates.checked,
      skipEmptyRows: elements.optSkipEmptyRows.checked,
      trimHeaders: elements.optTrimHeaders.checked
    };
  }

  /* ==========================================================================
     UI Rendering
     ========================================================================== */
  function renderUI() {
    const hasFiles = state.rawFiles.length > 0;
    elements.uploadedFilesSection.style.display = hasFiles ? 'block' : 'none';
    elements.resultsSection.style.display = hasFiles ? 'flex' : 'none';
    elements.resetAllBtn.style.display = hasFiles ? 'inline-flex' : 'none';

    if (!hasFiles) return;

    // Render Uploaded Files Count & Chips
    elements.uploadedFilesCount.textContent = state.rawFiles.length;
    renderUploadedFileCards();

    // Calculate Global Statistics
    let totalMergedRows = 0;
    let totalDedupRows = 0;

    state.schemaGroups.forEach(group => {
      const merged = ExcelMerger.mergeGroupData(group, state.options);
      totalMergedRows += merged.totalRows;
      totalDedupRows += merged.deduplicatedCount;
    });

    elements.statTotalFiles.textContent = `${state.parsedItems.length}개`;
    elements.statTotalGroups.textContent = `${state.schemaGroups.length}개 그룹`;
    elements.statTotalRows.textContent = `${totalMergedRows.toLocaleString()}행`;
    
    if (totalDedupRows > 0) {
      elements.statDedupWrap.style.display = 'inline-flex';
      elements.statDedupRows.textContent = `${totalDedupRows.toLocaleString()}행`;
    } else {
      elements.statDedupWrap.style.display = 'none';
    }

    // Render Schema Groups
    renderSchemaGroups();

    // Re-trigger Lucide icons
    if (window.lucide) {
      lucide.createIcons();
    }
  }

  function renderUploadedFileCards() {
    elements.fileCardsList.innerHTML = '';

    state.rawFiles.forEach((file, index) => {
      const chip = document.createElement('div');
      chip.className = 'file-card-mini';

      const ext = file.name.split('.').pop().toUpperCase();
      const sizeStr = formatFileSize(file.size);

      chip.innerHTML = `
        <span class="file-name" title="${file.name} (${ext}, ${sizeStr})">${file.name}</span>
        <span class="file-size">${sizeStr}</span>
        <button class="file-mini-del" title="삭제" data-index="${index}">
          <i data-lucide="x"></i>
        </button>
      `;

      chip.querySelector('.file-mini-del').addEventListener('click', (e) => {
        e.stopPropagation();
        deleteFileByIndex(index);
      });

      elements.fileCardsList.appendChild(chip);
    });
  }

  function deleteFileByIndex(index) {
    state.rawFiles.splice(index, 1);
    if (state.rawFiles.length === 0) {
      resetAll();
    } else {
      reprocessAllFiles();
    }
  }

  function renderSchemaGroups() {
    elements.schemaGroupsList.innerHTML = '';

    if (state.schemaGroups.length === 0) {
      elements.schemaGroupsList.innerHTML = `
        <div class="group-card" style="padding: 1.5rem; text-align: center; color: var(--text-muted);">
          <p>감지된 유효한 엑셀 데이터 규격이 없습니다.</p>
        </div>
      `;
      return;
    }

    state.schemaGroups.forEach((group, index) => {
      const merged = ExcelMerger.mergeGroupData(group, state.options);

      const groupCard = document.createElement('div');
      groupCard.className = 'group-card';

      // Header HTML - Compact
      const headerHTML = `
        <div class="group-header">
          <div class="group-title-left">
            <span class="group-number-badge">규격 #${index + 1}</span>
            <div class="group-title-texts">
              <h4>${group.groupName}</h4>
            </div>
            <div class="group-badges">
              <span class="badge badge-primary">${group.files.length}개 파일</span>
              <span class="badge badge-neutral">${group.masterHeaders.length}개 열</span>
              <span class="badge badge-neutral">병합 ${merged.totalRows.toLocaleString()}행</span>
              ${merged.deduplicatedCount > 0 ? `<span class="badge badge-privacy">중복제거 ${merged.deduplicatedCount}행</span>` : ''}
            </div>
          </div>
          <div class="group-actions">
            <button class="btn btn-xs btn-primary btn-download-group-xlsx" data-group-id="${group.groupId}" title="이 규격 데이터만 엑셀로 다운로드">
              <i data-lucide="download"></i>
              <span>엑셀(.xlsx)</span>
            </button>
            <button class="btn btn-xs btn-secondary btn-download-group-csv" data-group-id="${group.groupId}" title="CSV 파일로 다운로드">
              <span>CSV</span>
            </button>
            <button class="btn btn-xs btn-outline btn-open-modal" data-group-id="${group.groupId}" title="전체 데이터 검색 및 확인">
              <i data-lucide="maximize-2"></i>
              <span>미리보기</span>
            </button>
          </div>
        </div>
      `;

      // Columns Tag Bar HTML
      const tagsHTML = merged.headers.map(h => {
        const isSource = h.startsWith('_');
        return `<span class="column-tag ${isSource ? 'source-tag' : ''}">${h}</span>`;
      }).join('');

      const columnsBarHTML = `
        <div class="group-columns-bar">
          <span class="group-columns-label">열 목록:</span>
          <div class="columns-tag-list">${tagsHTML}</div>
        </div>
      `;

      // Included Files Mini List
      const filesChipsHTML = group.files.map(f => {
        return `<span class="file-chip">${f.fileName} <span class="file-chip-rows">(${f.rowCount.toLocaleString()}행)</span></span>`;
      }).join('');

      const filesBarHTML = `
        <div class="group-files-bar">
          <div class="group-files-chips">
            <strong style="color: var(--text-dim); margin-right: 0.2rem;">포함 파일:</strong>
            ${filesChipsHTML}
          </div>
        </div>
      `;

      // Preview Table (first 5 rows)
      const previewRows = merged.rows.slice(0, 5);
      const tableRowsHTML = previewRows.map(row => {
        const cellsHTML = merged.headers.map(h => {
          const val = row[h] !== undefined && row[h] !== null ? String(row[h]) : '';
          const isSource = h.startsWith('_');
          return `<td class="${isSource ? 'cell-source' : ''}" title="${val}">${escapeHtml(val)}</td>`;
        }).join('');
        return `<tr>${cellsHTML}</tr>`;
      }).join('');

      const previewBodyHTML = `
        <div class="group-preview-body">
          <div class="preview-table-header">
            <span class="preview-table-title">데이터 미리보기 (상위 ${previewRows.length}행)</span>
            ${merged.rows.length > 5 ? `<span style="font-size: 0.72rem; color: var(--text-dim);">외 ${(merged.rows.length - 5).toLocaleString()}개 행</span>` : ''}
          </div>
          <div class="table-responsive">
            <table class="data-table">
              <thead>
                <tr>
                  ${merged.headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}
                </tr>
              </thead>
              <tbody>
                ${tableRowsHTML.length > 0 ? tableRowsHTML : `<tr><td colspan="${merged.headers.length}" style="text-align: center; color: var(--text-dim);">데이터가 없습니다.</td></tr>`}
              </tbody>
            </table>
          </div>
        </div>
      `;

      groupCard.innerHTML = headerHTML + columnsBarHTML + filesBarHTML + previewBodyHTML;

      // Bind Individual Group Actions
      groupCard.querySelector('.btn-open-modal').addEventListener('click', () => {
        openModalForGroup(group);
      });

      groupCard.querySelector('.btn-download-group-xlsx').addEventListener('click', () => {
        ExcelMerger.exportGroupToXLSX(group, state.options);
        showToast(`'${group.groupName}' 엑셀 파일이 다운로드되었습니다.`, 'success');
      });

      groupCard.querySelector('.btn-download-group-csv').addEventListener('click', () => {
        ExcelMerger.exportGroupToCSV(group, state.options);
        showToast(`'${group.groupName}' CSV 파일이 다운로드되었습니다.`, 'success');
      });

      elements.schemaGroupsList.appendChild(groupCard);
    });
  }

  /* ==========================================================================
     Modal Full Data Preview
     ========================================================================== */
  function openModalForGroup(group) {
    state.modalGroup = group;
    state.modalFilterText = '';
    elements.modalSearchInput.value = '';

    elements.modalGroupTitle.textContent = `${group.groupName}`;
    elements.previewModal.style.display = 'flex';
    document.body.style.overflow = 'hidden';

    renderModalTable();

    if (window.lucide) {
      lucide.createIcons();
    }
  }

  function renderModalTable() {
    if (!state.modalGroup) return;

    const merged = ExcelMerger.mergeGroupData(state.modalGroup, state.options);
    let rows = merged.rows;

    // Apply Filter if text is entered
    if (state.modalFilterText) {
      const query = state.modalFilterText;
      rows = rows.filter(row => {
        return merged.headers.some(h => {
          const val = String(row[h] || '').toLowerCase();
          return val.includes(query);
        });
      });
    }

    elements.modalRowBadge.textContent = `${rows.length.toLocaleString()}행`;

    const displayLimit = 500;
    const limitedRows = rows.slice(0, displayLimit);
    elements.modalPreviewInfo.textContent = rows.length > displayLimit 
      ? `검색 결과 ${rows.length.toLocaleString()}행 중 상위 ${displayLimit}행 표시 중` 
      : `총 ${rows.length.toLocaleString()}행 표시 중`;

    const tableRowsHTML = limitedRows.map(row => {
      const cellsHTML = merged.headers.map(h => {
        const val = row[h] !== undefined && row[h] !== null ? String(row[h]) : '';
        const isSource = h.startsWith('_');
        return `<td class="${isSource ? 'cell-source' : ''}" title="${val}">${escapeHtml(val)}</td>`;
      }).join('');
      return `<tr>${cellsHTML}</tr>`;
    }).join('');

    elements.modalTableContainer.innerHTML = `
      <table class="data-table">
        <thead>
          <tr>
            ${merged.headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}
          </tr>
        </thead>
        <tbody>
          ${tableRowsHTML.length > 0 ? tableRowsHTML : `<tr><td colspan="${merged.headers.length}" style="text-align: center; padding: 1.5rem; color: var(--text-dim);">검색 결과가 없습니다.</td></tr>`}
        </tbody>
      </table>
    `;
  }

  function closeModal() {
    elements.previewModal.style.display = 'none';
    document.body.style.overflow = '';
    state.modalGroup = null;
  }

  function copyModalTableToClipboard() {
    if (!state.modalGroup) return;
    const merged = ExcelMerger.mergeGroupData(state.modalGroup, state.options);
    
    const lines = [];
    lines.push(merged.headers.join('\t'));
    merged.rows.forEach(row => {
      const rowVals = merged.headers.map(h => String(row[h] || '').replace(/\t|\r?\n/g, ' '));
      lines.push(rowVals.join('\t'));
    });

    navigator.clipboard.writeText(lines.join('\n')).then(() => {
      showToast('클립보드에 복사되었습니다 (Ctrl+V로 붙여넣기 가능).', 'success');
    }).catch(err => {
      showToast('복사 실패: ' + err.message, 'error');
    });
  }

  /* ==========================================================================
     Sample Demo Data Generator
     ========================================================================== */
  function loadSampleDemoData() {
    showOverlay(true, '샘플 데이터 로드 중...', '가상 엑셀 파일 5개 생성 중');

    setTimeout(() => {
      try {
        // Schema 1: 고객 명단 (이름, 전화번호, 이메일, 거주지역, 가입일자)
        const schema1File1 = createMockExcelFile('2026_1분기_수도권_고객명단.xlsx', [
          ['이름', '전화번호', '이메일', '거주지역', '가입일자'],
          ['김민수', '010-1234-5678', 'minsoo@example.com', '서울 강남구', '2026-01-15'],
          ['이서연', '010-2345-6789', 'seoyeon@example.com', '경기 성남시', '2026-02-01'],
          ['박지훈', '010-3456-7890', 'jihoon@example.com', '인천 연수구', '2026-02-14'],
          ['최유진', '010-4567-8901', 'yujin@example.com', '서울 마포구', '2026-03-05']
        ]);

        const schema1File2 = createMockExcelFile('2026_1분기_지방권_고객명단.xlsx', [
          ['전화번호', '이름', '거주지역', '가입일자', '이메일'],
          ['010-5678-9012', '정도윤', '부산 해운대구', '2026-01-20', 'doyun@example.com'],
          ['010-6789-0123', '한예은', '대전 유성구', '2026-02-18', 'yeeun@example.com'],
          ['010-7890-1234', '송하준', '광주 서구', '2026-03-10', 'hajun@example.com']
        ]);

        // Schema 2: 상품 주문 및 결제 내역 (주문번호, 주문일시, 상품명, 수량, 결제금액, 결제수단)
        const schema2File1 = createMockExcelFile('스마트스토어_1월_결제내역.xlsx', [
          ['주문번호', '주문일시', '상품명', '수량', '결제금액', '결제수단'],
          ['ORD-20260101-01', '2026-01-01 10:25', '무선 블루투스 마우스', 2, 58000, '신용카드'],
          ['ORD-20260102-02', '2026-01-02 14:10', '기계식 키보드 갈축', 1, 129000, '네이버페이'],
          ['ORD-20260103-03', '2026-01-03 16:45', 'C타입 고속충전 케이블', 3, 27000, '카카오페이']
        ]);

        const schema2File2 = createMockExcelFile('쿠팡로켓_1월_결제내역.xlsx', [
          ['주문번호', '주문일시', '상품명', '수량', '결제금액', '결제수단'],
          ['ORD-20260104-04', '2026-01-04 09:30', '모니터 암 싱글', 1, 45000, '신용카드'],
          ['ORD-20260105-05', '2026-01-05 18:20', '게이밍 패드 장패드', 2, 36000, '토스페이'],
          ['ORD-20260106-06', '2026-01-06 21:00', 'USB 허브 7포트', 1, 38000, '신용카드']
        ]);

        // Schema 3: 재고 관리 목록 (상품코드, 상품명, 현재고, 안전재고, 보관위치)
        const schema3File = createMockExcelFile('본사물류센터_재고현황.xlsx', [
          ['상품코드', '상품명', '현재고', '안전재고', '보관위치'],
          ['PRD-A01', '무선 블루투스 마우스', 150, 30, 'A-01-03'],
          ['PRD-B02', '기계식 키보드 갈축', 45, 20, 'B-02-01'],
          ['PRD-C03', 'C타입 고속충전 케이블', 320, 50, 'C-01-05'],
          ['PRD-D04', '모니터 암 싱글', 80, 15, 'D-03-02']
        ]);

        const demoFiles = [schema1File1, schema1File2, schema2File1, schema2File2, schema3File];
        handleFilesSelected(demoFiles);
      } catch (err) {
        showOverlay(false);
        showToast('샘플 생성 오류: ' + err.message, 'error');
      }
    }, 200);
  }

  function createMockExcelFile(fileName, dataAOA) {
    const ws = XLSX.utils.aoa_to_sheet(dataAOA);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return new File([blob], fileName, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', lastModified: Date.now() });
  }

  /* ==========================================================================
     Helper Functions
     ========================================================================== */
  function toggleTheme() {
    const nextTheme = state.theme === 'dark' ? 'light' : 'dark';
    state.theme = nextTheme;
    document.documentElement.setAttribute('data-theme', nextTheme);
    localStorage.setItem('excel_merger_theme', nextTheme);
  }

  function resetAll() {
    state.rawFiles = [];
    state.parsedItems = [];
    state.schemaGroups = [];
    renderUI();
    showToast('전체 초기화되었습니다.', 'info');
  }

  function toggleSettingsAccordion() {
    const isHidden = elements.settingsContent.style.display === 'none';
    elements.settingsContent.style.display = isHidden ? 'grid' : 'none';
    elements.settingsToggleText.textContent = isHidden ? '설정 접기' : '설정 펼치기';
    elements.settingsChevron.style.transform = isHidden ? 'rotate(180deg)' : 'rotate(0deg)';
  }

  function showOverlay(show, title = '', desc = '') {
    elements.parsingOverlay.style.display = show ? 'flex' : 'none';
    if (show) {
      elements.parsingTitle.textContent = title;
      elements.parsingDesc.textContent = desc;
      elements.parsingBar.style.width = '0%';
    }
  }

  function updateProgress(percentage, desc) {
    elements.parsingBar.style.width = `${percentage}%`;
    if (desc) elements.parsingDesc.textContent = desc;
  }

  function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    let iconName = 'info';
    if (type === 'success') iconName = 'check-circle-2';
    if (type === 'error') iconName = 'alert-triangle';

    toast.innerHTML = `
      <i data-lucide="${iconName}"></i>
      <span>${escapeHtml(message)}</span>
    `;

    elements.toastContainer.appendChild(toast);
    if (window.lucide) {
      lucide.createIcons();
    }

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      toast.style.transition = 'all 0.25s ease';
      setTimeout(() => {
        if (toast.parentNode) {
          toast.parentNode.removeChild(toast);
        }
      }, 250);
    }, 3000);
  }

  function formatFileSize(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

});
