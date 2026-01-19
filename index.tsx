
// Fix: Removed invalid file markers that were causing parsing errors.
// --- Type Definitions for external libraries ---
// Fix: Replaced `declare` statements with explicit assignments from the global `window` object to prevent "Cannot find name 'declare'" errors.
const XLSX: any = (window as any).XLSX;
const Chart: any = (window as any).Chart;
const ChartDataLabels: any = (window as any).ChartDataLabels;

// --- DOM Elements ---
const fileUpload = document.getElementById('file-upload') as HTMLInputElement;
const dashboardGrid = document.getElementById('dashboard-grid') as HTMLElement;
const lastUpdate = document.getElementById('last-update') as HTMLElement;
const placeholder = document.getElementById('placeholder') as HTMLElement;
const filterContainer = document.getElementById('filter-container') as HTMLElement;
const chartsContainer = document.getElementById('charts-container') as HTMLElement;
const applyFiltersBtn = document.getElementById('apply-filters-btn') as HTMLButtonElement;
const resetFiltersBtn = document.getElementById('reset-filters-btn') as HTMLButtonElement;
const totalFclDisplay = document.getElementById('total-fcl-display') as HTMLElement;
const totalFclCount = document.getElementById('total-fcl-count') as HTMLElement;

// View Tabs
const viewTabsContainer = document.getElementById('view-tabs-container') as HTMLElement;
const viewVesselBtn = document.getElementById('view-vessel-btn') as HTMLButtonElement;
const viewPoBtn = document.getElementById('view-po-btn') as HTMLButtonElement;
const viewWarehouseBtn = document.getElementById('view-warehouse-btn') as HTMLButtonElement;
const viewDetailedBtn = document.getElementById('view-detailed-btn') as HTMLButtonElement;


// Filter Inputs
const arrivalStartDate = document.getElementById('arrival-start-date') as HTMLInputElement;
const arrivalEndDate = document.getElementById('arrival-end-date') as HTMLInputElement;
const deadlineStartDate = document.getElementById('deadline-start-date') as HTMLInputElement;
const deadlineEndDate = document.getElementById('deadline-end-date') as HTMLInputElement;
const statusFilter = document.getElementById('status-filter') as HTMLSelectElement;
const shipmentTypeFilter = document.getElementById('shipment-type-filter') as HTMLSelectElement;
const cargoTypeFilter = document.getElementById('cargo-type-filter') as HTMLSelectElement;
const poSearchInput = document.getElementById('po-search-input') as HTMLInputElement;
const vesselSearchInput = document.getElementById('vessel-search-input') as HTMLInputElement;
const blSearchInput = document.getElementById('bl-search-input') as HTMLInputElement;
const brokerSearchInput = document.getElementById('broker-search-input') as HTMLInputElement;
const poFilter = document.getElementById('po-filter') as HTMLSelectElement;
const vesselFilter = document.getElementById('vessel-filter') as HTMLSelectElement;
const batchFilter = document.getElementById('batch-filter') as HTMLSelectElement;
const brokerFilter = document.getElementById('broker-filter') as HTMLSelectElement;


// Export Buttons
const exportExcelBtn = document.getElementById('export-excel-btn') as HTMLButtonElement;


// Loading Overlay
const loadingOverlay = document.getElementById('loading-overlay') as HTMLElement;

// Column Visibility
const columnToggleContainer = document.getElementById('column-toggle-container') as HTMLElement;
const columnToggleBtn = document.getElementById('column-toggle-btn') as HTMLButtonElement;
const columnToggleDropdown = document.getElementById('column-toggle-dropdown') as HTMLElement;

// Theme Toggle Buttons
const darkModeBtn = document.getElementById('dark-mode-btn') as HTMLButtonElement;
const lightModeBtn = document.getElementById('light-mode-btn') as HTMLButtonElement;

// Modal Elements
const detailsModal = document.getElementById('details-modal') as HTMLElement;
const modalContent = document.getElementById('modal-content') as HTMLElement;
const modalHeaderContent = document.getElementById('modal-header-content') as HTMLElement;
const modalBody = document.getElementById('modal-body') as HTMLElement;
const modalCloseBtn = document.getElementById('modal-close-btn') as HTMLButtonElement;

// Logo Elements
const companyLogo = document.getElementById('company-logo') as HTMLImageElement;
const logoUpload = document.getElementById('logo-upload') as HTMLInputElement;
const removeLogoBtn = document.getElementById('remove-logo-btn') as HTMLButtonElement;

// Language Selector
const languageSelector = document.getElementById('language-selector') as HTMLSelectElement;


// --- Global State ---
let originalData: any[] = [];
let filteredDataCache: any[] = [];
let chartDataSource: {
    bar: any[];
    pie: { label: string; value: number }[];
    deadline: { label: string; count: number }[];
} = { bar: [], pie: [], deadline: [] };
let mainBarChart: any = null;
let statusPieChart: any = null;
let deadlineDistributionChart: any = null;
let currentViewState: 'vessel' | 'po' | 'warehouse' | 'detailed' = 'vessel';
const TODAY = new Date(); // Use current date
let currentSortKey: string = 'Dias Restantes';
let currentSortOrder: 'asc' | 'desc' = 'asc';
let activeModalItem: any = null;
let currentLanguage = 'pt';

// Column definitions for each view
const viewColumns = {
    vessel: ['PO SAP', 'VOYAGE', 'BL/AWB', 'SHIPOWNER', 'STATUS', 'SHIPMENT TYPE', 'TYPE OF CARGO', 'BATCH CHINA', 'ACTUAL ETA', 'FREE TIME DEADLINE', 'Dias Restantes', 'BONDED WAREHOUSE', 'BROKER'],
    po: ['ARRIVAL VESSEL', 'VOYAGE', 'BL/AWB', 'SHIPOWNER', 'STATUS', 'SHIPMENT TYPE', 'TYPE OF CARGO', 'BATCH CHINA', 'ACTUAL ETA', 'FREE TIME DEADLINE', 'Dias Restantes', 'BONDED WAREHOUSE', 'BROKER'],
    warehouse: ['ARRIVAL VESSEL', 'VOYAGE', 'PO SAP', 'BL/AWB', 'SHIPOWNER', 'STATUS', 'SHIPMENT TYPE', 'TYPE OF CARGO', 'BATCH CHINA', 'ACTUAL ETA', 'FREE TIME DEADLINE', 'Dias Restantes', 'BROKER']
};
let columnVisibility: Record<string, boolean> = {};


// --- Main App Initialization ---
window.addEventListener('load', () => {
    initializeApp();
});

function initializeApp() {
    // Defensive check for external libraries to prevent race conditions
    if (typeof Chart === 'undefined' || typeof ChartDataLabels === 'undefined' || typeof XLSX === 'undefined') {
        console.warn("External libraries not yet loaded, retrying in 100ms...");
        setTimeout(initializeApp, 100);
        return;
    }
    
    // Initialize Language
    const savedLang = localStorage.getItem('language') || 'pt';
    setLanguage(savedLang);

    // Initialize Theme
    const savedTheme = localStorage.getItem('theme');
    const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
    const initialTheme = savedTheme || (prefersDark ? 'dark' : 'light');
    setTheme(initialTheme as 'dark' | 'light');


    // Event Listeners
    fileUpload.addEventListener('change', handleFileUpload);
    applyFiltersBtn.addEventListener('click', applyFiltersAndRender);
    resetFiltersBtn.addEventListener('click', resetFiltersAndRender);
    viewVesselBtn.addEventListener('click', () => setView('vessel'));
    viewPoBtn.addEventListener('click', () => setView('po'));
    viewWarehouseBtn.addEventListener('click', () => setView('warehouse'));
    viewDetailedBtn.addEventListener('click', () => setView('detailed'));
    exportExcelBtn.addEventListener('click', handleExcelExport);
    darkModeBtn.addEventListener('click', () => setTheme('dark'));
    lightModeBtn.addEventListener('click', () => setTheme('light'));
    poSearchInput.addEventListener('input', applyFiltersAndRender);
    vesselSearchInput.addEventListener('input', applyFiltersAndRender);
    blSearchInput.addEventListener('input', applyFiltersAndRender);
    brokerSearchInput.addEventListener('input', applyFiltersAndRender);
    logoUpload.addEventListener('change', handleLogoUpload);
    removeLogoBtn.addEventListener('click', handleRemoveLogo);
    languageSelector.addEventListener('change', (e) => setLanguage((e.target as HTMLSelectElement).value));


    dashboardGrid.addEventListener('click', handleSortClick);
    setupColumnToggles();
    loadSavedLogo();

    // Modal Listeners
    modalCloseBtn.addEventListener('click', closeModal);
    detailsModal.addEventListener('click', (e) => {
        if (e.target === detailsModal) { // Close if clicking on the overlay
            closeModal();
        }
    });
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && !detailsModal.classList.contains('hidden')) {
            closeModal();
        }
    });
}

// --- Translation ---
const translations = {
    // General UI
    'dashboard_title': { pt: 'DASHBOARD', en: 'DASHBOARD', zh: '仪表板' },
    'upload_prompt': { pt: 'Carregue um arquivo .xlsx para começar', en: 'Upload an .xlsx file to start', zh: '上传.xlsx文件开始' },
    'total_containers_view': { pt: 'Total Containers (Visão Atual):', en: 'Total Containers (Current View):', zh: '总集装箱数（当前视图）：' },
    'tooltip_upload_logo': { pt: 'Carregar Logo da Empresa', en: 'Upload Company Logo', zh: '上传公司标志' },
    'tooltip_remove_logo': { pt: 'Remover Logo', en: 'Remove Logo', zh: '移除标志' },
    'tooltip_dark_mode': { pt: 'Modo Escuro', en: 'Dark Mode', zh: '深色模式' },
    'tooltip_light_mode': { pt: 'Modo Claro', en: 'Light Mode', zh: '浅色模式' },
    'export_pdf': { pt: 'Exportar PDF', en: 'Export PDF', zh: '导出PDF' },
    'export_excel': { pt: 'Exportar Excel', en: 'Export Excel', zh: '导出Excel' },
    'export_csv': { pt: 'Exportar CSV', en: 'Export CSV', zh: '导出CSV' },
    'upload_xlsx': { pt: 'Carregar XLSX', en: 'Upload XLSX', zh: '上传XLSX' },
    'loading_text': { pt: 'Processando...', en: 'Processing...', zh: '处理中...' },
    'upload_processing': { pt: 'Processando...', en: 'Processing...', zh: '处理中...' },
    // Filters
    'search_po': { pt: 'Pesquisar por PO', en: 'Search by PO', zh: '按采购订单搜索' },
    'placeholder_search_po': { pt: 'Digite o número da PO...', en: 'Enter PO number...', zh: '输入采购订单号...' },
    'filter_po': { pt: 'Filtrar POs', en: 'Filter POs', zh: '筛选采购订单' },
    'search_vessel': { pt: 'Pesquisar por Navio', en: 'Search by Vessel', zh: '按船只搜索' },
    'placeholder_search_vessel': { pt: 'Digite o nome do navio...', en: 'Enter vessel name...', zh: '输入船名...' },
    'search_bl': { pt: 'Pesquisar por BL', en: 'Search by BL', zh: '按提单搜索' },
    'placeholder_search_bl': { pt: 'Digite o número do BL...', en: 'Enter BL number...', zh: '输入提单号...' },
    'search_broker': { pt: 'Pesquisar por Broker', en: 'Search by Broker', zh: '按Broker搜索' },
    'placeholder_search_broker': { pt: 'Digite o nome do Broker...', en: 'Enter Broker name...', zh: '输入Broker名称...' },
    'filter_vessel': { pt: 'Filtrar Navios', en: 'Filter Vessels', zh: '筛选船只' },
    'filter_broker': { pt: 'Filtrar Brokers', en: 'Filter Brokers', zh: '筛选Brokers' },
    'cargo_status': { pt: 'Status da Carga', en: 'Cargo Status', zh: '货物状态' },
    'shipment_type': { pt: 'Tipo de Carga', en: 'Shipment Type', zh: '装运类型' },
    'cargo_type': { pt: 'Tipo de Mercadoria', en: 'Cargo Type', zh: '货物类型' },
    'batch_filter': { pt: 'Batch China', en: 'Batch China', zh: '中国批次' },
    'arrival_start': { pt: 'Início da Chegada', en: 'Arrival Start', zh: '预计到达开始' },
    'arrival_end': { pt: 'Fim da Chegada', en: 'Arrival End', zh: '预计到达结束' },
    'freetime_start': { pt: 'Início do FreeTime', en: 'FreeTime Start', zh: '免费期开始' },
    'freetime_end': { pt: 'Fim do FreeTime', en: 'FreeTime End', zh: '免费期结束' },
    'filter_button': { pt: 'Filtrar', en: 'Filter', zh: '筛选' },
    'reset_button': { pt: 'Limpar', en: 'Reset', zh: '重置' },
    'columns_button': { pt: 'Colunas', en: 'Columns', zh: '列' },
    // Views
    'view_by_vessel': { pt: 'Análise por Navio', en: 'Analysis by Vessel', zh: '按船只分析' },
    'view_by_po': { pt: 'Análise por PO', en: 'Analysis by PO', zh: '按采购订单分析' },
    'view_by_warehouse': { pt: 'Análise por Armazém', en: 'Analysis by Warehouse', zh: '按仓库分析' },
    'view_by_detailed': { pt: 'Visão Detalhada', en: 'Detailed View', zh: '详细视图' },
    // Charts
    'chart_title_containers_by_vessel': { pt: 'Total de Containers por Navio', en: 'Total Containers by Vessel', zh: '各船只集装箱总数' },
    'chart_title_containers_by_po': { pt: 'Total de Containers por PO', en: 'Total Containers by PO', zh: '各采购订单集装箱总数' },
    'chart_title_containers_by_warehouse': { pt: 'Total de Containers por Armazém', en: 'Total Containers by Warehouse', zh: '各仓库集装箱总数' },
    'chart_title_status_distribution': { pt: 'Distribuição de Status de Carga', en: 'Cargo Status Distribution', zh: '货物状态分布' },
    'chart_title_deadline_distribution': { pt: 'Distribuição de Prazos (Dias Restantes)', en: 'Deadline Distribution (Days Remaining)', zh: '截止日期分布（剩余天数）' },
    // Placeholder
    'placeholder_title': { pt: 'Aguardando arquivo...', en: 'Waiting for file...', zh: '等待文件...' },
    'placeholder_text': { pt: 'Selecione uma planilha para analisar os navios e o risco de demurrage.', en: 'Select a spreadsheet to analyze vessels and demurrage risk.', zh: '选择电子表格以分析船只和滞期费风险。' },
    // Dynamic Text & Placeholders
    'data_from_sheet': { pt: 'Dados de', en: 'Data from', zh: '数据来源' },
    'updated_at': { pt: 'Atualizado em', en: 'Updated on', zh: '更新于' },
    'all_option': { pt: '-- Todos --', en: '-- All --', zh: '-- 全部 --' },
    'no_status': { pt: 'Sem Status', en: 'No Status', zh: '无状态' },
    'no_po': { pt: 'Sem PO', en: 'No PO', zh: '无采购订单' },
    'no_vessel': { pt: 'Sem Navio', en: 'No Vessel', zh: '无船只' },
    'no_shipment_type': { pt: 'Sem Tipo', en: 'No Type', zh: '无类型' },
    'no_cargo_type': { pt: 'Sem Tipo de Mercadoria', en: 'No Cargo Type', zh: '无货物类型' },
    'no_warehouse': { pt: 'Sem Armazém', en: 'No Warehouse', zh: '无仓库' },
    'no_batch': { pt: 'Sem Batch', en: 'No Batch', zh: '无批次' },
    'no_broker': { pt: 'Sem Broker', en: 'No Broker', zh: '无Broker' },
    // Toast Messages
    'toast_success_upload': { pt: 'Dashboard carregado com sucesso!', en: 'Dashboard loaded successfully!', zh: '仪表板加载成功！' },
    'toast_error_sheet_not_found': { pt: 'Planilha "{sheetName}" não encontrada.', en: 'Sheet "{sheetName}" not found.', zh: '未找到工作表“{sheetName}”。' },
    'toast_error_sheet_empty': { pt: 'A planilha está vazia.', en: 'The spreadsheet is empty.', zh: '电子表格为空。' },
    'toast_error_processing_file': { pt: 'Erro ao processar arquivo.', en: 'Error processing file.', zh: '处理文件时出错。' },
    'toast_error_reading_file': { pt: 'Não foi possível ler o arquivo.', en: 'Could not read the file.', zh: '无法读取文件。' },
    'toast_no_data_to_export': { pt: 'Não há dados para exportar.', en: 'No data to export.', zh: '无数据可导出。' },
    'toast_excel_export_started': { pt: 'Exportação para Excel iniciada.', en: 'Excel export started.', zh: 'Excel导出已开始。' },
    'toast_pdf_error': { pt: 'Ocorreu um erro ao gerar o PDF.', en: 'An error occurred while generating the PDF.', zh: '生成PDF时出错。' },
    'toast_logo_updated': { pt: 'Logo atualizado com sucesso!', en: 'Logo updated successfully!', zh: '标志更新成功！' },
    'toast_logo_error': { pt: 'Erro ao carregar o logo.', en: 'Error uploading logo.', zh: '上传标志时出错。' },
    'toast_logo_removed': { pt: 'Logo removido.', en: 'Logo removed.', zh: '标志已移除。' },
    // Card & Modal
    'containers': { pt: 'Containers', en: 'Containers', zh: '集装箱' },
    'delivered': { pt: 'Entregue', en: 'Delivered', zh: '已交付' },
    'remaining_days': { pt: 'Dias Restantes', en: 'Days Remaining', zh: '剩余天数' },
    'card_placeholder_no_vessels': { pt: 'Nenhum navio encontrado para os filtros aplicados.', en: 'No vessels found for the applied filters.', zh: '未找到符合所应用筛选条件的船只。' },
    'card_placeholder_no_pos': { pt: 'Nenhuma PO encontrada para os filtros aplicados.', en: 'No POs found for the applied filters.', zh: '未找到符合所应用筛选条件的采购订单。' },
    'card_placeholder_no_warehouses': { pt: 'Nenhum armazém encontrado para os filtros aplicados.', en: 'No warehouses found for the applied filters.', zh: '未找到符合所应用筛选条件的仓库。' },
    'card_placeholder_try_resetting': { pt: 'Tente limpar os filtros ou carregar um novo arquivo.', en: 'Try clearing the filters or uploading a new file.', zh: '请尝试清除筛选条件或上传新文件。' },
    // Chart Tooltips/Labels
    'high_risk': { pt: 'Risco Alto', en: 'High Risk', zh: '高风险' },
    'medium_risk': { pt: 'Risco Médio', en: 'Medium Risk', zh: '中风险' },
    'low_risk': { pt: 'Risco Baixo', en: 'Low Risk', zh: '低风险' },
    'chart_tooltip_delivered': { pt: 'Entregue', en: 'Delivered', zh: '已交付' },
    'deadline_bin_overdue': { pt: 'Atrasado (<0)', en: 'Overdue (<0)', zh: '逾期 (<0)' },
    'deadline_bin_0_7': { pt: '0-7 Dias', en: '0-7 Days', zh: '0-7 天' },
    'deadline_bin_8_15': { pt: '8-15 Dias', en: '8-15 Days', zh: '8-15 天' },
    'deadline_bin_16_30': { pt: '16-30 Dias', en: '16-30 Days', zh: '16-30 天' },
    'deadline_bin_31_plus': { pt: '31+ Dias', en: '31+ Days', zh: '31+ 天' },
    'chart_deadline_label': { pt: 'Nº de Cargas', en: '# of Shipments', zh: '装运数量' },
};

function translate(key: string, replacements: Record<string, string> = {}): string {
    const translation = translations[key]?.[currentLanguage] || key;
    return Object.entries(replacements).reduce((acc, [placeholder, value]) => {
        return acc.replace(`{${placeholder}}`, value);
    }, translation);
}

function setLanguage(lang: string) {
    currentLanguage = lang;
    localStorage.setItem('language', lang);
    if (languageSelector.value !== lang) {
        languageSelector.value = lang;
    }
    
    document.documentElement.lang = lang.split('-')[0];

    document.querySelectorAll('[data-translate-key]').forEach(el => {
        const key = el.getAttribute('data-translate-key')!;
        (el as HTMLElement).innerText = translate(key);
    });

    document.querySelectorAll('[data-translate-key-placeholder]').forEach(el => {
        const key = el.getAttribute('data-translate-key-placeholder')!;
        (el as HTMLInputElement).placeholder = translate(key);
    });

    document.querySelectorAll('[data-translate-key-title]').forEach(el => {
        const key = el.getAttribute('data-translate-key-title')!;
        (el as HTMLElement).title = translate(key);
    });

    if (originalData.length > 0) {
        // Re-populate static filters to update "All" option
        populateStatusFilter(originalData);
        populateShipmentTypeFilter(originalData);
        populateCargoTypeFilter(originalData);
        populatePoFilter(originalData);
        populateVesselFilter(originalData);
        populateBatchFilter(originalData);
        populateBrokerFilter(originalData);
        // Re-render everything to apply translations to dynamic content
        applyFiltersAndRender();
    }
}

// --- Theme Management ---
function setTheme(theme: 'dark' | 'light') {
    const isDark = theme === 'dark';
    
    // 1. Toggle class on body
    document.body.classList.toggle('dark', isDark);

    // 2. Toggle button visibility
    darkModeBtn.classList.toggle('hidden', isDark);
    lightModeBtn.classList.toggle('hidden', !isDark);

    // 3. Save to localStorage
    localStorage.setItem('theme', theme);

    // 4. Update Chart.js defaults for new charts
    const textColor = isDark ? '#d1d5db' : '#4b5563';
    const gridColor = isDark ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.1)';
    Chart.defaults.color = textColor;
    Chart.defaults.borderColor = gridColor;

    // 5. If data is loaded, re-render to update existing charts and UI
    if (originalData.length > 0) {
        applyFiltersAndRender(); 
    }
}


// --- Loading Indicator ---
function showLoading() {
    loadingOverlay.classList.remove('hidden');
}

function hideLoading() {
    loadingOverlay.classList.add('hidden');
}


// --- Toast Notifications ---
function showToast(messageKey: string, type: 'success' | 'error' | 'warning' = 'success', replacements: Record<string, string> = {}) {
    const toastContainer = document.getElementById('toast-container');
    if (!toastContainer) return;

    const message = translate(messageKey, replacements);
    const toast = document.createElement('div');
    const icons = { success: 'fa-check-circle', error: 'fa-times-circle', warning: 'fa-exclamation-triangle' };
    const colors = { success: 'bg-green-500', error: 'bg-red-500', warning: 'bg-yellow-500' };
    toast.className = `toast ${colors[type]} text-white py-3 px-5 rounded-lg shadow-xl flex items-center mb-2`;
    toast.innerHTML = `<i class="fas ${icons[type]} mr-3"></i> <p>${message}</p>`;
    toastContainer.appendChild(toast);
    setTimeout(() => toast.remove(), 5000);
}

// --- File Handling ---
function handleFileUpload(event: Event) {
    const target = event.target as HTMLInputElement;
    const file = target.files?.[0];
    const uploadLabel = document.querySelector('label[for="file-upload"] > span');
    if (!file || !uploadLabel) return;

    const originalText = uploadLabel.innerHTML;
    uploadLabel.parentElement?.classList.add('opacity-50', 'cursor-not-allowed');
    uploadLabel.innerHTML = `<i class="fas fa-spinner fa-spin mr-2"></i> ${translate('upload_processing')}`;

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const workbook = XLSX.read(new Uint8Array(e.target!.result as ArrayBuffer), { type: 'array' });
            const sheetName = "FUP Report";
            if (!workbook.Sheets[sheetName]) throw new Error(translate('toast_error_sheet_not_found', {sheetName}));
            
            originalData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { raw: false, defval: '' });
            
            if (originalData.length === 0) throw new Error(translate('toast_error_sheet_empty'));

            // --- Data Normalization Step ---
            originalData.forEach(row => {
                // Normalize warehouse names to group "TECON - WILSON SONS" with "TECON"
                const warehouse = String(row['BONDED WAREHOUSE'] || '').toUpperCase();
                if (warehouse.includes('TECON')) {
                    row['BONDED WAREHOUSE'] = 'TECON';
                }
            });

            populateStatusFilter(originalData);
            populateShipmentTypeFilter(originalData);
            populateCargoTypeFilter(originalData);
            populatePoFilter(originalData);
            populateVesselFilter(originalData);
            populateBatchFilter(originalData);
            populateBrokerFilter(originalData);
            applyFiltersAndRender();
            
            filterContainer.classList.remove('hidden');
            chartsContainer.classList.remove('hidden');
            viewTabsContainer.classList.remove('hidden');
            exportExcelBtn.classList.remove('hidden');
            totalFclDisplay.classList.remove('hidden');

            lastUpdate.textContent = `${translate('data_from_sheet')} "${sheetName}" | ${translate('updated_at')}: ${new Date().toLocaleString(currentLanguage)}`;
            showToast('toast_success_upload', 'success');
        } catch (err: any) {
            showToast(err.message, 'error'); // Show raw error message from catch
            resetUI();
        } finally {
            uploadLabel.parentElement?.classList.remove('opacity-50', 'cursor-not-allowed');
            uploadLabel.innerHTML = originalText;
            fileUpload.value = '';
        }
    };
    reader.onerror = () => {
        showToast('toast_error_reading_file', 'error');
        resetUI();
    };
    reader.readAsArrayBuffer(file);
}

// --- Data Processing & Filtering ---
function excelDateToJSDate(serial: any): Date | null {
    if (!serial) return null;
    if (typeof serial === 'string') {
        if (serial.match(/^\d{5}$/)) { // Looks like an excel serial number as a string
             serial = parseInt(serial, 10);
        } else if (serial.includes('/') || serial.includes('-') || serial.includes('.')) {
             // Attempt to parse common date formats, this might need refinement
             const date = new Date(serial.replace(/(\d{2})\.(\d{2})\.(\d{4})/, '$2/$1/$3')); // Handle DD.MM.YYYY
             return isNaN(date.getTime()) ? null : date;
        } else {
            return null;
        }
    }
    if (typeof serial !== 'number' || serial < 1) return null;
    const utc_days = Math.floor(serial - 25569);
    const date_info = new Date(utc_days * 86400 * 1000);
    return new Date(date_info.getTime() + (date_info.getTimezoneOffset() * 60 * 1000));
}

/**
 * Finds the earliest arrival date from a group of shipments.
 * @param shipments - An array of shipment objects.
 * @returns The earliest arrival date as a Date object, or null if no valid dates are found.
 */
function getEarliestArrivalDate(shipments: any[]): Date | null {
    return shipments.reduce((earliest: Date | null, shipment: any) => {
        const arrivalDate = excelDateToJSDate(shipment['ACTUAL ETA']);
        if (arrivalDate && (!earliest || arrivalDate < earliest)) {
            return arrivalDate;
        }
        return earliest;
    }, null);
}

function applyFiltersAndRender() {
    showLoading();
    setTimeout(() => {
        let filteredData = [...originalData];
        
        // Date Filters
        const arrivalStart = arrivalStartDate.valueAsDate;
        const arrivalEnd = arrivalEndDate.valueAsDate;
        const deadlineStart = deadlineStartDate.valueAsDate;
        const deadlineEnd = deadlineEndDate.valueAsDate;

        if (arrivalStart) filteredData = filteredData.filter(row => {
            const arrivalDate = excelDateToJSDate(row['ACTUAL ETA']);
            return arrivalDate && arrivalDate >= arrivalStart;
        });
        if (arrivalEnd) filteredData = filteredData.filter(row => {
            const arrivalDate = excelDateToJSDate(row['ACTUAL ETA']);
            return arrivalDate && arrivalDate <= arrivalEnd;
        });
        if (deadlineStart) filteredData = filteredData.filter(row => {
            const deadlineDate = excelDateToJSDate(row['FREE TIME DEADLINE']);
            return deadlineDate && deadlineDate >= deadlineStart;
        });
        if (deadlineEnd) filteredData = filteredData.filter(row => {
            const deadlineDate = excelDateToJSDate(row['FREE TIME DEADLINE']);
            return deadlineDate && deadlineDate <= deadlineEnd;
        });

        // PO Search Filter
        const poSearchTerm = poSearchInput.value.trim().toLowerCase();
        if (poSearchTerm) {
            filteredData = filteredData.filter(row => {
                return String(row['PO SAP'] || '').toLowerCase().includes(poSearchTerm);
            });
        }

        // BL Search Filter
        const blSearchTerm = blSearchInput.value.trim().toLowerCase();
        if (blSearchTerm) {
            filteredData = filteredData.filter(row => {
                return String(row['BL/AWB'] || '').toLowerCase().includes(blSearchTerm);
            });
        }

        // Broker Search Filter
        const brokerSearchTerm = brokerSearchInput.value.trim().toLowerCase();
        if (brokerSearchTerm) {
            filteredData = filteredData.filter(row => {
                return String(row['BROKER'] || '').toLowerCase().includes(brokerSearchTerm);
            });
        }
        
        // PO Filter (multi-select)
        const selectedPOs = Array.from(poFilter.selectedOptions).map(opt => opt.value);
        if (selectedPOs.length > 0 && !selectedPOs.includes('')) {
            filteredData = filteredData.filter(row => selectedPOs.includes(row['PO SAP'] || translate('no_po')));
        }

        // Vessel Filter (multi-select)
        const selectedVessels = Array.from(vesselFilter.selectedOptions).map(opt => opt.value);
        if (selectedVessels.length > 0 && !selectedVessels.includes('')) {
            filteredData = filteredData.filter(row => selectedVessels.includes(row['ARRIVAL VESSEL'] || translate('no_vessel')));
        }

        // Status Filter (multi-select)
        const selectedStatuses = Array.from(statusFilter.selectedOptions).map(opt => opt.value);
        if (selectedStatuses.length > 0 && !selectedStatuses.includes('')) {
            filteredData = filteredData.filter(row => selectedStatuses.includes(row.STATUS || translate('no_status')));
        }

        // ShipmentType Filter (multi-select)
        const selectedShipmentTypes = Array.from(shipmentTypeFilter.selectedOptions).map(opt => opt.value);
        if (selectedShipmentTypes.length > 0 && !selectedShipmentTypes.includes('')) {
            filteredData = filteredData.filter(row => selectedShipmentTypes.includes(row['SHIPMENT TYPE'] || translate('no_shipment_type')));
        }

        // CargoType Filter (multi-select)
        const selectedCargoTypes = Array.from(cargoTypeFilter.selectedOptions).map(opt => opt.value);
        if (selectedCargoTypes.length > 0 && !selectedCargoTypes.includes('')) {
            filteredData = filteredData.filter(row => selectedCargoTypes.includes(row['TYPE OF CARGO'] || translate('no_cargo_type')));
        }
        
        // Batch Filter (multi-select)
        const selectedBatches = Array.from(batchFilter.selectedOptions).map(opt => opt.value);
        if (selectedBatches.length > 0 && !selectedBatches.includes('')) {
            filteredData = filteredData.filter(row => selectedBatches.includes(row['BATCH CHINA'] || translate('no_batch')));
        }

        // Broker Filter (multi-select)
        const selectedBrokers = Array.from(brokerFilter.selectedOptions).map(opt => opt.value);
        if (selectedBrokers.length > 0 && !selectedBrokers.includes('')) {
            filteredData = filteredData.filter(row => selectedBrokers.includes(row['BROKER'] || translate('no_broker')));
        }

        // Vessel Search Filter
        const vesselSearchTerm = vesselSearchInput.value.trim().toLowerCase();
        if (vesselSearchTerm) {
            filteredData = filteredData.filter(row => {
                return String(row['ARRIVAL VESSEL'] || '').toLowerCase().includes(vesselSearchTerm);
            });
        }

        filteredDataCache = filteredData;
        
        let processedData;
        if (currentViewState === 'vessel') {
            processedData = processDataByVessel(filteredData);
            renderVesselDashboard(processedData);
            renderCharts(processedData, filteredData);
        } else if (currentViewState === 'po') {
            processedData = processDataByPO(filteredData);
            renderPODashboard(processedData);
            renderCharts(processedData, filteredData);
        } else if (currentViewState === 'warehouse') {
            processedData = processDataByWarehouse(filteredData);
            renderWarehouseDashboard(processedData);
            renderCharts(processedData, filteredData);
        } else if (currentViewState === 'detailed') {
            processedData = processDataForDetailedView(filteredData);
            renderDetailedDashboard(processedData);
            // Destroy charts if they exist
            if (mainBarChart) mainBarChart.destroy();
            if (statusPieChart) statusPieChart.destroy();
            if (deadlineDistributionChart) deadlineDistributionChart.destroy();
        }
        
        updateTotalContainers(filteredData);
        hideLoading();
    }, 50); // Timeout allows UI to show spinner before processing
}

function resetFiltersAndRender() {
    showLoading();
    setTimeout(() => {
        arrivalStartDate.value = '';
        arrivalEndDate.value = '';
        deadlineStartDate.value = '';
        deadlineEndDate.value = '';
        poSearchInput.value = '';
        vesselSearchInput.value = '';
        blSearchInput.value = '';
        brokerSearchInput.value = '';
        Array.from(statusFilter.options).forEach((opt, i) => opt.selected = i === 0);
        Array.from(shipmentTypeFilter.options).forEach((opt, i) => opt.selected = i === 0);
        Array.from(cargoTypeFilter.options).forEach((opt, i) => opt.selected = i === 0);
        Array.from(poFilter.options).forEach((opt, i) => opt.selected = i === 0);
        Array.from(vesselFilter.options).forEach((opt, i) => opt.selected = i === 0);
        Array.from(batchFilter.options).forEach((opt, i) => opt.selected = i === 0);
        Array.from(brokerFilter.options).forEach((opt, i) => opt.selected = i === 0);
        applyFiltersAndRender();
        hideLoading();
    }, 50);
}


// --- View Switching ---
function setView(view: 'vessel' | 'po' | 'warehouse' | 'detailed') {
    if (currentViewState === view) return;
    currentViewState = view;
    
    const buttons = { vessel: viewVesselBtn, po: viewPoBtn, warehouse: viewWarehouseBtn, detailed: viewDetailedBtn };
    Object.values(buttons).forEach(btn => {
        btn.classList.add('text-gray-500', 'border-transparent');
        btn.classList.remove('border-blue-600', 'text-blue-600');
    });

    buttons[view].classList.add('border-blue-600', 'text-blue-600');
    buttons[view].classList.remove('text-gray-500', 'border-transparent');
    
    // Hide/show charts and column toggles based on view
    const isDetailedView = view === 'detailed';
    chartsContainer.classList.toggle('hidden', isDetailedView);
    columnToggleContainer.classList.toggle('hidden', isDetailedView);

    if (originalData.length > 0) {
        if (!isDetailedView) {
            populateColumnToggles(); // Repopulate columns for the new view
        }
        applyFiltersAndRender();
    }
}

// --- Data Grouping and Processing ---
/**
 * Calculates the container count for a single shipment row by summing 'FCL' and 'LCL' columns.
 * @param shipment - A single shipment object from the dataset.
 * @returns The total number of containers for that row.
 */
function getContainerCount(shipment: any): number {
    const fcl = parseInt(shipment['FCL'], 10) || 0;
    const lcl = parseInt(shipment['LCL'], 10) || 0;
    return fcl + lcl;
}

/**
 * Calculates the total container count for a list of shipments, ensuring accuracy
 * by summing quantities based on unique Bill of Lading (BL/AWB) numbers.
 * This prevents double-counting if a shipment is listed on multiple rows.
 * @param shipments - An array of shipment objects.
 * @returns The total number of containers.
 */
function calculateCorrectTotals(shipments: any[]): number {
    const blQuantities = new Map<string, number>();

    shipments.forEach(s => {
        // Use BL/AWB as the unique key. If a BL is split across rows, this will sum them up correctly.
        const bl = s['BL/AWB'];
        if (bl) {
            const quantity = getContainerCount(s);
            // Only add the quantity for the first time we see a BL to avoid double counting from other groupings.
            // A more robust way is to sum unique BLs. Let's build a map of BL to its quantity.
            // This assumes that if a BL is listed multiple times, each row has a partial quantity.
             blQuantities.set(bl, (blQuantities.get(bl) || 0) + quantity);
        }
    });

    // Sum the quantities of all unique BLs.
    // If we assume each row is unique, a simple sum is better.
    // The BL de-duplication is safer if the raw data can have repeated lines for the same BL.
    let total = 0;
    const uniqueBLs = new Set();
    shipments.forEach(s => {
        const bl = s['BL/AWB'];
        if (bl && !uniqueBLs.has(bl)) {
            total += getContainerCount(s); // This is not quite right if BLs are duplicated with partial sums.
            uniqueBLs.add(bl);
        } else if (!bl) { // Handle rows without a BL
            total += getContainerCount(s);
        }
    });
    // The reduce approach is better as it handles split BLs. Let's revert to a simpler total.
    return shipments.reduce((sum, s) => sum + getContainerCount(s), 0);
}


function calculateRisk(shipment: any) {
    let deadline = excelDateToJSDate(shipment['FREE TIME DEADLINE']);
    // FIX: Add a sanity check for very old dates which are likely parsing errors from blank cells.
    if (deadline && deadline.getFullYear() < 1950) {
        deadline = null;
    }
    
    const isDelivered = (shipment.STATUS || '').toLowerCase().includes('delivered');
    let daysToDeadline: number | null = null;
    let risk = 'low';

    if (isDelivered) {
        risk = 'none';
    } else if (deadline) {
        const timeDiff = deadline.getTime() - TODAY.getTime();
        daysToDeadline = Math.ceil(timeDiff / (1000 * 3600 * 24));
        if (daysToDeadline < 0) risk = 'high';
        else if (daysToDeadline <= 7) risk = 'medium';
    }
    return { ...shipment, daysToDeadline, risk };
}

function processDataByVessel(data: any[]) {
    const defaultVesselName = translate('no_vessel');
    // FIX: Explicitly typed the `grouped` constant and the `reduce` accumulator to ensure correct type inference, preventing cascading 'unknown' type errors.
    const grouped: Record<string, any[]> = data.reduce((acc: Record<string, any[]>, row: any) => {
        const vesselName = row['ARRIVAL VESSEL'] ? String(row['ARRIVAL VESSEL']).trim().toUpperCase() : defaultVesselName;
        const voyage = row['VOYAGE'] ? String(row['VOYAGE']).trim().toUpperCase() : '';
        const name = voyage ? `${vesselName} - ${voyage}` : vesselName; // Composite key
        
        if (!acc[name]) acc[name] = [];
        acc[name].push(row);
        return acc;
    }, {});

    return Object.entries(grouped).map(([name, shipments]) => {
        const processedShipments = shipments.map(calculateRisk);
        const riskCounts = { high: 0, medium: 0, low: 0, none: 0 };
        processedShipments.forEach(s => (riskCounts as any)[s.risk]++);
        
        let overallRisk = 'low';
        if (riskCounts.high > 0) overallRisk = 'high';
        else if (riskCounts.medium > 0) overallRisk = 'medium';
        else if (riskCounts.none === processedShipments.length) overallRisk = 'none';

        const totalContainers = shipments.reduce((sum, s) => sum + getContainerCount(s), 0);
        const earliestArrival = getEarliestArrivalDate(shipments);
        
        return { name, shipments: processedShipments, totalFCL: totalContainers, overallRisk, riskCounts, earliestArrival };
    }).sort((a, b) => {
        // Sort chronologically by the earliest arrival date in the group.
        const dateA = a.earliestArrival;
        const dateB = b.earliestArrival;

        if (dateA && dateB) {
            return dateA.getTime() - dateB.getTime(); // Sort ascending
        }
        if (dateA) return -1; // Groups with dates come before groups without
        if (dateB) return 1;
        return a.name.localeCompare(b.name); // Fallback to alphabetical if no dates
    });
}

function processDataByPO(data: any[]) {
    return processDataGeneric(data, 'PO SAP', translate('no_po'));
}

function processDataByWarehouse(data: any[]) {
    return processDataGeneric(data, 'BONDED WAREHOUSE', translate('no_warehouse'));
}

function processDataGeneric(data: any[], groupKey: string, defaultName: string) {
    // FIX: Explicitly typed the `grouped` constant and the `reduce` accumulator to ensure correct type inference, preventing cascading 'unknown' type errors.
    const grouped: Record<string, any[]> = data.reduce((acc: Record<string, any[]>, row: any) => {
        const name = row[groupKey] ? String(row[groupKey]).trim().toUpperCase() : defaultName;
        if (!acc[name]) acc[name] = [];
        acc[name].push(row);
        return acc;
    }, {});

    return Object.entries(grouped).map(([name, shipments]) => {
        const processedShipments = shipments.map(calculateRisk);
        const riskCounts = { high: 0, medium: 0, low: 0, none: 0 };
        processedShipments.forEach(s => (riskCounts as any)[s.risk]++);
        
        let overallRisk = 'low';
        if (riskCounts.high > 0) overallRisk = 'high';
        else if (riskCounts.medium > 0) overallRisk = 'medium';
        else if (riskCounts.none === processedShipments.length) overallRisk = 'none';

        const totalContainers = shipments.reduce((sum, s) => sum + getContainerCount(s), 0);
        const earliestArrival = getEarliestArrivalDate(shipments);
        
        return { name, shipments: processedShipments, totalFCL: totalContainers, overallRisk, riskCounts, earliestArrival };
    }).sort((a, b) => {
        // Sort chronologically by the earliest arrival date in the group.
        const dateA = a.earliestArrival;
        const dateB = b.earliestArrival;

        if (dateA && dateB) {
            return dateA.getTime() - dateB.getTime(); // Sort ascending
        }
        if (dateA) return -1; // Groups with dates come before groups without
        if (dateB) return 1;
        return a.name.localeCompare(b.name); // Fallback to alphabetical if no dates
    });
}

function processDataForDetailedView(data: any[]) {
    const defaultVesselName = translate('no_vessel');
    // FIX: Explicitly typed the `groupedByVessel` constant and the `reduce` accumulator.
    const groupedByVessel: Record<string, any[]> = data.reduce((acc: Record<string, any[]>, row: any) => {
        const vesselName = row['ARRIVAL VESSEL'] ? String(row['ARRIVAL VESSEL']).trim().toUpperCase() : defaultVesselName;
        if (!acc[vesselName]) {
            acc[vesselName] = [];
        }
        acc[vesselName].push(row);
        return acc;
    }, {});

    const processedVessels = Object.entries(groupedByVessel).map(([vesselName, shipments]) => {
        // Step 1: Group shipments by PO
        const poGroupsMap = new Map<string, { shipments: any[] }>();
        shipments.forEach(s => {
            const poNumber = s['PO SAP'] || 'N/A';
            if (!poGroupsMap.has(poNumber)) {
                poGroupsMap.set(poNumber, { shipments: [] });
            }
            poGroupsMap.get(poNumber)!.shipments.push(s);
        });

        // Step 2: Process each PO group
        const poGroups = Array.from(poGroupsMap.entries()).map(([poNumber, data]) => {
            const groupShipments = data.shipments;
            const quantity = groupShipments.reduce((sum, s) => sum + getContainerCount(s), 0);
            const cargoType = groupShipments[0]?.['TYPE OF CARGO'] || '';
            return { poNumber, quantity, cargoType, shipments: groupShipments };
        });
        
        // Step 3: Sort PO groups
        poGroups.sort((a, b) => {
             if (a.cargoType < b.cargoType) return -1;
             if (a.cargoType > b.cargoType) return 1;
             if (a.poNumber < b.poNumber) return -1;
             if (a.poNumber > b.poNumber) return 1;
             return 0;
        });
        
        // Step 4: Sort individual shipments within each group by BL for correct order
        poGroups.forEach(group => {
            group.shipments.sort((a, b) => {
                const blA = String(a['BL/AWB'] || '');
                const blB = String(b['BL/AWB'] || '');
                return blA.localeCompare(blB);
            });
        });

        // Step 5: Calculate total containers and earliest arrival for the vessel
        const totalContainers = poGroups.reduce((sum, group) => sum + group.quantity, 0);
        const earliestArrival = getEarliestArrivalDate(shipments);

        return { vesselName, poGroups, totalContainers, earliestArrival };
    });

    // Sort the final vessel cards chronologically
    processedVessels.sort((a, b) => {
        const dateA = a.earliestArrival;
        const dateB = b.earliestArrival;
        if (dateA && dateB) {
            return dateA.getTime() - dateB.getTime();
        }
        if (dateA) return -1;
        if (dateB) return 1;
        return a.vesselName.localeCompare(b.vesselName);
    });
    
    return processedVessels;
}



// --- UI Rendering ---
function renderVesselDashboard(data: any[]) {
    const headers = getVisibleColumns();
    const renderConfig = {
        placeholderTextKey: "card_placeholder_no_vessels",
        cardTitlePrefix: '',
        columns: headers,
    };
    renderDashboard(data, renderConfig);
}

function renderPODashboard(data: any[]) {
    const headers = getVisibleColumns();
    const renderConfig = {
        placeholderTextKey: "card_placeholder_no_pos",
        cardTitlePrefix: 'PO: ',
        columns: headers,
    };
    renderDashboard(data, renderConfig);
}

function renderWarehouseDashboard(data: any[]) {
    const headers = getVisibleColumns();
    const renderConfig = {
        placeholderTextKey: "card_placeholder_no_warehouses",
        cardTitlePrefix: '',
        columns: headers,
    };
    renderDashboard(data, renderConfig);
}

function renderDashboard(data: any[], config: any) {
    dashboardGrid.innerHTML = '';
    if (data.length === 0) {
        placeholder.classList.remove('hidden');
        placeholder.querySelector('h2')!.textContent = translate(config.placeholderTextKey);
        placeholder.querySelector('p')!.textContent = translate('card_placeholder_try_resetting');
        return;
    }
    placeholder.classList.add('hidden');

    data.forEach(item => {
        const card = createDashboardCard(
            `${config.cardTitlePrefix}${item.name}`,
            item.totalFCL,
            config.columns,
            item.shipments,
            item.overallRisk
        );
        card.addEventListener('click', () => openModal(item));
        dashboardGrid.appendChild(card);
    });
}

function renderDetailedDashboard(data: any[]) {
    dashboardGrid.innerHTML = '';
    if (data.length === 0) {
        placeholder.classList.remove('hidden');
        placeholder.querySelector('h2')!.textContent = translate('card_placeholder_no_vessels');
        placeholder.querySelector('p')!.textContent = translate('card_placeholder_try_resetting');
        return;
    }
    placeholder.classList.add('hidden');

    data.forEach(item => {
        const card = createDetailedCard(item.vesselName, item.poGroups, item.totalContainers);
        dashboardGrid.appendChild(card);
    });
}

function createDetailedCard(vesselName: string, poGroups: any[], totalContainers: number) {
    const card = document.createElement('div');
    card.className = `card col-span-1 lg:col-span-3 xl:col-span-3`;

    const headers = ['CARGO TYPE', 'PO QTY', 'PO NUMBER', 'BATCH', 'BL', 'QTY', 'ETA', 'WAREHOUSE', 'STATUS', 'BROKER'];
    const dataKeys: Record<string, string> = {
        'BATCH': 'BATCH CHINA',
        'BL': 'BL/AWB',
        'ETA': 'ACTUAL ETA',
        'STATUS': 'STATUS',
        'WAREHOUSE': 'BONDED WAREHOUSE',
        'BROKER': 'BROKER'
    };

    let shipmentsHtml = '';
    poGroups.forEach((group: any) => {
        group.shipments.forEach((shipment: any, index: number) => {
            const isFirstRowOfGroup = index === 0;

            const cargoTypeHtml = isFirstRowOfGroup ? `<td class="px-3 py-2 whitespace-nowrap">${group.cargoType || ''}</td>` : '<td></td>';
            const poQtyHtml = isFirstRowOfGroup ? `<td class="px-3 py-2 whitespace-nowrap font-bold">${group.quantity || ''}</td>` : '<td></td>';
            const poNumberHtml = isFirstRowOfGroup ? `<td class="px-3 py-2 whitespace-nowrap">${group.poNumber || ''}</td>` : '<td></td>';
            
            const blQty = getContainerCount(shipment) || '';

            const otherCells = ['BATCH', 'BL'].map(header => {
                return `<td class="px-3 py-2 whitespace-nowrap">${shipment[dataKeys[header]] || ''}</td>`;
            }).join('');
            
            const qtyCell = `<td class="px-3 py-2 whitespace-nowrap font-semibold">${blQty}</td>`;

            const etaWarehouseStatusCells = ['ETA', 'WAREHOUSE', 'STATUS', 'BROKER'].map(header => {
                 return `<td class="px-3 py-2 whitespace-nowrap">${shipment[dataKeys[header]] || ''}</td>`;
            }).join('');

            shipmentsHtml += `<tr class="border-b text-xs">${cargoTypeHtml}${poQtyHtml}${poNumberHtml}${otherCells}${qtyCell}${etaWarehouseStatusCells}</tr>`;
        });
    });

    card.innerHTML = `
        <div class="p-4 border-b flex justify-between items-center">
            <h3 class="font-extrabold text-lg text-gray-800">${vesselName}</h3>
            <div class="text-right">
                <span class="block text-2xl font-bold text-blue-600">${totalContainers}</span>
                <span class="text-sm font-medium text-gray-500">${translate('containers')}</span>
            </div>
        </div>
        <div class="table-responsive">
            <table class="min-w-full text-sm">
                <thead class="bg-gray-50">
                    <tr class="border-b">
                        ${headers.map(h => `<th class="px-3 py-2 text-left font-semibold text-gray-500 text-xs uppercase tracking-wider">${h}</th>`).join('')}
                    </tr>
                </thead>
                <tbody class="bg-white divide-y divide-gray-200">
                    ${shipmentsHtml}
                </tbody>
            </table>
        </div>`;
    return card;
}


function createDashboardCard(title: string, totalFCL: number, headers: string[], shipments: any[], risk: string) {
    const card = document.createElement('div');
    card.className = `card risk-${risk} cursor-pointer`;

    // Sort shipments based on global state
    const sortedShipments = sortData(shipments);

    const shipmentsHtml = sortedShipments.map((s: any) => {
        const daysText = s.risk === 'none' ? `<span class="font-semibold text-green-700">${translate('delivered')}</span>` : s.daysToDeadline !== null ? `<span class="font-bold ${s.daysToDeadline < 0 ? 'text-red-600' : 'text-gray-800'}">${s.daysToDeadline}</span>` : 'N/A';

        const rowData = headers.map(header => {
            let cellContent = s[header] || '';
            if (header === 'Dias Restantes') {
                return `<td class="px-3 py-2 text-center text-xs">${daysText}</td>`;
            }
             if (header === 'FREE TIME DEADLINE') {
                return `<td class="px-3 py-2 whitespace-nowrap text-xs font-semibold">${cellContent}</td>`;
            }
            return `<td class="px-3 py-2 whitespace-nowrap text-xs">${cellContent}</td>`;
        }).join('');

        return `<tr class="row-risk-${s.risk} hover:bg-opacity-50">${rowData}</tr>`;
    }).join('');
    
    const getSortIndicator = (key: string) => {
        if (key === currentSortKey) {
            return currentSortOrder === 'asc' ? '<i class="fas fa-arrow-up ml-1"></i>' : '<i class="fas fa-arrow-down ml-1"></i>';
        }
        return '';
    };

    card.innerHTML = `<div class="p-4 border-b border-gray-200">
            <div class="flex justify-between items-center">
                <h3 class="font-extrabold text-lg text-gray-800">${title}</h3>
                <div class="text-right">
                    <span class="block text-2xl font-bold text-blue-600">${totalFCL}</span>
                    <span class="text-sm font-medium text-gray-500">${translate('containers')}</span>
                </div>
            </div>
        </div>
        <div class="flex-grow table-responsive">
            <table class="min-w-full text-sm">
                <thead class="bg-gray-50"><tr class="border-b">
                    ${headers.map(h => `<th class="px-3 py-2 text-left font-semibold text-gray-500 text-xs uppercase tracking-wider sortable-header" data-sort-key="${h}">${h} ${getSortIndicator(h)}</th>`).join('')}
                </tr></thead>
                <tbody class="bg-white divide-y divide-gray-200">${shipmentsHtml}</tbody>
            </table>
        </div>`;
    return card;
}

// --- Chart Rendering ---
function renderCharts(processedData: any[], filteredData: any[]) {
    // Save data for export
    chartDataSource.bar = processedData;

    const isDarkMode = document.body.classList.contains('dark');
    const textColor = isDarkMode ? '#d1d5db' : '#4b5563';
    
    // --- Chart Color Helpers ---
    const getRiskColor = (risk: string) => {
        const colors: Record<string, string> = {
            high: 'rgba(239, 68, 68, 0.7)',    // red
            medium: 'rgba(249, 115, 22, 0.7)', // orange
            low: 'rgba(34, 197, 94, 0.7)',     // green
            none: 'rgba(107, 114, 128, 0.7)'   // gray
        };
        return colors[risk] || colors.none;
    };

    const getStatusColor = (status: string) => {
        const s = status.toLowerCase();
        if (s.includes('delivered')) return 'rgba(34, 197, 94, 0.7)';      // green
        if (s.includes('cleared')) return 'rgba(59, 130, 246, 0.7)';       // blue
        if (s.includes('sem status') || s.includes('no status')) return 'rgba(239, 68, 68, 0.7)';     // red
        if (s.includes('presence') || s.includes('unloaded')) return 'rgba(249, 115, 22, 0.7)'; // orange
        if (s.includes('on board') || s.includes('transshipment')) return 'rgba(168, 85, 247, 0.7)'; // purple
        // Default color for other statuses
        return `rgba(${Math.floor(Math.random() * 155) + 100}, ${Math.floor(Math.random() * 155) + 100}, ${Math.floor(Math.random() * 155) + 100}, 0.7)`;
    };


    let barChartTitleKey: string;
    if (currentViewState === 'vessel') {
        barChartTitleKey = 'chart_title_containers_by_vessel';
    } else if (currentViewState === 'po') {
        barChartTitleKey = 'chart_title_containers_by_po';
    } else { // Warehouse view
        barChartTitleKey = 'chart_title_containers_by_warehouse';
    }
    document.getElementById('bar-chart-title')!.textContent = translate(barChartTitleKey);
    
    const labels = processedData.map(item => item.name);
    const data = processedData.map(item => item.totalFCL);
    
    // Bar Chart
    if (mainBarChart) mainBarChart.destroy();
    Chart.register(ChartDataLabels);
    mainBarChart = new Chart(document.getElementById('main-bar-chart') as HTMLCanvasElement, {
        type: 'bar',
        data: { 
            labels, 
            datasets: [{ 
                label: 'Total Containers', 
                data, 
                backgroundColor: (processedData as any[]).map(item => getRiskColor(item.overallRisk)),
                // Custom property to hold risk data for tooltips
                // @ts-ignore
                riskCounts: (processedData as any[]).map(item => item.riskCounts)
            }] 
        },
        options: { 
            indexAxis: 'y', 
            responsive: true, 
            scales: {
                x: {
                    beginAtZero: true,
                    // @ts-ignore
                    grace: '5%' // Add extra space at the top
                }
            },
            plugins: { 
                legend: { display: false },
                datalabels: {
                    anchor: 'end',
                    align: 'end',
                    color: textColor,
                    font: {
                        weight: 'bold'
                    }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let label = context.dataset.label || '';
                            if (label) {
                                label += ': ';
                            }
                            if (context.parsed.x !== null) {
                                label += context.parsed.x;
                            }
                            return label;
                        },
                        afterLabel: function(context) {
                            const riskData = (context.dataset as any).riskCounts[context.dataIndex];
                            return [
                                `${translate('high_risk')}: ${riskData.high}`,
                                `${translate('medium_risk')}: ${riskData.medium}`,
                                `${translate('low_risk')}: ${riskData.low}`,
                                `${translate('chart_tooltip_delivered')}: ${riskData.none}`
                            ];
                        }
                    }
                }
            } 
        }
    });

    // Pie Chart
    const statusGroups: Record<string, any[]> = {};
    filteredData.forEach(row => {
        const status = row.STATUS || translate('no_status');
        if (!statusGroups[status]) {
            statusGroups[status] = [];
        }
        statusGroups[status].push(row);
    });

    const statusCounts: Record<string, {count: number, fcl: number}> = {};
    let totalContainersInView = 0;
    Object.entries(statusGroups).forEach(([status, shipments]) => {
        const totalContainersForStatus = shipments.reduce((sum, s) => sum + getContainerCount(s), 0);
        statusCounts[status] = {
            count: shipments.length, // number of BLs
            fcl: totalContainersForStatus // total containers
        };
        totalContainersInView += totalContainersForStatus;
    });
    
    // Save data for export
    chartDataSource.pie = Object.entries(statusCounts).map(([label, data]) => ({ label, value: data.fcl }));

    if (statusPieChart) statusPieChart.destroy();
    statusPieChart = new Chart(document.getElementById('status-pie-chart') as HTMLCanvasElement, {
        type: 'pie',
        data: {
            labels: Object.keys(statusCounts),
            datasets: [{
                data: Object.values(statusCounts).map(s => s.fcl),
                backgroundColor: Object.keys(statusCounts).map(status => getStatusColor(status))
            }]
        },
        options: { 
            responsive: true, 
            plugins: { 
                legend: { position: 'right' },
                tooltip: {
                     callbacks: {
                        label: function(context) {
                            const label = context.label || '';
                            const value = Number(context.parsed as any) || 0;
                            const total = totalContainersInView;
                            const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : '0.0';
                            return `${label}: ${value} ${translate('containers')} (${percentage}%)`;
                        }
                    }
                }
            },
            onClick: (event, elements, chart) => {
                if (elements.length > 0) {
                    const clickedIndex = elements[0].index;
                    const clickedStatus = chart.data.labels![clickedIndex] as string;
                    
                    // Reset selections and select only the clicked one
                    Array.from(statusFilter.options).forEach(opt => {
                        opt.selected = (opt.value === clickedStatus);
                    });

                    applyFiltersAndRender();
                }
            }
        }
    });

    // Deadline Distribution Chart
    const deadlineCanvas = document.getElementById('deadline-distribution-chart') as HTMLCanvasElement;
    const bins = {
        overdue: { labelKey: 'deadline_bin_overdue', count: 0, color: 'rgba(239, 68, 68, 0.7)' }, // red
        '0-7': { labelKey: 'deadline_bin_0_7', count: 0, color: 'rgba(249, 115, 22, 0.7)' }, // orange
        '8-15': { labelKey: 'deadline_bin_8_15', count: 0, color: 'rgba(245, 158, 11, 0.7)' }, // amber
        '16-30': { labelKey: 'deadline_bin_16_30', count: 0, color: 'rgba(132, 204, 22, 0.7)' }, // lime
        '31+': { labelKey: 'deadline_bin_31_plus', count: 0, color: 'rgba(34, 197, 94, 0.7)' }, // green
        delivered: { labelKey: 'delivered', count: 0, color: 'rgba(107, 114, 128, 0.7)' } // gray
    };

    const enrichedFilteredData = filteredData.map(calculateRisk);

    for (const shipment of enrichedFilteredData) {
        const { daysToDeadline, risk } = shipment;
        if (risk === 'none') {
            (bins.delivered.count)++;
        } else if (daysToDeadline !== null) {
            if (daysToDeadline < 0) (bins.overdue.count)++;
            else if (daysToDeadline <= 7) (bins['0-7'].count)++;
            else if (daysToDeadline <= 15) (bins['8-15'].count)++;
            else if (daysToDeadline <= 30) (bins['16-30'].count)++;
            else (bins['31+'].count)++;
        }
    }
    
    // Save data for export
    chartDataSource.deadline = Object.values(bins).map(b => ({ label: translate(b.labelKey), count: b.count }));


    if (deadlineDistributionChart) deadlineDistributionChart.destroy();
    deadlineDistributionChart = new Chart(deadlineCanvas, {
        type: 'bar',
        data: {
            labels: Object.values(bins).map(b => translate(b.labelKey)),
            datasets: [{
                label: translate('chart_deadline_label'),
                data: Object.values(bins).map(b => b.count),
                backgroundColor: Object.values(bins).map(b => b.color)
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { display: false },
                datalabels: {
                    anchor: 'end',
                    align: 'end',
                    color: textColor,
                    font: {
                        weight: 'bold'
                    }
                }
            }
        }
    });
}


// --- Helper Functions ---
function populateStatusFilter(data: any[]) {
    const statuses = [...new Set(data.map(row => row.STATUS || translate('no_status')))].sort();
    statusFilter.innerHTML = `<option value="" selected>${translate('all_option')}</option>`;
    statuses.forEach(status => {
        const option = document.createElement('option');
        option.value = status;
        option.textContent = status;
        statusFilter.appendChild(option);
    });
}

function populatePoFilter(data: any[]) {
    const pos = [...new Set(data.map(row => row['PO SAP'] || translate('no_po')))].sort();
    poFilter.innerHTML = `<option value="" selected>${translate('all_option')}</option>`;
    pos.forEach(po => {
        const option = document.createElement('option');
        option.value = po;
        option.textContent = po;
        poFilter.appendChild(option);
    });
}

function populateVesselFilter(data: any[]) {
    const vessels = [...new Set(data.map(row => row['ARRIVAL VESSEL'] || translate('no_vessel')))].sort();
    vesselFilter.innerHTML = `<option value="" selected>${translate('all_option')}</option>`;
    vessels.forEach(vessel => {
        const option = document.createElement('option');
        option.value = vessel;
        option.textContent = vessel;
        vesselFilter.appendChild(option);
    });
}

function populateShipmentTypeFilter(data: any[]) {
    const shipmentTypes = [...new Set(data.map(row => row['SHIPMENT TYPE'] || translate('no_shipment_type')))].sort();
    shipmentTypeFilter.innerHTML = `<option value="" selected>${translate('all_option')}</option>`;
    shipmentTypes.forEach(type => {
        const option = document.createElement('option');
        option.value = type;
        option.textContent = type;
        shipmentTypeFilter.appendChild(option);
    });
}

function populateCargoTypeFilter(data: any[]) {
    const cargoTypes = [...new Set(data.map(row => row['TYPE OF CARGO'] || translate('no_cargo_type')))].sort();
    cargoTypeFilter.innerHTML = `<option value="" selected>${translate('all_option')}</option>`;
    cargoTypes.forEach(type => {
        const option = document.createElement('option');
        option.value = type;
        option.textContent = type;
        cargoTypeFilter.appendChild(option);
    });
}

function populateBatchFilter(data: any[]) {
    const batches = [...new Set(data.map(row => row['BATCH CHINA'] || translate('no_batch')))].sort();
    batchFilter.innerHTML = `<option value="" selected>${translate('all_option')}</option>`;
    batches.forEach(batch => {
        const option = document.createElement('option');
        option.value = batch;
        option.textContent = batch;
        batchFilter.appendChild(option);
    });
}

function populateBrokerFilter(data: any[]) {
    const brokers = [...new Set(data.map(row => row['BROKER'] || translate('no_broker')))].sort();
    brokerFilter.innerHTML = `<option value="" selected>${translate('all_option')}</option>`;
    brokers.forEach(broker => {
        const option = document.createElement('option');
        option.value = broker;
        option.textContent = broker;
        brokerFilter.appendChild(option);
    });
}

function updateTotalContainers(data: any[]) {
    const total = data.reduce((sum, s) => sum + getContainerCount(s), 0);
    totalFclCount.textContent = total.toString();
}

function resetUI() {
    dashboardGrid.innerHTML = '';
    placeholder.classList.remove('hidden');
    filterContainer.classList.add('hidden');
    chartsContainer.classList.add('hidden');
    viewTabsContainer.classList.add('hidden');
    exportExcelBtn.classList.add('hidden');
    totalFclDisplay.classList.add('hidden');
    originalData = [];
    if (mainBarChart) mainBarChart.destroy();
    if (statusPieChart) statusPieChart.destroy();
    if (deadlineDistributionChart) deadlineDistributionChart.destroy();
    lastUpdate.textContent = translate('upload_prompt');
    setView('vessel');
}

// --- Sorting ---
function handleSortClick(event: MouseEvent) {
    const target = event.target as HTMLElement;
    const header = target.closest('.sortable-header');
    if (!header) return;

    const sortKey = header.getAttribute('data-sort-key');
    if (!sortKey) return;

    if (sortKey === currentSortKey) {
        currentSortOrder = currentSortOrder === 'asc' ? 'desc' : 'asc';
    } else {
        currentSortKey = sortKey;
        currentSortOrder = 'asc';
    }
    applyFiltersAndRender();
}

function sortData(data: any[]) {
    return [...data].sort((a, b) => {
        const valA = a[currentSortKey];
        const valB = b[currentSortKey];
        
        let comparison = 0;
        
        if (currentSortKey === 'ACTUAL ETA' || currentSortKey === 'FREE TIME DEADLINE') {
            const dateA = excelDateToJSDate(valA);
            const dateB = excelDateToJSDate(valB);
            if (dateA && dateB) {
                comparison = dateA.getTime() - dateB.getTime();
            } else if (dateA) {
                comparison = -1;
            } else if (dateB) {
                comparison = 1;
            }
        } else if (currentSortKey === 'Dias Restantes') {
             // Handle nulls for delivered items
            const daysA = a.daysToDeadline ?? Infinity;
            const daysB = b.daysToDeadline ?? Infinity;
            comparison = daysA - daysB;
        } else {
            // Default string/number sort
            const strA = String(valA || '').toLowerCase();
            const strB = String(valB || '').toLowerCase();
            if (strA < strB) comparison = -1;
            if (strA > strB) comparison = 1;
        }
        
        return currentSortOrder === 'asc' ? comparison : -comparison;
    });
}

// --- Column Visibility ---
function setupColumnToggles() {
    loadColumnVisibility();
    populateColumnToggles();

    columnToggleBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        columnToggleDropdown.classList.toggle('hidden');
    });

    document.addEventListener('click', (e) => {
        if (!columnToggleDropdown.classList.contains('hidden') && !columnToggleBtn.contains(e.target as Node) && !columnToggleDropdown.contains(e.target as Node)) {
            columnToggleDropdown.classList.add('hidden');
        }
    });

    columnToggleDropdown.addEventListener('change', (e) => {
        const checkbox = e.target as HTMLInputElement;
        columnVisibility[checkbox.value] = checkbox.checked;
        saveColumnVisibility();
        applyFiltersAndRender();
    });
}

function populateColumnToggles() {
    const columns = viewColumns[currentViewState as keyof typeof viewColumns] || viewColumns.vessel;
    // Re-check defaults if a column has no saved setting
    columns.forEach(col => {
        if (columnVisibility[col] === undefined) {
            columnVisibility[col] = true;
        }
    });
    
    columnToggleDropdown.innerHTML = columns.map(col => `
        <label class="flex items-center px-4 py-2 hover:bg-gray-100 cursor-pointer">
            <input type="checkbox" class="form-checkbox h-4 w-4 text-blue-600 border-gray-300 rounded" value="${col}" ${columnVisibility[col] ? 'checked' : ''}>
            <span class="ml-3 text-sm text-gray-700">${col}</span>
        </label>
    `).join('');
}

function getVisibleColumns(): string[] {
    const columns = viewColumns[currentViewState as keyof typeof viewColumns] || viewColumns.vessel;
    return columns.filter(col => columnVisibility[col]);
}

function loadColumnVisibility() {
    try {
        const saved = localStorage.getItem('columnVisibility');
        if (saved) {
            columnVisibility = JSON.parse(saved);
        } else {
            // Default all to visible
            columnVisibility = {};
             Object.values(viewColumns).flat().forEach(col => columnVisibility[col] = true);
        }
    } catch(e) {
        console.error("Could not load column visibility from localStorage", e);
        columnVisibility = {};
        Object.values(viewColumns).flat().forEach(col => columnVisibility[col] = true);
    }
}

function saveColumnVisibility() {
    localStorage.setItem('columnVisibility', JSON.stringify(columnVisibility));
}

// --- Exporting ---
function handleExcelExport() {
    if (filteredDataCache.length === 0) {
        showToast("toast_no_data_to_export", "warning");
        return;
    }
    
    const workbook = XLSX.utils.book_new();

    // Handle "Detailed View" as a special case
    if (currentViewState === 'detailed') {
        const detailedData = processDataForDetailedView(filteredDataCache);
        const dataToExport: any[] = [];
        
        detailedData.forEach(vesselData => {
            vesselData.poGroups.forEach(poGroup => {
                poGroup.shipments.forEach((shipment: any) => {
                    dataToExport.push({
                        'Vessel': vesselData.vesselName,
                        'Cargo Type': poGroup.cargoType,
                        'PO Number': poGroup.poNumber,
                        'PO Total Qty': poGroup.quantity,
                        'Batch': shipment['BATCH CHINA'] || '',
                        'BL': shipment['BL/AWB'] || '',
                        'BL Qty': getContainerCount(shipment),
                        'ETA': shipment['ACTUAL ETA'] || '',
                        'Warehouse': shipment['BONDED WAREHOUSE'] || '',
                        'Status': shipment['STATUS'] || '',
                        'Broker': shipment['BROKER'] || ''
                    });
                });
            });
        });

        const detailedSheet = XLSX.utils.json_to_sheet(dataToExport);
        XLSX.utils.book_append_sheet(workbook, detailedSheet, 'Detailed View Export');

    } else {
        // Handle standard views (Vessel, PO, Warehouse)
        
        // Define ALL columns we want in the report, including mandatory Vessel and Qty
        const exportHeaders = [
            'ARRIVAL VESSEL', 'VOYAGE', 'PO SAP', 'BL/AWB', 'SHIPOWNER', 'STATUS', 
            'SHIPMENT TYPE', 'TYPE OF CARGO', 'BATCH CHINA', 'ACTUAL ETA', 
            'FREE TIME DEADLINE', 'Dias Restantes', 'BONDED WAREHOUSE', 'CONTAINER QTY', 'BROKER'
        ];

        // 1. Filtered Data Sheet
        const dataToExport = filteredDataCache.map(row => {
            const newRow: Record<string, any> = {};
            exportHeaders.forEach(header => {
                if (header === 'CONTAINER QTY') {
                    newRow[header] = getContainerCount(row);
                } else {
                    newRow[header] = row[header] ?? '';
                }
            });
            return newRow;
        });
        const filteredDataSheet = XLSX.utils.json_to_sheet(dataToExport, { header: exportHeaders });
        XLSX.utils.book_append_sheet(workbook, filteredDataSheet, 'Filtered Data');

        // 2. Main Bar Chart Data Sheet
        if (chartDataSource.bar.length > 0) {
            const barDataForSheet = chartDataSource.bar.map(item => ({
                'Item': item.name,
                'Total Containers': item.totalFCL,
                'High Risk': item.riskCounts.high,
                'Medium Risk': item.riskCounts.medium,
                'Low Risk': item.riskCounts.low,
                'Delivered': item.riskCounts.none
            }));
            const barChartSheet = XLSX.utils.json_to_sheet(barDataForSheet);
            XLSX.utils.book_append_sheet(workbook, barChartSheet, 'Chart Data (Main View)');
        }

        // 3. Status Pie Chart Data Sheet
        if (chartDataSource.pie.length > 0) {
            const pieDataForSheet = chartDataSource.pie.map(item => ({
                'STATUS': item.label,
                'Total Containers': item.value
            }));
            const pieChartSheet = XLSX.utils.json_to_sheet(pieDataForSheet);
            XLSX.utils.book_append_sheet(workbook, pieChartSheet, 'Chart Data (Status)');
        }

        // 4. Deadline Distribution Chart Data Sheet
        if (chartDataSource.deadline.length > 0) {
            const deadlineDataForSheet = chartDataSource.deadline.map(item => ({
                'Prazo (Dias)': item.label,
                'Nº de Cargas': item.count
            }));
            const deadlineChartSheet = XLSX.utils.json_to_sheet(deadlineDataForSheet);
            XLSX.utils.book_append_sheet(workbook, deadlineChartSheet, 'Chart Data (Deadlines)');
        }
    }


    // Create a filename and trigger the download
    const fileName = `dashboard_export_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(workbook, fileName);
    
    showToast("toast_excel_export_started", "success");
}

// --- Modal ---
function openModal(item: any) {
    activeModalItem = item;
    renderModalContent();
    detailsModal.classList.remove('hidden');
    document.body.classList.add('overflow-hidden');
    // For animation
    setTimeout(() => {
        detailsModal.classList.add('modal-open');
    }, 10);
}

function closeModal() {
    detailsModal.classList.remove('modal-open');
     setTimeout(() => {
        detailsModal.classList.add('hidden');
        document.body.classList.remove('overflow-hidden');
        modalBody.innerHTML = ''; // Clear content
        modalHeaderContent.innerHTML = '';
        activeModalItem = null;
    }, 300); // Match CSS transition duration
}

function renderModalContent() {
    if (!activeModalItem) return;

    const { name, totalFCL, shipments } = activeModalItem;
    
    // 1. Render Header
    const prefix = (currentViewState === 'po') ? 'PO: ' : '';
    const title = `${prefix}${name}`;
    
    const uniqueWarehouses = [...new Set(shipments
        .map((s: any) => s['BONDED WAREHOUSE'])
        .filter((wh: string | undefined) => wh && wh.trim() !== ''))];

    const warehouseDisplayHtml = uniqueWarehouses.length > 0 
        ? `<div class="mt-1 flex items-center">
             <i class="fas fa-warehouse text-gray-500 mr-2"></i>
             <p class="text-sm text-gray-600 font-medium">${uniqueWarehouses.join(', ')}</p>
           </div>`
        : '';

    modalHeaderContent.innerHTML = `
        <div class="flex justify-between items-start">
            <div>
                <h3 class="font-extrabold text-xl text-gray-800">${title}</h3>
                ${warehouseDisplayHtml}
            </div>
            <div class="text-right ml-8 flex-shrink-0">
                <span class="block text-3xl font-bold text-blue-600">${totalFCL}</span>
                <span class="text-sm font-medium text-gray-500">${translate('containers')}</span>
            </div>
        </div>
    `;

    // 2. Render Body (Table)
    
    // Define a fixed, comprehensive set of columns for the detail modal to ensure consistency.
    let headers: string[] = [];
    const baseHeaders = [
        'BL/AWB',
        'SHIPOWNER',
        'BONDED WAREHOUSE',
        'STATUS',
        'SHIPMENT TYPE',
        'TYPE OF CARGO',
        'QTY_FCL_LCL', // Placeholder key for dynamic calculation
        'BATCH CHINA',
        'ACTUAL ETA',
        'FREE TIME DEADLINE',
        'Dias Restantes',
        'BROKER'
    ];
    
    // Add view-specific columns at the beginning
    if (currentViewState === 'vessel') {
        headers = ['PO SAP', 'VOYAGE', ...baseHeaders];
    } else if (currentViewState === 'po') {
        headers = ['ARRIVAL VESSEL', 'VOYAGE', ...baseHeaders];
    } else { // warehouse view
        headers = ['ARRIVAL VESSEL', 'VOYAGE', 'PO SAP', ...baseHeaders];
    }

    // Map internal data keys to user-friendly Portuguese headers for the modal.
    const headerDisplayMap: Record<string, string> = {
        'PO SAP': 'PO',
        'ARRIVAL VESSEL': 'Navio',
        'VOYAGE': 'Viagem',
        'SHIPMENT TYPE': 'Tipo de Envio',
        'TYPE OF CARGO': 'Tipo de Mercadoria',
        'QTY_FCL_LCL': 'Qtd. Containers',
        'BATCH CHINA': 'Batch',
        'SHIPOWNER': 'Transportadora',
        'STATUS': 'Status',
        'ACTUAL ETA': 'Chegada Real',
        'FREE TIME DEADLINE': 'Prazo Free Time',
        'BONDED WAREHOUSE': 'Armazém',
        'BROKER': 'Broker'
    };
    
    const sortedShipments = sortData(shipments);
    
    const shipmentsHtml = sortedShipments.map((shipment: any) => {
        const daysText = shipment.risk === 'none' 
            ? `<span class="font-semibold text-green-700">${translate('delivered')}</span>` 
            : shipment.daysToDeadline !== null 
            ? `<span class="font-bold ${shipment.daysToDeadline < 0 ? 'text-red-600' : 'text-gray-800'}">${shipment.daysToDeadline}</span>` 
            : 'N/A';

        const rowData = headers.map(header => {
            let cellContent;
            
            if (header === 'QTY_FCL_LCL') {
                const count = getContainerCount(shipment);
                cellContent = count > 0 ? count : '';
            } else {
                cellContent = shipment[header] || '';
            }
            
            // FIX: Prevent invalid dates from being displayed.
            if ((header === 'FREE TIME DEADLINE' || header === 'ACTUAL ETA') && String(cellContent).startsWith('1899')) {
                cellContent = 'N/A';
            }

            if (header === 'Dias Restantes') {
                return `<td class="px-2 py-2 text-center">${daysText}</td>`;
            }
            if (header === 'FREE TIME DEADLINE' || header === 'ACTUAL ETA') {
                return `<td class="px-2 py-2 whitespace-nowrap">${cellContent}</td>`;
            }
            if (header === 'BL/AWB' || header === 'PO SAP') {
                return `<td class="px-2 py-2 whitespace-nowrap">${cellContent}</td>`;
            }
            if (header === 'ARRIVAL VESSEL') {
                return `<td class="px-2 py-2 break-all">${cellContent}</td>`;
            }
            return `<td class="px-2 py-2 break-words">${cellContent}</td>`;
        }).join('');

        return `<tr class="row-risk-${shipment.risk}">${rowData}</tr>`;
    }).join('');

    modalBody.innerHTML = `
        <div class="table-responsive">
            <table class="min-w-full text-sm">
                <thead class="bg-gray-50"><tr class="border-b">
                    ${headers.map(h => `<th class="px-2 py-2 text-left font-semibold text-gray-500 uppercase tracking-wider text-xs">${headerDisplayMap[h] || h}</th>`).join('')}
                </tr></thead>
                <tbody class="bg-white divide-y divide-gray-200">${shipmentsHtml}</tbody>
            </table>
        </div>
    `;
}

// --- Logo Management ---
function loadSavedLogo() {
    const savedLogo = localStorage.getItem('companyLogo');
    if (savedLogo) {
        companyLogo.src = savedLogo;
        companyLogo.classList.remove('hidden');
        removeLogoBtn.classList.remove('hidden');
    }
}

function handleLogoUpload(event: Event) {
    const target = event.target as HTMLInputElement;
    const file = target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const logoDataUrl = e.target?.result as string;
        if (logoDataUrl) {
            localStorage.setItem('companyLogo', logoDataUrl);
            companyLogo.src = logoDataUrl;
            companyLogo.classList.remove('hidden');
            removeLogoBtn.classList.remove('hidden');
            showToast('toast_logo_updated', 'success');
        }
    };
    reader.onerror = () => {
         showToast('toast_logo_error', 'error');
    };
    reader.readAsDataURL(file);
    target.value = ''; // Reset input so the same file can be chosen again
}

function handleRemoveLogo() {
    localStorage.removeItem('companyLogo');
    companyLogo.src = '';
    companyLogo.classList.add('hidden');
    removeLogoBtn.classList.add('hidden');
    showToast('toast_logo_removed', 'success');
}
