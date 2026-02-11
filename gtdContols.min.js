// ============================================================================
// DataGrid Class - AG Grid Implementation with Bootstrap 5
// Supports Multiple Instances
// ============================================================================

/**
 * DataGrid class that wraps AG Grid for use in Bootstrap 5 applications
 * Supports multiple instances on a single page
 * Designed for integration with Microsoft Dynamics 365 portals and Power Automate flows
 * 
 * Action handling is delegated to a developer-defined GridActionHandler function
 * that receives the action name, selected row(s) data, and grid instance for custom processing.
 * 
 * Actions are configured with selectionMode to control when they're enabled:
 * - 'single': Only enabled when exactly one row is selected
 * - 'multiple': Only enabled when multiple rows are selected
 * - 'any': Enabled when one or more rows are selected (default)
 * 
 * class DataGrid
 * example
 * // Define GridActionHandler function
 * function GridActionHandler(actionName, selectedRows, gridInstance) {
 *   // gridInstance.instanceName can identify which grid triggered the action
 *   console.log('Grid:', gridInstance.instanceName);
 *   switch(actionName) {
 *     case 'view':
 *       console.log('View:', selectedRows);
 *       break;
 *     case 'edit':
 *       console.log('Edit:', selectedRows);
 *       break;
 *   }
 * }
 * 
 * const columnDefs = [
 *   { field: 'make', sortable: true, filter: true },
 *   { field: 'model', sortable: true, filter: true },
 *   { field: 'price', sortable: true, filter: true }
 * ];
 * 
 * const actions = [
 *   { name: 'edit', text: 'Edit Record', selectionMode: 'single' },
 *   { name: 'delete', text: 'Delete Records', selectionMode: 'any' },
 *   { name: 'bulkUpdate', text: 'Bulk Update', selectionMode: 'multiple' }
 * ];
 * 
 * const grid = new DataGrid('myGrid', 'gridContainer', columnDefs, actions, 'https://flow.url', 'record123', GridActionHandler);
 */
class DataGrid {
  /**
   * Creates a new DataGrid instance
   * @param {string|Object} instanceName - Unique name for this grid instance OR options object with named parameters
   * @param {string} containerId - ID of the DIV element where the grid will be placed (if using positional params)
   * @param {Array} columnDefs - AG Grid column definitions array (if using positional params)
   * @param {Array<{name: string, text: string, selectionMode?: string}>} actions - Array of action definitions with optional selectionMode ('single', 'multiple', or 'any')
   * @param {string} flowUrl - URL to Power Automate Flow (not used in local development)
   * @param {string} recordId - Database row ID to pass to the flow
   * @param {Function} actionHandler - Optional function to handle grid actions (receives actionName, selectedRows array, and gridInstance)
   * @param {Object} fieldMapping - Optional mapping object to convert Dataverse field names to display field names (e.g., {'make': 'cr123_vehiclemake', 'model': 'cr123_vehiclemodel'})
   */
  constructor(instanceName, containerId, columnDefs, actions = [], flowUrl = '', recordId = '', actionHandler = null, fieldMapping = null) {
    // Support both positional parameters and named parameters via options object
    if (typeof instanceName === 'object') {
      const options = instanceName;
      this.instanceName = options.instanceName;
      this.containerId = options.containerId;
      this.columnDefs = options.columnDefs;
      this.actions = options.actions || [];
      this.flowUrl = options.flowUrl || '';
      this.recordId = options.recordId || '';
      this.actionHandler = options.actionHandler || null;
      this.fieldMapping = options.fieldMapping || null;
    } else {
      // Legacy positional parameters
      this.instanceName = instanceName;
      this.containerId = containerId;
      this.columnDefs = columnDefs;
      this.actions = actions;
      this.flowUrl = flowUrl;
      this.recordId = recordId;
      this.actionHandler = actionHandler;
      this.fieldMapping = fieldMapping;
    }
    this.gridApi = null;
    this.gridColumnApi = null;
    
    // Pagination state
    this.currentPage = 1;
    this.pageSize = 10;
    this.totalRecords = 0;
    this.totalPages = 0;
    this.pageToken = null; // For production use with fetchDocuments
    
    // Generate unique instance ID
    this.instanceId = `ag-grid-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
    
    // Make this instance globally accessible
    window[instanceName] = this;
    
    console.log(`=== DataGrid Instance Created: ${instanceName} ===`);
  }
  
  /**
   * Creates the action toolbar HTML above the grid
   * @returns {string} HTML string for action toolbar
   * @private
   */
  createActionToolbar() {
    if (!this.actions || this.actions.length === 0) {
      return '';
    }
    
    const toolbarHtml = `
      <div class="card-header bg-light border-bottom">
        <div class="btn-group">
          <button type="button" class="btn btn-sm btn-primary dropdown-toggle" 
                  id="${this.instanceId}-actions-btn" 
                  data-bs-toggle="dropdown" 
                  aria-expanded="false" 
                  aria-disabled="true">
            Actions
          </button>
          <ul class="dropdown-menu" id="${this.instanceId}-actions-menu">
            ${this.actions.map(action => `
              <li>
                <a class="dropdown-item" href="#" 
                   data-action="${action.name}" 
                   data-selection-mode="${action.selectionMode || 'any'}">
                  ${action.text}
                </a>
              </li>
            `).join('')}
          </ul>
        </div>
      </div>
    `;
    
    return toolbarHtml;
  }
  
  /**
   * Updates the action button state based on current selection
   * @private
   */
  updateActionButtonState() {
    const actionBtn = document.getElementById(`${this.instanceId}-actions-btn`);
    if (!actionBtn) return;
    
    const selectedRows = this.getSelectedRow();
    const selectionCount = selectedRows.length;
    
    // Enable/disable the actions button using custom attribute and styling
    if (selectionCount === 0) {
      actionBtn.setAttribute('data-disabled', 'true');
      actionBtn.classList.add('disabled');
      actionBtn.style.opacity = '0.65';
      actionBtn.style.cursor = 'not-allowed';
      actionBtn.textContent = 'Actions';
    } else {
      actionBtn.removeAttribute('data-disabled');
      actionBtn.classList.remove('disabled');
      actionBtn.style.opacity = '1';
      actionBtn.style.cursor = 'pointer';
      actionBtn.textContent = `Actions (${selectionCount} selected)`;
    }
    
    // Update individual action items based on selectionMode
    const actionMenu = document.getElementById(`${this.instanceId}-actions-menu`);
    if (actionMenu) {
      actionMenu.querySelectorAll('[data-action]').forEach(item => {
        const selectionMode = item.getAttribute('data-selection-mode');
        let shouldEnable = true;
        
        if (selectionMode === 'single' && selectionCount !== 1) {
          shouldEnable = false;
        } else if (selectionMode === 'multiple' && selectionCount < 2) {
          shouldEnable = false;
        } else if (selectionMode === 'any' && selectionCount === 0) {
          shouldEnable = false;
        }
        
        if (shouldEnable) {
          item.classList.remove('disabled');
          item.style.opacity = '1';
        } else {
          item.classList.add('disabled');
          item.style.opacity = '0.5';
        }
      });
    }
  }
  
  /**
   * Creates the pagination controls HTML below the grid
   * @returns {string} HTML string for pagination controls
   * @private
   */
  createPaginationControls() {
    const paginationHtml = `
      <div class="card-footer bg-light border-top" id="${this.instanceId}-pagination">
        <div class="row align-items-center">
          <div class="col-md-3">
            <div class="input-group input-group-sm">
              <label class="input-group-text" for="${this.instanceId}-page-size">Page Size:</label>
              <select class="form-select" id="${this.instanceId}-page-size">
                <option value="5">5</option>
                <option value="10" selected>10</option>
                <option value="25">25</option>
                <option value="50">50</option>
                <option value="100">100</option>
              </select>
            </div>
          </div>
          <div class="col-md-6 text-center">
            <div class="btn-group" role="group" id="${this.instanceId}-page-buttons">
              <button type="button" class="btn btn-sm btn-outline-primary" id="${this.instanceId}-first-page"
                      title="First Page">&laquo;</button>
              <button type="button" class="btn btn-sm btn-outline-primary" id="${this.instanceId}-prev-page"
                      title="Previous Page">&lsaquo;</button>
              <button type="button" class="btn btn-sm btn-outline-secondary" disabled id="${this.instanceId}-page-info">Page 1 of 1</button>
              <button type="button" class="btn btn-sm btn-outline-primary" id="${this.instanceId}-next-page"
                      title="Next Page">&rsaquo;</button>
              <button type="button" class="btn btn-sm btn-outline-primary" id="${this.instanceId}-last-page"
                      title="Last Page">&raquo;</button>
            </div>
          </div>
          <div class="col-md-3 text-end">
            <small class="text-muted" id="${this.instanceId}-record-info">0 records</small>
          </div>
        </div>
      </div>
    `;
    return paginationHtml;
  }
  
  /**
   * Updates the pagination controls based on current state
   * @private
   */
  updatePaginationControls() {
    const firstBtn = document.getElementById(`${this.instanceId}-first-page`);
    const prevBtn = document.getElementById(`${this.instanceId}-prev-page`);
    const nextBtn = document.getElementById(`${this.instanceId}-next-page`);
    const lastBtn = document.getElementById(`${this.instanceId}-last-page`);
    const pageInfo = document.getElementById(`${this.instanceId}-page-info`);
    const recordInfo = document.getElementById(`${this.instanceId}-record-info`);
    const pageSizeSelect = document.getElementById(`${this.instanceId}-page-size`);
    
    if (!firstBtn || !prevBtn || !nextBtn || !lastBtn || !pageInfo || !recordInfo) return;
    
    // Update page info text
    pageInfo.textContent = `Page ${this.currentPage} of ${this.totalPages}`;
    
    // Update record info
    const startRecord = this.totalRecords === 0 ? 0 : (this.currentPage - 1) * this.pageSize + 1;
    const endRecord = Math.min(this.currentPage * this.pageSize, this.totalRecords);
    recordInfo.textContent = `${startRecord}-${endRecord} of ${this.totalRecords} records`;
    
    // Update page size selector
    if (pageSizeSelect) {
      pageSizeSelect.value = this.pageSize.toString();
    }
    
    // Enable/disable buttons based on current page
    const isFirstPage = this.currentPage === 1;
    const isLastPage = this.currentPage >= this.totalPages;
    const onlyOnePage = this.totalPages <= 1;
    
    firstBtn.disabled = isFirstPage || onlyOnePage;
    prevBtn.disabled = isFirstPage || onlyOnePage;
    nextBtn.disabled = isLastPage || onlyOnePage;
    lastBtn.disabled = isLastPage || onlyOnePage;
    
    // Disable page size selector if no data
    if (pageSizeSelect) {
      pageSizeSelect.disabled = this.totalRecords === 0;
    }
  }
  
  /**
   * Binds event handlers for pagination controls
   * @private
   */
  bindPaginationHandlers() {
    const firstBtn = document.getElementById(`${this.instanceId}-first-page`);
    const prevBtn = document.getElementById(`${this.instanceId}-prev-page`);
    const nextBtn = document.getElementById(`${this.instanceId}-next-page`);
    const lastBtn = document.getElementById(`${this.instanceId}-last-page`);
    const pageSizeSelect = document.getElementById(`${this.instanceId}-page-size`);
    
    if (firstBtn) {
      firstBtn.addEventListener('click', () => this.goToPage(1));
    }
    
    if (prevBtn) {
      prevBtn.addEventListener('click', () => this.goToPage(this.currentPage - 1));
    }
    
    if (nextBtn) {
      nextBtn.addEventListener('click', () => this.goToPage(this.currentPage + 1));
    }
    
    if (lastBtn) {
      lastBtn.addEventListener('click', () => this.goToPage(this.totalPages));
    }
    
    if (pageSizeSelect) {
      pageSizeSelect.addEventListener('change', (e) => {
        this.pageSize = parseInt(e.target.value, 10);
        this.currentPage = 1; // Reset to first page when changing page size
        this.loadPageData();
      });
    }
  }
  
  /**
   * Navigates to a specific page
   * @param {number} pageNumber - Page number to navigate to
   * @private
   */
  async goToPage(pageNumber) {
    if (pageNumber < 1 || pageNumber > this.totalPages) {
      console.warn(`Invalid page number: ${pageNumber}`);
      return;
    }
    
    this.currentPage = pageNumber;
    await this.loadPageData();
  }
  
  /**
   * Extracts Dataverse field names from the field mapping
   * @returns {Array<string>} Array of Dataverse field names to request from the API
   * @private
   */
  getDataverseFieldNames() {
    if (!this.fieldMapping) {
      return undefined; // No mapping means request all fields or use default
    }
    
    // Extract Dataverse field names from the mapping values
    return Object.values(this.fieldMapping);
  }
  
  /**
   * Transforms Dataverse data using the field mapping
   * Converts Dataverse field names (e.g., cr123_vehiclemake) to display field names (e.g., make)
   * @param {Array} data - Raw data from Dataverse with Dataverse field names
   * @returns {Array} Transformed data with display field names
   * @private
   */
  transformDataverseData(data) {
    // If no field mapping is provided, return data as-is
    if (!this.fieldMapping || !data) {
      return data;
    }
    
    // Transform each row
    return data.map(row => {
      const transformedRow = {};
      
      // Map each display field to its Dataverse field
      for (const [displayField, dataverseField] of Object.entries(this.fieldMapping)) {
        transformedRow[displayField] = row[dataverseField];
      }
      
      // Include any unmapped fields as well
      for (const [key, value] of Object.entries(row)) {
        if (!Object.values(this.fieldMapping).includes(key)) {
          transformedRow[key] = value;
        }
      }
      
      return transformedRow;
    });
  }
  
  /**
   * Loads data for the current page and updates the grid
   * @private
   */
  async loadPageData() {
    if (!this.gridApi) {
      console.warn(`=== DataGrid ${this.instanceName}: Grid not initialized ===`);
      return;
    }
    
    try {
      console.log(`=== DataGrid ${this.instanceName}: Loading page ${this.currentPage} ===`);
      
      // Show loading overlay
      this.gridApi.showLoadingOverlay();
      
      // Fetch page data
      const result = await this.getData(this.currentPage, this.pageSize, this.pageToken);
      
      // Update pagination state
      this.totalRecords = result.pagination.totalRecords;
      this.totalPages = result.pagination.totalPages;
      this.pageToken = result.pagination.nextPageToken;
      
      // Transform data if field mapping is provided
      const transformedData = this.transformDataverseData(result.data);
      
      // Update the grid data
      this.gridApi.setGridOption('rowData', transformedData);
      
      // Update pagination controls
      this.updatePaginationControls();
      
      // Hide loading overlay
      this.gridApi.hideOverlay();
      
      console.log(`=== DataGrid ${this.instanceName}: Page ${this.currentPage} loaded successfully ===`);
    } catch (error) {
      console.error(`=== DataGrid ${this.instanceName}: Error loading page ===`, error);
      this.gridApi.hideOverlay();
      throw error;
    }
  }
  
  /**
   * Binds event handlers for action toolbar
   * @private
   */
  bindActionHandlers() {
    // Add click handler to the action button itself to show message when disabled
    const actionBtn = document.getElementById(`${this.instanceId}-actions-btn`);
    if (actionBtn) {
      actionBtn.addEventListener('click', (e) => {
        if (actionBtn.getAttribute('data-disabled') === 'true') {
          e.preventDefault();
          e.stopPropagation();
          alert('Please select at least one record to perform an action.');
          return false;
        }
      });
    }
    
    const actionMenu = document.getElementById(`${this.instanceId}-actions-menu`);
    if (!actionMenu) return;
    
    actionMenu.addEventListener('click', (e) => {
      const actionLink = e.target.closest('[data-action]');
      if (!actionLink || actionLink.classList.contains('disabled')) {
        e.preventDefault();
        return;
      }
      
      e.preventDefault();
      const actionName = actionLink.getAttribute('data-action');
      const selectedRows = this.getSelectedRow();
      
      console.log(`Action '${actionName}' triggered for ${selectedRows.length} row(s):`, selectedRows);
      
      // Call the actionHandler if provided
      if (this.actionHandler && typeof this.actionHandler === 'function') {
        this.actionHandler(actionName, selectedRows, this);
      } else {
        console.warn(`No actionHandler defined for grid instance '${this.instanceName}'`);
      }
    });
  }
  
  /**
   * Returns sample data for local development
   * In production, this will be replaced with Power Automate flow integration via fetchDocuments
   * @param {number} page - Page number (1-based)
   * @param {number} pageSize - Number of records per page
   * @param {string} pageToken - Continuation token (for production use)
   * @returns {Promise<Object>} Object with data array and pagination info
   */
  async getData(page = 1, pageSize = 10, pageToken = null) {
    // PRODUCTION CODE (currently disabled for local development):
    // Uncomment this block when integrating with Dataverse via fetchDocuments
    /*
    const dataverseFields = this.getDataverseFieldNames();
    
    const result = await fetchDocuments({
      fields: dataverseFields,
      accountId: this.recordId, // or use separate accountId/productId properties
      // productId: this.productId,
      // category: this.category,
      pageSize: pageSize,
      page: page,
      pageToken: pageToken
    });
    
    // fetchDocuments returns data in Dataverse field names,
    // which will be transformed by transformDataverseData() in loadPageData()
    return {
      data: result.data,
      pagination: result.pagination
    };
    */
    
    // LOCAL DEVELOPMENT CODE (remove when using production):
    // Simulate async data fetch
    return new Promise((resolve) => {
      setTimeout(() => {
        // Extended sample data set for pagination demo (50+ records)
        // In production, this will call fetchDocuments with accountId/productId filters
        const allSampleData = [
          { make: 'Toyota', model: 'Celica', price: 35000, year: 2020, color: 'Red', dateFirstReleased: '2020-03-15' },
          { make: 'Ford', model: 'Mondeo', price: 32000, year: 2019, color: 'Blue', dateFirstReleased: '2019-08-22' },
          { make: 'Porsche', model: 'Boxster', price: 72000, year: 2021, color: 'Silver', dateFirstReleased: '2021-05-10' },
          { make: 'BMW', model: 'M50', price: 60000, year: 2022, color: 'Black', dateFirstReleased: '2022-01-18' },
          { make: 'Aston Martin', model: 'DBX', price: 190000, year: 2023, color: 'Green', dateFirstReleased: '2023-06-30' },
          { make: 'Tesla', model: 'Model 3', price: 45000, year: 2023, color: 'White', dateFirstReleased: '2023-02-14' },
          { make: 'Honda', model: 'Civic', price: 28000, year: 2020, color: 'Gray', dateFirstReleased: '2020-09-05' },
          { make: 'Mercedes', model: 'C-Class', price: 55000, year: 2022, color: 'Silver', dateFirstReleased: '2022-11-12' },
          { make: 'Volkswagen', model: 'Golf', price: 25000, year: 2021, color: 'Blue', dateFirstReleased: '2021-07-20' },
          { make: 'Audi', model: 'A4', price: 48000, year: 2023, color: 'Black', dateFirstReleased: '2023-04-08' },
          { make: 'Mazda', model: 'MX-5', price: 32000, year: 2022, color: 'Red', dateFirstReleased: '2022-04-12' },
          { make: 'Nissan', model: 'GT-R', price: 115000, year: 2023, color: 'Gray', dateFirstReleased: '2023-01-20' },
          { make: 'Chevrolet', model: 'Corvette', price: 68000, year: 2021, color: 'Yellow', dateFirstReleased: '2021-08-15' },
          { make: 'Dodge', model: 'Challenger', price: 52000, year: 2020, color: 'Orange', dateFirstReleased: '2020-06-10' },
          { make: 'Jaguar', model: 'F-Type', price: 73000, year: 2022, color: 'Blue', dateFirstReleased: '2022-09-22' },
          { make: 'Lexus', model: 'LC500', price: 95000, year: 2023, color: 'White', dateFirstReleased: '2023-03-18' },
          { make: 'Alfa Romeo', model: 'Giulia', price: 46000, year: 2021, color: 'Red', dateFirstReleased: '2021-12-05' },
          { make: 'Subaru', model: 'WRX', price: 38000, year: 2022, color: 'Blue', dateFirstReleased: '2022-07-14' },
          { make: 'Mitsubishi', model: 'Lancer', price: 29000, year: 2020, color: 'Silver', dateFirstReleased: '2020-11-30' },
          { make: 'Hyundai', model: 'Veloster', price: 26000, year: 2021, color: 'Green', dateFirstReleased: '2021-05-25' },
          { make: 'Kia', model: 'Stinger', price: 42000, year: 2022, color: 'Gray', dateFirstReleased: '2022-02-08' },
          { make: 'Genesis', model: 'G70', price: 48000, year: 2023, color: 'Black', dateFirstReleased: '2023-08-16' },
          { make: 'Cadillac', model: 'CT5', price: 51000, year: 2022, color: 'White', dateFirstReleased: '2022-10-12' },
          { make: 'Lincoln', model: 'Corsair', price: 44000, year: 2021, color: 'Blue', dateFirstReleased: '2021-09-20' },
          { make: 'Volvo', model: 'S60', price: 47000, year: 2023, color: 'Silver', dateFirstReleased: '2023-05-14' },
          { make: 'Infiniti', model: 'Q50', price: 43000, year: 2022, color: 'Red', dateFirstReleased: '2022-03-28' },
          { make: 'Acura', model: 'TLX', price: 45000, year: 2023, color: 'Gray', dateFirstReleased: '2023-07-09' },
          { make: 'Buick', model: 'Regal', price: 35000, year: 2020, color: 'Blue', dateFirstReleased: '2020-12-15' },
          { make: 'Chrysler', model: '300', price: 39000, year: 2021, color: 'Black', dateFirstReleased: '2021-06-22' },
          { make: 'Ram', model: '1500', price: 48000, year: 2022, color: 'White', dateFirstReleased: '2022-08-05' },
          { make: 'GMC', model: 'Sierra', price: 52000, year: 2023, color: 'Silver', dateFirstReleased: '2023-09-11' },
          { make: 'Jeep', model: 'Grand Cherokee', price: 56000, year: 2022, color: 'Green', dateFirstReleased: '2022-05-18' },
          { make: 'Land Rover', model: 'Range Rover', price: 98000, year: 2023, color: 'Black', dateFirstReleased: '2023-10-25' },
          { make: 'Maserati', model: 'Ghibli', price: 82000, year: 2022, color: 'Blue', dateFirstReleased: '2022-12-08' },
          { make: 'Bentley', model: 'Continental', price: 215000, year: 2023, color: 'Silver', dateFirstReleased: '2023-11-14' },
          { make: 'Rolls Royce', model: 'Ghost', price: 350000, year: 2023, color: 'White', dateFirstReleased: '2023-12-20' },
          { make: 'Ferrari', model: '488', price: 280000, year: 2022, color: 'Red', dateFirstReleased: '2022-06-30' },
          { make: 'Lamborghini', model: 'Huracan', price: 265000, year: 2023, color: 'Yellow', dateFirstReleased: '2023-04-17' },
          { make: 'McLaren', model: '720S', price: 310000, year: 2022, color: 'Orange', dateFirstReleased: '2022-09-09' },
          { make: 'Bugatti', model: 'Chiron', price: 3000000, year: 2023, color: 'Blue', dateFirstReleased: '2023-01-05' },
          { make: 'Koenigsegg', model: 'Jesko', price: 2800000, year: 2023, color: 'Gray', dateFirstReleased: '2023-03-22' },
          { make: 'Pagani', model: 'Huayra', price: 2600000, year: 2022, color: 'Silver', dateFirstReleased: '2022-07-28' },
          { make: 'Lotus', model: 'Evora', price: 95000, year: 2021, color: 'Green', dateFirstReleased: '2021-10-14' },
          { make: 'Mini', model: 'Cooper', price: 32000, year: 2022, color: 'Red', dateFirstReleased: '2022-04-19' },
          { make: 'Fiat', model: '500', price: 22000, year: 2021, color: 'White', dateFirstReleased: '2021-11-08' },
          { make: 'Peugeot', model: '508', price: 38000, year: 2022, color: 'Blue', dateFirstReleased: '2022-01-26' },
          { make: 'Renault', model: 'Megane', price: 28000, year: 2021, color: 'Gray', dateFirstReleased: '2021-08-12' },
          { make: 'Citroen', model: 'C4', price: 26000, year: 2022, color: 'Silver', dateFirstReleased: '2022-02-15' },
          { make: 'Skoda', model: 'Octavia', price: 30000, year: 2023, color: 'Black', dateFirstReleased: '2023-06-03' },
          { make: 'Seat', model: 'Leon', price: 29000, year: 2022, color: 'Red', dateFirstReleased: '2022-11-21' },
          { make: 'Smart', model: 'ForTwo', price: 18000, year: 2021, color: 'Yellow', dateFirstReleased: '2021-04-09' },
          { make: 'Rivian', model: 'R1T', price: 73000, year: 2023, color: 'Green', dateFirstReleased: '2023-09-27' },
          { make: 'Lucid', model: 'Air', price: 92000, year: 2023, color: 'Silver', dateFirstReleased: '2023-07-18' }
        ];
        
        // Calculate pagination
        const totalRecords = allSampleData.length;
        const totalPages = Math.ceil(totalRecords / pageSize);
        const startIndex = (page - 1) * pageSize;
        const endIndex = startIndex + pageSize;
        const pageData = allSampleData.slice(startIndex, endIndex);
        
        // Simulate pageToken for production compatibility
        const nextPageToken = page < totalPages ? `token_page_${page + 1}` : null;
        
        console.log(`=== DataGrid ${this.instanceName}: Retrieved page ${page}/${totalPages} (${pageData.length} records) ===`);
        
        resolve({
          data: pageData,
          pagination: {
            currentPage: page,
            pageSize: pageSize,
            totalRecords: totalRecords,
            totalPages: totalPages,
            hasNextPage: page < totalPages,
            hasPreviousPage: page > 1,
            nextPageToken: nextPageToken
          }
        });
      }, 500); // Simulate network delay
    });
  }
  
  /**
   * Initializes the AG Grid and renders it in the container
   * @returns {Promise<void>}
   */
  async initialize() {
    const container = document.getElementById(this.containerId);
    if (!container) {
      throw new Error(`Container with id "${this.containerId}" not found`);
    }
    
    // Create the grid HTML structure with action toolbar and pagination
    const actionToolbarHtml = this.createActionToolbar();
    const paginationHtml = this.createPaginationControls();
    const gridHtml = `
      <div class="ag-grid-wrapper card shadow-sm">
        ${actionToolbarHtml}
        <div class="card-body p-0">
          <div id="${this.instanceId}" class="ag-theme-alpine" style="width: 100%; height: 500px;"></div>
        </div>
        ${paginationHtml}
      </div>
    `;
    
    container.innerHTML = gridHtml;
    
    // Get initial page data
    const result = await this.getData(this.currentPage, this.pageSize, this.pageToken);
    const rowData = result.data;
    
    // Update pagination state
    this.totalRecords = result.pagination.totalRecords;
    this.totalPages = result.pagination.totalPages;
    this.pageToken = result.pagination.nextPageToken;
    
    // Prepare column definitions - add checkbox selection column
    let finalColumnDefs = [...this.columnDefs];
    
    // Add checkbox selection to the first column or create a dedicated selection column
    if (finalColumnDefs.length > 0) {
      // Add checkboxSelection to the first user-defined column
      finalColumnDefs[0] = {
        ...finalColumnDefs[0],
        checkboxSelection: true,
        headerCheckboxSelection: true,
        headerCheckboxSelectionFilteredOnly: true,
        width: finalColumnDefs[0].width || 150,
        pinned: 'left'
      };
    }
    
    // Grid Options
    const gridOptions = {
      columnDefs: finalColumnDefs,
      rowData: rowData,
      rowSelection: 'multiple',
      animateRows: true,
      suppressRowClickSelection: true,
      defaultColDef: {
        flex: 1,
        minWidth: 100,
        resizable: true,
      },
      onGridReady: (params) => {
        this.gridApi = params.api;
        this.gridColumnApi = params.columnApi;
        this.updateActionButtonState();
        this.updatePaginationControls();
        console.log(`=== DataGrid ${this.instanceName}: Grid initialized ===`);
      },
      onSelectionChanged: () => {
        const selectedRows = this.getSelectedRow();
        this.updateActionButtonState();
        console.log(`=== DataGrid ${this.instanceName}: ${selectedRows.length} row(s) selected ===`);
      }
    };
    
    // Create the grid
    const gridDiv = document.getElementById(this.instanceId);
    agGrid.createGrid(gridDiv, gridOptions);
    
    // Bind action handlers
    this.bindActionHandlers();
    
    // Bind pagination handlers
    this.bindPaginationHandlers();
    
    console.log(`=== DataGrid ${this.instanceName}: Initialization complete ===`);
  }
  
  /**
   * Clears all data from the grid
   */
  clear() {
    if (this.gridApi) {
      this.gridApi.setGridOption('rowData', []);
      console.log(`=== DataGrid ${this.instanceName}: Grid cleared ===`);
    } else {
      console.warn(`=== DataGrid ${this.instanceName}: Grid not initialized ===`);
    }
  }
  
  /**
   * Updates the grid by fetching new data and refreshing the display
   * This happens in an AJAX manner without reloading the entire page
   * Reloads the current page to reflect any backend changes
   * @returns {Promise<void>}
   */
  async update() {
    if (!this.gridApi) {
      console.warn(`=== DataGrid ${this.instanceName}: Grid not initialized ===`);
      return;
    }
    
    console.log(`=== DataGrid ${this.instanceName}: Updating grid data ===`);
    await this.loadPageData();
  }
  
  /**
   * Gets the currently selected row(s)
   * @returns {Array} Array of selected row data objects
   */
  getSelectedRow() {
    if (!this.gridApi) {
      console.warn(`=== DataGrid ${this.instanceName}: Grid not initialized ===`);
      return [];
    }
    
    const selectedRows = this.gridApi.getSelectedRows();
    return selectedRows;
  }
  
  /**
   * Sets the row selection programmatically
   * @param {Function} predicateFn - Function that returns true for rows to select
   */
  selectRows(predicateFn) {
    if (!this.gridApi) {
      console.warn(`=== DataGrid ${this.instanceName}: Grid not initialized ===`);
      return;
    }
    
    this.gridApi.forEachNode((node) => {
      if (predicateFn(node.data)) {
        node.setSelected(true);
      }
    });
  }
  
  /**
   * Exports grid data to CSV
   * @param {string} fileName - Name of the CSV file (optional)
   */
  exportToCSV(fileName = `${this.instanceName}-export.csv`) {
    if (!this.gridApi) {
      console.warn(`=== DataGrid ${this.instanceName}: Grid not initialized ===`);
      return;
    }
    
    this.gridApi.exportDataAsCsv({ fileName: fileName });
    console.log(`=== DataGrid ${this.instanceName}: Data exported to ${fileName} ===`);
  }
  
  /**
   * Gets all current grid data
   * @returns {Array} Array of all row data objects
   */
  getAllData() {
    if (!this.gridApi) {
      console.warn(`=== DataGrid ${this.instanceName}: Grid not initialized ===`);
      return [];
    }
    
    const allData = [];
    this.gridApi.forEachNode((node) => {
      allData.push(node.data);
    });
    return allData;
  }
  
  /**
   * Updates column definitions and refreshes the grid
   * @param {Array} newColumnDefs - New column definitions
   */
  updateColumns(newColumnDefs) {
    if (!this.gridApi) {
      console.warn(`=== DataGrid ${this.instanceName}: Grid not initialized ===`);
      return;
    }
    
    this.columnDefs = newColumnDefs;
    this.gridApi.setGridOption('columnDefs', newColumnDefs);
    console.log(`=== DataGrid ${this.instanceName}: Columns updated ===`);
  }
}

// ============================================================================
// DataGridConfig Class - Simplified Configuration Wrapper for DataGrid
// Automatically generates action handlers and reduces boilerplate code
// ============================================================================

/**
 * DataGridConfig class that simplifies DataGrid creation by providing
 * auto-generated action handlers and sensible defaults.
 * 
 * This wrapper class reduces designer effort from ~50 lines to ~10 lines
 * while maintaining full flexibility through custom handler support.
 * 
 * Actions support selectionMode parameter:
 * - 'single': Only enabled when exactly one row is selected
 * - 'multiple': Only enabled when multiple rows are selected  
 * - 'any': Enabled when one or more rows are selected (default)
 * 
 * @class DataGridConfig
 * @example
 * // Simple usage with default handlers
 * const gridConfig = new DataGridConfig('Car Inventory', {
 *     containerId: 'gridContainer1',
 *     columns: [
 *         { field: 'make', sortable: true, filter: true },
 *         { field: 'model', sortable: true, filter: true }
 *     ],
 *     actions: [
 *         { name: 'view', selectionMode: 'single' },
 *         { name: 'edit', selectionMode: 'single' },
 *         { name: 'delete', selectionMode: 'any' }
 *     ],
 *     flowUrl: 'https://flow.example.com/cars',
 *     recordId: 'car_record_123'
 * });
 * await gridConfig.render();
 * 
 * @example
 * // Advanced usage with custom handler
 * const gridConfig = new DataGridConfig('Car Inventory', {
 *     containerId: 'gridContainer1',
 *     columns: [...],
 *     actions: [
 *         { name: 'view', selectionMode: 'single' },     // Use default handler for single row
 *         { name: 'delete', selectionMode: 'any' },      // Use default handler for any selection
 *         {
 *             name: 'approve',
 *             text: 'Approve Cars',
 *             selectionMode: 'multiple',
 *             handler: (selectedRows) => {
 *                 // Custom logic here for multiple rows
 *             }
 *         }
 *     ]
 * });
 */
class DataGridConfig {
    /**
     * Creates a new DataGridConfig instance
     * @param {string} title - Display title for the grid
     * @param {Object} config - Configuration object
     * @param {string} config.containerId - ID of the container DIV element
     * @param {Array} config.columns - AG Grid column definitions
     * @param {Array<string|Object>} config.actions - Action definitions. Can be strings (use defaults) or objects with {name, text?, selectionMode?, handler?}
     *   - selectionMode: 'single' (one row), 'multiple' (2+ rows), 'any' (1+ rows, default)
     * @param {string} [config.flowUrl=''] - URL to Power Automate Flow
     * @param {string} [config.recordId=''] - Record ID to pass to flows
     * @param {string} [config.instanceName] - Optional custom instance name (auto-generated if not provided)
     * @param {Object} [config.gridOptions={}] - Additional AG Grid options to merge
     * @param {Function} [config.actionHandler] - Optional function to handle all grid actions (receives actionName and selectedRows array)
     * @param {Object} [config.fieldMapping] - Optional mapping object to convert Dataverse field names to display field names (e.g., {'make': 'cr123_vehiclemake'})
     */
    constructor(title, config) {
        this.title = title;
        this.containerId = config.containerId;
        this.columns = config.columns;
        this.actions = config.actions || [];
        this.flowUrl = config.flowUrl || '';
        this.recordId = config.recordId || '';
        this.instanceName = config.instanceName || this.generateInstanceName();
        this.gridOptions = config.gridOptions || {};
        this.actionHandler = config.actionHandler || null;
        this.fieldMapping = config.fieldMapping || null;
        
        // Will hold the underlying DataGrid instance
        this.grid = null;
        
        // Storage for action handlers
        this.actionHandlers = {};
        
        console.log(`=== DataGridConfig Created: ${this.title} ===`);
    }
    
    /**
     * Generates a unique instance name based on the title
     * @returns {string} Instance name
     * @private
     */
    generateInstanceName() {
        const safeName = this.title
            .replace(/[^a-zA-Z0-9]/g, '')
            .replace(/\s+/g, '_')
            .toLowerCase();
        return `grid_${safeName}_${Date.now().toString(36)}`;
    }
    
    /**
     * Auto-generates handlers for all configured actions
     * Supports both string actions (use defaults) and object actions (custom handlers)
     * @private
     */
    generateActionHandlers() {
        this.actions.forEach(action => {
            let actionName;
            if (typeof action === 'string') {
                actionName = action;
                // String action - use built-in default handler
                this.actionHandlers[actionName] = this.getDefaultHandler(actionName);
            } else if (typeof action === 'object' && action.name) {
                actionName = action.name;
                // Object action - check if custom handler provided
                if (typeof action.handler === 'function') {
                    // Use custom handler
                    this.actionHandlers[actionName] = action.handler;
                } else {
                    // No custom handler - use default
                    this.actionHandlers[actionName] = this.getDefaultHandler(actionName);
                }
            }
        });
        
        console.log(`=== DataGridConfig ${this.title}: Generated ${Object.keys(this.actionHandlers).length} action handlers ===`);
    }
    
    /**
     * Returns the default handler function for a given action name
     * @param {string} actionName - Name of the action
     * @returns {Function} Handler function that accepts selectedRows array
     * @private
     */
    getDefaultHandler(actionName) {
        const handlers = {
            'view': (selectedRows) => this.defaultViewHandler(selectedRows),
            'edit': (selectedRows) => this.defaultEditHandler(selectedRows),
            'delete': (selectedRows) => this.defaultDeleteHandler(selectedRows),
            'reserve': (selectedRows) => this.defaultReserveHandler(selectedRows),
            'approve': (selectedRows) => this.defaultApproveHandler(selectedRows),
            'reject': (selectedRows) => this.defaultRejectHandler(selectedRows),
            'download': (selectedRows) => this.defaultDownloadHandler(selectedRows),
            'print': (selectedRows) => this.defaultPrintHandler(selectedRows),
            'duplicate': (selectedRows) => this.defaultDuplicateHandler(selectedRows)
        };
        
        // Return specific handler or generic one
        return handlers[actionName] || ((selectedRows) => {
            this.defaultGenericHandler(actionName, selectedRows);
        });
    }
    
    /**
     * Default VIEW handler - shows row data in a Bootstrap modal
     * @param {Array} selectedRows - Array of selected row data objects
     * @private
     */
    defaultViewHandler(selectedRows) {
        if (!selectedRows || selectedRows.length === 0) {
            alert('No rows selected');
            return;
        }
        
        console.log(`View action triggered for ${selectedRows.length} row(s):`, selectedRows);
        
        // For single row, show detailed view
        if (selectedRows.length === 1) {
            const rowData = selectedRows[0];
            let content = '<div class="table-responsive"><table class="table table-striped">';
            Object.entries(rowData).forEach(([key, value]) => {
                content += `
                    <tr>
                        <th style="width: 30%">${this.formatFieldName(key)}</th>
                        <td>${this.formatFieldValue(value)}</td>
                    </tr>
                `;
            });
            content += '</table></div>';
            
            this.showModal('View Details', content, [
                { text: 'Close', class: 'btn-secondary', action: 'close' }
            ]);
        } else {
            // For multiple rows, show summary
            let content = `<p>Viewing ${selectedRows.length} selected records.</p>`;
            content += '<div class="table-responsive"><table class="table table-sm table-striped">';
            content += '<thead><tr>';
            
            // Get all keys from first row
            const keys = Object.keys(selectedRows[0]);
            keys.forEach(key => {
                content += `<th>${this.formatFieldName(key)}</th>`;
            });
            content += '</tr></thead><tbody>';
            
            // Add rows (limit to first 10 for performance)
            selectedRows.slice(0, 10).forEach(row => {
                content += '<tr>';
                keys.forEach(key => {
                    content += `<td>${this.formatFieldValue(row[key])}</td>`;
                });
                content += '</tr>';
            });
            
            if (selectedRows.length > 10) {
                content += `<tr><td colspan="${keys.length}"><em>... and ${selectedRows.length - 10} more</em></td></tr>`;
            }
            
            content += '</tbody></table></div>';
            
            this.showModal(`View ${selectedRows.length} Records`, content, [
                { text: 'Close', class: 'btn-secondary', action: 'close' }
            ]);
        }
    }
    
    /**
     * Default EDIT handler - prompts for field edits
     * @param {Array} selectedRows - Array of selected row data objects
     * @private
     */
    defaultEditHandler(selectedRows) {
        if (!selectedRows || selectedRows.length === 0) {
            alert('No rows selected');
            return;
        }
        
        if (selectedRows.length > 1) {
            alert('Please select only one row to edit. For bulk updates, use a bulk action.');
            return;
        }
        
        const rowData = selectedRows[0];
        console.log(`Edit action triggered for:`, rowData);
        
        // For now, use simple prompt - can be enhanced with modal form
        const fields = Object.keys(rowData);
        const firstEditableField = fields.find(f => f !== 'id' && f !== 'actions');
        
        if (firstEditableField) {
            const currentValue = rowData[firstEditableField];
            const newValue = prompt(
                `Edit ${this.formatFieldName(firstEditableField)}:`, 
                currentValue
            );
            
            if (newValue !== null && newValue !== currentValue) {
                rowData[firstEditableField] = newValue;
                // Update grid would happen here
                console.log(`Updated ${firstEditableField} to:`, newValue);
                
                // In production, would call Power Automate flow here
                // await this.callFlow('update', rowData);
                alert(`Field updated! (In production, this would save to Dataverse)`);
            }
        } else {
            alert('No editable fields found');
        }
    }
    
    /**
     * Default DELETE handler - confirms and logs deletion
     * @param {Array} selectedRows - Array of selected row data objects
     * @private
     */
    defaultDeleteHandler(selectedRows) {
        if (!selectedRows || selectedRows.length === 0) {
            alert('No rows selected');
            return;
        }
        
        console.log(`Delete action triggered for ${selectedRows.length} row(s):`, selectedRows);
        
        const message = selectedRows.length === 1 
            ? `Are you sure you want to delete this record?\n\n${this.getRowDisplayValue(selectedRows[0])}`
            : `Are you sure you want to delete ${selectedRows.length} records?`;
        
        if (confirm(message)) {
            console.log('Delete confirmed for:', selectedRows);
            
            // In production, would call Power Automate flow here
            // await this.callFlow('delete', selectedRows);
            // Then remove from grid:
            // this.grid.gridApi.applyTransaction({ remove: selectedRows });
            
            alert(`Delete action would be performed for ${selectedRows.length} record(s).\n(Removed from database via Power Automate flow)`);
        }
    }
    
    /**
     * Default RESERVE handler - prompts for reservation details
     * @param {Array} selectedRows - Array of selected row data objects
     * @private
     */
    defaultReserveHandler(selectedRows) {
        if (!selectedRows || selectedRows.length === 0) {
            alert('No rows selected');
            return;
        }
        
        console.log(`Reserve action triggered for ${selectedRows.length} row(s):`, selectedRows);
        
        const message = selectedRows.length === 1
            ? `Reserve this item for customer:\n\n${this.getRowDisplayValue(selectedRows[0])}`
            : `Reserve ${selectedRows.length} items for customer:`;
        
        const customerName = prompt(message);
        
        if (customerName) {
            console.log('Reserved for:', customerName, 'Rows:', selectedRows);
            
            // In production, would call Power Automate flow here
            // await this.callFlow('reserve', { rows: selectedRows, customerName });
            
            alert(`Reserved ${selectedRows.length} item(s) for: ${customerName}`);
        }
    }
    
    /**
     * Default APPROVE handler
     * @param {Array} selectedRows - Array of selected row data objects
     * @private
     */
    defaultApproveHandler(selectedRows) {
        if (!selectedRows || selectedRows.length === 0) {
            alert('No rows selected');
            return;
        }
        
        console.log(`Approve action triggered for ${selectedRows.length} row(s):`, selectedRows);
        
        const message = selectedRows.length === 1
            ? 'Approve this record?'
            : `Approve ${selectedRows.length} records?`;
        
        if (confirm(message)) {
            alert(`Approval action would be performed for ${selectedRows.length} record(s).\n(Sent to Power Automate flow)`);
            console.log('Approved:', selectedRows);
        }
    }
    
    /**
     * Default REJECT handler
     * @param {Array} selectedRows - Array of selected row data objects
     * @private
     */
    defaultRejectHandler(selectedRows) {
        if (!selectedRows || selectedRows.length === 0) {
            alert('No rows selected');
            return;
        }
        
        console.log(`Reject action triggered for ${selectedRows.length} row(s):`, selectedRows);
        
        const reason = prompt(`Reason for rejecting ${selectedRows.length} record(s):`);
        if (reason) {
            alert(`Rejection action would be performed for ${selectedRows.length} record(s).\n(Sent to Power Automate flow)`);
            console.log('Rejected:', selectedRows, 'Reason:', reason);
        }
    }
    
    /**
     * Default DOWNLOAD handler
     * @param {Array} selectedRows - Array of selected row data objects
     * @private
     */
    defaultDownloadHandler(selectedRows) {
        if (!selectedRows || selectedRows.length === 0) {
            alert('No rows selected');
            return;
        }
        
        console.log(`Download action triggered for ${selectedRows.length} row(s):`, selectedRows);
        alert(`Download functionality would be implemented for ${selectedRows.length} record(s).\n(Fetch files from Dataverse)`);
    }
    
    /**
     * Default PRINT handler
     * @param {Array} selectedRows - Array of selected row data objects
     * @private
     */
    defaultPrintHandler(selectedRows) {
        if (!selectedRows || selectedRows.length === 0) {
            alert('No rows selected');
            return;
        }
        
        console.log(`Print action triggered for ${selectedRows.length} row(s):`, selectedRows);
        
        // Create printable content
        let content = `<h3>Record Details (${selectedRows.length} record(s))</h3>`;
        
        selectedRows.forEach((rowData, index) => {
            if (selectedRows.length > 1) {
                content += `<h5>Record ${index + 1}</h5>`;
            }
            content += '<table class="table table-sm">';
            Object.entries(rowData).forEach(([key, value]) => {
                content += `<tr><th style="width: 30%">${this.formatFieldName(key)}</th><td>${this.formatFieldValue(value)}</td></tr>`;
            });
            content += '</table>';
            if (index < selectedRows.length - 1) {
                content += '<hr>';
            }
        });
        
        const printWindow = window.open('', '_blank');
        printWindow.document.write(`
            <html>
                <head>
                    <title>Print Records</title>
                    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
                </head>
                <body class="p-4">${content}</body>
            </html>
        `);
        printWindow.document.close();
        printWindow.print();
    }
    
    /**
     * Default DUPLICATE handler
     * @param {Array} selectedRows - Array of selected row data objects
     * @private
     */
    defaultDuplicateHandler(selectedRows) {
        if (!selectedRows || selectedRows.length === 0) {
            alert('No rows selected');
            return;
        }
        
        console.log(`Duplicate action triggered for ${selectedRows.length} row(s):`, selectedRows);
        
        const message = selectedRows.length === 1
            ? 'Create a duplicate of this record?'
            : `Create duplicates of ${selectedRows.length} records?`;
        
        if (confirm(message)) {
            alert(`Duplicate action would be performed for ${selectedRows.length} record(s).\n(Create new records via Power Automate flow)`);
            console.log('Duplicating:', selectedRows);
        }
    }
    
    /**
     * Generic handler for actions without specific defaults
     * @param {string} actionName - Action name
     * @param {Array} selectedRows - Array of selected row data objects
     * @private
     */
    defaultGenericHandler(actionName, selectedRows) {
        if (!selectedRows || selectedRows.length === 0) {
            alert('No rows selected');
            return;
        }
        
        console.log(`${actionName} action triggered for ${selectedRows.length} row(s):`, selectedRows);
        alert(`Action: ${actionName.charAt(0).toUpperCase() + actionName.slice(1)}\n\nSelected ${selectedRows.length} row(s):\n${JSON.stringify(selectedRows.slice(0, 3), null, 2)}${selectedRows.length > 3 ? '\n... and more' : ''}`);
    }
    
    /**
     * Helper: Formats field name for display (converts camelCase/snake_case to Title Case)
     * @param {string} fieldName - Field name
     * @returns {string} Formatted name
     * @private
     */
    formatFieldName(fieldName) {
        return fieldName
            .replace(/([A-Z])/g, ' $1')
            .replace(/_/g, ' ')
            .replace(/^./, str => str.toUpperCase())
            .trim();
    }
    
    /**
     * Helper: Formats field value for display
     * @param {*} value - Field value
     * @returns {string} Formatted value
     * @private
     */
    formatFieldValue(value) {
        if (value === null || value === undefined) return '<em>N/A</em>';
        if (typeof value === 'boolean') return value ? 'Yes' : 'No';
        if (typeof value === 'number') return value.toLocaleString();
        if (typeof value === 'object') return JSON.stringify(value);
        return value;
    }
    
    /**
     * Helper: Gets a human-readable display value for a row
     * @param {Object} rowData - Row data
     * @returns {string} Display value
     * @private
     */
    getRowDisplayValue(rowData) {
        // Try to find name/title field
        const nameFields = ['name', 'title', 'make', 'model', 'description', 'label'];
        const foundField = nameFields.find(f => rowData[f]);
        
        if (foundField) {
            return `${this.formatFieldName(foundField)}: ${rowData[foundField]}`;
        }
        
        // Fallback: show first 3 non-ID fields
        const fields = Object.entries(rowData)
            .filter(([key]) => key !== 'id' && key !== 'actions')
            .slice(0, 3);
        
        return fields.map(([key, value]) => `${this.formatFieldName(key)}: ${value}`).join('\n');
    }
    
    /**
     * Shows a Bootstrap modal (creates if doesn't exist)
     * @param {string} title - Modal title
     * @param {string} content - HTML content
     * @param {Array<Object>} buttons - Array of button definitions
     * @private
     */
    showModal(title, content, buttons = []) {
        // Check if modal already exists
        let modal = document.getElementById('dataGridConfigModal');
        
        if (!modal) {
            // Create modal HTML
            const modalHtml = `
                <div class="modal fade" id="dataGridConfigModal" tabindex="-1">
                    <div class="modal-dialog modal-lg">
                        <div class="modal-content">
                            <div class="modal-header">
                                <h5 class="modal-title"></h5>
                                <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                            </div>
                            <div class="modal-body"></div>
                            <div class="modal-footer"></div>
                        </div>
                    </div>
                </div>
            `;
            
            document.body.insertAdjacentHTML('beforeend', modalHtml);
            modal = document.getElementById('dataGridConfigModal');
        }
        
        // Update modal content
        modal.querySelector('.modal-title').textContent = title;
        modal.querySelector('.modal-body').innerHTML = content;
        
        // Create footer buttons
        const footer = modal.querySelector('.modal-footer');
        footer.innerHTML = '';
        
        buttons.forEach(btn => {
            const button = document.createElement('button');
            button.type = 'button';
            button.className = `btn ${btn.class}`;
            button.textContent = btn.text;
            
            if (btn.action === 'close') {
                button.setAttribute('data-bs-dismiss', 'modal');
            } else if (typeof btn.action === 'function') {
                button.addEventListener('click', btn.action);
            }
            
            footer.appendChild(button);
        });
        
        // Show modal
        const bootstrapModal = new bootstrap.Modal(modal);
        bootstrapModal.show();
    }
    
    /**
     * Renders the grid with all handlers registered
     * @returns {Promise<void>}
     */
    async render() {
        console.log(`=== DataGridConfig ${this.title}: Starting render ===`);
        
        // Convert action definitions to action objects for DataGrid (preserve selectionMode)
        const actionDefs = this.actions.map(action => {
            if (typeof action === 'string') {
                return {
                    name: action,
                    text: action.charAt(0).toUpperCase() + action.slice(1),
                    selectionMode: 'any' // Default for string actions
                };
            } else {
                return {
                    name: action.name,
                    text: action.text || (action.name.charAt(0).toUpperCase() + action.name.slice(1)),
                    selectionMode: action.selectionMode || 'any' // Preserve or default
                };
            }
        });
        
        // Determine which action handler to use
        let handlerToUse = this.actionHandler;
        
        // If no custom actionHandler provided, generate default handlers
        if (!handlerToUse) {
            this.generateActionHandlers();
            
            // Create a wrapper function that dispatches to the appropriate handler
            handlerToUse = (actionName, selectedRows) => {
                const handler = this.actionHandlers[actionName];
                if (handler && typeof handler === 'function') {
                    handler(selectedRows); // Pass selectedRows array
                } else {
                    console.warn(`No handler found for action: ${actionName}`);
                }
            };
        }
        
        // Create underlying DataGrid instance
        this.grid = new DataGrid(
            this.instanceName,
            this.containerId,
            this.columns,
            actionDefs,
            this.flowUrl,
            this.recordId,
            handlerToUse,
            this.fieldMapping  // Pass field mapping to DataGrid
        );
        
        // Initialize the grid
        await this.grid.initialize();
        
        console.log(`=== DataGridConfig ${this.title}: Render complete ===`);
    }
    
    /**
     * Updates the grid data by fetching from the data source
     * @returns {Promise<void>}
     */
    async update() {
        if (this.grid) {
            await this.grid.update();
        } else {
            console.warn(`=== DataGridConfig ${this.title}: Grid not initialized ===`);
        }
    }
    
    /**
     * Clears all data from the grid
     */
    clear() {
        if (this.grid) {
            this.grid.clear();
        } else {
            console.warn(`=== DataGridConfig ${this.title}: Grid not initialized ===`);
        }
    }
    
    /**
     * Exports grid data to CSV
     * @param {string} fileName - CSV file name
     */
    exportToCSV(fileName) {
        if (this.grid) {
            this.grid.exportToCSV(fileName || `${this.instanceName}.csv`);
        } else {
            console.warn(`=== DataGridConfig ${this.title}: Grid not initialized ===`);
        }
    }
    
    /**
     * Gets the currently selected row(s)
     * @returns {Array} Array of selected row data
     */
    getSelectedRow() {
        if (this.grid) {
            return this.grid.getSelectedRow();
        } else {
            console.warn(`=== DataGridConfig ${this.title}: Grid not initialized ===`);
            return [];
        }
    }
    
    /**
     * Gets all current grid data
     * @returns {Array} Array of all row data
     */
    getAllData() {
        if (this.grid) {
            return this.grid.getAllData();
        } else {
            console.warn(`=== DataGridConfig ${this.title}: Grid not initialized ===`);
            return [];
        }
    }
    
    /**
     * Provides direct access to the underlying AG Grid API
     * @returns {Object|null} AG Grid API object
     */
    getGridApi() {
        return this.grid ? this.grid.gridApi : null;
    }


}


/**
 * Front-end helper: fetch documents from server logic API.
 *
 * Parameters (pass as an object):
 *  - fields: string[] | undefined
 *      Array of column logical names to return.
 *      Example: ["cr5ad_name","createdon","gtdv1_documentcategory"]
 *
 *  - accountId: string | undefined
 *      GUID for primary filter on gtdv1_linktoaccount (use either accountId OR productId)
 *
 *  - productId: string | undefined
 *      GUID for primary filter on cr5ad_linktoproduct (use either productId OR accountId)
 *
 *  - category: number | undefined
 *      Optional integer secondary filter on gtdv1_documentcategory
 *
 *  - pageSize: number | undefined
 *      Page size (default handled by server: 50)
 *
 *  - page: number | undefined
 *      1-based page index (used only if pageToken not supplied)
 *
 *  - pageToken: string | undefined
 *      Continuation token from prior response.pagination.nextPageToken
 * 
 * Example URIs will be
 * 
 * 
 * A) Filter by account + category + specific fields
 *     /_api/serverlogics/v1-documents?accountId=<GUID>&category=2&fields=cr5ad_name,createdon
 *  B) Filter by product, default fields (all), page 2, 25 per page      
 *     /_api/serverlogics/v1-documents?productId=<GUID>&pageSize=25&page=2
 *  C) Continue paging using pageToken
 *     /_api/serverlogics/v1-documents?productId=<GUID>&pageSize=50&pageToken=<tokenFromPreviousResponse>
 * 
 */
function fetchDocuments({
  fields,
  accountId,
  productId,
  category,
  pageSize,
  page,
  pageToken
} = {}) {

  const params = new URLSearchParams();

  if (Array.isArray(fields) && fields.length) {
    params.set("fields", fields.join(","));
  }

  if (accountId) params.set("accountId", accountId);
  if (productId) params.set("productId", productId);

  if (Number.isInteger(category)) params.set("category", String(category));
  if (Number.isInteger(pageSize)) params.set("pageSize", String(pageSize));
  if (Number.isInteger(page)) params.set("page", String(page));
  if (pageToken) params.set("pageToken", pageToken);

  const url = "/_api/serverlogics/v1-documents" + (params.toString() ? ("?" + params.toString()) : "");

  return new Promise((resolve, reject) => {
    shell.safeAjax({
      type: "GET",
      url,
      contentType: "application/json",
      success: resolve,
      error: reject
    });
  });
}



// ============================================================================
// RTF Editor Class - Supports Multiple Instances
// ============================================================================

class RTFEditor {
  // Static method to ensure the image upload modal exists (called once, shared by all instances)
  static ensureImageUploadModal() {
    if (document.getElementById('imageUploadModal')) {
      return; // Modal already exists
    }
    
    const modalHtml = `
      <div class="modal fade" id="imageUploadModal" tabindex="-1" aria-labelledby="imageUploadModalLabel" aria-hidden="true">
        <div class="modal-dialog modal-dialog-centered">
          <div class="modal-content">
            <div class="modal-header">
              <h5 class="modal-title" id="imageUploadModalLabel">Image Upload Not Supported</h5>
              <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div class="modal-body">
              If you wish to upload an image then please upload this as a file and state that in this list. Uploading into the RTF is currently unsupported.
            </div>
            <div class="modal-footer">
              <button type="button" class="btn btn-primary" data-bs-dismiss="modal">OK</button>
            </div>
          </div>
        </div>
      </div>
    `;
    
    // Append modal to body
    document.body.insertAdjacentHTML('beforeend', modalHtml);
  }
  
  constructor(containerId, title = 'RTF Editor', initialRtfContent = '') {
    this.containerId = containerId;
    this.title = title;
    this.currentInstance = null; // Track which instance is being acted upon
    
    // Convert RTF to HTML if provided
    const initialHtml = initialRtfContent ? RTFEditor.rtfToHtml(initialRtfContent) : '';
    
    // Create the HTML structure
    this.createHTML(initialHtml);
    
    // Ensure the image upload modal exists (shared across all instances)
    RTFEditor.ensureImageUploadModal();
    
    // Bind event handlers
    this.bindEvents();
  }
  
  createHTML(initialContent) {
    const container = document.getElementById(this.containerId);
    if (!container) {
      throw new Error(`Container with id "${this.containerId}" not found`);
    }
    
    const instanceId = `rtf-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
    this.instanceId = instanceId;
    
    // Use provided content, or fall back to default only if content is null/undefined (not empty string)
    const defaultContent = (initialContent !== null && initialContent !== undefined && initialContent !== '') 
      ? initialContent 
      : `<p><b>Welcome!</b> This is a simple HTML editor that can export <u>RTF</u>.</p>
<p>Try <i>formatting</i> text, making lists, changing colors, etc.</p>
<ul><li>Bullet one</li><li>Bullet two</li></ul>`;
    
    container.innerHTML = `
      <div class="rtf-instance card shadow-sm">
        <div class="card-header bg-primary text-white">
          <h5 class="mb-0">${this.title}</h5>
        </div>
        <div class="card-body">
          <div class="toolbar btn-toolbar mb-3 p-2 bg-light border rounded" id="toolbar-${instanceId}" role="toolbar">
            <div class="btn-group me-2" role="group" aria-label="Text formatting">
              <button type="button" class="btn btn-sm btn-outline-secondary" data-cmd="bold" title="Bold"><b>B</b></button>
              <button type="button" class="btn btn-sm btn-outline-secondary" data-cmd="italic" title="Italic"><i>I</i></button>
              <button type="button" class="btn btn-sm btn-outline-secondary" data-cmd="underline" title="Underline"><u>U</u></button>
            </div>

            <div class="btn-group me-2" role="group" aria-label="Font size">
              <select class="form-select form-select-sm fontSize" style="width: auto;">
                <option value="">Font size</option>
                <option value="2">Small</option>
                <option value="3">Normal</option>
                <option value="4">Large</option>
                <option value="5">Larger</option>
                <option value="6">Huge</option>
                <option value="7">Gigantic</option>
              </select>
            </div>

            <div class="btn-group me-2" role="group" aria-label="Text color">
              <button type="button" class="btn btn-sm btn-outline-secondary colorBtn" title="Text Color">🎨 Color</button>
            </div>

            <div class="btn-group me-2" role="group" aria-label="Lists">
              <button type="button" class="btn btn-sm btn-outline-secondary" data-cmd="insertUnorderedList" title="Bullet List">• List</button>
              <button type="button" class="btn btn-sm btn-outline-secondary" data-cmd="insertOrderedList" title="Numbered List">1. List</button>
            </div>

            <div class="btn-group me-2" role="group" aria-label="Alignment">
              <button type="button" class="btn btn-sm btn-outline-secondary" data-cmd="justifyLeft" title="Align Left">⟸</button>
              <button type="button" class="btn btn-sm btn-outline-secondary" data-cmd="justifyCenter" title="Center">⇔</button>
              <button type="button" class="btn btn-sm btn-outline-secondary" data-cmd="justifyRight" title="Align Right">⟹</button>
            </div>

            <div class="btn-group me-2" role="group" aria-label="Insert">
              <button type="button" class="btn btn-sm btn-outline-secondary tableBtn" title="Insert Table">Table</button>
              <button type="button" class="btn btn-sm btn-outline-secondary linkBtn" title="Insert Link">Link</button>
              <button type="button" class="btn btn-sm btn-outline-secondary imageBtn" title="Insert Image">Image</button>
            </div>
            <input type="file" class="imageFileInput" accept="image/*" style="display: none;" />
          </div>

          <div class="rtf-editor border rounded p-3 mb-3 bg-white" contenteditable="true" id="editor-${instanceId}" style="min-height: 250px; max-height: 500px; overflow-y: auto;">
            ${defaultContent}
          </div>
        </div>
      </div>
    `;
    
    // Get references to elements
    this.editor = document.getElementById(`editor-${instanceId}`);
    this.toolbar = document.getElementById(`toolbar-${instanceId}`);
  }
  
  bindEvents() {
    // Toolbar buttons
    this.toolbar.addEventListener('click', (e) => {
      const btn = e.target.closest('button[data-cmd]');
      if (!btn) return;
      const cmd = btn.getAttribute('data-cmd');
      document.execCommand(cmd, false, null);
      this.editor.focus();
    });
    
    // Font size
    const fontSizeSelect = this.toolbar.querySelector('.fontSize');
    fontSizeSelect.addEventListener('change', (e) => {
      const v = e.target.value;
      if (v) document.execCommand('fontSize', false, v);
      this.editor.focus();
      e.target.value = '';
    });
    
    // Color button
    const colorBtn = this.toolbar.querySelector('.colorBtn');
    colorBtn.addEventListener('click', () => {
      RTFEditor.currentColorInstance = this;
      RTFEditor.showColorModal();
    });
    
    // Link button
    const linkBtn = this.toolbar.querySelector('.linkBtn');
    linkBtn.addEventListener('click', () => {
      const url = prompt('Enter URL (https://...)');
      if (url) document.execCommand('createLink', false, url);
      this.editor.focus();
    });
    
    // Table button
    const tableBtn = this.toolbar.querySelector('.tableBtn');
    tableBtn.addEventListener('click', () => {
      RTFEditor.currentTableInstance = this; // Set current instance for table creation
      RTFEditor.showTableModal();
    });
    
    // Image button
    const imageBtn = this.toolbar.querySelector('.imageBtn');
    const imageFileInput = this.toolbar.querySelector('.imageFileInput');
    imageBtn.addEventListener('click', () => {
      // Show Bootstrap modal instead of file input
      const imageModal = document.getElementById('imageUploadModal');
      if (imageModal) {
        const modal = new bootstrap.Modal(imageModal);
        modal.show();
      } else {
        // Fallback to alert if modal not found
        alert('If you wish to upload an image then please upload this as a file and state that in this list. Uploading into the RTF is currently unsupported.');
      }
    });
    
    imageFileInput.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;
      
      const reader = new FileReader();
      reader.onload = (event) => {
        const base64Data = event.target.result;
        const img = `<img src="${base64Data}" alt="Inserted image" />`;
        document.execCommand('insertHTML', false, img);
        this.editor.focus();
        e.target.value = '';
      };
      reader.readAsDataURL(file);
    });
  }
  
  // ============================================================================
  // Public API Methods
  // ============================================================================
  
  /**
   * Save method - returns RTF version of the content
   * Can be called programmatically to get RTF output
   */
  Save() {
    const html = this.editor.innerHTML;
    const rtf = RTFEditor.htmlToRtf(html);
    return rtf;
  }
  
  /**
   * Load method - accepts RTF string and converts to HTML
   */
  Load(rtfContent) {
    const html = RTFEditor.rtfToHtml(rtfContent);
    this.editor.innerHTML = html;
  }
  
  /**
   * Get HTML content
   */
  getHTML() {
    return this.editor.innerHTML;
  }
  
  /**
   * Set HTML content
   */
  setHTML(html) {
    this.editor.innerHTML = html;
  }
  
  /**
   * Get RTF content
   */
  getRTF() {
    const html = this.editor.innerHTML;
    return RTFEditor.htmlToRtf(html);
  }
  
  /**
   * Clear content
   */
  clear() {
    this.editor.innerHTML = '<p><br></p>';
  }
  
  // ============================================================================
  // Static Methods (Shared utilities)
  // ============================================================================
  
  static currentTableInstance = null;
  static modalInitialized = false;
  static colorModalInitialized = false;
  
  static initializeTableModal() {
    if (RTFEditor.modalInitialized) return;
    
    // Create modal HTML
    const modalHTML = `
      <div id="tableModal" class="modal" style="display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.4);">
        <div class="modal-content" style="background-color: #fff; margin: 15% auto; padding: 20px; border: 1px solid #888; border-radius: 8px; width: 300px;">
          <h3>Create Table</h3>
          <div>
            <label>Rows: <input type="number" id="tableRows" min="1" max="20" value="3" style="width: 60px; padding: 5px; margin: 5px;"></label>
          </div>
          <div>
            <label>Columns: <input type="number" id="tableCols" min="1" max="10" value="3" style="width: 60px; padding: 5px; margin: 5px;"></label>
          </div>
          <div style="margin-top: 15px;">
            <button id="createTableBtn" style="margin: 5px; padding: 8px 15px; border: 1px solid #ccc; background: #fff; border-radius: 6px; cursor: pointer;">Create</button>
            <button id="cancelTableBtn" style="margin: 5px; padding: 8px 15px; border: 1px solid #ccc; background: #fff; border-radius: 6px; cursor: pointer;">Cancel</button>
          </div>
        </div>
      </div>
    `;
    
    // Append modal to body
    document.body.insertAdjacentHTML('beforeend', modalHTML);
    
    // Initialize modal event listeners
    const tableModal = document.getElementById('tableModal');
    const createTableBtn = document.getElementById('createTableBtn');
    const cancelTableBtn = document.getElementById('cancelTableBtn');
    const tableRows = document.getElementById('tableRows');
    const tableCols = document.getElementById('tableCols');
    
    cancelTableBtn.addEventListener('click', () => {
      RTFEditor.hideTableModal();
    });
    
    createTableBtn.addEventListener('click', () => {
      const rows = parseInt(tableRows.value) || 3;
      const cols = parseInt(tableCols.value) || 3;
      
      let tableHTML = '<table class="table table-bordered table-striped table-hover">';
      for (let i = 0; i < rows; i++) {
        tableHTML += '<tr>';
        for (let j = 0; j < cols; j++) {
          tableHTML += '<td>&nbsp;</td>';
        }
        tableHTML += '</tr>';
      }
      tableHTML += '</table>';
      
      // Insert table into the stored instance
      if (RTFEditor.currentTableInstance) {
        RTFEditor.currentTableInstance.editor.focus();
        document.execCommand('insertHTML', false, tableHTML);
      }
      
      RTFEditor.hideTableModal();
    });
    
    // Close modal when clicking outside
    window.addEventListener('click', (event) => {
      if (event.target === tableModal) {
        RTFEditor.hideTableModal();
      }
    });
    
    RTFEditor.modalInitialized = true;
  }
  
  static showTableModal() {
    RTFEditor.initializeTableModal();
    const tableModal = document.getElementById('tableModal');
    tableModal.style.display = 'block';
  }
  
  static hideTableModal() {
    const tableModal = document.getElementById('tableModal');
    if (tableModal) {
      tableModal.style.display = 'none';
    }
  }

  // ============================================================================
  // Color Picker Modal
  // ============================================================================
  
  static initializeColorModal() {
    if (RTFEditor.colorModalInitialized) return;
    
    // Create color modal HTML
    const modalHTML = `
      <div id="colorModal" class="modal" style="display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.4);">
        <div class="modal-content" style="background-color: #fff; margin: 10% auto; padding: 20px; border: 1px solid #888; border-radius: 8px; width: 400px;">
          <h3 style="margin-top: 0;">Text Color Picker</h3>
          
          <div id="colorSelectionWarning" style="display: none; background: #fff3cd; border: 1px solid #ffc107; padding: 10px; border-radius: 4px; margin-bottom: 15px; color: #856404;">
            ⚠️ No text selected. Please select text before changing color.
          </div>
          
          <div style="margin-bottom: 20px;">
            <label style="display: block; margin-bottom: 10px; font-weight: bold;">Choose Color:</label>
            <input type="color" id="colorPickerInput" value="#000000" style="width: 100%; height: 60px; cursor: pointer; border: 1px solid #ccc; border-radius: 4px;">
          </div>
          
          <div style="margin-bottom: 20px;">
            <label style="display: block; margin-bottom: 10px; font-weight: bold;">RGB Values:</label>
            <div style="margin-bottom: 10px;">
              <label style="display: inline-block; width: 30px;">R:</label>
              <input type="range" id="colorR" min="0" max="255" value="0" style="width: 200px; margin-right: 10px;">
              <input type="number" id="colorRValue" min="0" max="255" value="0" style="width: 60px; padding: 5px;">
            </div>
            <div style="margin-bottom: 10px;">
              <label style="display: inline-block; width: 30px;">G:</label>
              <input type="range" id="colorG" min="0" max="255" value="0" style="width: 200px; margin-right: 10px;">
              <input type="number" id="colorGValue" min="0" max="255" value="0" style="width: 60px; padding: 5px;">
            </div>
            <div style="margin-bottom: 10px;">
              <label style="display: inline-block; width: 30px;">B:</label>
              <input type="range" id="colorB" min="0" max="255" value="0" style="width: 200px; margin-right: 10px;">
              <input type="number" id="colorBValue" min="0" max="255" value="0" style="width: 60px; padding: 5px;">
            </div>
          </div>
          
          <div style="margin-bottom: 20px;">
            <label style="display: block; margin-bottom: 5px; font-weight: bold;">Preview:</label>
            <div id="colorPreview" style="width: 100%; height: 50px; border: 2px solid #ccc; border-radius: 4px; background-color: #000000; display: flex; align-items: center; justify-content: center; color: white; font-weight: bold;">
              Sample Text
            </div>
          </div>
          
          <div style="margin-top: 15px; display: flex; gap: 10px;">
            <button id="applyColorBtn" style="flex: 1; padding: 10px 15px; border: 1px solid #28a745; background: #28a745; color: white; border-radius: 6px; cursor: pointer; font-weight: bold;">Apply Color</button>
            <button id="cancelColorBtn" style="flex: 1; padding: 10px 15px; border: 1px solid #ccc; background: #fff; border-radius: 6px; cursor: pointer;">Cancel</button>
          </div>
        </div>
      </div>
    `;
    
    // Append modal to body
    document.body.insertAdjacentHTML('beforeend', modalHTML);
    
    // Get modal elements
    const colorModal = document.getElementById('colorModal');
    const colorPickerInput = document.getElementById('colorPickerInput');
    const colorR = document.getElementById('colorR');
    const colorG = document.getElementById('colorG');
    const colorB = document.getElementById('colorB');
    const colorRValue = document.getElementById('colorRValue');
    const colorGValue = document.getElementById('colorGValue');
    const colorBValue = document.getElementById('colorBValue');
    const colorPreview = document.getElementById('colorPreview');
    const applyColorBtn = document.getElementById('applyColorBtn');
    const cancelColorBtn = document.getElementById('cancelColorBtn');
    const colorSelectionWarning = document.getElementById('colorSelectionWarning');
    
    // Function to convert RGB to Hex
    function rgbToHex(r, g, b) {
      return '#' + [r, g, b].map(x => {
        const hex = x.toString(16);
        return hex.length === 1 ? '0' + hex : hex;
      }).join('');
    }
    
    // Function to convert Hex to RGB
    function hexToRgb(hex) {
      const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
      return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
      } : null;
    }
    
    // Function to update preview
    function updatePreview() {
      const r = parseInt(colorR.value);
      const g = parseInt(colorG.value);
      const b = parseInt(colorB.value);
      const hex = rgbToHex(r, g, b);
      colorPreview.style.backgroundColor = hex;
      // Adjust text color for better visibility
      const brightness = (r * 299 + g * 587 + b * 114) / 1000;
      colorPreview.style.color = brightness > 128 ? '#000000' : '#ffffff';
    }
    
    // Sync RGB sliders with number inputs
    colorR.addEventListener('input', (e) => {
      colorRValue.value = e.target.value;
      const hex = rgbToHex(parseInt(colorR.value), parseInt(colorG.value), parseInt(colorB.value));
      colorPickerInput.value = hex;
      updatePreview();
    });
    
    colorG.addEventListener('input', (e) => {
      colorGValue.value = e.target.value;
      const hex = rgbToHex(parseInt(colorR.value), parseInt(colorG.value), parseInt(colorB.value));
      colorPickerInput.value = hex;
      updatePreview();
    });
    
    colorB.addEventListener('input', (e) => {
      colorBValue.value = e.target.value;
      const hex = rgbToHex(parseInt(colorR.value), parseInt(colorG.value), parseInt(colorB.value));
      colorPickerInput.value = hex;
      updatePreview();
    });
    
    // Sync number inputs with RGB sliders
    colorRValue.addEventListener('input', (e) => {
      let val = parseInt(e.target.value) || 0;
      val = Math.max(0, Math.min(255, val));
      colorR.value = val;
      colorRValue.value = val;
      const hex = rgbToHex(parseInt(colorR.value), parseInt(colorG.value), parseInt(colorB.value));
      colorPickerInput.value = hex;
      updatePreview();
    });
    
    colorGValue.addEventListener('input', (e) => {
      let val = parseInt(e.target.value) || 0;
      val = Math.max(0, Math.min(255, val));
      colorG.value = val;
      colorGValue.value = val;
      const hex = rgbToHex(parseInt(colorR.value), parseInt(colorG.value), parseInt(colorB.value));
      colorPickerInput.value = hex;
      updatePreview();
    });
    
    colorBValue.addEventListener('input', (e) => {
      let val = parseInt(e.target.value) || 0;
      val = Math.max(0, Math.min(255, val));
      colorB.value = val;
      colorBValue.value = val;
      const hex = rgbToHex(parseInt(colorR.value), parseInt(colorG.value), parseInt(colorB.value));
      colorPickerInput.value = hex;
      updatePreview();
    });
    
    // Color picker input updates RGB sliders
    colorPickerInput.addEventListener('input', (e) => {
      const rgb = hexToRgb(e.target.value);
      if (rgb) {
        colorR.value = rgb.r;
        colorG.value = rgb.g;
        colorB.value = rgb.b;
        colorRValue.value = rgb.r;
        colorGValue.value = rgb.g;
        colorBValue.value = rgb.b;
        updatePreview();
      }
    });
    
    // Apply button
    applyColorBtn.addEventListener('click', () => {
      if (RTFEditor.currentColorInstance) {
        const editor = RTFEditor.currentColorInstance.editor;
        editor.focus();
        
        // Check if there's a selection
        const selection = window.getSelection();
        if (!selection || selection.toString().trim() === '') {
          colorSelectionWarning.style.display = 'block';
          return;
        }
        
        const color = colorPickerInput.value;
        document.execCommand('foreColor', false, color);
        RTFEditor.hideColorModal();
      }
    });
    
    // Cancel button
    cancelColorBtn.addEventListener('click', () => {
      RTFEditor.hideColorModal();
    });
    
    // Close modal when clicking outside
    window.addEventListener('click', (event) => {
      if (event.target === colorModal) {
        RTFEditor.hideColorModal();
      }
    });
    
    RTFEditor.colorModalInitialized = true;
  }
  
  static showColorModal() {
    RTFEditor.initializeColorModal();
    const colorModal = document.getElementById('colorModal');
    const colorSelectionWarning = document.getElementById('colorSelectionWarning');
    const applyColorBtn = document.getElementById('applyColorBtn');
    
    // Hide warning initially
    colorSelectionWarning.style.display = 'none';
    
    // Check if text is selected when modal opens
    if (RTFEditor.currentColorInstance) {
      const editor = RTFEditor.currentColorInstance.editor;
      const selection = window.getSelection();
      const selectionText = selection.toString().trim();
      
      if (selectionText === '') {
        colorSelectionWarning.style.display = 'block';
        applyColorBtn.disabled = true;
        applyColorBtn.style.opacity = '0.5';
        applyColorBtn.style.cursor = 'not-allowed';
      } else {
        applyColorBtn.disabled = false;
        applyColorBtn.style.opacity = '1';
        applyColorBtn.style.cursor = 'pointer';
      }
    }
    
    colorModal.style.display = 'block';
  }
  
  static hideColorModal() {
    const colorModal = document.getElementById('colorModal');
    if (colorModal) {
      colorModal.style.display = 'none';
    }
  }
  
  static rtfToHtml(rtf) {
  // Basic RTF to HTML parser
  let html = '';
  let text = rtf;
  
  // Extract color table first before removing it
  const colorTable = [];
  const colorTableMatch = text.match(/\{\\colortbl([^}]*)\}/);
  if (colorTableMatch) {
    const colorDefs = colorTableMatch[1];
    // Split by semicolons and parse each color definition
    // Note: RTF color tables are 1-based, with entry 0 being auto/default (usually empty)
    const colors = colorDefs.split(';');
    colors.forEach(colorDef => {
      const redMatch = colorDef.match(/\\red(\d+)/);
      const greenMatch = colorDef.match(/\\green(\d+)/);
      const blueMatch = colorDef.match(/\\blue(\d+)/);
      
      if (redMatch && greenMatch && blueMatch) {
        const r = parseInt(redMatch[1]);
        const g = parseInt(greenMatch[1]);
        const b = parseInt(blueMatch[1]);
        colorTable.push(`rgb(${r}, ${g}, ${b})`);
      } else {
        colorTable.push(''); // Empty entry for default/auto color
      }
    });
  }
  
  // Remove RTF header and common control groups
  text = text.replace(/^\{\\rtf1\\[a-z]+\\[a-z]+\d*\s*/, ''); // Remove header like {\rtf1\ansi\deff0
  text = text.replace(/\{\\fonttbl[^}]*\}\s*/g, ''); // Remove font table
  text = text.replace(/\{\\colortbl[^}]*\}\s*/g, ''); // Remove color table
  text = text.replace(/\}\s*$/, ''); // Remove closing brace at the end
  
  // State tracking
  let bold = false, italic = false, underline = false;
  let fontSize = null; // Track current font size (in half-points, null = default)
  let currentColor = null; // Track current color index
  let inParagraph = false;
  let inList = false;
  let inListItem = false;
  let isOrderedList = false; // Track if current list is ordered (OL) or unordered (UL)
  let inTable = false; // Track if we're inside a table
  let inTableCell = false; // Track if we're inside a table cell
  let currentRow = []; // Current table row cells
  let tableRows = []; // All table rows
  let cellBuffer = []; // Buffer for current cell content
  let result = [];
  
  // Helper function to add content to the appropriate buffer
  const addContent = (content) => {
    if (inTableCell) {
      cellBuffer.push(content);
    } else {
      result.push(content);
    }
  };
  
  // Stack for handling RTF groups (braces)
  const stateStack = [];
  
  // Parse RTF control words and text
  // Updated to handle \uN? unicode sequences and table commands
  // Use non-capturing groups (?:...) to avoid numbers appearing as separate tokens
  const tokens = text.split(/(\\u-?\d+\?|\\[a-z]+(?:-?\d+)?\s*|\{|\}|\\par|\\line|\\bullet|\\tab|\\trowd|\\cell|\\row|\\intbl)/);
  
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (!token) continue;
    
    if (token.startsWith('\\')) {
      // Handle unicode sequences first (\uN?)
      const unicodeMatch = token.match(/\\u(-?\d+)\?/);
      if (unicodeMatch) {
        const code = parseInt(unicodeMatch[1]);
        const char = String.fromCharCode(code);
        if (!inTableCell && !inParagraph && !inListItem && !inList) {
          addContent('<p>');
          inParagraph = true;
        }
        addContent(char.replace(/</g, '&lt;').replace(/>/g, '&gt;'));
        continue;
      }
      
      const cmd = token.match(/\\([a-z]+)(-?\d+)?/);
      if (!cmd) continue;
      
      const command = cmd[1];
      const param = cmd[2] ? parseInt(cmd[2]) : null;
      
      // List of RTF commands to ignore (structural/formatting we don't use)
      const ignoredCommands = ['f', 'ql', 'qr', 'qc', 'qj', 'deff', 'ansi', 
                               'fonttbl', 'colortbl', 'red', 'green', 'blue', 
                               'viewkind', 'uc', 'pard', 'sa', 'sb', 'fi', 'li', 'ri',
                               'trgaph', 'trleft', 'clbrdrt', 'clbrdrl', 'clbrdrb', 'clbrdrr',
                               'brdrs', 'brdrw', 'cellx'];
      
      if (ignoredCommands.includes(command)) {
        continue; // Skip these commands
      }
      
      switch (command) {
        case 'fs':
          // Font size in half-points (e.g., \fs24 = 12pt, \fs72 = 36pt)
          if (param !== null && param > 0) {
            // Close previous font size span if any
            if (fontSize !== null) {
              addContent('</span>');
            }
            // Convert half-points to points
            const pointSize = param / 2;
            addContent(`<span style="font-size: ${pointSize}pt;">`);
            fontSize = param;
          } else if (param === 0 || param === null) {
            // Reset to default font size
            if (fontSize !== null) {
              addContent('</span>');
              fontSize = null;
            }
          }
          break;
        case 'cf':
          // Color foreground - close previous color span if any
          if (currentColor !== null) {
            addContent('</span>');
          }
          // Apply new color if param is valid and color exists in table
          // RTF color table: split by ';' gives [' ', 'color1', 'color2', ...]
          // so \cf1 means colorTable[1], \cf2 means colorTable[2], etc.
          if (param !== null && param > 0 && colorTable[param] && colorTable[param] !== '') {
            addContent(`<span style="color: ${colorTable[param]};">`);
            currentColor = param;
          } else {
            currentColor = null;
          }
          break;
        case 'b':
          if (param === 0) {
            if (bold) { addContent('</b>'); bold = false; }
          } else {
            if (!bold) { addContent('<b>'); bold = true; }
          }
          break;
        case 'i':
          if (param === 0) {
            if (italic) { addContent('</i>'); italic = false; }
          } else {
            if (!italic) { addContent('<i>'); italic = true; }
          }
          break;
        case 'ul':
          if (param === 0) {
            if (underline) { addContent('</u>'); underline = false; }
          } else {
            if (!underline) { addContent('<u>'); underline = true; }
          }
          break;
        case 'par':
          // Skip \par if we're in a table - tables have their own structure
          if (inTable) {
            // Check if we're processing content after a table (table completed)
            // If tableRows has content and we're not in a cell, the table is complete
            if (tableRows.length > 0 && !inTableCell) {
              // Output the complete table
              result.push('<table>');
              for (const row of tableRows) {
                result.push('<tr>');
                for (const cell of row) {
                  result.push('<td>' + cell + '</td>');
                }
                result.push('</tr>');
              }
              result.push('</table>');
              // Reset table state
              inTable = false;
              tableRows = [];
              currentRow = [];
            } else if (inTableCell) {
              // We're in a table cell - handle list items and formatting within the cell
              if (inListItem) {
                if (bold) { addContent('</b>'); bold = false; }
                if (italic) { addContent('</i>'); italic = false; }
                if (underline) { addContent('</u>'); underline = false; }
                if (fontSize !== null) { addContent('</span>'); fontSize = null; }
                if (currentColor !== null) { addContent('</span>'); currentColor = null; }
                addContent('</li>');
                inListItem = false;
              }
            }
            break;
          }
          // Close any open list item
          if (inListItem) {
            if (bold) { addContent('</b>'); bold = false; }
            if (italic) { addContent('</i>'); italic = false; }
            if (underline) { addContent('</u>'); underline = false; }
            if (fontSize !== null) { addContent('</span>'); fontSize = null; }
            if (currentColor !== null) { addContent('</span>'); currentColor = null; }
            addContent('</li>');
            inListItem = false;
          } else if (inList) {
            // We're in a list but not in a list item, so close the list
            addContent(isOrderedList ? '</ol>' : '</ul>');
            inList = false;
            isOrderedList = false;
            // Now start a new paragraph
            addContent('<p>');
            inParagraph = true;
          } else {
            if (bold) { addContent('</b>'); bold = false; }
            if (italic) { addContent('</i>'); italic = false; }
            if (underline) { addContent('</u>'); underline = false; }
            if (fontSize !== null) { addContent('</span>'); fontSize = null; }
            if (currentColor !== null) { addContent('</span>'); currentColor = null; }
            if (inParagraph) {
              addContent('</p>');
              inParagraph = false;
            }
            addContent('<p>');
            inParagraph = true;
          }
          break;
        case 'line':
          addContent('<br>');
          break;
        case 'tab':
          // Check if this tab follows a number pattern (ordered list)
          // Look back in the appropriate buffer to see if we just added a number followed by '.'
          let foundNumberPattern = false;
          const targetBuffer = inTableCell ? cellBuffer : result;
          if (targetBuffer.length >= 1) {
            // Check the last item in buffer for number pattern
            const lastItem = targetBuffer[targetBuffer.length - 1];
            if (typeof lastItem === 'string') {
              const numberedMatch = lastItem.match(/^(\d+)\.\s*$/);
              if (numberedMatch) {
                // Remove the number and period we just added
                targetBuffer.pop();
                foundNumberPattern = true;
              }
            }
          }
          
          if (foundNumberPattern) {
              
              // Start ordered list if not already in one, or switch from UL to OL
              if (!inList || !isOrderedList) {
                // Close unordered list if switching
                if (inList && !isOrderedList) {
                  if (inListItem) {
                    if (bold) { addContent('</b>'); bold = false; }
                    if (italic) { addContent('</i>'); italic = false; }
                    if (underline) { addContent('</u>'); underline = false; }
                    if (fontSize !== null) { addContent('</span>'); fontSize = null; }
                    if (currentColor !== null) { addContent('</span>'); currentColor = null; }
                    addContent('</li>');
                    inListItem = false;
                  }
                  addContent('</ul>');
                }
                
                if (inParagraph) {
                  addContent('</p>');
                  inParagraph = false;
                }
                addContent('<ol>');
                inList = true;
                isOrderedList = true;
              }
              // Close previous list item if open
              if (inListItem) {
                if (bold) { addContent('</b>'); bold = false; }
                if (italic) { addContent('</i>'); italic = false; }
                if (underline) { addContent('</u>'); underline = false; }
                if (fontSize !== null) { addContent('</span>'); fontSize = null; }
                if (currentColor !== null) { addContent('</span>'); currentColor = null; }
                addContent('</li>');
              }
              // Start new list item
              addContent('<li>');
              inListItem = true;
            } else {
              // Skip tabs that immediately follow bullets (they're just RTF formatting)
              // Only output tabs in regular paragraph content
              if (!inListItem) {
                addContent('&nbsp;&nbsp;&nbsp;&nbsp;');
              }
            }
          break;
        case 'bullet':
          // Start list if not already in one
          if (!inList || isOrderedList) {
            // Close ordered list if switching
            if (inList && isOrderedList) {
              if (inListItem) {
                if (bold) { addContent('</b>'); bold = false; }
                if (italic) { addContent('</i>'); italic = false; }
                if (underline) { addContent('</u>'); underline = false; }
                if (fontSize !== null) { addContent('</span>'); fontSize = null; }
                if (currentColor !== null) { addContent('</span>'); currentColor = null; }
                addContent('</li>');
                inListItem = false;
              }
              addContent('</ol>');
            }
            
            if (inParagraph) {
              addContent('</p>');
              inParagraph = false;
            }
            addContent('<ul>');
            inList = true;
            isOrderedList = false;
          }
          // Close previous list item if open
          if (inListItem) {
            if (bold) { addContent('</b>'); bold = false; }
            if (italic) { addContent('</i>'); italic = false; }
            if (underline) { addContent('</u>'); underline = false; }
            if (fontSize !== null) { addContent('</span>'); fontSize = null; }
            if (currentColor !== null) { addContent('</span>'); currentColor = null; }
            addContent('</li>');
          }
          // Start new list item
          addContent('<li>');
          inListItem = true;
          break;
        case 'trowd':
          // Table row definition - marks start of a new table or row
          if (!inTable) {
            // Starting a new table
            if (inParagraph) {
              result.push('</p>');
              inParagraph = false;
            }
            inTable = true;
            tableRows = [];
          }
          // Start a new row
          currentRow = [];
          break;
        case 'intbl':
          // Inside table marker - indicates we're in a cell
          if (inTable && !inTableCell) {
            inTableCell = true;
            cellBuffer = [];
          }
          break;
        case 'cell':
          // End of table cell - close any open lists/formatting and save the cell content
          if (inTableCell) {
            // Close any open formatting
            if (bold) { cellBuffer.push('</b>'); bold = false; }
            if (italic) { cellBuffer.push('</i>'); italic = false; }
            if (underline) { cellBuffer.push('</u>'); underline = false; }
            if (fontSize !== null) { cellBuffer.push('</span>'); fontSize = null; }
            if (currentColor !== null) { cellBuffer.push('</span>'); currentColor = null; }
            // Close any open list item
            if (inListItem) {
              cellBuffer.push('</li>');
              inListItem = false;
            }
            // Close any open list
            if (inList) {
              cellBuffer.push(isOrderedList ? '</ol>' : '</ul>');
              inList = false;
              isOrderedList = false;
            }
            currentRow.push(cellBuffer.join(''));
            cellBuffer = [];
            inTableCell = false;
          }
          break;
        case 'row':
          // End of table row
          if (inTable && currentRow.length > 0) {
            tableRows.push([...currentRow]);
            currentRow = [];
          }
          break;
      }
    } else if (token === '{') {
      // Push current state onto stack when entering a group
      stateStack.push({
        bold: bold,
        italic: italic,
        underline: underline,
        fontSize: fontSize,
        currentColor: currentColor
      });
    } else if (token === '}') {
      // Pop state when exiting a group
      if (stateStack.length > 0) {
        const prevState = stateStack.pop();
        
        // Close any tags that were opened in this group
        if (bold && !prevState.bold) { addContent('</b>'); }
        if (italic && !prevState.italic) { addContent('</i>'); }
        if (underline && !prevState.underline) { addContent('</u>'); }
        
        // Handle font size changes
        if (fontSize !== prevState.fontSize) {
          if (fontSize !== null) { addContent('</span>'); }
          if (prevState.fontSize !== null) {
            const pointSize = prevState.fontSize / 2;
            addContent(`<span style="font-size: ${pointSize}pt;">`);
          }
        }
        
        // Handle color changes
        if (currentColor !== prevState.currentColor) {
          if (currentColor !== null) { addContent('</span>'); }
          if (prevState.currentColor !== null) {
            addContent(`<span style="color: ${colorTable[prevState.currentColor]};">`);
          }
        }
        
        // Restore previous state
        bold = prevState.bold;
        italic = prevState.italic;
        underline = prevState.underline;
        fontSize = prevState.fontSize;
        currentColor = prevState.currentColor;
      }
    } else {
      // Regular text - clean up RTF escape sequences
      let cleanText = token
        .replace(/\\u(-?\d+)\?/g, (m, code) => String.fromCharCode(parseInt(code))) // unicode - must be first!
        .replace(/\\'/g, "'")  // escaped apostrophe
        .replace(/\\\\/g, '\\')  // escaped backslash
        .replace(/\\\{/g, '{')  // escaped brace
        .replace(/\\\}/g, '}');  // escaped brace
      
      // Only skip if the token is entirely whitespace/empty
      if (cleanText && cleanText.trim()) {
        // Don't open a paragraph if we're in a list item, in a list, or in a table cell
        if (!inParagraph && !inListItem && !inList && !inTableCell && cleanText.length > 0) {
          addContent('<p>');
          inParagraph = true;
        }
        addContent(cleanText.replace(/</g, '&lt;').replace(/>/g, '&gt;'));
      }
    }
  }
  
  // Close any open tags
  if (bold) result.push('</b>');
  if (italic) result.push('</i>');
  if (underline) result.push('</u>');
  if (fontSize !== null) result.push('</span>');
  if (currentColor !== null) result.push('</span>');
  if (inListItem) {
    result.push('</li>');
    inListItem = false;
  }
  if (inList) {
    result.push(isOrderedList ? '</ol>' : '</ul>');
    inList = false;
    isOrderedList = false;
  }
  // Output any remaining table
  if (inTable && tableRows.length > 0) {
    result.push('<table>');
    for (const row of tableRows) {
      result.push('<tr>');
      for (const cell of row) {
        result.push('<td>' + cell + '</td>');
      }
      result.push('</tr>');
    }
    result.push('</table>');
  }
  if (inParagraph) result.push('</p>');
  
  html = result.join('');
  
  // If no content, return a default paragraph
  if (!html.trim()) {
    html = '<p><br></p>';
  }
  
  return html;
}


static htmlToRtf(html) {
  const doc = new DOMParser().parseFromString(html, 'text/html');

  // Color table: collect unique text colors used
  const colorSet = new Map(); // key: rgb string -> index
  let colorIndex = 1; // RTF color table is 1-based
  function addColor(rgb) {
    if (!rgb) return 0;
    if (!colorSet.has(rgb)) colorSet.set(rgb, colorIndex++);
    return colorSet.get(rgb);
  }
  // Convert CSS color like 'rgb(255, 0, 0)' or '#rrggbb' to r,g,b
  function parseColorToRgbTriplet(value) {
    if (!value) return null;
    
    // Handle hex colors directly
    if (value.startsWith('#')) {
      const hex = value.substring(1);
      if (hex.length === 6) {
        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);
        if (!isNaN(r) && !isNaN(g) && !isNaN(b)) {
          return { r, g, b };
        }
      } else if (hex.length === 3) {
        const r = parseInt(hex[0] + hex[0], 16);
        const g = parseInt(hex[1] + hex[1], 16);
        const b = parseInt(hex[2] + hex[2], 16);
        if (!isNaN(r) && !isNaN(g) && !isNaN(b)) {
          return { r, g, b };
        }
      }
    }
    
    // Handle rgb(r, g, b) format
    const m = value.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
    if (m) {
      return { r: +m[1], g: +m[2], b: +m[3] };
    }
    
    return null;
  }

  // Escape RTF special chars and unicode
  function escapeText(s) {
    return s.replace(/[\\{}]/g, m => '\\' + m)
            .replace(/\n/g, '\\par ');
  }
  function unicodeToRtf(s) {
    // convert string to RTF unicode (\uN?)
    let out = '';
    for (const ch of s) {
      const code = ch.codePointAt(0);
      if (code < 128) {
        out += escapeText(ch);
      } else {
        // \uN? where N is signed 16-bit. Use fallback '?'
        const signed = (code & 0x7FFF) - (code & 0x8000);
        out += `\\u${signed}?`;
      }
    }
    return out;
  }

  function styleToRtfOpenClose(el) {
    // Return arrays of [openTags, closeTags] based on inline styles and tags
    const open = [];
    const close = [];

    const tag = el.tagName;

    // Bold/Italic/Underline via tag or style
    const isB = tag === 'B' || tag === 'STRONG';
    const isI = tag === 'I' || tag === 'EM';
    const isU = tag === 'U';

    if (isB) { open.push('\\b '); close.unshift('\\b0 '); }
    if (isI) { open.push('\\i '); close.unshift('\\i0 '); }
    if (isU) { open.push('\\ul '); close.unshift('\\ul0 '); }

    // Text color - check color attribute (for FONT tags) and style attribute
    let colorValue = null;
    if (tag === 'FONT' && el.hasAttribute('color')) {
      colorValue = el.getAttribute('color');
    } else if (el.style && el.style.color) {
      colorValue = el.style.color;
    }
    
    if (colorValue) {
      const rgb = parseColorToRgbTriplet(colorValue);
      if (rgb) {
        const key = `rgb(${rgb.r},${rgb.g},${rgb.b})`;
        const idx = addColor(key);
        if (idx > 0) open.push(`\\cf${idx} `);
      }
    }

    // Font size - check size attribute (for FONT tags) and style attribute
    let fontSizeValue = null;
    if (tag === 'FONT' && el.hasAttribute('size')) {
      // HTML font sizes are 1-7, convert to approximate point sizes
      const htmlSize = parseInt(el.getAttribute('size'));
      const sizeMap = { 1: 8, 2: 10, 3: 12, 4: 14, 5: 18, 6: 24, 7: 36 };
      fontSizeValue = sizeMap[htmlSize] || 12;
    } else if (el.style && el.style.fontSize) {
      // Parse CSS font-size (e.g., "18px", "2em", "large")
      const fontSize = el.style.fontSize;
      if (fontSize.endsWith('px')) {
        fontSizeValue = parseInt(fontSize);
      } else if (fontSize.endsWith('pt')) {
        fontSizeValue = parseInt(fontSize);
      } else if (fontSize.endsWith('em')) {
        fontSizeValue = Math.round(parseFloat(fontSize) * 12); // Assume base 12pt
      } else {
        // Named sizes
        const namedSizes = {
          'xx-small': 8, 'x-small': 10, 'small': 12, 
          'medium': 14, 'large': 18, 'x-large': 24, 'xx-large': 36
        };
        fontSizeValue = namedSizes[fontSize] || null;
      }
    }
    
    if (fontSizeValue && fontSizeValue > 0) {
      // RTF font size is in half-points, so multiply by 2
      const halfPoints = fontSizeValue * 2;
      open.push(`\\fs${halfPoints} `);
    }

    // Alignment for block elements - check style attribute
    if (tag === 'P' || tag === 'DIV' || tag === 'LI') {
      const ta = el.style.textAlign || el.getAttribute('align');
      if (ta === 'center') { open.push('\\qc '); }
      else if (ta === 'right') { open.push('\\qr '); }
      else if (ta === 'justify') { open.push('\\qj '); }
      else { open.push('\\ql '); }
    }

    return [open.join(''), close.join('')];
  }

  function convertNode(node) {
    if (node.nodeType === Node.TEXT_NODE) {
      return unicodeToRtf(node.nodeValue || '');
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return '';

    const el = node;
    const tag = el.tagName;

    // Handle line/paragraph structure
    let rtf = '';

    // Lists
    if (tag === 'UL' || tag === 'OL') {
      const isOrdered = tag === 'OL';
      let counter = 1;
      for (const li of el.children) {
        if (li.tagName !== 'LI') continue;
        const [open, close] = styleToRtfOpenClose(li);
        const marker = isOrdered ? `${counter}. \\tab ` : '\\bullet\\tab ';
        // Convert child nodes, but skip paragraph markers for P tags inside LI
        const content = Array.from(li.childNodes).map(node => {
          if (node.nodeType === Node.ELEMENT_NODE && (node.tagName === 'P' || node.tagName === 'DIV')) {
            // For P/DIV inside LI, just get the content without the paragraph break
            return Array.from(node.childNodes).map(convertNode).join('');
          }
          return convertNode(node);
        }).join('');
        rtf += `{${open}${marker}${content}${close}}\\par\n`;
        if (isOrdered) counter++;
      }
      return rtf;
    }

    // Tables
    if (tag === 'TABLE') {
      rtf += '\\par\n'; // paragraph before table
      for (const row of el.querySelectorAll('tr')) {
        const cells = row.querySelectorAll('td, th');
        const cellCount = cells.length;
        
        // RTF table row starts with row definition
        let rowDef = '\\trowd\\trgaph108\\trleft-108'; // table row defaults
        
        // Define cell boundaries (cumulative widths in twips, 1 inch = 1440 twips)
        const cellWidth = Math.floor(6480 / cellCount); // ~4.5 inches total width
        for (let i = 0; i < cellCount; i++) {
          const rightBoundary = (i + 1) * cellWidth;
          rowDef += `\\clbrdrt\\brdrs\\brdrw10\\clbrdrl\\brdrs\\brdrw10\\clbrdrb\\brdrs\\brdrw10\\clbrdrr\\brdrs\\brdrw10\\cellx${rightBoundary}`;
        }
        
        rtf += rowDef;
        
        // Add cell content
        for (const cell of cells) {
          const cellText = Array.from(cell.childNodes).map(convertNode).join('');
          rtf += `\\pard\\intbl ${cellText}\\cell\n`;
        }
        
        rtf += '\\row\n';
      }
      rtf += '\\pard\\par\n'; // paragraph after table
      return rtf;
    }

    // Line break
    if (tag === 'BR') return '\\line ';
    // Images: embed as PNG in RTF
    if (tag === 'IMG') {
      const src = el.getAttribute('src');
      if (src && src.startsWith('data:image/')) {
        // Extract base64 data and convert to hex for RTF
        const base64Match = src.match(/^data:image\/[^;]+;base64,(.+)$/);
        if (base64Match) {
          try {
            const base64 = base64Match[1];
            const binaryString = atob(base64);
            let hexData = '';
            for (let i = 0; i < binaryString.length; i++) {
              const hex = binaryString.charCodeAt(i).toString(16).padStart(2, '0');
              hexData += hex;
            }
            
            // Get image dimensions if available
            const width = el.naturalWidth || el.width || 200;
            const height = el.naturalHeight || el.height || 200;
            
            // Convert pixels to twips (1 pixel ≈ 15 twips for 96 DPI)
            const widthTwips = Math.round(width * 15);
            const heightTwips = Math.round(height * 15);
            
            // RTF picture syntax
            rtf += `{\\pict\\pngblip\\picw${width}\\pich${height}\\picwgoal${widthTwips}\\pichgoal${heightTwips} ${hexData}}\\par\n`;
            return rtf;
          } catch (err) {
            console.error('Error converting image to RTF:', err);
            return ''; // Skip image if conversion fails
          }
        }
      }
      return '';
    }
    // FONT tag: handle color and other attributes
    if (tag === 'FONT') {
      const [open, close] = styleToRtfOpenClose(el);
      const text = Array.from(el.childNodes).map(convertNode).join('');
      if (open || close) {
        return `{${open}${text}${close}}`;
      }
      return text;
    }
    
    // Links: we'll keep the text, and append URL in parentheses (simple approach)
    if (tag === 'A') {
      const [open, close] = styleToRtfOpenClose(el);
      const text = Array.from(el.childNodes).map(convertNode).join('');
      const href = el.getAttribute('href') || '';
      const hrefText = href ? ` ${unicodeToRtf('(' + href + ')')}` : '';
      return `{${open}${text}${hrefText}${close}}`;
    }

    // Block elements: P, DIV, H1..H3 (basic mapping)
    const blockLike = ['P','DIV','H1','H2','H3'];
    if (blockLike.includes(tag)) {
      const [open, close] = styleToRtfOpenClose(el);
      let prefix = '';
      if (tag === 'H1') prefix = '\\fs48 \\b ';
      if (tag === 'H2') prefix = '\\fs36 \\b ';
      if (tag === 'H3') prefix = '\\fs28 \\b ';
      rtf += `{${open}${prefix}${Array.from(el.childNodes).map(convertNode).join('')}${close}}\\par\n`;
      return rtf;
    }

    // Inline: B, I, U, SPAN, STRONG, EM, and other inline elements
    const [open, close] = styleToRtfOpenClose(el);
    const childContent = Array.from(el.childNodes).map(convertNode).join('');
    // Wrap in RTF group if there are any formatting codes
    if (open || close) {
      return `{${open}${childContent}${close}}`;
    }
    return childContent;
  }

  // Build body RTF
  const bodyRtf = Array.from(doc.body.childNodes).map(convertNode).join('');

  // Build color table
  let colorTbl = '';
  if (colorSet.size) {
    const entries = Array.from(colorSet.keys()).map(rgb => {
      const m = rgb.match(/^rgb\((\d+),(\d+),(\d+)\)$/);
      const r = +m[1], g = +m[2], b = +m[3];
      return `\\red${r}\\green${g}\\blue${b};`;
    }).join('');
    colorTbl = `{\\colortbl ;${entries}}\n`;
  }

  // RTF header/footer
  const rtf =
    `{\\rtf1\\ansi\\deff0\n` +
    `{\\fonttbl{\\f0 Arial;}}\n` +
    (colorTbl ? colorTbl : '') +
    `\\fs24 \\f0\n` +
    bodyRtf +
    `\n}`;


  return rtf;
}
}
