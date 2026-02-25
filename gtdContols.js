//////////////////////////////////////////////////////
//DocMan Controls - used to manage the upload of documents
//this section is the back end. there is a septarte class for the pages
//////////////////////////////////////////////////////
const GTDCTRLVer = "12.03";
    
    const DocManVer = "14.03";
    
    function _UpdateDocCategory(DocCat, passedGUID) {
            console.log("UpdateDocCategory: called ver=" + DocManVer);
            const flowUrl = "/_api/cloudflow/v1.0/trigger/4c999c4e-3495-f011-b41b-002248c5d639";
            const payload = {
              ActionType: 'UpdateDocCat', // group upload: adds multiple files to one group
              DocCategory: DocCat,
              AccGUID: '',
              DocManRowGuid: passedGUID
            };

            _GenericFlowHandler(flowUrl, payload)
            .done(function (response) {
              console.log("UpdateCat: Passed-GUID-"+  passedGUID+ " -response:" + response);
            })
            .fail(function (error) {
              console.log("UpdateCat: Error-GUID-"+  passedGUID+ " -response:" + error);
            })
            .always(function () {  
            });

            // const requestObj = {
            //   eventData: JSON.stringify(payload)
            // };

            // shell.ajaxSafePost({
            //   type: "POST",
            //   contentType: "application/json",
            //   url: "/_api/cloudflow/v1.0/trigger/4c999c4e-3495-f011-b41b-002248c5d639",
            //   data: JSON.stringify(requestObj),
            //   processData: false,
            //   global: false
            // })
            // .done(function (response) {
            //   console.log("Flow response for DocMan row", passedGUID);
            // })
            // .fail(function (error) {
            //   console.error("Flow response for DocMan row", passedGUID);
            // })
            // .always(function () {

            // });

    }

    /**
    * SHARED FUNCTION: extract SAS token/url from flow response.
    * @param {any} flowResponse
    * @returns {string}
    */
    function extractSasToken(flowResponse) {
    var response = flowResponse;

    if (typeof response === 'string') {
      try {
        response = JSON.parse(response);
      } catch (e) {
        if (response.indexOf('sig=') !== -1 || response.indexOf('?sv=') !== -1) {
          return response;
        }
        return '';
      }
    }

    if (!response || typeof response !== 'object') {
      return '';
    }

    var directCandidates = [
      response.sasToken,
      response.sas,
      response.token,
      response.sas_token,
      response.SASToken,
      response.url,
      response.sasUrl
    ];

    for (var i = 0; i < directCandidates.length; i++) {
      if (typeof directCandidates[i] === 'string' && directCandidates[i]) {
        return directCandidates[i];
      }
    }

    var keys = Object.keys(response);
    for (var j = 0; j < keys.length; j++) {
      var value = response[keys[j]];
      if (typeof value === 'string' && (value.indexOf('sig=') !== -1 || value.indexOf('?sv=') !== -1)) {
        return value;
      }
    }

    return '';
    }

    /**
    * SHARED FUNCTION: parse CorrelationId and Error ID from Power Pages HTML error page.
    * @param {string} responseText
    * @returns {{correlationId:string, errorId:string, flowTrackingId:string, hasPowerAutomateErrorView:boolean}}
    * SHOULD THIS FUNCTION BE REMOVED. NOT SURE THAT ITS IS BEING USED	 
    */
    function _ParsePowerPagesErrorMetadata(responseText) {
    var html = responseText || '';
    var correlationId = '';
    var errorId = '';
    var flowTrackingId = '';
    var parsedJson = null;

    if (html && (html.charAt(0) === '{' || html.charAt(0) === '[')) {
      try {
        parsedJson = JSON.parse(html);
      } catch (jsonParseError) {
        parsedJson = null;
      }
    }

    if (parsedJson && typeof parsedJson === 'object') {
      var messageText = '';
      if (typeof parsedJson.Message === 'string') {
        messageText = parsedJson.Message;
      } else if (typeof parsedJson.message === 'string') {
        messageText = parsedJson.message;
      }

      var trackingMatchFromJson = /tracking\s+id\s+is\s+["'\u0027]*([0-9a-fA-F-]{36})/i.exec(messageText);
      if (trackingMatchFromJson && trackingMatchFromJson[1]) {
        flowTrackingId = trackingMatchFromJson[1];
      }
    }

    var correlationMatch = /correlationId:\s*'([0-9a-fA-F-]{36})'/i.exec(html)
      || /Error\s*ID\s*#\s*\[([0-9a-fA-F-]{36})\]/i.exec(html)
      || /CorrelationId\s*:\s*([0-9a-fA-F-]{36})/i.exec(html);

    if (correlationMatch && correlationMatch[1]) {
      correlationId = correlationMatch[1];
    }

    var errorIdMatch = /Error\s*ID\s*#\s*\[([^\]]+)\]/i.exec(html);
    if (errorIdMatch && errorIdMatch[1]) {
      errorId = errorIdMatch[1];
    }

    if (!flowTrackingId) {
      var trackingMatchFromText = /tracking\s+id\s+is\s+["'\u0027]*([0-9a-fA-F-]{36})/i.exec(html);
      if (trackingMatchFromText && trackingMatchFromText[1]) {
        flowTrackingId = trackingMatchFromText[1];
      }
    }

    if (!errorId && correlationId) {
      errorId = correlationId;
    }

    return {
      correlationId: correlationId,
      errorId: errorId,
      flowTrackingId: flowTrackingId,
      hasPowerAutomateErrorView: html.indexOf('~/Areas/PowerAutomate/Views/PowerAutomate/Error') !== -1
    };
    }

    /**
      * SHARED FUNCTION: preflight diagnostics for cloud flow calls.
      * Logs request context and common configuration checks before invoking the endpoint.
      * @param {string} flowUrl
      * @param {any} payload
      * @returns {object}
      */
    function _PreflightCloudFlowCheck(flowUrl, payload) {
      var portal = window.Microsoft && window.Microsoft.Dynamic365 && window.Microsoft.Dynamic365.Portal
        ? window.Microsoft.Dynamic365.Portal
        : null;
      var flowMatch = /^\/_api\/cloudflow\/v1\.0\/trigger\/([0-9a-fA-F-]{36})$/.exec(flowUrl || '');
      var triggerId = flowMatch ? flowMatch[1] : '';
      var payloadKeys = payload && typeof payload === 'object' ? Object.keys(payload) : [];

      var preflight = {
        timestamp: new Date().toISOString(),
        flowUrl: flowUrl,
        triggerId: triggerId,
        isCloudFlowEndpointPattern: !!flowMatch,
        pageUrl: window.location.href,
        origin: window.location.origin,
        hasShellAjaxSafePost: typeof shell !== 'undefined' && shell && typeof shell.ajaxSafePost === 'function',
        hasJQuery: typeof $ !== 'undefined',
        payloadType: typeof payload,
        payloadKeys: payloadKeys,
        portalContextAvailable: !!portal,
        portalId: portal && portal.id ? portal.id : '',
        portalType: portal && portal.type ? portal.type : '',
        portalVersion: portal && portal.version ? portal.version : '',
        tenant: portal && portal.tenant ? portal.tenant : '',
        geo: portal && portal.geo ? portal.geo : '',
        contactId: portal && portal.User && portal.User.contactId ? portal.User.contactId : ''
      };

      console.log('_PreflightCloudFlowCheck: summary', preflight);

      if (!preflight.isCloudFlowEndpointPattern) {
        console.error('_PreflightCloudFlowCheck: flowUrl format does not match /_api/cloudflow/v1.0/trigger/{guid}', { flowUrl: flowUrl });
      }

      if (!preflight.triggerId) {
        console.error('_PreflightCloudFlowCheck: no triggerId detected in URL.');
      }

      if (!preflight.portalContextAvailable) {
        console.error('_PreflightCloudFlowCheck: Microsoft.Dynamic365.Portal context not found on page.');
      }

      if (!preflight.contactId) {
        console.warn('_PreflightCloudFlowCheck: no contactId found in portal context. Authenticated user may be required by flow permissions.');
      }

      console.log('_PreflightCloudFlowCheck: checklist', {
        step1_enableCloudFlowIntegration: 'Verify Power Pages cloud flow integration is enabled in site settings',
        step2_flowBinding: 'Verify this trigger ID belongs to current site/environment and is bound to this page context',
        step3_permissions: 'Verify current user has permission to run the flow',
        step4_triggerSchema: 'Verify trigger expects eventData with text/number in current stage',
        step5_publishAndRetry: 'Republish flow and retry from same portal environment'
      });

      return preflight;
    }

     /**
      * SHARED FUNCTION: generic flow transport used to communicate with Power Automate.
      * Shared functions must only use console logging and return status/results.
      * @param flowUrl
      * @param payload
      * @returns
      */

    function _GenericFlowHandler(flowUrl, payload) {
    //https://learn.microsoft.com/en-us/power-pages/configure/cloud-flow-integration
      var startedAt = Date.now();
      var handlerVersion = "DocManVer";

      console.log("_GenericFlowHandler: ver", handlerVersion);
      console.log("_GenericFlowHandler: flowUrl", flowUrl);
      console.log("_GenericFlowHandler: payload (raw)", payload);
      console.log("_GenericFlowHandler: payload typeof", typeof payload);

      var payloadString = "";
      try {
        payloadString = JSON.stringify(payload);
        console.log("_GenericFlowHandler: payload stringify success", payloadString);
      } catch (stringifyError) {
        console.error("_GenericFlowHandler: payload stringify failed", stringifyError);
      }

      //here is the section where we acutally configure what is being sent to the flow
      //Send the object directly and let jQuery handle serialization with proper Content-Type
      //This ensures Power Pages cloud flow receives valid JSON, not a JSON string
      var requestConfig = {
        type: "POST",
        url: flowUrl,
        data: { "eventData": JSON.stringify(payload) },
      };

      console.log("_GenericFlowHandler: requestConfig", {
        type: requestConfig.type,
        contentType: requestConfig.contentType,
        url: requestConfig.url,
        processData: requestConfig.processData,
        global: requestConfig.global,
        dataPreview: requestConfig.data
      });

      //all of this code for this line!!!!
      var request = shell.ajaxSafePost(requestConfig);

      request.done(function(responseData, textStatus, jqXHR) {
        console.log("_GenericFlowHandler: DONE", {
          durationMs: Date.now() - startedAt,
          textStatus: textStatus,
          status: jqXHR && typeof jqXHR.status === "number" ? jqXHR.status : null,
          responseData: responseData
        });
      });

      request.fail(function(jqXHR, textStatus, errorThrown) {
        var responseText = jqXHR && jqXHR.responseText ? jqXHR.responseText : "";
        var responseSnippet = responseText ? responseText.substring(0, 700) : "";
        var errorMeta = _ParsePowerPagesErrorMetadata(responseText);
        var isPowerAutomateErrorView = errorMeta.hasPowerAutomateErrorView;

        console.error("_GenericFlowHandler: FAIL", {
          durationMs: Date.now() - startedAt,
          textStatus: textStatus,
          errorThrown: errorThrown ? String(errorThrown) : "",
          status: jqXHR && typeof jqXHR.status === "number" ? jqXHR.status : null,
          statusText: jqXHR && jqXHR.statusText ? jqXHR.statusText : "",
          responseText: responseText,
          isPowerAutomateErrorView: isPowerAutomateErrorView,
          correlationId: errorMeta.correlationId,
          errorId: errorMeta.errorId,
          flowTrackingId: errorMeta.flowTrackingId
        });

        console.error("_GenericFlowHandler: FAIL response snippet", responseSnippet);

        if (isPowerAutomateErrorView) {
          console.error("_GenericFlowHandler: The cloud flow trigger endpoint exists but failed server-side in Power Pages. Check site settings for Power Automate integration, flow binding/permissions, and whether the trigger ID belongs to the current site/environment.");
        }
        
        // Run preflight becuase there was an error
        console.error("_GenericFlowHandler: Running preflight diagnostics due to flow failure");
        _PreflightCloudFlowCheck(flowUrl, payload);
      });

      console.log("_GenericFlowHandler: request sent");
      return request;
    }

///////////////////////////////////////////////////////
//PowerPagesListMenuEnhancer is a class used to add controls to the standard list
///////////////////////////////////////////////////////
  class PowerPagesListMenuEnhancer {
    constructor(options) {
      options = options || {};
      this.logPrefix = options.logPrefix || "[PowerPagesListMenuEnhancer]";
      this.triggerSelector = options.triggerSelector || "[data-automation-key='ppNativeListContextualMenu']";
      this.listKey = options.listKey || null;
      this.listSelector = options.listSelector || (this.listKey ? "pages-grid[list-id='" + this.escapeCssAttrValue(this.listKey) + "']" : null);
      this.guidSelector = options.guidSelector || null;
      this.instanceKey = options.instanceKey || (this.logPrefix + (this.listSelector || "")).replace(/[^a-zA-Z0-9_-]/g, "_");
      this.pendingRowContext = null;
      this.menuObserver = null;
      this.actions = [];
      this.isAwaitingMenuForClick = false;
      this.debug = !!options.debug;
      this.nativeActionsConfig = options.nativeActionsConfig || {};
      this.nativeActionLabelsToHide = Array.isArray(this.nativeActionsConfig.labelsToHide) ? this.nativeActionsConfig.labelsToHide : [];
      this.icons = {
        gear: '<path d="M8 4.754a3.246 3.246 0 1 0 0 6.492 3.246 3.246 0 0 0 0-6.492zM5.754 8a2.246 2.246 0 1 1 4.492 0 2.246 2.246 0 0 1-4.492 0z"/>' +
          '<path d="M9.796 1.343c-.527-1.79-3.065-1.79-3.592 0l-.094.319a1.873 1.873 0 0 1-2.292 1.226l-.314-.093c-1.79-.527-3.065 1.748-1.933 3.165l.2.25a1.873 1.873 0 0 1 0 2.58l-.2.25c-1.132 1.417.143 3.692 1.933 3.165l.314-.093a1.873 1.873 0 0 1 2.292 1.226l.094.319c.527 1.79 3.065 1.79 3.592 0l.094-.319a1.873 1.873 0 0 1 2.292-1.226l.314.093c1.79.527 3.065-1.748 1.933-3.165l-.2-.25a1.873 1.873 0 0 1 0-2.58l.2-.25c1.132-1.417-.143-3.692-1.933-3.165l-.314.093a1.873 1.873 0 0 1-2.292-1.226l-.094-.319z"/>',
        pencil: '<path d="M12.146.146a.5.5 0 0 1 .708 0l3 3a.5.5 0 0 1 0 .708L5.207 14.5H2v-3.207z"/><path fill-rule="evenodd" d="M1 13.5V16h2.5l8.793-8.793-2.5-2.5z"/>'
      };
      this.ensureStylesInjected();
    }

    log() {
      var args = Array.prototype.slice.call(arguments);
      args.unshift(this.logPrefix);
      console.log.apply(console, args);
    }

    debugLog() {
      if (!this.debug) return;
      var args = Array.prototype.slice.call(arguments);
      args.unshift(this.logPrefix);
      args.unshift("[debug]");
      console.log.apply(console, args);
    }

    normalizeText(value) {
      return String(value || "").replace(/\s+/g, " ").trim().toLowerCase();
    }

    decodeUriSafe(value) {
      if (value == null) return "";
      var text = String(value);
      try {
        return decodeURIComponent(text);
      } catch (e) {
        return text;
      }
    }

    formatCompactGuid(compact) {
      var normalized = String(compact || "").replace(/[^0-9a-fA-F]/g, "");
      if (normalized.length !== 32) return null;
      return (
        normalized.slice(0, 8) + "-" +
        normalized.slice(8, 12) + "-" +
        normalized.slice(12, 16) + "-" +
        normalized.slice(16, 20) + "-" +
        normalized.slice(20)
      ).toLowerCase();
    }

    extractGuid(value) {
      var text = String(value || "");
      var candidates = [text, this.decodeUriSafe(text)];

      for (var i = 0; i < candidates.length; i++) {
        var candidate = candidates[i];
        if (!candidate) continue;

        var hyphenated = candidate.match(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/);
        if (hyphenated) return hyphenated[0].toLowerCase();

        var compact = candidate.match(/\b[0-9a-fA-F]{32}\b/);
        if (compact) return this.formatCompactGuid(compact[0]);
      }

      return null;
    }

    extractGuidFromElement(element) {
      if (!element) return null;

      var attributesToCheck = ["data-guid", "data-id", "data-record-id", "data-entityid", "data-row-id", "value", "id", "href", "onclick", "aria-label", "name"];

      for (var i = 0; i < attributesToCheck.length; i++) {
        var attrValue = element.getAttribute(attributesToCheck[i]);
        var guidFromAttr = this.extractGuid(attrValue);
        if (guidFromAttr) return guidFromAttr;
      }

      if (element.attributes && element.attributes.length) {
        for (var j = 0; j < element.attributes.length; j++) {
          var attr = element.attributes[j];
          if (!attr) continue;
          var guidFromAnyAttr = this.extractGuid(attr.value);
          if (guidFromAnyAttr) return guidFromAnyAttr;
        }
      }

      if (typeof element.value === "string") {
        var guidFromValue = this.extractGuid(element.value);
        if (guidFromValue) return guidFromValue;
      }

      var guidFromText = this.extractGuid(element.textContent);
      if (guidFromText) return guidFromText;

      return this.extractGuid(element.innerHTML);
    }

    getGuidFromRow(row) {
      if (!row) return null;

      var guid = null;

      if (this.guidSelector) {
        var selectedElement = null;

        if (typeof row.matches === "function" && row.matches(this.guidSelector)) {
          selectedElement = row;
        } else {
          selectedElement = row.querySelector(this.guidSelector);
        }

        guid = this.extractGuidFromElement(selectedElement);
        if (guid) return guid;
      }

      guid = this.extractGuidFromElement(row);
      if (guid) return guid;

      var hiddenInputs = row.querySelectorAll("input[type='hidden']");
      for (var i = 0; i < hiddenInputs.length; i++) {
        guid = this.extractGuidFromElement(hiddenInputs[i]);
        if (guid) return guid;
      }

      var cells = row.querySelectorAll("[role='gridcell'], td");
      for (var j = 0; j < cells.length; j++) {
        guid = this.extractGuidFromElement(cells[j]);
        if (guid) return guid;
      }

      return null;
    }

    getGuidFromTrigger(trigger) {
      if (!trigger) return null;

      var node = trigger;
      while (node && node !== document.body) {
        var guidFromNode = this.extractGuidFromElement(node);
        if (guidFromNode) return guidFromNode;
        node = node.parentElement;
      }

      var guidFromTriggerHtml = this.extractGuid(trigger.outerHTML);
      if (guidFromTriggerHtml) return guidFromTriggerHtml;

      return null;
    }

    getGuidFromMenuRoot(menuRoot) {
      if (!menuRoot) return null;

      var candidates = menuRoot.querySelectorAll("*");

      var guidFromRoot = this.extractGuidFromElement(menuRoot);
      if (guidFromRoot) return guidFromRoot;

      for (var i = 0; i < candidates.length; i++) {
        var candidate = candidates[i];
        var guidFromElement = this.extractGuidFromElement(candidate);
        if (guidFromElement) return guidFromElement;

        var hrefGuid = this.extractGuid(candidate.getAttribute("href"));
        if (hrefGuid) return hrefGuid;

        var onClickGuid = this.extractGuid(candidate.getAttribute("onclick"));
        if (onClickGuid) return onClickGuid;
      }

      return this.extractGuid(menuRoot.outerHTML);
    }

    resolveContextGuid(context, menuRoot) {
      var baseContext = context || {};
      if (baseContext.guid) return baseContext;

      var resolvedGuid = null;
      var resolvedSource = null;

      if (baseContext.row) {
        resolvedGuid = this.getGuidFromRow(baseContext.row);
        if (resolvedGuid) resolvedSource = "row";
      }

      if (!resolvedGuid && baseContext.trigger) {
        resolvedGuid = this.getGuidFromTrigger(baseContext.trigger);
        if (resolvedGuid) resolvedSource = "trigger";
      }

      if (!resolvedGuid) {
        resolvedGuid = this.getGuidFromMenuRoot(menuRoot);
        if (resolvedGuid) resolvedSource = "native-menu";
      }

      if (!resolvedGuid) return baseContext;

      this.log("Resolved GUID from", resolvedSource + ":", resolvedGuid);
      return Object.assign({}, baseContext, { guid: resolvedGuid, guidSource: resolvedSource });
    }

    escapeCssAttrValue(value) {
      return String(value).replace(/\\/g, "\\\\").replace(/'/g, "\\'");
    }

    ensureStylesInjected() {
      var styleId = "pp-custom-actions-styles";
      if (document.getElementById(styleId)) return;

      var customActionStyles = document.createElement("style");
      customActionStyles.id = styleId;
      customActionStyles.textContent =
        ".custom-menu-action { justify-content: flex-start; }" +
        ".custom-menu-action:hover, .custom-menu-action:focus-visible {" +
        " background-color: var(--bs-dropdown-link-hover-bg, rgba(0, 0, 0, 0.05));" +
        " color: var(--bs-dropdown-link-hover-color, inherit);" +
        "}";

      document.head.appendChild(customActionStyles);
    }

    createSvgIcon(svgPaths, color) {
      var icon = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      icon.setAttribute("xmlns", "http://www.w3.org/2000/svg");
      icon.setAttribute("width", "16");
      icon.setAttribute("height", "16");
      icon.setAttribute("fill", "currentColor");
      icon.setAttribute("viewBox", "0 0 16 16");
      icon.setAttribute("aria-hidden", "true");
      icon.innerHTML = svgPaths;
      if (color) {
        icon.style.color = color;
      }
      return icon;
    }

    createIcon(name, color) {
      var svgPaths = this.icons[name];
      if (!svgPaths) return null;
      return this.createSvgIcon(svgPaths, color);
    }

    collectMatchesDeep(root, selector) {
      var results = [];
      if (!root || !selector) return results;

      function visit(node) {
        if (!node) return;

        if (node.nodeType === 1 && typeof node.matches === "function" && node.matches(selector)) {
          results.push(node);
        }

        var children = node.children || [];
        for (var i = 0; i < children.length; i++) {
          visit(children[i]);
        }

        if (node.shadowRoot) {
          visit(node.shadowRoot);
        }
      }

      visit(root);
      return results;
    }

    findFirstDeep(root, selector) {
      var matches = this.collectMatchesDeep(root, selector);
      return matches.length ? matches[0] : null;
    }

    findScopedHostFromTrigger(trigger) {
      if (!trigger) return null;

      if (this.listSelector) {
        var selectorMatch = trigger.closest(this.listSelector);
        if (selectorMatch) return selectorMatch;
      }

      if (this.listKey) {
        var keySelector = "pages-grid[list-id='" + this.escapeCssAttrValue(this.listKey) + "']";
        var containers = document.querySelectorAll("pages-native-container");

        for (var i = 0; i < containers.length; i++) {
          var container = containers[i];
          if (container.querySelector(keySelector) && container.contains(trigger)) {
            return container;
          }
        }
      }

      return null;
    }

    getRowContextFromTrigger(trigger) {
      var row = trigger.closest("[role='row'], tr");
      if (!row) return null;

      var firstCell = row.querySelector("[role='gridcell'], td");
      var label = firstCell ? (firstCell.innerText || "").trim() : "Row";
      var listRoot = this.findScopedHostFromTrigger(trigger);
      var guid = this.getGuidFromRow(row);

      return { row: row, label: label, guid: guid, trigger: trigger, listRoot: listRoot, listKey: this.listKey || null };
    }

    getRowContextFromRow(row, listRoot) {
      if (!row) return null;

      var firstCell = row.querySelector("[role='gridcell'], td");
      var label = firstCell ? (firstCell.innerText || "").trim() : "Row";
      var guid = this.getGuidFromRow(row);

      return { row: row, label: label, guid: guid, listRoot: listRoot || null, listKey: this.listKey || null };
    }

    isTriggerInScope(trigger) {
      if (!this.listSelector && !this.listKey) return true;
      return !!this.findScopedHostFromTrigger(trigger);
    }

    findOpenMenuRoot() {
      var menus = Array.from(document.querySelectorAll("[role='menu']"));
      if (!menus.length) return null;

      for (var i = menus.length - 1; i >= 0; i--) {
        var menu = menus[i];
        var style = window.getComputedStyle(menu);
        if (style.display !== "none" && style.visibility !== "hidden") return menu;
      }

      return null;
    }

    getScopedListRoots() {
      if (this.listSelector) {
        return this.collectMatchesDeep(document, this.listSelector);
      }

      if (this.listKey) {
        var keySelector = "pages-grid[list-id='" + this.escapeCssAttrValue(this.listKey) + "']";
        return this.collectMatchesDeep(document, keySelector);
      }

      return this.collectMatchesDeep(document, "pages-grid");
    }

    listHasNativeActions(listRoot) {
      if (!listRoot) return false;
      return !!this.findFirstDeep(listRoot, this.triggerSelector);
    }

    addAction(config) {
      if (!config || !config.id || !config.label || typeof config.onClick !== "function") {
        throw new Error("Action must include: id, label, onClick(context)");
      }

      this.actions.push({
        id: config.id,
        label: config.label,
        ariaLabel: config.ariaLabel || config.label,
        icon: config.icon || null,
        onClick: config.onClick
      });

      return this;
    }

    addActions(configs) {
      if (!Array.isArray(configs)) {
        throw new Error("addActions expects an array of action configs");
      }

      for (var i = 0; i < configs.length; i++) {
        this.addAction(configs[i]);
      }

      return this;
    }

    createMenuItem(action, context) {
      var item = document.createElement("button");
      item.type = "button";
      item.setAttribute("role", "menuitem");
      item.className = "dropdown-item custom-menu-action custom-menu-action-" + action.id;
      item.setAttribute("data-custom-action-id", this.instanceKey + "__" + action.id);
      item.setAttribute("aria-label", action.ariaLabel);

      item.style.display = "flex";
      item.style.alignItems = "center";
      item.style.gap = "8px";
      item.style.width = "100%";
      item.style.textAlign = "left";
      item.style.padding = "8px 12px";
      item.style.cursor = "pointer";

      if (action.icon) {
        item.appendChild(action.icon.cloneNode(true));
      }

      var text = document.createElement("span");
      text.textContent = action.label;
      item.appendChild(text);

      item.addEventListener("click", function (e) {
        e.preventDefault();
        e.stopPropagation();
        action.onClick(context || {});
      });

      return item;
    }

    injectIntoMenu(menuRoot, context) {
      if (!menuRoot || !this.actions.length) return false;

      var resolvedContext = this.resolveContextGuid(context, menuRoot);

      this.applyNativeActionVisibility(menuRoot);

      var existingMenuItems = menuRoot.querySelectorAll("[role='menuitem']");
      var container = existingMenuItems.length
        ? (existingMenuItems[existingMenuItems.length - 1].parentElement || menuRoot)
        : menuRoot;
      var injectedCount = 0;

      for (var i = 0; i < this.actions.length; i++) {
        var action = this.actions[i];
        var actionDomId = this.instanceKey + "__" + action.id;
        if (menuRoot.querySelector("[data-custom-action-id='" + actionDomId + "']")) {
          continue;
        }
        container.appendChild(this.createMenuItem(action, resolvedContext));
        injectedCount++;
      }

      if (injectedCount > 0) {
        this.log("Injected", injectedCount, "custom action(s) into dropdown menu.");
      }

      return injectedCount > 0;
    }

    isCustomMenuItem(menuItem) {
      if (!menuItem) return false;
      var customActionId = menuItem.getAttribute("data-custom-action-id") || "";
      return customActionId.indexOf(this.instanceKey + "__") === 0;
    }

    matchesNativeActionLabel(menuItem) {
      if (!menuItem || !this.nativeActionLabelsToHide.length) return false;
      var candidates = [
        this.normalizeText(menuItem.textContent),
        this.normalizeText(menuItem.getAttribute("aria-label")),
        this.normalizeText(menuItem.getAttribute("title")),
        this.normalizeText(menuItem.getAttribute("data-original-title"))
      ];
      var hasCandidate = false;

      for (var c = 0; c < candidates.length; c++) {
        if (candidates[c]) {
          hasCandidate = true;
          break;
        }
      }

      if (!hasCandidate) return false;

      for (var i = 0; i < this.nativeActionLabelsToHide.length; i++) {
        var label = this.nativeActionLabelsToHide[i];
        if (!label || typeof label !== "string") continue;

        var normalizedLabel = this.normalizeText(label);
        for (var j = 0; j < candidates.length; j++) {
          var candidate = candidates[j];
          if (!candidate) continue;
          if (candidate === normalizedLabel || candidate.indexOf(normalizedLabel) === 0) {
            return true;
          }
        }
      }

      return false;
    }

    shouldHideNativeAction(menuItem) {
      return this.matchesNativeActionLabel(menuItem);
    }

    applyNativeActionVisibility(menuRoot) {
      if (!menuRoot) return;
      if (!this.nativeActionLabelsToHide.length) return;

      var menuItems = menuRoot.querySelectorAll("[role='menuitem']");
      var affected = 0;

      this.debugLog("Scanning menu items:", menuItems.length);

      for (var i = 0; i < menuItems.length; i++) {
        var menuItem = menuItems[i];
        if (this.isCustomMenuItem(menuItem)) continue;
        var matched = this.shouldHideNativeAction(menuItem);
        this.debugLog(
          "Native menu item",
          i,
          {
            text: this.normalizeText(menuItem.textContent),
            ariaLabel: this.normalizeText(menuItem.getAttribute("aria-label")),
            title: this.normalizeText(menuItem.getAttribute("title")),
            matched: matched
          }
        );
        if (!matched) continue;

        menuItem.style.display = "none";
        menuItem.setAttribute("aria-hidden", "true");

        affected++;
      }

      if (affected > 0) {
        this.log("Native action(s) hidden:", affected);
      } else {
        this.debugLog("No native actions matched hide rules.");
      }
    }

    tryInjectAfterOpen() {
      if (!this.isAwaitingMenuForClick) return;

      var attempts = 0;
      var maxAttempts = 20;
      var self = this;

      var timer = setInterval(function () {
        attempts++;
        var menu = self.findOpenMenuRoot();
        if (menu && self.injectIntoMenu(menu, self.pendingRowContext)) {
          clearInterval(timer);
          self.isAwaitingMenuForClick = false;
        } else if (attempts >= maxAttempts) {
          clearInterval(timer);
          self.isAwaitingMenuForClick = false;
          self.log("No open menu detected for injection.");
        }
      }, 100);
    }

    startMenuObserver() {
      if (this.menuObserver) return;

      var self = this;
      this.menuObserver = new MutationObserver(function () {
        if (!self.isAwaitingMenuForClick) return;
        var menu = self.findOpenMenuRoot();
        if (menu) {
          var injected = self.injectIntoMenu(menu, self.pendingRowContext);
          if (injected) {
            self.isAwaitingMenuForClick = false;
          }
        }
      });

      this.menuObserver.observe(document.body, { childList: true, subtree: true });
      this.log("Menu observer started.");
    }

    init() {
      var self = this;

      document.addEventListener("click", function (e) {
        var trigger = e.target.closest(self.triggerSelector);
        if (!trigger) return;
        if (!self.isTriggerInScope(trigger)) return;

        self.pendingRowContext = self.getRowContextFromTrigger(trigger);
        self.isAwaitingMenuForClick = true;
        self.log("Context menu trigger clicked for row:", self.pendingRowContext ? self.pendingRowContext.label : "(unknown)");
        self.tryInjectAfterOpen();
      });

      self.startMenuObserver();
      return self;
    }
  }

  