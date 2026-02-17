sap.ui.define(
  [
    "sap/ui/core/mvc/Controller",
    "sap/m/MessageBox",
  ],
  function (BaseController, MessageBox) {
    "use strict";

    return BaseController.extend("obsolerteworkflow.workflowuimodule.controller.App", {

      _REQUIRED_COLUMNS: [
        "RFQ-ID",
        "PLANT",
        "COMPONENT",
        "DESCRIPTION",
        "MANUFACTURERPARTNR.",
        "AVAILABLESTOCK(FREESTOCK),M,PC",
        "FREESTOCKFULLCOPPER",
        "CURRENCY",
        "LASTCONSUMPTION",
        "RANGEOFCOVERAGEINMONTHS",
        "PN",
        "CUSTOMER",
        "ENDCUSTOMER",
        "REASON",
        "CAUSED",
        "WEIGHT(ONBASEUNIT)",
        "TOTALAMOUNT"
      ],


      onInit() {
        const oModel = this.getView().getModel("obsolete");

        const oItemModel = new sap.ui.model.json.JSONModel({
          items: []
        });

        this.getView().setModel(oItemModel, "WorkflowItem");

        this._sSearchQuery = "";
        this._bShowOnlyErrors = false;


        const oViewModel = new sap.ui.model.json.JSONModel({
          fileUploadEnabled: false
        });
        this.getView().setModel(oViewModel, "view");

      },
      _normalizeColumnName: function (col) {
        return col
          .replace(/\r?\n/g, " ")
          .replace(/\s+/g, " ")
          .trim()
          .toUpperCase();
      },


      startWorkflowInstance: function (folderIdCmis, docId) {
        var that = this;

        return new Promise(function (resolve, reject) {
          let data = {
            definitionId: "eu10.builddevlapp.obsolete.obsoleteCreationProcess",
            context: {
              company: that.byId("companySelect").getSelectedKey(),
              file: folderIdCmis,
              documentid: docId
            }
          };

          $.ajax({
            url: that._getWorkflowRuntimeBaseURL() +
              "/workflow-instances?environmentId=businessprocessworkflow",
            method: "POST",
            contentType: "application/json",
            headers: {
              "X-CSRF-Token": that._fetchToken()
            },
            data: JSON.stringify(data),

            success: function (result) {
              sap.m.MessageBox.success(
                "Submitted successfully",
                {
                  title: "Success",
                  onClose: function () {
                    that._resetForm();
                  }
                }
              );
              resolve(result);   //  Promise resolved
            },

            error: function (request) {
              try {
                var response = JSON.parse(request.responseText);
                MessageBox.error(response.error.message);
                reject(response);
              } catch (e) {
                MessageBox.error("Workflow start failed");
                reject(e);
              }
            }
          });
        });
      },
      _fetchToken: function () {
        var fetchedToken;

        jQuery.ajax({
          url: this._getWorkflowRuntimeBaseURL() + "/xsrf-token",
          method: "GET",
          async: false,
          headers: {
            "X-CSRF-Token": "Fetch",
          },
          success(result, xhr, data) {
            fetchedToken = data.getResponseHeader("X-CSRF-Token");
          },
        });
        return fetchedToken;
      },

      _getWorkflowRuntimeBaseURL: function () {
        var ui5CloudService = this.getOwnerComponent().getManifestEntry("/sap.cloud/service").replaceAll(".", "");
        var ui5ApplicationName = this.getOwnerComponent().getManifestEntry("/sap.app/id").replaceAll(".", "");
        var appPath = `${ui5CloudService}.${ui5ApplicationName}`;
        return `/${appPath}/api/public/workflow/rest/v1`

      },
      _resetForm: function () {

        // Reset Company dropdown
        var oCompanySelect = this.byId("companySelect");
        if (oCompanySelect) {
          oCompanySelect.setSelectedKey("");
        }

        // Reset FileUploader
        var oUploader = this.byId("excelUploader");
        if (oUploader) {
          oUploader.clear(); //  removes last attached file
        }

        // Clear table data
        const oModel = this.getView().getModel("WorkflowItem");
        if (oModel) {
          oModel.setProperty("/items", []);
        }

        // Reset search and checkbox state (optional but recommended)
        this._sSearchQuery = "";
        this._bShowOnlyErrors = false;

        const oViewModel = this.getView().getModel("view");
        if (oViewModel) {
          oViewModel.setProperty("/fileUploadEnabled", false);
        }


        // Optional: clear stored file reference
        this._selectedFile = null;
      },



      onSubmit: function () {

        const oView = this.getView();
        oView.setBusy(true);

        // -------------------------------
        // Check table for validation errors
        // -------------------------------
        const oModel = oView.getModel("WorkflowItem");
        const aItems = oModel.getProperty("/items") || [];

        const totalErrors = aItems.reduce(
          (sum, item) => sum + (item.errorCount || 0),
          0
        );

        if (totalErrors > 0) {
          oView.setBusy(false);
          sap.m.MessageBox.error(
            `Table contains ${totalErrors} validation errors.\nPlease re-upload the file.`
          );
          return;
        }

        // -------------------------------
        // File validation
        // -------------------------------
        const oUploader = this.byId("excelUploader");

        if (!oUploader || !oUploader.oFileUpload || !oUploader.oFileUpload.files.length) {
          oView.setBusy(false);
          sap.m.MessageBox.error("Please upload an Excel file");
          return;
        }

        const oFile = oUploader.oFileUpload.files[0];
        const sFileName = oFile.name + "_" + this._generateUUID();
        const sOrgFileName = oFile.name;

        try {
          const selectedCompany = this.byId("companySelect").getSelectedKey();

          oView.setBusy(false);

          if (selectedCompany === "") {
            sap.m.MessageBox.error("Select the company code");
            return;
          }

          sap.m.MessageBox.confirm(
            "Excel validated successfully.\n\nAre you sure you want to submit?",
            {
              title: "Confirm Submission",
              actions: [
                sap.m.MessageBox.Action.YES,
                sap.m.MessageBox.Action.NO
              ],
              emphasizedAction: sap.m.MessageBox.Action.YES,
              onClose: async function (oAction) {

                if (oAction === sap.m.MessageBox.Action.YES) {
                  oView.setBusy(true);

                  try {
                    const folderId = await this.onUpload(sFileName);

                    if (folderId) {
                      const docId = await this.onUploadDocument(
                        oFile,
                        sOrgFileName,
                        sFileName
                      );

                      const folderIdCmis = `spa-res:cmis:folderid:${folderId}`;

                      await this.startWorkflowInstance(
                        folderIdCmis,
                        docId
                      );
                    } else {
                      sap.m.MessageToast.show("Failed to create folder");
                    }

                  } catch (err) {
                    console.error(err);
                    sap.m.MessageBox.error("Submission failed");
                  } finally {
                    oView.setBusy(false);
                  }
                }
              }.bind(this)
            }
          );

        } catch (err) {
          oView.setBusy(false);
          sap.m.MessageBox.error(err.message);
        }
      },
      _validateComponent: function (aPayload) {
        const oModel = this.getView().getModel("obsolete");
        const csrfToken = oModel.getSecurityToken();
        const serviceUrl = oModel.sServiceUrl;

        // Get company code from view model
        const companyCode = this.byId("companySelect").getSelectedKey();

        // Transform payload to required structure
        const aCheckItems = aPayload.map(item => ({
          companyCode: companyCode,
          plant: item.Plant,
          component: item.Component
        }));

        return new Promise((resolve, reject) => {
          jQuery.ajax({
            url: `${serviceUrl}/checkDuplicate`,
            method: "POST",
            contentType: "application/json",
            data: JSON.stringify({
              items: aCheckItems
            }),
            headers: {
              "X-CSRF-Token": csrfToken
            },
            success: resolve,
            error: reject
          });
        });
      },

      _validateLocalDuplicates: function (aPayload) {
        const map = {};
        const duplicates = {};
        const result = [];

        aPayload.forEach(item => {
          const key = `${item.Plant}__${item.Component}`;

          if (map[key]) {
            duplicates[key] = {
              plant: item.Plant,
              component: item.Component
            };
          } else {
            map[key] = true;
          }
        });

        return Object.values(duplicates);
      },




      _normalizeExcelData: function (excelData) {

        return excelData.map(row => {
          var normalizedRow = {};

          Object.keys(row).forEach(key => {

            var normalizedKey = key
              .replace(/\r?\n/g, "")   // remove line breaks
              .replace(/\s+/g, "")     // remove ALL spaces
              .toUpperCase();          // uppercase

            normalizedRow[normalizedKey] = row[key];
          });

          return normalizedRow;
        });
      },




      _normalizeRowKeys: function (row) {
        var normalized = {};

        Object.keys(row).forEach(key => {
          var cleanKey = key
            .replace(/\n/g, " ")      // remove line breaks
            .replace(/\r/g, " ")
            .replace(/\s+/g, " ")     // collapse multiple spaces
            .trim()
            .toUpperCase();

          normalized[cleanKey] = row[key];
        });

        return normalized;
      },


      _validateCaused: function (value, rowIndex) {
        if (!value) {
          throw new Error(`Empty value in column "CAUSED" at row ${rowIndex + 2}`);
        }

        var normalized = value.toString().trim().toUpperCase();

        if (normalized !== "PLANT" && normalized !== "CUSTOMER") {
          throw new Error(
            `Invalid value "${value}" in column "CAUSED" at row ${rowIndex + 2}. Allowed values: Plant, Customer`
          );
        }

        return normalized; // optional: return normalized value
      },


      _formatExcelDate1: function (value) {

        // If already a JS Date
        if (value instanceof Date) {
          return value.toISOString().split("T")[0]; // YYYY-MM-DD
        }

        // If Excel serial number
        if (typeof value === "number") {
          var excelEpoch = new Date(Date.UTC(1899, 11, 30));
          var resultDate = new Date(excelEpoch.getTime() + value * 86400000);
          return resultDate.toISOString().split("T")[0];
        }

        // If string (already formatted)
        return value;
      },
      _formatExcelDate: function (excelDate) {
        if (!excelDate) {
          return null;
        }

        // If already a date, return it
        if (excelDate instanceof Date) {
          return excelDate;
        }

        // If it's a number (Excel serial date)
        if (!isNaN(excelDate)) {
          // Excel base date: Jan 1, 1900
          const excelEpoch = new Date(Date.UTC(1899, 11, 30));
          const jsDate = new Date(excelEpoch.getTime() + excelDate * 86400000);
          return jsDate;
        }

        // If string date, try parsing
        const parsed = new Date(excelDate);
        return isNaN(parsed) ? null : parsed;
      },


      _toDecimal: function (value) {

        if (value === null || value === undefined || value === "") {
          return 0;
        }

        // If already a number
        if (typeof value === "number") {
          return Number(value.toFixed(3));
        }

        // If string → clean it
        if (typeof value === "string") {
          var cleaned = value
            .replace(/€/g, "")
            .replace(/,/g, "")
            .trim();

          var num = parseFloat(cleaned);
          return isNaN(num) ? 0 : Number(num.toFixed(3));
        }

        return 0;
      },


      _toString: function (value) {

        if (value === null || value === undefined) {
          return "";
        }

        // If number → convert to string (preserve value)
        if (typeof value === "number") {
          return value.toString();
        }

        // If already string → trim
        if (typeof value === "string") {
          return value.trim();
        }

        // Fallback
        return String(value);
      },



      onUpload: function (sFolderName) {
        return new Promise((resolve, reject) => {
          try {
            const repositoryId = "cc918620-3f34-4544-b260-cb5ad8a568d7";
            // const folderName = sFileName + "_" + this._generateUUID();

            const formData = new FormData();
            formData.append("cmisaction", "createFolder");
            formData.append("propertyId[0]", "cmis:name");
            formData.append("propertyValue[0]", sFolderName);
            formData.append("propertyId[1]", "cmis:objectTypeId");
            formData.append("propertyValue[1]", "cmis:folder");
            formData.append("succinct", "true");

            $.ajax({
              url: this._getWorkflowRuntimeBaseURLTest() + `/${repositoryId}/root`,
              method: "POST",
              data: formData,
              processData: false,
              contentType: false,
              headers: {
                "X-CSRF-Token": this._fetchToken()
              },
              success: function (data) {
                const folderId = data.succinctProperties["cmis:objectId"];
                resolve(folderId); //  THIS is the real return
              },
              error: function (err) {
                reject(err);
              }
            });

          } catch (e) {
            reject(e);
          }
        });
      },
      _generateUUID: function () {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
          const r = Math.random() * 16 | 0;
          const v = c === 'x' ? r : (r & 0x3 | 0x8);
          return v.toString(16);
        });
      },


      onUploadDocument: function (oFile, sFileName, sFolderObjectId) {
        return new Promise((resolve, reject) => {
          try {

            //When deploying to   PROD chnage the repositoryID
            const repositoryId = "cc918620-3f34-4544-b260-cb5ad8a568d7";

            const formData = new FormData();
            formData.append("cmisaction", "createDocument");

            // CMIS properties
            formData.append("propertyId[0]", "cmis:name");
            formData.append("propertyValue[0]", sFileName);

            formData.append("propertyId[1]", "cmis:objectTypeId");
            formData.append("propertyValue[1]", "cmis:document");

            // formData.append("_charset_", "UTF-8");
            formData.append("filename", sFileName);
            formData.append("succinct", "true");
            formData.append("includeAllowableActions", "true");
            // File content
            formData.append("media", oFile);

            $.ajax({
              url: this._getWorkflowRuntimeBaseURLTest() +
                `/${repositoryId}/root/${sFolderObjectId}`,
              method: "POST",
              data: formData,
              processData: false,
              contentType: false,
              headers: {
                "X-CSRF-Token": this._fetchToken()
              },
              success: function (data) {
                const documentId = data.succinctProperties["cmis:objectId"];
                console.log("Document ID:", documentId);
                resolve(documentId);
              },
              error: function (err) {
                reject(err);
              }
            });

          } catch (e) {
            reject(e);
          }
        });
      },


      _getWorkflowRuntimeBaseURLTest: function () {
        var ui5CloudService = this.getOwnerComponent().getManifestEntry("/sap.cloud/service").replaceAll(".", "");
        var ui5ApplicationName = this.getOwnerComponent().getManifestEntry("/sap.app/id").replaceAll(".", "");
        var appPath = `${ui5CloudService}.${ui5ApplicationName}`;
        return `/${appPath}/dms_api/browser`

      },


      onFileChange: async function (oEvent) {
        const oView = this.getView();
        const oFile = oEvent.getParameter("files")[0];

        // If user removed the file (no file selected)
        if (!oFile) {
          const oTableModel = oView.getModel("WorkflowItem");
          if (oTableModel) {
            oTableModel.setProperty("/items", []);
          }

          // Reset search and error filter
          this._sSearchQuery = "";
          this._bShowOnlyErrors = false;

          return;
        }

        oView.setBusy(true);   // START BUSY

        try {
          // Read file
          const arrayBuffer = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = reject;
            reader.readAsArrayBuffer(oFile);
          });

          const data = new Uint8Array(arrayBuffer);
          const workbook = XLSX.read(data, { type: "array" });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          let jsonData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          jsonData = this._normalizeExcelData(jsonData);

          // Async validation (Excel + backend)
          const aMapped = await this._validateAndMapExcel(jsonData);

          // Load into table
          this.getView()
            .getModel("WorkflowItem")
            .setProperty("/items", aMapped);

        } catch (err) {
          sap.m.MessageBox.error(err.message || "Error processing Excel file");
        }

        oView.setBusy(false);   // STOP BUSY
      },
      _validateAndMapExcel: async function (excelData) {

        // HARD VALIDATION
        if (!excelData.length) {
          throw new Error("Excel file is empty");
        }

        this._REQUIRED_COLUMNS.forEach(col => {
          if (!(col in excelData[0])) {
            throw new Error(`Missing required column: ${col}`);
          }
        });

        // Build duplicate map (Excel level)
        const oKeyMap = {};
        excelData.forEach((r, index) => {
          const plant = this._toString(r["PLANT"]);
          const component = this._toString(r["COMPONENT"]);
          const key = plant + "||" + component;

          if (!oKeyMap[key]) {
            oKeyMap[key] = [];
          }
          oKeyMap[key].push(index);
        });

        const aResult = [];

        // ROW-LEVEL VALIDATION
        excelData.forEach((r, index) => {
          const aErrors = [];
          const rowNo = index + 2;

          const plant = this._toString(r["PLANT"]);
          const component = this._toString(r["COMPONENT"]);
          const key = plant + "||" + component;

          // Mandatory validation
          this._REQUIRED_COLUMNS.forEach(col => {
            if (r[col] === "" || r[col] === null || r[col] === undefined) {
              aErrors.push(`Empty value in column "${col}"`);
            }
          });

          // Excel duplicate validation
          if (oKeyMap[key].length > 1) {
            aErrors.push("Duplicate Plant + Component combination in Excel");
          }

          // Caused validation
          try {
            this._validateCaused(r["CAUSED"], index);
          } catch (e) {
            aErrors.push(e.message);
          }

          // Plant length
          if (plant && plant.length > 4) {
            aErrors.push("Plant value cannot be greater than 4 characters");
          }

          // availableStock > 0 decimal
          const availableStock = this._toDecimal(
            r["AVAILABLESTOCK(FREESTOCK),M,PC"]
          );
          if (isNaN(availableStock) || availableStock <= 0) {
            aErrors.push("Available Stock must be a decimal greater than 0");
          }

          // freeStock > 0 decimal
          const freeStock = this._toDecimal(
            r["FREESTOCKFULLCOPPER"]
          );
          if (isNaN(freeStock) || freeStock <= 0) {
            aErrors.push("Free stock full copper must be a decimal greater than 0");
          }

          // totalAmount decimal
          const totalAmount = this._toDecimal(r["TOTALAMOUNT"]);
          if (isNaN(totalAmount)) {
            aErrors.push("Total Amount must be a decimal value");
          }

          // weight decimal
          const weight = this._toDecimal(r["WEIGHT(ONBASEUNIT)"]);
          if (isNaN(weight)) {
            aErrors.push("Weight must be a decimal value");
          }

          // currency length
          const currency = r["CURRENCY"];
          if (!currency || currency.toString().trim().length !== 3) {
            aErrors.push("Currency must be exactly 3 characters");
          }

          // reason integer
          const reason = r["REASON"];
          if (!/^\d+$/.test(reason)) {
            aErrors.push("Reason must be an integer value");
          }

          // lastConsumption date
          const lastConsumption = this._formatExcelDate(r["LASTCONSUMPTION"]);
          if (!lastConsumption) {
            aErrors.push("Last Consumption must be a valid date");
          }

          aResult.push({
            plant: plant,
            component: component,
            rfq: this._toString(r["RFQ-ID"]),
            description: r["DESCRIPTION"],
            manufacturerPart: this._toString(r["MANUFACTURERPARTNR."]),
            availableStock: availableStock,
            freeStock: freeStock,
            currency: currency,
            lastConsumption: lastConsumption,
            rangeCoverage: this._toDecimal(
              r["RANGEOFCOVERAGEINMONTHS"]
            ),
            pn: r["PN"],
            customer: r["CUSTOMER"],
            endCustomer: r["ENDCUSTOMER"],
            reason: reason,
            caused: r["CAUSED"],
            totalAmount: totalAmount,
            weight: weight,
            errors: aErrors,
            errorCount: aErrors.length
          });
        });

        // ----------------------------
        // BACKEND DUPLICATE VALIDATION
        // ----------------------------
        const aPayload = aResult.map(item => ({
          Plant: item.plant,
          Component: item.component
        }));

        const aDuplicateComponent = await this._validateComponent(aPayload);
        const aDuplicates = aDuplicateComponent?.d?.results || [];

        if (aDuplicates.length > 0) {
          aDuplicates.forEach(dup => {
            aResult.forEach(row => {
              if (
                row.plant === dup.plant &&
                row.component === dup.component
              ) {
                row.errors.push("Component already exists in system");
                row.errorCount = row.errors.length;
              }
            });
          });
        }

        return aResult;
      },
      onErrorPress: function (oEvent) {
        const oContext = oEvent.getSource().getBindingContext("WorkflowItem");
        const aErrors = oContext.getProperty("errors");

        if (!aErrors || !aErrors.length) {
          sap.m.MessageToast.show("No errors for this row");
          return;
        }

        const sMessage = aErrors
          .map((msg, i) => `${i + 1}. ${msg}`)
          .join("\n");

        sap.m.MessageBox.error(sMessage);
      },
      onToggleErrorFilter: function (oEvent) {
        const bShowOnlyErrors = oEvent.getParameter("selected");
        this._bShowOnlyErrors = oEvent.getParameter("selected");
        this._applyTableFilters();

        const oTable = this.byId("errorTable");
        const oBinding = oTable.getBinding("rows");

        if (!oBinding) return;

        if (bShowOnlyErrors) {
          const oFilter = new sap.ui.model.Filter(
            "errorCount",
            sap.ui.model.FilterOperator.GT,
            0
          );
          oBinding.filter([oFilter]);
        } else {
          oBinding.filter([]); // clear filter
        }
      },
      onSearchTable: function (oEvent) {
        this._sSearchQuery = oEvent.getParameter("newValue") || "";
        this._applyTableFilters();
      }, _applyTableFilters: function () {
        const oTable = this.byId("errorTable");
        const oBinding = oTable.getBinding("rows");

        if (!oBinding) {
          return;
        }

        const aFilters = [];

        // Error filter
        if (this._bShowOnlyErrors) {
          aFilters.push(
            new sap.ui.model.Filter(
              "errorCount",
              sap.ui.model.FilterOperator.GT,
              0
            )
          );
        }

        // Search filter
        // Search filter
        if (this._sSearchQuery) {
          const sQuery = this._sSearchQuery;
          const aSearchFilters = [
            new sap.ui.model.Filter("plant", sap.ui.model.FilterOperator.Contains, sQuery),
            new sap.ui.model.Filter("component", sap.ui.model.FilterOperator.Contains, sQuery),
            new sap.ui.model.Filter("rfq", sap.ui.model.FilterOperator.Contains, sQuery),
            new sap.ui.model.Filter("description", sap.ui.model.FilterOperator.Contains, sQuery),
            new sap.ui.model.Filter("manufacturerPart", sap.ui.model.FilterOperator.Contains, sQuery),
            new sap.ui.model.Filter("currency", sap.ui.model.FilterOperator.Contains, sQuery),
            new sap.ui.model.Filter("pn", sap.ui.model.FilterOperator.Contains, sQuery),
            new sap.ui.model.Filter("customer", sap.ui.model.FilterOperator.Contains, sQuery),
            new sap.ui.model.Filter("endCustomer", sap.ui.model.FilterOperator.Contains, sQuery),
            new sap.ui.model.Filter("caused", sap.ui.model.FilterOperator.Contains, sQuery)
          ];

          aFilters.push(
            new sap.ui.model.Filter({
              filters: aSearchFilters,
              and: false
            })
          );
        }

        // Apply all filters together
        oBinding.filter(aFilters);
      },

      onCompanyChange: function (oEvent) {
        const sCompany = oEvent.getSource().getSelectedKey();
        const oView = this.getView();
        const oViewModel = oView.getModel("view");

        // Enable uploader only if company selected
        oViewModel.setProperty("/fileUploadEnabled", !!sCompany);

        // Clear table data
        const oTableModel = oView.getModel("WorkflowItem");
        if (oTableModel) {
          oTableModel.setProperty("/items", []);
        }

        // Clear file uploader
        const oUploader = this.byId("excelUploader");
        if (oUploader) {
          oUploader.clear();
        }

        // Reset search + error filter states
        this._sSearchQuery = "";
        this._bShowOnlyErrors = false;
      }

    });
  }
);
