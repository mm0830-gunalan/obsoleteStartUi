sap.ui.define(
  [
    "sap/ui/core/mvc/Controller",
    "sap/m/MessageBox",
  ],
  function (BaseController, MessageBox) {
    "use strict";

    return BaseController.extend("obsolerteworkflow.workflowuimodule.controller.App", {
      // _REQUIRED_COLUMNS: [
      //   "RFQ - ID",
      //   "PLANT",
      //   "COMPONENT",
      //   "DESCRIPTION",
      //   "MANUFACTURER PART NR.",
      //   "AVAILABLE STOCK (FREE STOCK), M, PC",
      //   "AVAILABLE €\r\nCU INCLUDED",
      //   // "AVAILABLE € CU INCLUDED",
      //   "LAST CONSUMPTION",
      //   "RANGE OF COVERAGE IN MONTHS",
      //   "PN",
      //   "CUSTOMER",
      //   "END CUSTOMER",
      //   "REASON",
      //   "CAUSED",
      //   "WEIGHT (ON BASE UNIT)",
      //   "TOTAL AMOUNT"
      // ],
      _REQUIRED_COLUMNS: [
        "RFQ-ID",
        "PLANT",
        "COMPONENT",
        "DESCRIPTION",
        "MANUFACTURERPARTNR.",
        "AVAILABLESTOCK(FREESTOCK),M,PC",
        "AVAILABLE€CUINCLUDED",
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
      },
      _normalizeColumnName: function (col) {
        return col
          .replace(/\r?\n/g, " ")
          .replace(/\s+/g, " ")
          .trim()
          .toUpperCase();
      },

      startWorkflowInstance: function (oPayload) {
        var model = this.getView().getModel();
        let data = {
          "definitionId": "eu10.builddevlapp.obsolete.obsoleteCreationProcess",
          "context": {
            "obsoleteitems": oPayload,
            "company": this.byId("companySelect").getSelectedKey()
          }
        }

        $.ajax({
          url: this._getWorkflowRuntimeBaseURL() + "/workflow-instances?environmentId=businessprocessworkflow",
          method: "POST",
          async: false,
          contentType: "application/json",
          headers: {
            "X-CSRF-Token": this._fetchToken(),
          },
          data: JSON.stringify(data),
          // success: function (result, xhr, data) {
          // },
          success: function (result) {
            sap.m.MessageBox.success(
              "Submitted successfully",
              {
                title: "Success",
                onClose: function () {
                  this._resetForm();
                }.bind(this)
              }
            );
          }.bind(this),
          error: function (request, status, error) {
            var response = JSON.parse(request.responseText);
            MessageBox.error(response.error.message);
          },
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
          oUploader.clear(); // ✅ removes last attached file
        }

        // Optional: clear stored file reference
        this._selectedFile = null;
      },



      onSubmit: function () {
        // var oUploader = this.byId("excelUploader");
        var oUploader = this.byId("excelUploader");

        if (!oUploader || !oUploader.oFileUpload || !oUploader.oFileUpload.files.length) {
          MessageBox.error("Please upload an Excel file");
          return;
        }

        var oFile = oUploader.oFileUpload.files[0];

        if (!oFile) {
          MessageBox.error("Please upload an Excel file");
          return;
        }

        var reader = new FileReader();
        reader.onload = (e) => {
          var workbook = XLSX.read(e.target.result, { type: "binary" });
          var sheetName = workbook.SheetNames[0];
          var sheet = workbook.Sheets[sheetName];

          // var excelData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          var excelData = XLSX.utils.sheet_to_json(sheet, {
            defval: "",
            cellDates: true
          });

          // 🔥 NORMALIZE ONCE HERE
          excelData = this._normalizeExcelData(excelData);


          try {
            var payloadData = this._validateAndMapExcel(excelData);
            console.log("Mapped Payload:", payloadData);

            // 🔥 assign mapped data to workflow payload
            // this.startWorkflowInstance(payloadData);

            // sap.m.MessageToast.show("Excel validated successfully");
            // this.startWorkflowInstance(payloadData);
            MessageBox.confirm(
              "Excel validated successfully.\n\nAre you sure you want to submit?",
              {
                title: "Confirm Submission",
                actions: [sap.m.MessageBox.Action.YES, sap.m.MessageBox.Action.NO],
                emphasizedAction: sap.m.MessageBox.Action.YES,
                onClose: function (oAction) {
                  if (oAction === sap.m.MessageBox.Action.YES) {
                    this.startWorkflowInstance(payloadData);
                  }
                }.bind(this)
              }
            );


          } catch (err) {
            MessageBox.error(err.message);
          }
        };

        reader.readAsBinaryString(oFile);
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

      // _validateAndMapExcel: function (excelData) {

      //   if (!excelData.length) {
      //     throw new Error("Excel file is empty");
      //   }

      //   // Normalize headers to UPPER CASE
      //   var headers = Object.keys(excelData[0]).map(h => h.toUpperCase());

      //   // Validate all required columns exist
      //   this._REQUIRED_COLUMNS.forEach(col => {
      //     if (!headers.includes(col)) {
      //       throw new Error(`Missing required column: ${col}`);
      //     }
      //   });

      //   // Validate rows + map data
      //   var mappedItems = excelData.map((row, index) => {

      //     // Check empty cells
      //     this._REQUIRED_COLUMNS.forEach(col => {
      //       var value = row[col] || row[col.toLowerCase()];
      //       if (value === "" || value === null) {
      //         throw new Error(`Empty value in column "${col}" at row ${index + 2}`);
      //       }
      //     });

      //     return {
      //       RFQID: row["RFQ - ID"],
      //       Plant: row["PLANT"],
      //       Component: row["COMPONENT"],
      //       Description: row["DESCRIPTION"],
      //       Manufacturer: row["MANUFACTURER PART NR."],
      //       AvailableStock: row["AVAILABLE STOCK (FREE STOCK), M, PC"],
      //       AvailableCU: row[" AVAILABLE €\r\nCU INCLUDED "],
      //       RangeOfCoverage: row["RANGE OF COVERAGE IN MONTHS"],
      //       PN: row["PN"],
      //       Customer: row["CUSTOMER"],
      //       EndCustomer: row["END CUSTOMER"],
      //       LastConsumptionDate: row["LAST CONSUMPTION"],
      //       Reason: row["REASON"],
      //       Caused: row["CAUSED"],
      //       Weight: row["WEIGHT (ON BASE UNIT)"],
      //       TotalAmount: row["TOTAL AMOUNT"]
      //     };
      //   });

      //   return mappedItems;
      // },


      // _normalizeRowKeys: function (row) {
      //   var normalized = {};
      //   Object.keys(row).forEach(key => {
      //     normalized[key.trim().toUpperCase()] = row[key];
      //   });
      //   return normalized;
      // },


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


      // _validateAndMapExcel: function (excelData) {

      //   if (!excelData.length) {
      //     throw new Error("Excel file is empty");
      //   }

      //   // Normalize header names
      //   var normalizedHeaders = Object.keys(excelData[0])
      //     .map(h => h.trim().toUpperCase());

      //   // Validate required columns
      //   this._REQUIRED_COLUMNS.forEach(col => {
      //     if (!normalizedHeaders.includes(col)) {
      //       throw new Error(`Missing required column: ${col}`);
      //     }
      //   });






      //   // Validate rows + map
      //   var mappedItems = excelData.map((row, index) => {

      //     var r = this._normalizeRowKeys(row); // ✅ IMPORTANT

      //     // Validate empty cells
      //     this._REQUIRED_COLUMNS.forEach(col => {
      //       // if (col == 'AVAILABLE €\r\nCU INCLUDED') {
      //       //   col = "AVAILABLE € CU INCLUDED"
      //       // }
      //       if (r[col] === "" || r[col] === null || r[col] === undefined) {
      //         throw new Error(`Empty value in column "${col}" at row ${index + 2}`);
      //       }
      //     });

      //     // ✅ FINAL CORRECT MAPPING
      //     // return {
      //     //   RFQID: r["RFQ - ID"],
      //     //   Plant: r["PLANT"],
      //     //   Component: r["COMPONENT"],
      //     //   Description: r["DESCRIPTION"],
      //     //   Manufacturer: r["MANUFACTURER PART NR."],
      //     //   AvailableStock: r["AVAILABLE STOCK (FREE STOCK), M, PC"],
      //     //   AvailableCU: r["AVAILABLE €CU INCLUDED"],
      //     //   RangeOfCoverage: r["RANGE OF COVERAGE IN MONTHS"],
      //     //   PN: r["PN"],
      //     //   Customer: r["CUSTOMER"],
      //     //   EndCustomer: r["END CUSTOMER"],
      //     //   LastConsumptionDate: this._formatExcelDate(r["LAST CONSUMPTION"]),
      //     //   Reason: r["REASON"],
      //     //   Caused: r["CAUSED"],
      //     //   Weight: r["WEIGHT (ON BASE UNIT)"],
      //     //   TotalAmount: r["TOTAL AMOUNT"]
      //     // };
      //     return {
      //       RFQID: this._toString(r["RFQ - ID"]),
      //       Plant: r["PLANT"],
      //       Component: r["COMPONENT"],
      //       Description: r["DESCRIPTION"],
      //       Manufacturer: this._toString(
      //         r["MANUFACTURER PART NR."]
      //       ),
      //       AvailableStock: this._toDecimal(
      //         r["AVAILABLE STOCK (FREE STOCK), M, PC"]
      //       ),

      //       AvailableCU: this._toDecimal(
      //         r["AVAILABLE € CU INCLUDED"]
      //       ),

      //       //   AvailableCU: this._toDecimal(
      //       //   r["\\"AVAILABLE € ↵CU INCLUDED\\"]
      //       // ),

      //       RangeOfCoverage: this._toDecimal(
      //         r["RANGE OF COVERAGE IN MONTHS"]
      //       ),

      //       PN: r["PN"],
      //       Customer: r["CUSTOMER"],
      //       EndCustomer: r["END CUSTOMER"],
      //       LastConsumptionDate: this._formatExcelDate(
      //         r["LAST CONSUMPTION"]
      //       ),
      //       Reason: r["REASON"],
      //       Caused: r["CAUSED"],
      //       Weight: this._toDecimal(
      //         r["WEIGHT (ON BASE UNIT)"]
      //       ),
      //       TotalAmount: this._toDecimal(
      //         r["TOTAL AMOUNT"]
      //       )
      //     };

      //   });

      //   return mappedItems;
      // },


      _validateAndMapExcel: function (excelData) {

        if (!excelData.length) {
          throw new Error("Excel file is empty");
        }

        this._REQUIRED_COLUMNS.forEach(col => {
          if (!(col in excelData[0])) {
            throw new Error(`Missing required column: ${col}`);
          }
        });

        return excelData.map((r, index) => {

          this._REQUIRED_COLUMNS.forEach(col => {
            if (r[col] === "" || r[col] === null || r[col] === undefined) {
              throw new Error(
                `Empty value in column "${col}" at row ${index + 2}`
              );
            }
          });

          return {
            RFQID: this._toString(r["RFQ-ID"]),
            Plant: r["PLANT"],
            Component: r["COMPONENT"],
            Description: r["DESCRIPTION"],
            Manufacturer: this._toString(r["MANUFACTURERPARTNR."]),
            AvailableStock: this._toDecimal(
              r["AVAILABLESTOCK(FREESTOCK),M,PC"]
            ),
            AvailableCU: this._toDecimal(
              r["AVAILABLE€CUINCLUDED"]
            ),
            RangeOfCoverage: this._toDecimal(
              r["RANGEOFCOVERAGEINMONTHS"]
            ),
            PN: r["PN"],
            Customer: r["CUSTOMER"],
            EndCustomer: r["ENDCUSTOMER"],
            LastConsumptionDate: this._formatExcelDate(
              r["LASTCONSUMPTION"]
            ),
            Reason: r["REASON"],
            Caused: this._validateCaused(
              r["CAUSED"],
              index
            ),
            Weight: this._toDecimal(
              r["WEIGHT(ONBASEUNIT)"]
            ),
            TotalAmount: this._toDecimal(
              r["TOTALAMOUNT"]
            )
          };
        });
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


      _formatExcelDate: function (value) {

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
      // _formatExcelDate: function (value) {

      //   var dateObj;

      //   // Case 1: JS Date object
      //   if (value instanceof Date) {
      //     dateObj = value;
      //   }
      //   // Case 2: Excel serial number
      //   else if (typeof value === "number") {
      //     var excelEpoch = new Date(Date.UTC(1899, 11, 30));
      //     dateObj = new Date(excelEpoch.getTime() + value * 86400000);
      //   }
      //   // Case 3: Already string (assume correct format)
      //   else if (typeof value === "string") {
      //     return value; // e.g. "17.06.2024"
      //   }
      //   else {
      //     return "";
      //   }

      //   // Format → DD.MM.YYYY
      //   var day = String(dateObj.getUTCDate()).padStart(2, "0");
      //   var month = String(dateObj.getUTCMonth() + 1).padStart(2, "0");
      //   var year = dateObj.getUTCFullYear();

      //   return `${day}.${month}.${year}`;
      // },


      // _toDecimal: function (value) {

      //   if (value === null || value === undefined || value === "") {
      //     return 0;
      //   }

      //   // If already a number
      //   if (typeof value === "number") {
      //     return Number(value.toFixed(2));
      //   }

      //   // If string → clean it
      //   if (typeof value === "string") {
      //     var cleaned = value
      //       .replace(/€/g, "")
      //       .replace(/,/g, "")
      //       .trim();

      //     var num = parseFloat(cleaned);
      //     return isNaN(num) ? 0 : Number(num.toFixed(2));
      //   }

      //   return 0;
      // },

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

      // onUpload: function () {
      //   // var oUploader = this.byId("excelUploader");
      //   var oUploader = this.byId("excelUploader");

      //   if (!oUploader || !oUploader.oFileUpload || !oUploader.oFileUpload.files.length) {
      //     sap.m.MessageBox.error("Please upload an Excel file");
      //     return;
      //   }

      //   var oFile = oUploader.oFileUpload.files[0];


      //   // let data = {
      //   //   "cmisaction": "createFolder",
      //   //   "propertyId[0]": "cmis:name",
      //   //   "propertyValue[0]": "test",
      //   //   "propertyId[1]": "cmis:objectTypeId",
      //   //   "propertyValue[1]": "cmis:folder",
      //   //   "succinct": true
      //   // }
      //   fetch("/dmsrepo/root", {
      //     method: "POST",
      //     headers: {
      //       "Content-Type": "application/json"
      //     },
      //     body: JSON.stringify({
      //       properties: {
      //         "cmis:name": "testFolder",
      //         "cmis:objectTypeId": "cmis:folder"
      //       }
      //     })
      //   })
      //     .then(res => res.json())
      //     .then(data => {
      //       MessageToast.show("Folder created: " + data.succinctProperties["cmis:objectId"]);
      //     })
      //     .catch(err => {
      //       MessageToast.show("Error creating folder");
      //     });



      //   // $.ajax({
      //   //   url: this._getWorkflowRuntimeBaseURL1(),
      //   //   method: "POST",
      //   //   async: false,
      //   //   contentType: "application/json",
      //   //   headers: {
      //   //     "X-CSRF-Token": this._fetchToken(),
      //   //   },
      //   //   data: JSON.stringify(data),
      //   //   success: function (result, xhr, data) {
      //   //     // model.setProperty(
      //   //     console.log(result);
      //   //     //   "/apiResponse",
      //   //     //   JSON.stringify(result, null, 4)
      //   //     // );
      //   //   },
      //   //   error: function (request, status, error) {
      //   //     var response = JSON.parse(request.responseText);
      //   //     // model.setProperty(
      //   //     //   "/apiResponse",
      //   //     //   JSON.stringify(response, null, 4)
      //   //     // );
      //   //   },
      //   // });
      // },
      // _getWorkflowRuntimeBaseURL1: function () {
      //   var ui5CloudService = this.getOwnerComponent().getManifestEntry("/sap.cloud/service").replaceAll(".", "");
      //   var ui5ApplicationName = this.getOwnerComponent().getManifestEntry("/sap.app/id").replaceAll(".", "");
      //   var appPath = `${ui5CloudService}.${ui5ApplicationName}`;
      //   return `/${appPath}/browser/cc918620-3f34-4544-b260-cb5ad8a568d7/root`

      // },

      onUpload: async function () {
        try {
          const repositoryId = "cc918620-3f34-4544-b260-cb5ad8a568d7"; // from /browser
          const folderName = "sapui";


          const formData = new FormData();
          formData.append("cmisaction", "createFolder");
          formData.append("propertyId[0]", "cmis:name");
          formData.append("propertyValue[0]", folderName);
          formData.append("propertyId[1]", "cmis:objectTypeId");
          formData.append("propertyValue[1]", "cmis:folder");
          formData.append("succinct", true);



          $.ajax({
            url: `/dms/browser/${repositoryId}/root`,
            method: "POST",
            data: formData,
            processData: false,
            contentType: false,
            headers: {
              "Accept": "application/json"
            },


            // headers: {
            //   "X-CSRF-Token": csrfToken
            // },


            // headers: {
            //   // OAuth token copied from Postman
            //   "Authorization": `Bearer ${token}`
            // },
            success: function (data) {
              sap.m.MessageToast.show("Folder created successfully");
              console.log(data);
            },
            error: function (err) {
              MessageBox.error("Failed to create folder");
            }
          });


          if (!response.ok) {
            throw new Error(await response.text());
          }

          const result = await response.json();
          console.log("Folder Created:", result);

          sap.m.MessageToast.show("Folder created successfully");
        } catch (err) {
          console.error(err);
          MessageBox.error("Failed to create folder");
        }
      },


      onUpload1: function () {


        var sRepositoryId = "your_repository_id_here"; // e.g., "abc123-4567"

        // Folder name: Hardcoded or from an input field
        var sFolderName = "MyNewFolder";
        // Example with input: this.byId("folderNameInput").getValue();

        if (!sFolderName) {
          MessageBox.error("Please enter a folder name.");
          return;
        }

        // Proxied URL: /dms/browser/<repoId>/root for root folder creation
        var sUrl = "/dms/browser/" + sRepositoryId + "/root";

        // CMIS createFolder parameters (form-urlencoded)
        var oData = {
          cmisaction: "createFolder",
          "propertyId[0]": "cmis:name",
          "propertyValue[0]": sFolderName,
          "propertyId[1]": "cmis:objectTypeId",
          "propertyValue[1]": "cmis:folder",
          succinct: "true" // Optional: returns simplified JSON response
        };

        // AJAX POST call
        jQuery.ajax({
          url: sUrl,
          type: "POST",
          data: oData,
          contentType: "application/x-www-form-urlencoded",
          dataType: "json", // Expect JSON response
          success: function (oResponse) {
            MessageToast.show("Folder '" + sFolderName + "' created successfully!\nObject ID: " + oResponse.succinctProperties["cmis:objectId"]);
            // Optional: Refresh your folder list/table here
          },
          error: function (oXHR, sStatus, sError) {
            var sMsg = "Error creating folder: " + sStatus + " - " + sError;
            if (oXHR.responseJSON && oXHR.responseJSON.message) {
              sMsg += "\nDetails: " + oXHR.responseJSON.message;
            }
            MessageBox.error(sMsg);
          }
        });
      },
      onTestRepositories: function () {
        var sUrl = "/dms/browser";  // Proxies to /rest/v2/repositories

        jQuery.ajax({
          url: sUrl,
          type: "GET",
          dataType: "json",
          success: function (oResponse) {
            console.log("Repositories:", oResponse);
            sap.m.MessageToast.show("Repositories fetched! Check console for details.");
            // Look for your repo and note the "cmisRepositoryId" field
          },
          error: function (oXHR) {
            console.error("Error:", oXHR.responseText);
            sap.m.MessageBox.error("Failed to fetch repositories: " + oXHR.status);
          }
        });
      }







    });
  }
);
