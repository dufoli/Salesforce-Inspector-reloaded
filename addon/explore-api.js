/* global React ReactDOM */
import {sfConn, apiVersion} from "./inspector.js";
/* global initButton */
import {QueryHistory} from "./data-load.js";

class Model {
  constructor(sfHost, args) {
    this.sfHost = sfHost;
    this.sfLink = "https://" + sfHost;
    this.spinnerCount = 0;
    this.title = "API Request";
    this.userInfo = "...";
    this.expandSavedOptions = false;
    this.requestName = "";
    this.bodyType = "json";
    function compare(a, b) {
      return a.request == b.request && a.requestType == b.requestType && a.httpMethod == b.httpMethod && a.apiUrl == b.apiUrl && a.soapType == b.soapType && a.name == b.name;
    }
    function sort(a, b) {
      return (a.request > b.request) ? 1 : ((b.request > a.request) ? -1 : 0);
    }
    this.requestHistory = new QueryHistory("insextRequestHistory", 100, compare, sort);
    this.savedHistory = new QueryHistory("insextSavedRequestHistory", 50, compare, sort);
    this.apiResponse = null;
    this.selectedTextView = null;
    this.requestType = "REST";
    this.soapMethods = [];
    this.soapType = "Partner";
    this.operationToParams = {};
    this.apiUrl;
    this.payload = "";
    this.httpMethod = "GET";

    if (args.has("apiUrls")) {
      let apiUrls = args.getAll("apiUrls");
      this.title = apiUrls.length + " API requests, e.g. " + apiUrls[0];
      this.apiUrl = apiUrls[0];
      let apiPromise = Promise.all(apiUrls.map(url => sfConn.rest(url, {withoutCache: true})));
      this.performRequest(apiPromise);
    } else if (args.has("checkDeployStatus")) {
      let wsdl = sfConn.wsdl(apiVersion, "Metadata");
      this.title = "checkDeployStatus: " + args.get("checkDeployStatus");
      this.apiUrl = wsdl.servicePortAddress;
      let apiPromise = sfConn.soap(wsdl, "checkDeployStatus", {id: args.get("checkDeployStatus"), includeDetails: true});
      this.performRequest(apiPromise);
    } else {
      let apiUrl = args.get("apiUrl") || "/services/data/";
      this.title = apiUrl;
      this.apiUrl = apiUrl;
      let apiPromise = sfConn.rest(apiUrl, {withoutCache: true});
      this.performRequest(apiPromise);
    }
    let requestTemplatesRawValue = localStorage.getItem("requestTemplates");
    if (requestTemplatesRawValue && requestTemplatesRawValue != "[]") {
      try {
        this.requestTemplates = JSON.parse(requestTemplatesRawValue);
      } catch (err) {
        //try old format which do not support comments
        this.requestTemplates = requestTemplatesRawValue.split("//");
      }
    } else {
      this.requestTemplates = [
        {requestType: "REST", httpMethod: "GET", request: "", soapType: "Partner", apiUrl: "/services/data/", name: "Services list"},
        {requestType: "REST", httpMethod: "POST", request: "", soapType: "Partner", apiUrl: `/services/data/v${apiVersion}/sobjects/Account`, name: "Update account rest"},
        {requestType: "REST", httpMethod: "GET", request: "", soapType: "Partner", apiUrl: `/services/data/v${apiVersion}/query/?q=SELECT+Id,+Name+FROM+Account+LIMIT+10`, name: "Select query"},
        {requestType: "REST", httpMethod: "POST", request: "{ \"query\": \"query accounts { uiapi { query { Account { edges { node { Id  Name { value } } } } } } }\"}", soapType: "Partner", apiUrl: `/services/data/v${apiVersion}/graphql`, name: "Services list"},
        {requestType: "REST", httpMethod: "GET", request: "", soapType: "Partner", apiUrl: `/services/data/v${apiVersion}/metadata/deployRequest/deployRequestId?includeDetails=true`, name: "Deploy status"},
        {requestType: "REST", httpMethod: "POST", request: "{\"object\": \"Account\", \"contentType\" : \"CSV\", \"operation\" : \"insert\", \"lineEnding\" : \"CRLF\"}", soapType: "Partner", apiUrl: `/services/data/v${apiVersion}/jobs/ingest/`, name: "Bulk create job"},
        {requestType: "REST", httpMethod: "POST", request: "Name,ShippingCity,NumberOfEmployees,AnnualRevenue,Website,Description\r\nLorem Ipsum,Milano,2676,912260031,https://ft.com/lacus/at.jsp,\"Lorem ipsum dolor sit amet\"", soapType: "Partner", bodyType: "csv", apiUrl: `/services/data/v${apiVersion}/jobs/ingest/[jobId]/batches/`, name: "Bulk insert job"},
        {requestType: "REST", httpMethod: "GET", request: "", soapType: "Partner", apiUrl: `/services/data/v${apiVersion}/chatter/feeds/news/me/feed-elements`, name: "chatter News feed"},
        {requestType: "REST", httpMethod: "GET", request: "", soapType: "Partner", apiUrl: `/services/data/v${apiVersion}/analytics/reports/[ReportId]?includeDetails=true`, name: "Report data"},
        {requestType: "REST", httpMethod: "POST", request: "{\"FullName\": \"Carbon_Comparison_Channel__chn\", \"Metadata\": { \"channelType\": \"event\", \"label\": \"Carbon Comparison Channel\"}", soapType: "Partner", apiUrl: `/services/data/v${apiVersion}/tooling/sobjects/PlatformEventChannel`, name: "Create platform event channel"},
        {requestType: "REST", httpMethod: "POST", request: "{\"FullName\": \"Carbon_Comparison_Channel_chn_Carbon_Comparison_e\", \"Metadata\": {\"eventChannel\": \"Carbon_Comparison_Channel__chn\", \"selectedEntity\": \"Carbon_Comparison__e\"}}", soapType: "Partner", apiUrl: `/services/data/v${apiVersion}/tooling/sobjects/PlatformEventChannelMember`, name: "Create platform event channel member"},
      ];
    }
    this.spinFor(sfConn.soap(sfConn.wsdl(apiVersion, "Partner"), "getUserInfo", {}).then(res => {
      this.userInfo = res.userFullName + " / " + res.userName + " / " + res.organizationName;
    }));

  }
  /**
   * Notify React that we changed something, so it will rerender the view.
   * Should only be called once at the end of an event or asynchronous operation, since each call can take some time.
   * All event listeners (functions starting with "on") should call this function if they update the model.
   * Asynchronous operations should use the spinFor function, which will call this function after the asynchronous operation completes.
   * Other functions should not call this function, since they are called by a function that does.
   * @param cb A function to be called once React has processed the update.
   */
  didUpdate(cb) {
    if (this.reactCallback) {
      this.reactCallback(cb);
    }
  }
  /**
   * Show the spinner while waiting for a promise.
   * didUpdate() must be called after calling spinFor.
   * didUpdate() is called when the promise is resolved or rejected, so the caller doesn't have to call it, when it updates the model just before resolving the promise, for better performance.
   * @param promise The promise to wait for.
   */
  spinFor(promise) {
    this.spinnerCount++;
    promise
      .catch(err => {
        console.error("spinFor", err);
      })
      .then(() => {
        this.spinnerCount--;
        this.didUpdate();
      })
      .catch(err => console.log("error handling failed", err));
  }
  openSubUrl(subUrl) {
    let args = new URLSearchParams();
    args.set("host", this.sfHost);
    args.set("apiUrl", subUrl.apiUrl);
    return "explore-api.html?" + args;
  }
  openGroupUrl(groupUrl) {
    let args = new URLSearchParams();
    args.set("host", this.sfHost);
    for (let url of groupUrl.apiUrls) {
      args.append("apiUrls", url);
    }
    return "explore-api.html?" + args;
  }
  performRequest(apiPromise) {
    this.requestHistory.add({request: this.payload, requestType: this.requestType, httpMethod: this.httpMethod, bodyType: this.bodyType, apiUrl: this.apiUrl, soapType: this.soapType});
    this.spinFor(apiPromise.then(result => {
      this.parseResponse(result, "Success");
    }, err => {
      this.parseResponse(err.detail || err.message, "Error");
    }));
  }
  parseResponse(result, status) {
    /*
    Transform an arbitrary JSON structure (the `result` vaiable) into a list of two-dimensional TSV tables (the `textViews` variable), that can easily be copied into for example Excel.
    Each two-dimensional table corresponds to an array or set of related arrays in the JSON data.

    For example in a Sobject Describe, the list of fields is one table. Each row is a field, and each column is a property of that field.
    The list of picklist values is another table. Each column is a property of the picklist value, or a property of the field to which the picklist value belongs (i.e. a column inherited from the parent table).

    Map<String,TableView> tViews; // Map of all tables, keyed by the name of each table
    interface TableView {
      String name; // Name of the table, a JSON path that matches each row of the table
      TableView? parent; // For nested tables, contains a link to the parent table. A child table inherits all columns from its parent. Inherited columns are added to the end of a table.
      TableRow[] rows;
      Map<String,void> columnMap; // The set of all columns in this table, excluding columns inherited from the parent table
      String[]? columnList; // The list of all columns in this table, including columns inherited from the parent table
    }
    interface TableRow {
      JsonValue value; // The value of the row, as a JSON structure not yet flattened into row format.
      TableRow parent; // For nested tables, contains a link to the corresponding row in the parent table. A child row inherits all columns from its parent. Inherited columns are added to the end of a row.
      any[]? cells; // The list of all cells in this row, matching the columns in the table, including data inherited from the parent row
    }
    TextView[] textViews;
    interface TextView {
      String name; // Name of the table
      String value; // The table serialized in TSV format
      any[][]? table; // The table
    }

    In addition to building the table views of the JSON structure, we also scan it for values that look like API resource URLs, so we can display links to these.
    ApiSubUrl[] apiSubUrls;
    interface ApiSubUrl {
      String jsonPath; // The JSON path where the resource URL was found
      String apiUrl; // The URL
      String label; // A label describing the URL
    }

    We also group these URLs the same way we build tables, allowing the user to request all related resources in one go. For example, given a global describe result, the user can fetch object describes for all objects in one click.
    ApiGroupUrl[] apiGroupUrls;
    interface ApiGroupUrl {
      String jsonPath; // The JSON path where the resource URLs were found
      String[] apiUrls; // The related URLs
      String label; // A label describing the URLs
    }

    TODO: This transformation does not work in an ideal way on SOAP responses, since for those we can only detect an array if it has two or more elements.
    For example, "@.x.*.y.*" in the following shows two rows "2" and "3", where it should show three rows "1", "2", and "3":
    display(XML.parse(new DOMParser().parseFromString("<root><x><y>1</y></x><x><y>2</y><y>3</y></x></root>", "text/xml").documentElement))
    */

    // Recursively explore the JSON structure, discovering tables and their rows and columns.
    let apiSubUrls = [];
    let groupUrls = {};
    let textViews = [
      {name: "Raw JSON", value: JSON.stringify(result, null, "    ")}
    ];
    let tRow = {value: result, cells: null, parent: null}; // The root row
    let tViews = {
      "@": {name: "@", parent: null, rows: [tRow], columnMap: {}, columnList: null} // Dummy root table, always contains one row
    };
    exploreObject2(result, tRow, "", tViews["@"], "@");
    function exploreObject2(object /*JsonValue*/, tRow /*TableRow*/, columnName /*String, JSON path relative to tView.name*/, tView /*TableView*/, fullName /*String, JSON path including array indexes*/) {
      // Create the new column, if we have not created it already
      tView.columnMap[columnName] = true;

      if (object instanceof Array) {
        // Create a new table, if we have not created it already
        let childViewName = tView.name + columnName + ".*";
        let childView;
        tViews[childViewName] = childView = tViews[childViewName] || {name: childViewName, parent: tView, rows: [], columnMap: {}, columnList: null};

        for (let i = 0; i < object.length; i++) {
          if (object[i] && typeof object[i] == "object") {
            object[i]["#"] = i;
          }

          // Create the new row
          let childRow = {value: object[i], cells: null, parent: tRow};
          childView.rows.push(childRow);

          exploreObject2(object[i], childRow, "", childView, fullName + "." + i);
        }
      } else if (object && typeof object == "object") {
        for (let key in object) {
          exploreObject2(object[key], tRow, columnName + "." + key, tView, fullName + "." + key);
        }
      }

      if (typeof object == "string" && object.startsWith("/services/data/")) {
        apiSubUrls.push({jsonPath: fullName, apiUrl: object, label: object});
        if (tView.name != "@") {
          if (!groupUrls[tView.name + columnName]) {
            groupUrls[tView.name + columnName] = [];
          }
          groupUrls[tView.name + columnName].push(object);
        }
      }
    }

    // Build each of the discovered tables. Turn columns into a list, turn each row into a list matching the columns, and serialize as TSV.
    // Note that the tables are built in the order they are discovered. This means that a child table is always built after its parent table.
    // We can therefore re-use the build of the parent table when building the child table.
    for (let tView of Object.values(tViews)) {
      // Add own columns
      tView.columnList = Object.keys(tView.columnMap).map(column => tView.name + column);
      // Copy columns from parent table
      if (tView.parent) {
        tView.columnList = [...tView.columnList, ...tView.parent.columnList];
      }
      if (tView.rows && tView.rows.length == 0) {
        continue;
      }
      let table = [tView.columnList];
      // Add rows
      for (let row of tView.rows) {
        // Add cells to the row, matching the found columns
        row.cells = Object.keys(tView.columnMap).map(column => {
          // Find the value of the cell
          let fields = column.split(".");
          fields.splice(0, 1);
          let value = row.value;
          for (let field of fields) {
            if (typeof value != "object") {
              value = null;
            }
            if (value != null) {
              value = value[field];
            }
          }
          if (value instanceof Array) {
            value = "[Array " + value.length + "]";
          }
          return value;
        });
        // Add columns from parent row
        if (row.parent) {
          row.cells = [...row.cells, ...row.parent.cells];
        }
        table.push(row.cells);
      }
      let csvSignature = csvSerialize([
        ["Salesforce Inspector - API Explorer"],
        ["URL", this.title],
        ["Rows", tView.name],
        ["Extract time", new Date().toISOString()]
      ], "\t") + "\r\n\r\n";
      textViews.push({name: "Rows: " + tView.name + " (for copying to Excel)", value: csvSignature + csvSerialize(table, "\t")});
      textViews.push({name: "Rows: " + tView.name + " (for viewing)", table});
    }
    this.apiResponse = {
      status,
      textViews,
      // URLs to further explore the REST API, not grouped
      apiSubUrls,
      // URLs to further explore the REST API, grouped by table columns
      apiGroupUrls: Object.entries(groupUrls).map(([groupKey, apiUrls]) => ({jsonPath: groupKey, apiUrls, label: apiUrls.length + " API requests, e.g. " + apiUrls[0]})),
    };
    if (Array.isArray(result) && result.length == 1 && result[0] && result[0].errorCode) {
      this.selectedTextView = textViews[0];
    } else if (Array.isArray(result) && result.length < 100 && result[0] && result[0].url) {
      this.selectedTextView = null;
    } else if (textViews[0].value.length < 10000) {
      this.selectedTextView = textViews[0];
    } else {
      this.selectedTextView = null;
    }
    // Don't update selectedTextView. No radio button will be selected, leaving the text area blank.
    // The results can be quite large and take a long time to render, so we only want to render a result once the user has explicitly selected it.
  }
  setRequestType(requestType) {
    this.requestType = requestType;
    if (requestType == "SOAP") {
      //force refresh of operation
      this.setSoapType(this.soapType);
    } else {
      this.didUpdate();
    }
  }
  setSoapType(soapType) {
    this.soapType = soapType;
    sfConn.rest(sfConn.wsdl(apiVersion, soapType).wsdlUrl, {responseType: "document"}).then(wsdl => {
      let messages = {};
      this.operationToParams = {};
      this.soapMethods = [];
      let elementToTypes = {};

      for (let complexType of wsdl.querySelectorAll("complexType")) {
        let elementName = complexType.getAttribute("name");
        if (!elementName) {
          elementName = complexType.parentElement.getAttribute("name");
        }
        let params = {};
        for (let element of complexType.querySelectorAll("element")) {
          params[element.getAttribute("name")] = "";
        }
        elementToTypes[elementName] = params;
      }
      for (let message of wsdl.getElementsByTagName("message")) {
        let params = {};
        for (let part of message.getElementsByTagName("part")) {
          let element = part.getAttribute("element");
          if (element.includes(":")) {
            element = element.split(":", 2)[1];
          }
          if (elementToTypes[element]) {
            params = {...params, ...elementToTypes[element]};
          } else {
            params[part.getAttribute("name")] = "";
          }
        }
        messages[message.getAttribute("name")] = params;
      }
      let portTypes = wsdl.getElementsByTagName("portType");
      if (portTypes.length == 0) {
        return;
      }
      let operations = portTypes[0].getElementsByTagName("operation");
      for (let op of operations) {
        let opName = op.getAttribute("name");
        this.soapMethods.push(opName);
        let inputs = op.getElementsByTagName("input");
        if (inputs.length == 0) {
          continue;
        }
        let input = inputs[0];
        let msg = input.getAttribute("message");
        if (!msg || !opName) {
          continue;
        }
        if (msg.includes(":")) {
          msg = msg.split(":", 2)[1];
        }
        this.operationToParams[opName] = messages[msg];
      }
      if (this.soapMethods.length > 0) {
        this.setSoapMethod(this.soapMethods[0]);
      } else {
        this.didUpdate();
      }
    });
    this.didUpdate();
  }
  toggleSavedOptions() {
    this.expandSavedOptions = !this.expandSavedOptions;
  }
  setRequestName(value) {
    this.requestName = value;
  }
  selectHistoryEntry(selectedHistoryEntry) {
    if (selectedHistoryEntry != null) {
      this.requestType = selectedHistoryEntry.requestType;
      this.httpMethod = selectedHistoryEntry.httpMethod;
      this.soapType = selectedHistoryEntry.soapType;
      this.payload = selectedHistoryEntry.request;
      this.apiUrl = selectedHistoryEntry.apiUrl;
      this.bodyType = selectedHistoryEntry.bodyType || "json";
    }
  }
  selectRequestTemplate(selectedTemplate) {
    if (selectedTemplate != null) {
      this.requestType = selectedTemplate.requestType;
      this.httpMethod = selectedTemplate.httpMethod;
      this.payload = selectedTemplate.request;
      this.soapType = selectedTemplate.soapType;
      this.apiUrl = selectedTemplate.apiUrl;
      this.bodyType = selectedTemplate.bodyType || "json";
    }
    //this.editor.focus();
  }
  clearHistory() {
    this.requestHistory.clear();
  }
  selectSavedEntry(selectedSavedEntry) {
    if (selectedSavedEntry != null) {
      this.requestType = selectedSavedEntry.requestType;
      this.httpMethod = selectedSavedEntry.httpMethod;
      this.payload = selectedSavedEntry.request;
      this.soapType = selectedSavedEntry.soapType;
      this.apiUrl = selectedSavedEntry.apiUrl;
      this.bodyType = selectedSavedEntry.bodyType || "json";
    }
  }
  clearSavedHistory() {
    this.savedHistory.clear();
  }
  addToHistory() {
    this.savedHistory.add({request: this.payload, requestType: this.requestType, httpMethod: this.httpMethod, bodyType: this.bodyType, apiUrl: this.apiUrl, soapType: this.soapType, name: this.requestName});
  }
  removeFromHistory() {
    this.savedHistory.remove({request: this.payload, requestType: this.requestType, httpMethod: this.httpMethod, bodyType: this.bodyType, apiUrl: this.apiUrl, soapType: this.soapType, name: this.requestName});
  }
  /*getRequestToSave() {
    return this.requestName != "" ? this.requestName + ":" + this.payload : this.payload;
  }*/
  setHttpMethod(httpMethod) {
    this.httpMethod = httpMethod;
    this.didUpdate();
  }
  setBodyType(bodyType) {
    this.bodyType = bodyType;
    this.didUpdate();
  }
  formatforHuman(src) {
    let indent = -1;
    let matchTag;
    let result = "";
    let startIdx = 0;
    const tagRegExp = RegExp("<[\\/]?", "g");
    while ((matchTag = tagRegExp.exec(src)) !== null) {
      result += src.substring(startIdx, matchTag.index);
      startIdx = matchTag.index + matchTag[0].length;
      switch (matchTag[0]) {
        case "<":
          indent++;
          result += "\n" + "  ".repeat(indent) + matchTag[0];
          break;
        case "</":
          result += "\n" + "  ".repeat(indent) + matchTag[0];
          indent--;
          break;
        case "<\\":
          result += "\n" + "  ".repeat(indent) + matchTag[0];
          break;
        default:
          break;
      }
    }
    if (startIdx < src.length) {
      result += src.substring(startIdx, src.length - 1);
    }
    return result;
  }
  setSoapMethod(soapMethod) {
    //this.soapMethod = soapMethod;
    this.payload = this.formatforHuman(sfConn.formatSoapMessage(sfConn.wsdl(apiVersion, this.soapType), soapMethod, this.operationToParams[soapMethod], {}));
    this.didUpdate();
  }
  setUrl(url) {
    this.apiUrl = url;
    this.didUpdate();
  }
  setPayload(payload) {
    this.payload = payload;
    this.didUpdate();
  }
  execute() {
    switch (this.requestType) {
      case "SOAP":
        this.performRequest(sfConn.soap(sfConn.wsdl(apiVersion, this.soapType), null, this.payload));
        break;
      case "REST":
        try {
          if (this.httpMethod != "GET" && this.httpMethod != "DELETE") {
            if (this.bodyType == "json") {
              let body = JSON.parse(this.payload);
              this.performRequest(sfConn.rest(this.apiUrl, {method: this.httpMethod, bodyType: "json", body, withoutCache: true}));
            } else {
              this.performRequest(sfConn.rest(this.apiUrl, {method: this.httpMethod, bodyType: this.bodyType, body: this.payload, withoutCache: true}));
            }
          } else {
            this.performRequest(sfConn.rest(this.apiUrl, {method: this.httpMethod, withoutCache: true}));
          }
        } catch (e) {
          // ignore
          this.performRequest(sfConn.rest(this.apiUrl, {method: this.httpMethod, bodyType: "raw", body: ((this.httpMethod != "GET" && this.httpMethod != "DELETE") ? this.payload : null), withoutCache: true}));
        }

        break;
      default:
        break;
    }
  }
}

function csvSerialize(table, separator) {
  return table.map(row => row.map(text => "\"" + ("" + (text == null ? "" : text)).split("\"").join("\"\"") + "\"").join(separator)).join("\r\n");
}

let h = React.createElement;

class App extends React.Component {
  constructor(props) {
    super(props);
    this.setRequestType = this.setRequestType.bind(this);
    this.setSoapType = this.setSoapType.bind(this);
    this.setHttpMethod = this.setHttpMethod.bind(this);
    this.setBodyType = this.setBodyType.bind(this);
    this.setSoapMethod = this.setSoapMethod.bind(this);
    this.setUrl = this.setUrl.bind(this);
    this.setPayload = this.setPayload.bind(this);
    this.onExecute = this.onExecute.bind(this);
    this.onSelectHistoryEntry = this.onSelectHistoryEntry.bind(this);
    this.onSelectRequestTemplate = this.onSelectRequestTemplate.bind(this);
    this.onClearHistory = this.onClearHistory.bind(this);
    this.onSelectSavedEntry = this.onSelectSavedEntry.bind(this);
    this.onAddToHistory = this.onAddToHistory.bind(this);
    this.onRemoveFromHistory = this.onRemoveFromHistory.bind(this);
    this.onClearSavedHistory = this.onClearSavedHistory.bind(this);
    this.onSetRequestName = this.onSetRequestName.bind(this);
    this.onToggleSavedOptions = this.onToggleSavedOptions.bind(this);
    this.onSelectTextView = this.onSelectTextView.bind(this);
  }
  setRequestType(e) {
    let {model} = this.props;
    model.setRequestType(e.target.value);
  }
  setSoapType(e) {
    let {model} = this.props;
    model.setSoapType(e.target.value);
  }
  setSoapMethod(e) {
    let {model} = this.props;
    model.setSoapMethod(e.target.value);
  }
  setHttpMethod(e) {
    let {model} = this.props;
    model.setHttpMethod(e.target.value);
  }
  setBodyType(e) {
    let {model} = this.props;
    model.setBodyType(e.target.value);
  }
  setUrl(e) {
    let {model} = this.props;
    model.setUrl(e.target.value);
  }
  setPayload(e) {
    let {model} = this.props;
    model.setPayload(e.target.value);
  }
  onExecute() {
    let {model} = this.props;
    model.execute();
  }
  cleanCell(cell){
    return ((!cell || cell.toString() == "[object Object]") ? "" : cell.toString());
  }

  onSelectHistoryEntry(e) {
    let {model} = this.props;
    let selectedHistoryEntry = JSON.parse(e.target.value);
    model.selectHistoryEntry(selectedHistoryEntry);
    model.didUpdate();
  }
  onSelectRequestTemplate(e) {
    let {model} = this.props;
    let selectedTemplate = JSON.parse(e.target.value);
    model.selectRequestTemplate(selectedTemplate);
    model.didUpdate();
  }
  onClearHistory(e) {
    e.preventDefault();
    let r = confirm("Are you sure you want to clear the request history?");
    if (r == true) {
      let {model} = this.props;
      model.clearHistory();
      model.didUpdate();
    }
  }
  onSelectSavedEntry(e) {
    let {model} = this.props;
    let selectedSavedEntry = JSON.parse(e.target.value);
    model.selectSavedEntry(selectedSavedEntry);
    model.didUpdate();
  }
  onSelectTextView(e) {
    let {model} = this.props;
    model.selectedTextView = JSON.parse(e.target.value);
    model.didUpdate();
  }
  onAddToHistory(e) {
    e.preventDefault();
    let {model} = this.props;
    model.addToHistory();
    model.didUpdate();
  }
  onRemoveFromHistory(e) {
    e.preventDefault();
    let r = confirm("Are you sure you want to remove this saved request?");
    let {model} = this.props;
    if (r == true) {
      model.removeFromHistory();
    }
    model.toggleSavedOptions();
    model.didUpdate();
  }
  onClearSavedHistory(e) {
    e.preventDefault();
    let r = confirm("Are you sure you want to remove all saved requests?");
    let {model} = this.props;
    if (r == true) {
      model.clearSavedHistory();
    }
    model.toggleSavedOptions();
    model.didUpdate();
  }
  onSetRequestName(e) {
    let {model} = this.props;
    model.setRequestName(e.target.value);
    model.didUpdate();
  }
  onToggleSavedOptions(e) {
    e.preventDefault();
    let {model} = this.props;
    model.toggleSavedOptions();
    model.didUpdate();
  }
  render() {
    let {model} = this.props;
    document.title = model.title;
    let soapTypes = ["Enterprise", "Partner", "Apex", "Metadata", "Tooling"];
    let httpMethods = ["GET", "POST", "PUT", "PATCH", "DELETE"]; // not needed: "HEAD", "CONNECT", "OPTIONS", "TRACE"
    let bodyTypes = ["raw", "json", "csv", "xml"];
    return h("div", {},
      h("div", {id: "user-info"},
        h("a", {href: model.sfLink, className: "sf-link"},
          h("svg", {viewBox: "0 0 24 24"},
            h("path", {d: "M18.9 12.3h-1.5v6.6c0 .2-.1.3-.3.3h-3c-.2 0-.3-.1-.3-.3v-5.1h-3.6v5.1c0 .2-.1.3-.3.3h-3c-.2 0-.3-.1-.3-.3v-6.6H5.1c-.1 0-.3-.1-.3-.2s0-.2.1-.3l6.9-7c.1-.1.3-.1.4 0l7 7v.3c0 .1-.2.2-.3.2z"})
          ),
          " Salesforce Home"
        ),
        h("h1", {}, "Explore API"),
        h("span", {}, " / " + model.userInfo),
        h("div", {className: "flex-right"},
          h("div", {id: "spinner", role: "status", className: "slds-spinner slds-spinner_small slds-spinner_inline", hidden: model.spinnerCount == 0},
            h("span", {className: "slds-assistive-text"}),
            h("div", {className: "slds-spinner__dot-a"}),
            h("div", {className: "slds-spinner__dot-b"}),
          ),
        ),
      ),
      h("div", {className: "area", id: "query-area"},
        h("div", {className: "query-controls"},
          h("h1", {}, "Execute Request"),
          h("div", {className: "query-history-controls"},
            h("select", {value: "", onChange: this.onSelectRequestTemplate, className: "request-history", title: "Check documentation to customize templates"},
              h("option", {value: null, disabled: true, defaultValue: true, hidden: true}, "Templates"),
              model.requestTemplates.map(q => h("option", {key: JSON.stringify(q), value: JSON.stringify(q)}, (q.httpMethod + " " + q.apiUrl + " " + q.request).substring(0, 300)))
            ),
            h("div", {className: "button-group"},
              h("select", {value: "", onChange: this.onSelectHistoryEntry, className: "request-history"},
                h("option", {value: JSON.stringify(null), disabled: true}, "Request History"),
                model.requestHistory.list.map(q => h("option", {key: JSON.stringify(q), value: JSON.stringify(q)}, (q.httpMethod + " " + q.apiUrl + " " + q.request).substring(0, 300)))
              ),
              h("button", {onClick: this.onClearHistory, title: "Clear Request History"}, "Clear")
            ),
            h("div", {className: "pop-menu saveOptions", hidden: !model.expandSavedOptions},
              h("a", {href: "#", onClick: this.onRemoveFromHistory, title: "Remove request from saved history"}, "Remove Saved Request"),
              h("a", {href: "#", onClick: this.onClearSavedHistory, title: "Clear saved history"}, "Clear Saved Requests")
            ),
            h("div", {className: "button-group"},
              h("select", {value: "", onChange: this.onSelectSavedEntry, className: "request-history"},
                h("option", {value: JSON.stringify(null), disabled: true}, "Saved Requests"),
                model.savedHistory.list.map(q => h("option", {key: JSON.stringify(q), value: JSON.stringify(q)}, (q.httpMethod + " " + q.apiUrl + " " + q.request).substring(0, 300)))
              ),
              h("input", {placeholder: "Request Label", type: "save", value: model.requestName, onInput: this.onSetRequestName}),
              h("button", {onClick: this.onAddToHistory, title: "Add request to saved history"}, "Save Request"),
              h("button", {className: model.expandSavedOptions ? "toggle contract" : "toggle expand", title: "Show More Options", onClick: this.onToggleSavedOptions}, h("div", {className: "button-toggle-icon"}))
            ),
          ),
        ),
        h("div", {className: "form-line"},
          h("label", {className: "form-input"},
            h("span", {className: "form-label"}, "Type")),
          h("span", {className: "form-value"},
            h("select", {name: "requestType", onChange: this.setRequestType}, h("option", {value: "REST"}, "REST"), h("option", {value: "SOAP"}, "SOAP"))),
          h("button", {className: "highlighted", onClick: this.onExecute}, "Execute")
        ),
        h("div", {hidden: model.requestType != "SOAP", className: "form-line"},
          h("label", {className: "form-input"},
            h("span", {className: "form-label"}, "WSDL")),
          h("span", {className: "form-value"},
            h("select", {name: "soapType", onChange: this.setSoapType, value: model.soapType}, soapTypes.map(s => h("option", {value: s, key: s}, s))))),
        h("div", {hidden: model.requestType != "SOAP", className: "form-line"},
          h("label", {className: "form-input"},
            h("span", {className: "form-label"}, "Soap Method")),
          h("span", {className: "form-value"},
            h("select", {name: "soapMethod", onChange: this.setSoapMethod}, model.soapMethods.map(s => h("option", {value: s, key: s}, s))))),
        h("div", {hidden: model.requestType != "REST", className: "form-line"},
          h("label", {className: "form-input"},
            h("span", {className: "form-label"}, "Method")),
          h("span", {className: "form-value"},
            h("select", {name: "httpMethod", onChange: this.setHttpMethod, value: model.httpMethod}, httpMethods.map(s => h("option", {value: s, key: s}, s))))),
        h("div", {hidden: (model.requestType != "REST" || (model.httpMethod == "GET" || model.httpMethod == "DELETE")), className: "form-line"},
          h("label", {className: "form-input"},
            h("span", {className: "form-label"}, "Body type")),
          h("span", {className: "form-value"},
            h("select", {name: "bodyType", onChange: this.setBodyType, value: model.bodyType}, bodyTypes.map(s => h("option", {value: s, key: s}, s))))),
        h("div", {className: "form-line", hidden: (model.requestType == "REST" && (model.httpMethod == "GET" || model.httpMethod == "DELETE"))},
          h("label", {className: "form-input"},
            h("span", {className: "form-label"}, "Payload")),
          h("span", {className: "form-value"},
            h("textarea", {name: "httpBody", value: model.payload, onChange: this.setPayload}))),
        //TODO HTTP headers
        h("div", {hidden: model.requestType != "REST", className: "form-line"},
          h("label", {className: "form-input"},
            h("span", {className: "form-label"}, "URL")),
          h("span", {className: "form-value"},
            h("input", {name: "url", onChange: this.setUrl, value: model.apiUrl})))
      ),
      h("div", {className: "area", id: "result-area"},
        h("div", {className: "result-bar"},
          h("h1", {}, "Request Result"),
          model.apiResponse && h("div", {},
            h("select", {value: JSON.stringify(model.selectedTextView), onChange: this.onSelectTextView, className: "textview-format"},
              h("option", {value: JSON.stringify(null), disabled: true}, "Result format"),
              model.apiResponse.textViews.map(q => h("option", {key: JSON.stringify(q), value: JSON.stringify(q)}, q.name))
            ),
            h("span", {className: model.apiResponse.status == "Error" ? "status-error" : "status-success"}, "Status: " + model.apiResponse.status),
          ),
        ),
        h("div", {id: "result-table", ref: "scroller"},
          model.apiResponse && h("div", {},
            model.selectedTextView && !model.selectedTextView.table && h("div", {},
              h("textarea", {readOnly: true, value: model.selectedTextView.value})
            ),
            model.selectedTextView && model.selectedTextView.table && h("div", {},
              h("table", {className: "scrolltable-scrolled"},
                h("tbody", {},
                  model.selectedTextView.table.map((row, key) =>
                    h("tr", {key},
                      row.map((cell, key) =>
                        h("td", {key, className: "scrolltable-cell"}, this.cleanCell(cell))
                      )
                    )
                  )
                )
              )
            ),
            model.apiResponse.apiGroupUrls && h("ul", {},
              model.apiResponse.apiGroupUrls.map((apiGroupUrl, key) =>
                h("li", {key},
                  h("a", {href: model.openGroupUrl(apiGroupUrl)}, apiGroupUrl.jsonPath),
                  " - " + apiGroupUrl.label
                )
              )
            ),
            model.apiResponse.apiSubUrls && h("ul", {},
              model.apiResponse.apiSubUrls.map((apiSubUrl, key) =>
                h("li", {key},
                  h("a", {href: model.openSubUrl(apiSubUrl)}, apiSubUrl.jsonPath),
                  " - " + apiSubUrl.label
                )
              )
            )
          ),
          h("a", {href: "https://www.salesforce.com/us/developer/docs/api_rest/", target: "_blank"}, "REST API documentation"),
          " Open your browser's ",
          h("b", {}, "F12 Developer Tools"),
          " and select the ",
          h("b", {}, "Console"),
          " tab to make your own API calls."
        ),
      )
    );
  }

}

{

  let args = new URLSearchParams(location.search.slice(1));
  let sfHost = args.get("host");
  initButton(sfHost, true);
  sfConn.getSession(sfHost).then(() => {

    let root = document.getElementById("root");
    let model = new Model(sfHost, args);
    window.sfConn = sfConn;
    window.display = apiPromise => {
      if (model.spinnerCount > 0) {
        throw new Error("API call already in progress");
      }
      model.performRequest(Promise.resolve(apiPromise));
    };
    model.reactCallback = cb => {
      ReactDOM.render(h(App, {model}), root, cb);
    };
    ReactDOM.render(h(App, {model}), root);

  });

}
