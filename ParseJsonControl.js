import {
  LitElement,
  css,
  html,
  unsafeHTML,
} from "https://cdn.jsdelivr.net/gh/lit/dist@2/all/lit-all.min.js";
import { JSONPath } from "https://cdn.jsdelivr.net/npm/jsonpath-plus@10.1.0/dist/index-browser-esm.min.js";
import Mustache from "https://cdnjs.cloudflare.com/ajax/libs/mustache.js/4.2.0/mustache.min.js";

export class ParseJson extends LitElement {
  static properties = {
    pluginLoaded: { type: Boolean },
    message: { type: String },
    jsonResponse: { type: String },
    jsonPath: { type: String },
    displayAs: { type: String },
    mustacheTemplate: { type: String },
    currentPageMode: { type: String },
    outcome: { type: String },
    sortOrder: { type: String },
    defaultMessage: { type: String },
  };

  static getMetaConfig() {
    return {
      groupName: "Custom",
      controlName: "Parse JSON",
      description: "Get the properties from JSON Data",
      iconUrl: "data-lookup",
      searchTerms: ["parse", "json"],
      fallbackDisableSubmit: false,
      version: "2.0",
      pluginAuthor: "Preetha Ponnusamy",
      standardProperties: {
        fieldLabel: true,
        description: true,
        visibility: true,
      },
      properties: {
        jsonResponse: {
          type: "string",
          title: "JSON Data",
          description: "Provide JSON Data from api response",
          defaultValue: "",
        },
        jsonPath: {
          type: "string",
          title: "JSON Path",
          description: "Provide JSON Path to filter out data",
          defaultValue: "$.",
        },
        displayAs: {
          type: "string",
          title: "Display As",
          enum: ["Label", "Dropdown", "Label using Mustache Template"],
          description: "Provide display type of the control",
          defaultValue: "Label",
        },
        mustacheTemplate: {
          type: "string",
          title: "Mustache Template",
          description:
            "Provide Mustache template (applicable for selected display type)",
          defaultValue: "",
        },
        defaultMessage: {
          type: "string",
          title: "Default Option for Dropdown",
          description:
            "Provide default selected text (applicable when Display As is set to Dropdown)",
          defaultValue: "Please select an option",
        },
        sortOrder: {
          type: "string",
          title: "Sort Order",
          description: "Sorting order for selected display type",
          enum: ["As Is", "Asc", "Desc"],
          defaultValue: "",
        },
        outcome: {
          type: "string",
          title: "Outcome",
          description: "If set, the value will be overridden by api response",
          isValueField: true,
          standardProperties: {
            visibility: false,
          },
        },
      },
      events: ["ntx-value-change"],
    };
  }

  static styles = css`
    select.webapi-control {
      border-radius: var(--ntx-form-theme-border-radius);
      font-size: var(--ntx-form-theme-text-input-size);
      caret-color: var(--ntx-form-theme-color-input-text);
      color: var(--ntx-form-theme-color-input-text);
      border-color: var(--ntx-form-theme-color-border);
      font-family: var(--ntx-form-theme-font-family);
      background-color: var(--ntx-form-theme-color-input-background);
      line-height: var(--ntx-form-control-line-height, 1.25);
      min-height: 33px;
      height: auto;
      padding: 0.55rem;
      border: 1px solid #898f94;
      min-width: 70px;
      position: relative;
      display: block;
      box-sizing: border-box;
      width: 100%;
      appearance: none;
      background-image: url(data:image/svg+xml;charset=US-ASCII,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%22292.4%22%20height%3D%22292.4%22%3E%3Cpath%20fill%3D%22%23131313%22%20d%3D%22M287%2069.4a17.6%2017.6%200%200%200-13-5.4H18.4c-5%200-9.3%201.8-12.9%205.4A17.6%2017.6%200%200%200%200%2082.2c0%205%201.8%209.3%205.4%2012.9l128%20127.9c3.6%203.6%207.8%205.4%2012.8%205.4s9.2-1.8%2012.8-5.4L287%2095c3.5-3.5%205.4-7.8%205.4-12.8%200-5-1.9-9.2-5.5-12.8z%22%2F%3E%3C%2Fsvg%3E);
      background-repeat: no-repeat;
      background-position: right 0.7rem top 50%;
      background-size: 0.65rem auto;
    }
    div.webapi-control {
      padding: 4px 0px 3px;
      color: #000;
    }
  `;

  constructor() {
    super();
    this.message = "Loading...";
    this.webApi = "";
  }

  render() {
    return html` <div>${this.message}</div> `;
  }

  _propagateOutcomeChanges(targetValue) {
    const args = {
      bubbles: true,
      cancelable: false,
      composed: true,
      detail: targetValue,
    };
    const event = new CustomEvent("ntx-value-change", args);
    this.dispatchEvent(event);
  }
  updated(changedProperties) {
    super.updated(changedProperties);
    if (changedProperties.has("jsonResponse")) {
      this.getProperty();
    }
  }
  connectedCallback() {
    if (this.pluginLoaded) {
      return;
    }
    this.pluginLoaded = true;
    super.connectedCallback();
    var currentPageModeIndex = this.queryParam("mode");
    this.currentPageMode =
      currentPageModeIndex == 0
        ? "New"
        : currentPageModeIndex == 1
        ? "Edit"
        : "Display";
    if (event.target.innerText == "Edit") this.currentPageMode = "Edit";
    if (!this.jsonResponse) {
      this.message = html`Please provide valid jsonResponse`;
    }
    if (this.jsonResponse && this.jsonPath && this.displayAs) {
      this.getProperty();
    } else {
      this.message = html`Please configure control`;
      return;
    }
  }

  getProperty() {
    var jsonData = this.filterJson(JSON.parse(this.jsonResponse));
    this.plugToForm(jsonData);
  }

  plugToForm(jsonData) {
    if (this.displayAs == "Label") {
      this.constructLabelTemplate(jsonData);
    } else if (this.displayAs == "Dropdown") {
      this.constructDropdownTemplate(jsonData);
    } else if (this.displayAs == "Label using Mustache Template") {
      this.constructLabelUsingMustacheTemplate(jsonData);
    }
    this._propagateOutcomeChanges(this.outcome);
  }

  constructLabelTemplate(jsonData) {
    var outputTemplate = "";
    var htmlTemplate = html``;

    if (typeof jsonData === "string" || jsonData instanceof String) {
      outputTemplate = jsonData;
    }
    if (this.isInt(jsonData)) {
      outputTemplate = jsonData.toString();
    }
    if (typeof jsonData == "boolean") {
      outputTemplate = jsonData == true ? "true" : "false";
    }
    htmlTemplate = html`<div class="form-control webapi-control">
      ${outputTemplate}
    </div>`;

    this.outcome = outputTemplate;
    this.message = html`${htmlTemplate}`;
  }

  constructDropdownTemplate(items) {
    if (typeof items === "string") {
      items = [items];
    }

    if (Array.isArray(items)) {
      if (this.sortOrder === "Asc") {
        items.sort((a, b) => (a > b ? 1 : -1));
      } else if (this.sortOrder === "Desc") {
        items.sort((a, b) => (a < b ? 1 : -1));
      }
    }
    if (this.currentPageMode == "New" || this.currentPageMode == "Edit") {
      if (Array.isArray(items)) {
        var itemTemplates = [];
        itemTemplates.push(
          html`<option value="" disabled selected>
            ${this.defaultMessage || "Select an option"}
          </option>`
        );
        for (var i of items) {
          if (this.currentPageMode == "Edit" && i == this.outcome) {
            itemTemplates.push(html`<option selected>${i}</option>`);
          } else {
            itemTemplates.push(html`<option>${i}</option>`);
          }
        }

        this.message = html`<select
          class="form-control webapi-control"
          @change=${(e) => this._propagateOutcomeChanges(e.target.value)}
        >
          ${itemTemplates}
        </select> `;
      } else {
        this.message = html`<p>
          WebApi response not in array. Check WebApi Configuration
        </p>`;
      }
    } else {
      this.constructLabelTemplate(this.outcome);
    }
  }

  constructLabelUsingMustacheTemplate(jsonData) {
    var rawValue = "";
    var htmlTemplate = html``;

    if (typeof jsonData === "string" || jsonData instanceof String) {
      rawValue = jsonData;
    }
    if (this.isInt(jsonData)) {
      rawValue = jsonData.toString();
    }
    if (typeof jsonData == "boolean") {
      rawValue = jsonData == true ? "true" : "false";
    }
    if (Array.isArray(jsonData)) {
      rawValue = jsonData;
    }

    var outputTemplate = Mustache.render(this.mustacheTemplate, rawValue);

    htmlTemplate = html`<div class="form-control webapi-control">
      ${unsafeHTML(outputTemplate)}
    </div>`;

    this.outcome = rawValue;
    this.message = html`${htmlTemplate}`;
  }

  isInt(value) {
    return (
      !isNaN(value) &&
      (function (x) {
        return (x | 0) === x;
      })(parseFloat(value))
    );
  }

  filterJson(jsonData) {
    if (!this.jsonPath) {
      this.jsonPath = "$.";
    }
    if (jsonData) {
      var result = JSONPath({ path: this.jsonPath, json: jsonData });
      if (result.length == 1 && this.jsonPath.endsWith(".")) {
        result = result[0];
      }
      return result;
    }
  }

  queryParam(param) {
    const urlParams = new URLSearchParams(
      decodeURIComponent(window.location.search.replaceAll("amp;", ""))
    );
    return urlParams.get(param);
  }
}

customElements.define("parse-json", ParseJson);
