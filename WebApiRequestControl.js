import {LitElement, css, html, unsafeHTML} from 'https://cdn.jsdelivr.net/gh/lit/dist@2/all/lit-all.min.js';

export class WebApiRequest extends LitElement {
    
  static properties = {
    pluginLoaded: { type: Boolean },
    message: { type: String },
    webApiUrl: { type: String },
    headers: { type: String },
    isIntegratedAuth: { type: Boolean },
    outcome: { type: String }
} 
      
  static getMetaConfig() {    
    return {
      groupName : "Custom",
      controlName: 'WebApi Request',
      description : 'Make Web Api request including OnPrem, SPO and return JSON Response',
      iconUrl : 'data-lookup',
      searchTerms : ['web', 'webapi'],
      fallbackDisableSubmit: false,
      version: '2.0',
      pluginAuthor : 'Preetha Ponnusamy',
      standardProperties: {
        fieldLabel:true,
        description:true,
        visibility: true        
      },
      properties: {
        webApiUrl: {
            type: 'string',
            title: 'WebApi Url',
            description: 'Provide Web api Url',
            required: true,
            defaultValue: ''
        },
        headers: {
            type: 'string',
            title: 'Request header',
            description: 'Provide headers as json object',
            defaultValue: '{ "Accept" : "application/json" }'
        },
        isIntegratedAuth: {
            type: 'boolean',
            title: 'Is Integrated Authentication',
            description: 'Check yes for Windows Integrated Auth',
            defaultValue: false
        },
        outcome: {
            type: 'string',
            title: 'Outcome',
            description: 'If set, the value will be overridden by api response',
            isValueField: true
        }
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
      width:100%;
      appearance: none;
      background-image: url(data:image/svg+xml;charset=US-ASCII,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%22292.4%22%20height%3D%22292.4%22%3E%3Cpath%20fill%3D%22%23131313%22%20d%3D%22M287%2069.4a17.6%2017.6%200%200%200-13-5.4H18.4c-5%200-9.3%201.8-12.9%205.4A17.6%2017.6%200%200%200%200%2082.2c0%205%201.8%209.3%205.4%2012.9l128%20127.9c3.6%203.6%207.8%205.4%2012.8%205.4s9.2-1.8%2012.8-5.4L287%2095c3.5-3.5%205.4-7.8%205.4-12.8%200-5-1.9-9.2-5.5-12.8z%22%2F%3E%3C%2Fsvg%3E);
      background-repeat: no-repeat;
      background-position: right 0.7rem top 50%;
      background-size: 0.65rem auto;
    }
    div.webapi-control{
      padding: 4px 0px 3px;
      color: #000;
    }
  `;

  constructor() {
    super()
    this.message = 'Loading...';
    this.webApi = '';
}

render() {
    return html`        
    <div>${this.message && typeof this.message === 'object' ? html`<pre>${JSON.stringify(this.message, null, 2)}</pre>` : this.message}</div>
`
}

_propagateOutcomeChanges(targetValue) {
    const args = {
        bubbles: true,
        cancelable: false,
        composed: true,
        detail: targetValue,
    };
    const event = new CustomEvent('ntx-value-change', args);
    this.dispatchEvent(event);
}
updated(changedProperties) {
    super.updated(changedProperties);
    if (changedProperties.has('webApiUrl')) {
        this.callApi();
    }
}
connectedCallback() {
    if (this.pluginLoaded) {
        return;
    }
    this.pluginLoaded = true;
    super.connectedCallback();
    if (window.location.pathname == "/") {
        this.message = html`Please configure control`
        return;
    }

    if (!this.headers) {
        this.headers = '{ "Accept" : "application/json" }'
    }
    if (this.webApiUrl) {
        if (this.isValidJSON(this.headers)) {
            this.callApi();
        }
        else {
            this.message = html`Invalid Headers`
        }
    }
    else {
        this.message = html`Invalid WebApi Url`
    }
}

async callApi() {
    var inputWebApi = this.webApiUrl;
    if (inputWebApi.indexOf("/_api/web/") == -1 && inputWebApi.indexOf("/_api/site/") == -1) {
        await this.loadWebApi();
    }
    else {
        var hostWebUrl = this.queryParam("SPHostUrl");
        var appWebUrl = this.queryParam("SPAppWebUrl");
        var spoApiUrl = appWebUrl + inputWebApi.replace(hostWebUrl, "").replace("/_api/", "/_api/SP.AppContextSite(@target)/")
        if (inputWebApi.indexOf("?") == -1) {
            spoApiUrl = spoApiUrl + "?@target='" + hostWebUrl + "'";
        }
        else {
            spoApiUrl = spoApiUrl + "&@target='" + hostWebUrl + "'";
        }
        await this.loadSPOApi(appWebUrl, spoApiUrl);
    }

}

async executeAsyncWithPromise(appWebUrl, requestInfo) {
    return new Promise((resolve, reject) => {
        const executor = new SP.RequestExecutor(appWebUrl);
        executor.executeAsync({
            ...requestInfo,
            success: (response) => resolve(response),
            error: (response) => reject(response),
        });
    });
}

async loadSPOApi(appWebUrl, spoApiUrl) {
    const requestInfo = {
        url: spoApiUrl,
        method: "GET",
        headers: { "Accept": "application/json; odata=verbose" }
    };

    var response;
    try {
        response = await this.executeAsyncWithPromise(appWebUrl, requestInfo);
    }
    catch (e) {
        response = {}
        response.status = "500"
        response.statusText = e + ", Try checking end point";
    }

    if (response.body != undefined && response.statusCode == 200) {
        try {
            var jsonData = JSON.parse(response.body);
        }
        catch (e) {
            this.message = html`Invalid JSON response`
        }
        this.message=jsonData;
        this._propagateOutcomeChanges(JSON.stringify(jsonData));
    }
    else {
        this.message = html`WebApi request failed: ${response.status} - ${response.statusText == '' ? 'Error!' : response.statusText}`
    }
}

async loadWebApi() {
    var headers = { 'accept': 'application/json' }
    var fetchAttributes = { "headers": headers };
    if (this.isIntegratedAuth) {
        fetchAttributes = { "headers": headers, "credentials": "include" }
    }

    var response;
    try {
        response = await fetch(`${this.webApiUrl}`, fetchAttributes);
    }
    catch (e) {
        response = {}
        response.status = "500"
        response.statusText = e + ", Try checking authentication";
    }

    if (response != undefined && response.status == 200) {
        try {
            var jsonData = await response.json();
            // jsonData = this.filterJson(jsonData);        
        }
        catch (e) {
            this.message = html`Invalid JSON response`
        }
        this.message=jsonData;
        this._propagateOutcomeChanges(JSON.stringify(jsonData));
    }
    else {
        this.message = html`WebApi request failed: ${response.status} - ${response.statusText == '' ? 'Error!' : response.statusText}`
    }

}

 isValidJSON(str) {
    try {
        JSON.parse(str);
        return true;
    } catch (e) {
        return false;
    }
}

queryParam(param) {
    const urlParams = new URLSearchParams(decodeURIComponent(window.location.search.replaceAll("amp;", "")));
    return urlParams.get(param);
}

}

customElements.define('webapi-request', WebApiRequest);