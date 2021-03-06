"use strict";
const node_fetch_1 = require("node-fetch");
const URL = require("url");
class SharePointRestClient {
    constructor(url, authToken) {
        this.url = url;
        this.authToken = authToken;
        this.odataVerbose = true;
    }
    // private getHeaders(requestDigest: string = null, contentLength: number = null) : any {
    getHeaders(xVerb = null, requestDigest = null) {
        let headers = {
            "Authorization": `Bearer ${this.authToken}`
        };
        if (this.odataVerbose) {
            console.log("ODATA VERBOSE CONTENT TYPE");
            headers["Accept"] = 'application/json; odata=verbose';
            headers["Content-Type"] = 'application/json; odata=verbose';
        }
        else {
            console.log("JSON CONTENT TYPE");
            headers["Accept"] = 'application/json';
            headers["Content-Type"] = 'application/json';
        }
        // if (contentLength) {
        //     headers["Content-Length"] = contentLength;
        // }
        if (requestDigest) {
            headers["X-RequestDigest"] = requestDigest;
        }
        if (xVerb) {
            headers["X-HTTP-Method"] = xVerb;
            headers["If-Match"] = "*";
        }
        return headers;
    }
    getFullUrl(urlPart) {
        return URL.resolve(this.url, urlPart);
    }
    retrieve(relativeUrl) {
        return new Promise((resolve, reject) => {
            node_fetch_1.default(this.getFullUrl(relativeUrl), {
                headers: this.getHeaders()
            }).then(r => {
                resolve(r.json());
            }).catch(error => {
                console.log("[FETCH::ERROR] " + error);
                reject(error);
            });
        });
    }
    getContextInfo() {
        return new Promise((resolve, reject) => {
            node_fetch_1.default(this.getFullUrl(SharePointRestClient.ContextInfoRelativeUrl), {
                headers: this.getHeaders(),
                method: 'POST'
            }).then(r => {
                resolve(r.json());
            }).catch(error => {
                console.log("[FETCH::ERROR] " + error);
                reject(error);
            });
        });
    }
    getFormDigestValue(contextInfo) {
        if (this.odataVerbose)
            return contextInfo.d.FormDigestValue;
        else
            return contextInfo.FormDigestValue;
    }
    postRequest(verb, relativeUrl, data) {
        return new Promise((resolve, reject) => {
            return this.getContextInfo()
                .then(contextInfo => {
                let args = {
                    headers: this.getHeaders(verb, this.getFormDigestValue(contextInfo)),
                    method: 'POST'
                };
                if (data) {
                    args["body"] = typeof data !== "string" ? JSON.stringify(data) : data;
                }
                node_fetch_1.default(this.getFullUrl(relativeUrl), args)
                    .then(r => {
                    let jsonResult = r.json();
                    console.log("[FETCH SUCCESS] " + JSON.stringify(jsonResult));
                    resolve(jsonResult);
                }).catch(error => {
                    console.log("[FETCH::ERROR] " + error);
                    reject(error);
                });
            });
        });
    }
    create(relativeUrl, data) {
        return this.postRequest(null, relativeUrl, data);
    }
    update(relativeUrl, data) {
        return this.postRequest('MERGE', relativeUrl, data);
    }
    delete(relativeUrl) {
        return this.postRequest('DELETE', relativeUrl);
    }
}
SharePointRestClient.ContextInfoRelativeUrl = '_api/contextinfo';
exports.SharePointRestClient = SharePointRestClient;
//# sourceMappingURL=client.js.map