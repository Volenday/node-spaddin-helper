"use strict";
const node_fetch_1 = require("node-fetch");
const URL = require("url");
class SharePointRestClient {
    constructor(url, authToken) {
        this.url = url;
        this.authToken = authToken;
    }
    // private getHeaders(requestDigest: string = null, contentLength: number = null) : any {
    getHeaders(xVerb = null, requestDigest = null) {
        let headers = {
            "Authorization": `Bearer ${this.authToken}`,
            "Accept": 'application/json; odata=verbose',
            "Content-Type": 'application/json; odata=verbose'
        };
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
    get(relativeUrl) {
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
    postRequest(verb, relativeUrl, data) {
        return new Promise((resolve, reject) => {
            return this.getContextInfo()
                .then(contextInfo => {
                let args = {
                    headers: this.getHeaders(verb, contextInfo.d.FormDigestValue),
                    method: 'POST'
                };
                if (data) {
                    args["body"] = data;
                }
                node_fetch_1.default(this.getFullUrl(relativeUrl), args)
                    .then(r => {
                    resolve(r.json());
                }).catch(error => {
                    console.log("[FETCH::ERROR] " + error);
                    reject(error);
                });
            });
        });
    }
    post(relativeUrl, data) {
        return this.postRequest(null, relativeUrl, data);
    }
    put(relativeUrl, data) {
        return this.postRequest('PUT', relativeUrl, data);
    }
    patch(relativeUrl, data) {
        return this.postRequest('PATCH', relativeUrl, data);
    }
    delete(relativeUrl) {
        return this.postRequest('DELETE', relativeUrl);
    }
}
SharePointRestClient.ContextInfoRelativeUrl = '_api/contextinfo';
exports.SharePointRestClient = SharePointRestClient;
//# sourceMappingURL=client.js.map