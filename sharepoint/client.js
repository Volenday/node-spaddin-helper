"use strict";
const node_fetch_1 = require("node-fetch");
const URL = require("url");
class SharePointRestClient {
    constructor(url, authToken) {
        this.url = url;
        this.authToken = authToken;
    }
    // private getHeaders(requestDigest: string = null, contentLength: number = null) : any {
    getHeaders(requestDigest = null) {
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
    issueWriteRequest(verb, relativeUrl, data) {
        return new Promise((resolve, reject) => {
            return this.getContextInfo()
                .then(contextInfo => {
                let args = {
                    headers: this.getHeaders(contextInfo.d.FormDigestValue),
                    method: verb
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
        return this.issueWriteRequest('POST', relativeUrl, data);
    }
    put(relativeUrl, data) {
        return this.issueWriteRequest('PUT', relativeUrl, data);
    }
    patch(relativeUrl, data) {
        return this.issueWriteRequest('PATCH', relativeUrl, data);
    }
    delete(relativeUrl) {
        return this.issueWriteRequest('DELETE', relativeUrl);
    }
}
SharePointRestClient.ContextInfoRelativeUrl = '_api/contextinfo';
exports.SharePointRestClient = SharePointRestClient;
//# sourceMappingURL=client.js.map