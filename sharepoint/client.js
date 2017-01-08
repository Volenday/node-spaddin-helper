"use strict";
const node_fetch_1 = require("node-fetch");
const URL = require("url");
class SharePointRestClient {
    constructor(url, authToken) {
        this.url = url;
        this.authToken = authToken;
    }
    getHeaders() {
        return {
            "Authorization": `Bearer ${this.authToken}`,
            // "Accept": 'application/json;odata=verbose',
            // "Content-Type": 'application/json;odata=verbose'
            "Accept": 'application/json',
            "Content-Type": 'application/json'
        };
    }
    getFullUrl(urlPart) {
        return URL.resolve(this.url, urlPart);
    }
    get(relativeUrl) {
        return node_fetch_1.default(this.getFullUrl(relativeUrl), {
            headers: this.getHeaders()
        }).then(r => {
            return r.json();
        }).catch(error => {
            console.log("[FETCH::ERROR] " + error);
        });
    }
    post(relativeUrl, data) {
        return node_fetch_1.default(this.getFullUrl(relativeUrl), {
            headers: this.getHeaders(),
            body: data,
            method: 'POST'
        })
            .then(r => {
            return r.json();
        }).catch(error => {
            console.log("[FETCH::ERROR] " + error);
        });
    }
    put(relativeUrl, data) {
        return node_fetch_1.default(this.getFullUrl(relativeUrl), {
            headers: this.getHeaders(),
            body: data,
            method: 'PUT'
        })
            .then(r => {
            return r.json();
        }).catch(error => {
            console.log("[FETCH::ERROR] " + error);
        });
    }
    patch(relativeUrl, data) {
        return node_fetch_1.default(this.getFullUrl(relativeUrl), {
            headers: this.getHeaders(),
            body: data,
            method: 'PATCH'
        })
            .then(r => {
            return r.json();
        }).catch(error => {
            console.log("[FETCH::ERROR] " + error);
        });
    }
    delete(relativeUrl) {
        return node_fetch_1.default(this.getFullUrl(relativeUrl), {
            headers: this.getHeaders(),
            method: 'DELETE'
        })
            .then(r => {
            return r.json();
        }).catch(error => {
            console.log("[FETCH::ERROR] " + error);
        });
    }
}
exports.SharePointRestClient = SharePointRestClient;
//# sourceMappingURL=client.js.map