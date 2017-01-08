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
    post(relativeUrl, data) {
        return new Promise((resolve, reject) => {
            node_fetch_1.default(this.getFullUrl(relativeUrl), {
                headers: this.getHeaders(),
                body: data,
                method: 'POST'
            }).then(r => {
                resolve(r.json());
            }).catch(error => {
                console.log("[FETCH::ERROR] " + error);
                reject(error);
            });
        });
    }
    put(relativeUrl, data) {
        return new Promise((resolve, reject) => {
            node_fetch_1.default(this.getFullUrl(relativeUrl), {
                headers: this.getHeaders(),
                body: data,
                method: 'PUT'
            }).then(r => {
                resolve(r.json());
            }).catch(error => {
                console.log("[FETCH::ERROR] " + error);
                reject(error);
            });
        });
    }
    patch(relativeUrl, data) {
        return new Promise((resolve, reject) => {
            node_fetch_1.default(this.getFullUrl(relativeUrl), {
                headers: this.getHeaders(),
                body: data,
                method: 'PATCH'
            }).then(r => {
                resolve(r.json());
            }).catch(error => {
                console.log("[FETCH::ERROR] " + error);
                reject(error);
            });
        });
    }
    delete(relativeUrl) {
        return new Promise((resolve, reject) => {
            node_fetch_1.default(this.getFullUrl(relativeUrl), {
                headers: this.getHeaders(),
                method: 'DELETE'
            }).then(r => {
                resolve(r.json());
            }).catch(error => {
                console.log("[FETCH::ERROR] " + error);
                reject(error);
            });
        });
    }
}
exports.SharePointRestClient = SharePointRestClient;
//# sourceMappingURL=client.js.map