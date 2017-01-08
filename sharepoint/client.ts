import * as nodeFetch from "node-fetch";
import callNodeFetch from "node-fetch";
import * as URL from "url";


export class SharePointRestClient {

    private static ContextInfoRelativeUrl = '_api/contextinfo';

    constructor(private url: string, private authToken: string) {

    }

    // private getHeaders(requestDigest: string = null, contentLength: number = null) : any {
    private getHeaders(xVerb: string=null, requestDigest: string = null) : any {
        let headers = {
            "Authorization": `Bearer ${this.authToken}`,
            "Accept": 'application/json; odata=verbose',
            "Content-Type": 'application/json; odata=verbose'
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

    private getFullUrl(urlPart: string) : string {
        return URL.resolve(this.url, urlPart);
    }


    public retrieve(relativeUrl: string) : Promise<any> {
        return new Promise((resolve, reject) => {
                callNodeFetch(this.getFullUrl(relativeUrl), {
                headers: this.getHeaders()
            }).then(r => {
                resolve(r.json());
            }).catch(error => {
                console.log("[FETCH::ERROR] " + error);
                reject(error);
            });
        });
    }

    private getContextInfo() : Promise<any> {
        return new Promise((resolve, reject) => {
                callNodeFetch(this.getFullUrl(SharePointRestClient.ContextInfoRelativeUrl), {
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

    private postRequest(verb: string, relativeUrl: string, data?: any) : Promise<any> {
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
                callNodeFetch(this.getFullUrl(relativeUrl), args)
                .then(r => {
                    resolve(r.json());
                }).catch(error => {
                    console.log("[FETCH::ERROR] " + error);
                    reject(error);
                });
            }); 
        });
    }

    public create(relativeUrl: string, data: any) : Promise<any> {
         return this.postRequest(null, relativeUrl, data);
    }

    public update(relativeUrl: string, data: any) : Promise<any> {
       return this.postRequest('MERGE', relativeUrl, data);
    }

    public delete(relativeUrl: string) : Promise<any> {
        return this.postRequest('DELETE', relativeUrl);
    }
}