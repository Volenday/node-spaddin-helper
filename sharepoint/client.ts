import * as nodeFetch from "node-fetch";
import callNodeFetch from "node-fetch";
import * as URL from "url";


export class SharePointRestClient {

    private static ContextInfoRelativeUrl = '_api/contextinfo';

    public odataVerbose = true;

    constructor(private url: string, private authToken: string) {

    }

    // private getHeaders(requestDigest: string = null, contentLength: number = null) : any {
    private getHeaders(xVerb: string=null, requestDigest: string = null) : any {
        let headers = {
            "Authorization": `Bearer ${this.authToken}`
        };

        if (this.odataVerbose) {
            console.log("ODATA VERBOSE CONTENT TYPE");
            headers["Accept"] =  'application/json; odata=verbose';
            headers["Content-Type"] = 'application/json; odata=verbose';
        } else {
            console.log("JSON CONTENT TYPE");
            headers["Accept"] =  'application/json';
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

    private getFormDigestValue(contextInfo: any) : string {
        if (this.odataVerbose)
            return contextInfo.d.FormDigestValue;
        else
            return contextInfo.FormDigestValue;
    }

    private postRequest(verb: string, relativeUrl: string, data?: any) : Promise<any> {
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
                callNodeFetch(this.getFullUrl(relativeUrl), args)
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