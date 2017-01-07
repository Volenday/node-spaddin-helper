import * as nodeFetch from "node-fetch";
import callNodeFetch from "node-fetch";
import * as URL from "url";


export class SharePointRestClient {

    constructor(private url: string, private authToken: string) {

    }

    private getHeaders() : any {
        return {
            "Authorization": `Bearer ${this.authToken}`,
            // "Accept": 'application/json;odata=verbose',
            // "Content-Type": 'application/json;odata=verbose'
            "Accept": 'application/json',
            "Content-Type": 'application/json'
        }
    }

    private getFullUrl(urlPart: string) : string {
        return URL.resolve(this.url, urlPart);
    }


    public get(relativeUrl: string) : Promise<any> {
        return callNodeFetch(this.getFullUrl(relativeUrl), {
            headers: this.getHeaders()
        }).then(r => {
            return r.json();
        });
    }

    public post(relativeUrl: string, data: any) : Promise<any> {
        return callNodeFetch(this.getFullUrl(relativeUrl), {
            headers: this.getHeaders(),
            body:data,
            method: 'POST'
        }).then(r => r.json());
    }

    public put(relativeUrl: string, data: any) : Promise<any> {
        return callNodeFetch(this.getFullUrl(relativeUrl), {
            headers: this.getHeaders(),
            body:data,
            method: 'PUT'
        }).then(r => r.json());
    }

    public patch(relativeUrl: string, data: any) : Promise<any> {
        return callNodeFetch(this.getFullUrl(relativeUrl), {
            headers: this.getHeaders(),
            body:data,
            method: 'PATCH'
        }).then(r => r.json());
    }

    public delete(relativeUrl: string) : Promise<any> {
        return callNodeFetch(this.getFullUrl(relativeUrl), {
            headers: this.getHeaders(),
            method: 'DELETE'
        }).then(r => r.json());
    }
}