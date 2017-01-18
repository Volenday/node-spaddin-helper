import * as common from "./common";
import {TokenHelper} from  "./token-helper";
import {SharePointRestClient} from "./client";
import {Url} from "../helpers/url";

    export interface ISharePointContextInfo {
        SPHostUrl: string, 
        SPAppWebUrl: string, 
        SPLanguage: string, 
        SPClientTag: string, 
        SPProductNumber: string,
        contextTokenStr?: string,
        contextToken?: any
    }

    interface ITupleTokenExpiration { token: string,  expired: Date };
    interface IContextCacheHandler {
        save(req: any, contextInfo: ISharePointContextInfo) : void;
        load(req: any) : ISharePointContextInfo;
    }

    export class SharePointContext {
        public static SPHostUrlKey = "SPHostUrl";
        public static SPAppWebUrlKey = "SPAppWebUrl";
        public static SPLanguageKey = "SPLanguage";
        public static SPClientTagKey = "SPClientTag";
        public static SPProductNumberKey = "SPProductNumber";
        public static ContextCookieName = "SpContextParameters";

        protected static AccessTokenLifetimeToleranceInMilliSeconds = 5 * 60 * 1000; // 5 Minutes

        protected userAccessTokenForSPHost :ITupleTokenExpiration; 
        protected appOnlyAccessTokenForSPHost:ITupleTokenExpiration;

        public static ContextCacheHandler: IContextCacheHandler;

        private cacheKey: string = null;

        public static getSPHostUrl(request: common.ISPHttpRequest) : string {
            if (!request) throw new Error("httpRequest is undefined or null");

            let urlWithEnsuredSlash = Url.ensureTrailingSlash(request.query.SPHostUrl);
            if (!urlWithEnsuredSlash) {
                let sphostUrlFromCookie = request.cookies.SpContextParameters && Url.parseQueryString(request.cookies.SpContextParameters).SPHostUrl
                urlWithEnsuredSlash = Url.ensureTrailingSlash(sphostUrlFromCookie);
            }

            // Check if well formed HTTP URL
            if (urlWithEnsuredSlash){
                if (urlWithEnsuredSlash.indexOf("http://") == 0 || urlWithEnsuredSlash.indexOf("https://") == 0) {
                    console.log("SPHostUrl = " + urlWithEnsuredSlash);
                    return urlWithEnsuredSlash;
                }
            }
                
            return null;
        }

        private static loadFromRequest(req) : SharePointContext {
            // If no context cache mechanism is specified, return nothing
            if (!SharePointContext.ContextCacheHandler)
                return null;

            let info: ISharePointContextInfo = SharePointContext.ContextCacheHandler.load(req);
            return info && SharePointContext.createFromContextInfo(info);
        }

        private static validateContext(context: SharePointContext, req) : boolean {
            // TODO Implement this
            // Compare current request and context
            let spHostUrl = SharePointContext.getSPHostUrl(req);
            let contextTokenStr = TokenHelper.getContextTokenFromRequest(req);
            let spCacheKey = req.cookies.SPCacheKey;
        
            return spHostUrl == context.SPHostUrl 
            && (!spCacheKey || spCacheKey == context.cacheKey)
            && context.contextToken && (!contextTokenStr || contextTokenStr == context.contextTokenStr); 
        }

        private static save(context: SharePointContext, req) : void {
            // req.session.SPContext = context;
            // If no context cache mechanism is specified, don't do anything
            if (!SharePointContext.ContextCacheHandler)
                return;

            return SharePointContext.ContextCacheHandler.save(req, context.getContextInfo());
        }

        public static getFromRequest(req) : SharePointContext {
            if (!req) throw new Error("The HTTP request cannot be found");

            console.log("Get context from request");
            let spHostUrl = SharePointContext.getSPHostUrl(req);
            if (!spHostUrl) return null;

            console.log("Load context from server according to request");
            let spContext = SharePointContext.loadFromRequest(req);

            if (!spContext || !SharePointContext.validateContext(spContext, req))
            {
                console.log("Context not loaded or invalid");
                spContext = SharePointContext.createFromRequest(req);
                if (spContext)
                {
                    console.log("Context created");
                    SharePointContext.save(spContext, req);
                }
                else {
                    console.log("Context not created");
                }
            }

            return spContext;
        }

        constructor(private SPHostUrl: string, 
                    private SPAppWebUrl: string, 
                    private SPLanguage: string, 
                    private SPClientTag: string, 
                    private SPProductNumber: string,
                    private contextTokenStr?: string,
                    private contextToken?: any) {

            if (!SPHostUrl) throw new Error("SPHostUrl is required.");
            if (!SPProductNumber) throw new Error("SPProductNumber is required.");
            if (!SPLanguage) throw new Error("SPLanguage is required.");
            if (!SPClientTag) throw new Error("SPCLientTag is required.");
        }

        public getContextInfo() : ISharePointContextInfo {
            return {
                SPHostUrl: this.SPHostUrl,
                SPAppWebUrl: this.SPAppWebUrl,
                SPLanguage: this.SPLanguage,
                SPClientTag: this.SPClientTag,
                SPProductNumber: this.SPProductNumber,
                contextToken: this.contextToken,
                contextTokenStr: this.contextTokenStr
            };
        }

        public static createFromContextInfo(info: ISharePointContextInfo) : SharePointContext {
            return new SharePointContext(info.SPHostUrl, 
                info.SPAppWebUrl, 
                info.SPLanguage, 
                info.SPClientTag, 
                info.SPProductNumber, 
                info.contextTokenStr, 
                info.contextToken);
        }

        public static createFromRequest(req) : SharePointContext {
            if (!req) throw new Error("Request is not specified");

            // SPHostUrl
            let spHostUrl = SharePointContext.getSPHostUrl(req);
            if (!spHostUrl) return null;
            
            var query: common.ISPQueryStringArgs = req.query;

            // SPAppWebUrl
            let spAppWebUrl = Url.ensureTrailingSlash(query.SPAppWebUrl);
            if (!Url.validateHttpSchemes(spAppWebUrl, ['http', 'https']))
                spAppWebUrl = null;
                
            if (!query.SPLanguage) return null;
            if (!query.SPClientTag) return null;
            if (!query.SPProductNumber) return null;

            return SharePointContext.create(spHostUrl, query.SPAppWebUrl, query.SPLanguage, query.SPClientTag, query.SPProductNumber, req);
        }

        private static create(spHostUrl: string, 
                            spAppWebUrl: string, 
                            spLanguage: string, 
                            spClientTag: string, 
                            spProductNumber: string, 
                            request) : SharePointContext {
            let contextTokenStr = TokenHelper.getContextTokenFromRequest(request);
            if (!contextTokenStr)
                return null;

            try {
                var contextTokenObj = TokenHelper.readAndValidateContext(contextTokenStr, request.hostname);
                return new SharePointContext(spHostUrl,spAppWebUrl,spLanguage,spClientTag,spProductNumber, contextTokenStr, contextTokenObj);
            } catch (error) {
                return null;
            }
        }

        private static createRESTClient(spSiteUrl: string, accessToken: string) : SharePointRestClient {
            if (spSiteUrl && accessToken)
                return new SharePointRestClient(spSiteUrl, accessToken);

            return null;
        }

        public createClientForSPHost() : Promise<SharePointRestClient> {
            // If the token is already in cache and stil valid
            if (this.userAccessTokenForSPHost 
            && this.userAccessTokenForSPHost.token
            && this.userAccessTokenForSPHost.expired < new Date()) {
                let promise = new Promise<SharePointRestClient>(resolve => {
                    let client = SharePointContext.createRESTClient(this.SPHostUrl, this.userAccessTokenForSPHost.token);
                    resolve(client);
                });
            }

            return TokenHelper.getAccessToken(this.contextToken, this.SPHostUrl).then(token => {
                this.userAccessTokenForSPHost = {expired:new Date(Date.parse(token.expires_on)), token: token.access_token};
                return SharePointContext.createRESTClient(this.SPHostUrl, token.access_token);
            });   
        }

        public createAppOnlyClientForSPHost() : Promise<SharePointRestClient> {
            // If the token is already in cache and stil valid
            if (this.appOnlyAccessTokenForSPHost 
            && this.appOnlyAccessTokenForSPHost.token
            && this.appOnlyAccessTokenForSPHost.expired < new Date()) {
                let promise = new Promise<SharePointRestClient>(resolve => {
                    let client = SharePointContext.createRESTClient(this.SPHostUrl, this.appOnlyAccessTokenForSPHost.token);
                    resolve(client);
                });
            }

            return TokenHelper.getAccessToken(this.contextToken, this.SPHostUrl, true).then(token => {
                this.appOnlyAccessTokenForSPHost = {expired:new Date(Date.parse(token.expires_on)), token: token.access_token};
                return SharePointContext.createRESTClient(this.SPHostUrl, token.access_token);
            });
        }

        
        
    }