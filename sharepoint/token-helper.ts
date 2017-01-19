
import * as nodeFetch from "node-fetch";
import callNodeFetch from "node-fetch";
import * as JWT from "jsonwebtoken";

import { Url } from "../helpers/url";
import { SharePointAddinConfiguration } from "./configuration";
import { IAuthToken } from './common';


export class TokenHelper {

    // Constants
    static SharePointServicePrincipal: string = "00000003-0000-0ff1-ce00-000000000000";
    static authorizationPage = "_layouts/15/OAuthAuthorize.aspx"                         ;
    static redirectPage = "_layouts/15/AppRedirect.aspx"                                 ;
    static acsPrincipalName = "00000001-0000-0000-c000-000000000000"                     ;
    static acsMetadataEndPointRelativeUrl = "metadata/json/1"                            ;
    static S2SProtocol = "OAuth2"                                                       ;
    static delegationIssuance = "DelegationIssuance1.0"                                  ;
    static nameIdentifierClaimType = "nameid"; 
    static trustedForImpersonationClaimType = "trustedfordelegation"                       ;
    static actorTokenClaimType = "actortoken"           ;

    

    public static getAppContextTokenRequestUrl(contextUrl: string, redirectUri: string) : string
    {
        const config = SharePointAddinConfiguration.getInstance();
        return `${Url.ensureTrailingSlash(contextUrl)}${TokenHelper.redirectPage}?client_id=${config.clientId}&redirect_uri=${redirectUri}`;
    }

    public static getContextTokenFromRequest(request) : string {
        let paramNames = [ "AppContext", "AppContextToken", "AccessToken", "SPAppToken" ];
        let token = null;
        paramNames.forEach(p => {
            if (request.param[p]) {
                token = request.param[p];
            }
            if (request.body[p]) {
                token = request.body[p];
            }

        });
        
        return token;
    }

    public static readAndValidateContext(contextTokenStr: string, appHostName: string) : any {
        const config = SharePointAddinConfiguration.getInstance();
        let token = JWT.decode(contextTokenStr, config.clientSecret);
        token.appctx = JSON.parse(token.appctx);
        // TODO Implement the validation of the token here
        // TODO Return type should be a typed interface
        return token;
    }

public static getAccessToken(contextToken: any, siteUrl: string, appOnly:boolean = false): Promise<IAuthToken> {

        return new Promise<IAuthToken>((resolve, reject) => {
            const config = SharePointAddinConfiguration.getInstance();
            let realm = null;
            let principals = null;
            return this.getRealm(siteUrl)
                .then(retrievedRealm => {
                    realm = retrievedRealm;
                    principals = {
                        resource: TokenHelper.getFormattedPrincipal(TokenHelper.SharePointServicePrincipal, Url.parse(siteUrl).hostname, realm),
                        formattedClientId: TokenHelper.getFormattedPrincipal(config.clientId, "", realm)
                    };
            })
            .then(() => TokenHelper.getAuthUrl(realm))
            .then((authUrl:string) => {
                let body = [];
    
                if (!appOnly) {
                    body.push("grant_type=refresh_token");
                    body.push(`refresh_token=${contextToken.refreshtoken}`);
                } else {
                    body.push("grant_type=client_credentials");
                }
                body.push(`client_id=${principals.formattedClientId}`);
                body.push(`client_secret=${encodeURIComponent(config.clientSecret)}`);
                body.push(`resource=${principals.resource}`);
    
                return callNodeFetch(authUrl, {
                    body: body.join("&"),
                    headers: {
                        "Content-Type": "application/x-www-form-urlencoded",
                    },
                    method: "POST",
                });
            })
            .then((r: nodeFetch.Response) => {
                
                let token = r.json();
                console.log("getAccessToken : token = " + JSON.stringify(token));
                if (token) {
                    resolve(token);
                }
                else {
                    reject("Token cannot be retrieved");
                }
            })
            .catch(error => {
                reject(error);
            });
        });
    }

    private static _realm: string = null;
    private static getRealm(siteUrl: string): Promise<string> {

        return new Promise(resolve => {

            if (TokenHelper._realm) {
                resolve(this._realm);
            }

            let url = siteUrl + "/vti_bin/client.svc";

            callNodeFetch(url, {
                "method": "POST",
                "headers": {
                    "Authorization": "Bearer ",
                },
            }).then((r) => {

                let data: string = r.headers.get("www-authenticate");
                let index = data.indexOf("Bearer realm=\"");
                TokenHelper._realm = data.substring(index + 14, index + 50);
                resolve(this._realm);
            });
        });
    }

    private static getAuthUrl(realm: string): Promise<string> {

        let url = `https://accounts.accesscontrol.windows.net/metadata/json/1?realm=${realm}`;

        return callNodeFetch(url).then((r: nodeFetch.Response) => r.json()).then((json: any) => {

            for (let i = 0; i < json.endpoints.length; i++) {
                if (json.endpoints[i].protocol === "OAuth2") {
                    return json.endpoints[i].location;
                }
            }

            throw new Error("Auth URL Endpoint could not be determined from data.");
        });
    }

    private static getFormattedPrincipal(principalName: string, hostName: string, realm: string) : string {
        var name;
        if (hostName) {
            name = principalName + "/" + hostName + "@" + realm;
        } else {
            name = principalName + "@" + realm
        }
        console.log('Formated princial :', name)
        return name;
    }
}