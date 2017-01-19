"use strict";
const node_fetch_1 = require("node-fetch");
const JWT = require("jsonwebtoken");
const url_1 = require("../helpers/url");
const configuration_1 = require("./configuration");
class TokenHelper {
    static getAppContextTokenRequestUrl(contextUrl, redirectUri) {
        const config = configuration_1.SharePointAddinConfiguration.getInstance();
        return `${url_1.Url.ensureTrailingSlash(contextUrl)}${TokenHelper.redirectPage}?client_id=${config.clientId}&redirect_uri=${redirectUri}`;
    }
    static getContextTokenFromRequest(request) {
        let paramNames = ["AppContext", "AppContextToken", "AccessToken", "SPAppToken"];
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
    static readAndValidateContext(contextTokenStr, appHostName) {
        const config = configuration_1.SharePointAddinConfiguration.getInstance();
        let token = JWT.decode(contextTokenStr, config.clientSecret);
        token.appctx = JSON.parse(token.appctx);
        // TODO Implement the validation of the token here
        // TODO Return type should be a typed interface
        return token;
    }
    static getAccessToken(contextToken, siteUrl, appOnly = false) {
        return new Promise((resolve, reject) => {
            const config = configuration_1.SharePointAddinConfiguration.getInstance();
            let realm = null;
            let principals = null;
            return this.getRealm(siteUrl)
                .then(retrievedRealm => {
                realm = retrievedRealm;
                principals = {
                    resource: TokenHelper.getFormattedPrincipal(TokenHelper.SharePointServicePrincipal, url_1.Url.parse(siteUrl).hostname, realm),
                    formattedClientId: TokenHelper.getFormattedPrincipal(config.clientId, "", realm)
                };
            })
                .then(() => TokenHelper.getAuthUrl(realm))
                .then((authUrl) => {
                let body = [];
                if (!appOnly) {
                    body.push("grant_type=refresh_token");
                    body.push(`refresh_token=${contextToken.refreshtoken}`);
                }
                else {
                    body.push("grant_type=client_credentials");
                }
                body.push(`client_id=${principals.formattedClientId}`);
                body.push(`client_secret=${encodeURIComponent(config.clientSecret)}`);
                body.push(`resource=${principals.resource}`);
                return node_fetch_1.default(authUrl, {
                    body: body.join("&"),
                    headers: {
                        "Content-Type": "application/x-www-form-urlencoded",
                    },
                    method: "POST",
                });
            })
                .then((r) => {
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
    static getRealm(siteUrl) {
        return new Promise(resolve => {
            if (TokenHelper._realm) {
                resolve(this._realm);
            }
            let url = siteUrl + "/vti_bin/client.svc";
            node_fetch_1.default(url, {
                "method": "POST",
                "headers": {
                    "Authorization": "Bearer ",
                },
            }).then((r) => {
                let data = r.headers.get("www-authenticate");
                let index = data.indexOf("Bearer realm=\"");
                TokenHelper._realm = data.substring(index + 14, index + 50);
                resolve(this._realm);
            });
        });
    }
    static getAuthUrl(realm) {
        let url = `https://accounts.accesscontrol.windows.net/metadata/json/1?realm=${realm}`;
        return node_fetch_1.default(url).then((r) => r.json()).then((json) => {
            for (let i = 0; i < json.endpoints.length; i++) {
                if (json.endpoints[i].protocol === "OAuth2") {
                    return json.endpoints[i].location;
                }
            }
            throw new Error("Auth URL Endpoint could not be determined from data.");
        });
    }
    static getFormattedPrincipal(principalName, hostName, realm) {
        var name;
        if (hostName) {
            name = principalName + "/" + hostName + "@" + realm;
        }
        else {
            name = principalName + "@" + realm;
        }
        console.log('Formated princial :', name);
        return name;
    }
}
// Constants
TokenHelper.SharePointServicePrincipal = "00000003-0000-0ff1-ce00-000000000000";
TokenHelper.authorizationPage = "_layouts/15/OAuthAuthorize.aspx";
TokenHelper.redirectPage = "_layouts/15/AppRedirect.aspx";
TokenHelper.acsPrincipalName = "00000001-0000-0000-c000-000000000000";
TokenHelper.acsMetadataEndPointRelativeUrl = "metadata/json/1";
TokenHelper.S2SProtocol = "OAuth2";
TokenHelper.delegationIssuance = "DelegationIssuance1.0";
TokenHelper.nameIdentifierClaimType = "nameid";
TokenHelper.trustedForImpersonationClaimType = "trustedfordelegation";
TokenHelper.actorTokenClaimType = "actortoken";
TokenHelper._realm = null;
exports.TokenHelper = TokenHelper;
//# sourceMappingURL=token-helper.js.map