
    export interface ISPQueryStringArgs {
        SPHostUrl: string;
        SPAppWebUrl?: string;
        SPLanguage: string;
        SPClientTag: string;
        SPProductNumber: string;
        SPHasRedirectedToSharePoint?: string;
    }
    export interface ISPContextParametersCookie {
        SPHostUrl: string;
    }
    export interface ISPCookies {
        SpContextParameters:ISPContextParametersCookie;
    }
    export interface ISPHttpRequest {
        query: ISPQueryStringArgs;
        cookies: any;
    }

    export interface IAuthToken {
        token_type: string;
        expires_in: string;
        not_before: string;
        expires_on: string;
        resource: string;
        access_token: string;
    }
