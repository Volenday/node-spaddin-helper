export class Url {
    public static ensureTrailingSlash(url: string) : string
    {
        if (url && url[url.length-1] != '/')
            url = url + "/";
        
        return url;
    }

    public static parse(url: string) {
        var match = url.match(/^(https?\:)\/\/(([^:\/?#]*)(?:\:([0-9]+))?)([\/]{0,1}[^?#]*)(\?[^#]*|)(#.*|)$/);
        return match && {
                protocol: match[1],
                host: match[2],
                hostname: match[3],
                port: match[4],
                pathname: match[5],
                search: match[6],
                hash: match[7]
            };
    }

    public static validateHttpSchemes(url: string, schemes: string[]) : boolean {
            if (!url) return false;

            for (let i=0,to=schemes.length; i < to; i++) {
                let currentScheme = schemes[i];
                currentScheme = currentScheme.indexOf("://") < 0 ? (currentScheme + "://") : currentScheme;
                if (url.indexOf(currentScheme) == 0)
                    return true;
            }

            return false;
        }
}