export class SharePointAddinConfiguration {
    public clientId: string;
    public clientSecret: string;
    public spHostUrl: string;

    private static _instance = null;
    public static getInstance() : SharePointAddinConfiguration {
        return SharePointAddinConfiguration._instance || (SharePointAddinConfiguration._instance = new SharePointAddinConfiguration());
    }

    public static init(clientId: string, clientSecret: string) : void {
        let config = SharePointAddinConfiguration.getInstance();
        config.clientId = clientId;
        config.clientSecret = clientSecret;
    }
}