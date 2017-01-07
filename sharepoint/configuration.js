"use strict";
class SharePointAddinConfiguration {
    static getInstance() {
        return SharePointAddinConfiguration._instance || (SharePointAddinConfiguration._instance = new SharePointAddinConfiguration());
    }
    static init(clientId, clientSecret) {
        let config = SharePointAddinConfiguration.getInstance();
        config.clientId = clientId;
        config.clientSecret = clientSecret;
    }
}
SharePointAddinConfiguration._instance = null;
exports.SharePointAddinConfiguration = SharePointAddinConfiguration;
//# sourceMappingURL=configuration.js.map