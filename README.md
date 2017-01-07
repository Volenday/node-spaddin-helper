SharePoint provider-hosted Addins with Node.js
==============================================

Platform
--------
- Supports only SharePoint Online

How to install
--------------

> npm install spaddin-helper --save

How to use
----------

In a Node http request handler

> const handler = (req, res) => {
>        let ctx = SharePointContext.getFromRequest(req);
>        ctx.createAppOnlyClientForSPHost().then(client => {
>            client.get('/_api/web/currentuser').then(user => {
>                console.log(`username = ${user.Title}`);
>            });
>        });
>};


