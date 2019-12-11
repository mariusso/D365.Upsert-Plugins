"use strict";

const adal = require("./adal");
const superagent = require("superagent");
const findConfig = require("find-config");
const d365Config = findConfig.require("d365.json", { dir: "d365configuration" });

if (d365Config == null) {
    throw new Error("Missing d365.json");
}

exports.DefaultNodeExecution = async () => {

    const tokenResponse = await adal.acquireTokenWithUserNameAndPassword();
    const headers = adal.getRequestHeaders(tokenResponse);

    const whoAmIUrl = d365Config.webApiUrl + "WhoAmI"

    superagent.get(whoAmIUrl).set(headers).then((res) => {
        console.log(res.body);
    }).catch((error) => {
        console.error(error);
    });
}