"use strict";

const AcquireTokenWithClientCredentials = require("./AdalNode").AcquireTokenWithClientCredentials;
const BaseService = require("./Services/BaseService").BaseService;
const url = require("url");

exports.Execute = async (resource) => {

    try {

        const tokenResponse = await AcquireTokenWithClientCredentials();

        const webApiUrl = url.format(resource + "api/data/v9.1/");

        var baseService = new BaseService(tokenResponse, webApiUrl);
    }
    catch(error) {
        console.error(error.message);
    }
}