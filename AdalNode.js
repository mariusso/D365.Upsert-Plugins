"use strict";

const AuthenticationContext = require('adal-node').AuthenticationContext;
const clientConfig  = require("./clientConfig.json");

if(clientConfig == null) {
    throw new Error("Missing clientConfig.json");
}

exports.AcquireTokenWithUserNameAndPassword = () => {
    return new Promise((resolve, reject) => {
        const authContext = new AuthenticationContext(clientConfig.authority, true);
        authContext.acquireTokenWithUsernamePassword(clientConfig.resource, clientConfig.username, clientConfig.password, clientConfig.clientId, (error, tokenResponse) => {
            if(error) {
                reject(error);
            }
            resolve(tokenResponse);
        });
    })
 };

 exports.AcquireTokenWithClientCredentials = () => {
    return new Promise((resolve, reject) => {
        const authContext = new AuthenticationContext(clientConfig.authority, true);
        authContext.acquireTokenWithClientCredentials(clientConfig.resource, clientConfig.clientId, clientConfig.clientSecret, (error, tokenResponse) => {
            if(error) {
                reject(error);
            }
            resolve(tokenResponse);
        });
    })
 };