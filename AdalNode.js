"use strict";

const AuthenticationContext = require('adal-node').AuthenticationContext;

exports.AcquireTokenWithUserNameAndPassword = (clientConfig) => {

    if(clientConfig == null) {
        throw new Error("Missing clientConfig.json");
    }


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

 exports.AcquireTokenWithClientCredentials = (clientConfig) => {

    if(clientConfig == null) {
        throw new Error("Missing clientConfig.json");
    }

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