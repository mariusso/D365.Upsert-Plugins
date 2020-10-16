"use strict";

const url = require("url");
const superagent = require("superagent");
const { webApiurl } = require("../clientConfig.json");

exports.BaseService = class BaseService {

    constructor(tokenResponse, entitySetName) {

        this.RequestHeaders = {
            "Authorization": "Bearer " + tokenResponse.accessToken,
            "Accept": "application/json",
            "Content-Type": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "If-None-Match": null,
            "Prefer": "return=representation,odata.include-annotations=\"*\""
        };

        this.WebApiUrl = webApiurl;
        this.EntitySetName = entitySetName;
    }

    async Create(entity) {

        let response = undefined;
        try {
            response = await superagent.post(url.resolve(this.WebApiUrl, this.EntitySetName)).set(this.RequestHeaders).send(entity);
        }
        catch (errorResponse) {
            const message = errorResponse.response.body.error.message || errorResponse.message;
            throw new Error(message);
        }

        return response.body;
    }

    async Retrieve(id, select = "") {

        const operators = ["?"];

        if (select) {
            select = `${operators.shift()}$select=${select}`;
        }

        const address = encodeURI(`${url.resolve(this.WebApiUrl, this.EntitySetName)}(${id})${select}`);

        let response = undefined;
        try {
            response = await superagent.get(address).set(this.RequestHeaders);
        }
        catch (errorResponse) {
            const message = errorResponse.response.body.error.message || errorResponse.message;
            throw new Error(message);
        }

        return response.body;
    }

    async RetrieveMultiple(select = "", filter = "", orderBy = "", top = "", expand = "") {

        const operators = ["?", "&", "&", "&", "&"];

        if (select) {
            select = `${operators.shift()}$select=${encodeURIComponent(select)}`;
        }

        if (filter) {
            filter = `${operators.shift()}$filter=${encodeURIComponent(filter)}`;
        }

        if (orderBy) {
            orderBy = `${operators.shift()}$orderby=${encodeURIComponent(orderBy)}`;
        }

        if (top) {
            top = `${operators.shift()}$top=${encodeURIComponent(top)}`;
        }

        if (expand) {
            expand = `${operators.shift()}$expand=${encodeURIComponent(expand)}`;
        }

        const address = `${url.resolve(this.WebApiUrl, this.EntitySetName)}${select}${filter}${orderBy}${top}`;

        let response = undefined;
        try {
            response = await superagent.get(address).set(this.RequestHeaders);
        }
        catch (errorResponse) {
            const message = errorResponse.response.body.error.message || errorResponse.message;
            throw new Error(message);
        }

        return response.body.value;
    }

    async Update(id, entity) {

        const address = encodeURI(`${url.resolve(this.WebApiUrl, this.EntitySetName)}(${id})`);

        let response = undefined;
        try {
            response = await superagent.patch(address).set(this.RequestHeaders).send(entity);
        }
        catch (errorResponse) {
            const message = errorResponse.response.body.error.message || errorResponse.message;
            throw new Error(message);
        }

        return response.body;
    }

    async Delete(id) {

        const address = encodeURI(`${url.resolve(this.WebApiUrl, this.EntitySetName)}(${id})`);

        let response = undefined;
        try {
            response = await superagent.del(address).set(this.RequestHeaders);
        }
        catch (errorResponse) {
            const message = errorResponse.response.body.error.message || errorResponse.message;
            throw new Error(message);
        }
    }

    async ExecuteFetchXml(fetchXml) {

        const address = `${url.resolve(this.WebApiUrl, this.EntitySetName)}?fetchXml=${encodeURI(fetchXml)}`;

        let response = undefined;
        try {
            response = await superagent.get(address).set(this.RequestHeaders);
        }
        catch (errorResponse) {
            const message = errorResponse.response.body.error.message || errorResponse.message;
            throw new Error(message);
        }

        return response.body.value;
    }

    async Associate(primaryEntityId, secondaryEntitySetName, secondaryEntityId, relationshipName)
    {
        const address = `${url.resolve(this.WebApiUrl, this.EntitySetName)}(${primaryEntityId})/${relationshipName}/$ref`;

        const payload = {
            "@odata.id": `${this.WebApiUrl}/${secondaryEntitySetName}(${secondaryEntityId})`
        };

        let response = undefined;
        try {
            response = await superagent.post(address).set(this.RequestHeaders).send(payload);
        }
        catch (errorResponse) {
            const message = errorResponse.response.body.error.message || errorResponse.message;
            throw new Error(message);
        }
    }
}