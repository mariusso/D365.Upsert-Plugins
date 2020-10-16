"use strict";

const url = require("url");
const superagent = require("superagent");

exports.BaseService = class BaseService {

    constructor(tokenResponse, webApiUrl) {

        this.RequestHeaders = {
            "Authorization": "Bearer " + tokenResponse.accessToken,
            "Accept": "application/json",
            "Content-Type": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "If-None-Match": null,
            "Prefer": "return=representation,odata.include-annotations=\"*\""
        };

        this.WebApiUrl = webApiUrl;
    }

    async Create(entitySet, entity) {

        const url = encodeURI(`${this.WebApiUrl}${entitySet}`);

        let response = undefined;
        try {
            response = await superagent.post(url).set(this.RequestHeaders).send(entity);
        }
        catch (errorResponse) {
            throw new Error(errorResponse.response.body.error.message);
        }

        return response.body;
    }

    async Retrieve(entitySet, id, select = "") {

        const operators = ["?"];

        if (select) {
            select = `${operators.shift()}$select=${select}`;
        }

        const url = encodeURI(`${this.WebApiUrl}${entitySet}(${id})${select}`);

        let response = undefined;
        try {
            response = await superagent.get(url).set(this.RequestHeaders);
        }
        catch (errorResponse) {
            throw new Error(errorResponse.response.body.error.message);
        }

        return response.body;
    }

    async RetrieveMultiple(entitySet, select = "", filter = "", orderBy = "", top = "", expand = "") {

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

        const url = `${this.WebApiUrl}${entitySet}${select}${filter}${orderBy}${top}`;

        let response = undefined;
        try {
            response = await superagent.get(url).set(this.RequestHeaders);
        }
        catch (errorResponse) {
            throw new Error(errorResponse.response.body.error.message);
        }

        return response.body.value;
    }

    async Update(entitySet, id, entity) {

        const url = encodeURI(`${this.WebApiUrl}${entitySet}(${id})`);

        let response = undefined;
        try {
            response = await superagent.patch(url).set(this.RequestHeaders).send(entity);
        }
        catch (errorResponse) {
            throw new Error(errorResponse.response.body.error.message);
        }

        return response.body;
    }

    async Delete(entitySet, id) {

        const url = encodeURI(`${this.WebApiUrl}${entitySet}(${id})`);

        let response = undefined;
        try {
            response = await superagent.del(url).set(this.RequestHeaders);
        }
        catch (errorResponse) {
            throw new Error(errorResponse.response.body.error.message);
        }
    }

    async ExecuteFetchXml(entitySet, fetchXml) {

        const url = `${this.WebApiUrl}${entitySet}?fetchXml=${encodeURI(fetchXml)}`;

        let response = undefined;
        try {
            response = await superagent.get(url).set(this.RequestHeaders);
        }
        catch (errorResponse) {
            throw new Error(errorResponse.response.body.error.message);
        }

        return response.body.value;
    }
}