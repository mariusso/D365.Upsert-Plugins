"use strict";

const { AcquireTokenWithClientCredentials } = require("./AdalNode");
const { BaseService } = require("./Services/BaseService");
const pluginConfig = require("./pluginConfig.json");
const fs = require('fs').promises;
const clientConfig = require("./clientConfig.json");
const path = require("path");
const { Console } = require("console");

let assemblyService;
let pluginService;
let stepService;
let sdkMessageService;
let sdkMessageFilterService;
let secureConfigService;
let imageService;

async function EntryPoint() {

    const args = process.argv.slice(2);

    if (args.length < 1) {
        RegisterPlugins();
    }
    else {

        const command = args[0];

        if (command === "RegisterPlugins") {
            RegisterPlugins();
        }
        else if (command === "CreateConfigFromD365") {

            if(args.length < 2) {
                console.error("At least one assembly name is required.");
                return;
            }

            const assemblyNames = args.slice(1);
            CreateConfigFromD365(assemblyNames);
        }
        else {
            Console.error(`Unknown command: ${command}`);
        }
    }
}

async function RegisterPlugins() {

    try {

        const tokenResponse = await AcquireTokenWithClientCredentials();

        assemblyService = new BaseService(tokenResponse, "pluginassemblies");
        pluginService = new BaseService(tokenResponse, "plugintypes");
        stepService = new BaseService(tokenResponse, "sdkmessageprocessingsteps");
        sdkMessageService = new BaseService(tokenResponse, "sdkmessages");
        sdkMessageFilterService = new BaseService(tokenResponse, "sdkmessagefilters");
        secureConfigService = new BaseService(tokenResponse, "sdkmessageprocessingstepsecureconfigs");
        imageService = new BaseService(tokenResponse, "sdkmessageprocessingstepimages");

        for (const assembly of pluginConfig.Assemblies) {

            assembly.Id = await UpsertAssembly(assembly);

            for (const plugin of assembly.Plugins) {

                plugin.Id = await UpsertPlugin(plugin, assembly.Id);

                for (const step of plugin.Steps) {

                    step.Id = await UpsertStep(step, plugin.name, plugin.Id);

                    if (step.secureconfiguration) {
                        step.secureconfiguration.Id = await UpsertSecureConfig(step.secureconfiguration, step.Id);
                    }

                    for (const image of step.Images) {

                        image.Id = await UpsertImage(image, step.Id);
                    }
                }
            }
        }

        var pluginConfigJson = JSON.stringify(pluginConfig);

        fs.writeFile("./pluginConfig.json", pluginConfigJson);

        console.info("Plugin Registration - Done.");
    }
    catch (error) {
        console.error(error.message);
    }
}

async function CreateConfigFromD365(assemblyNames) {

    try {
        const tokenResponse = await AcquireTokenWithClientCredentials();

        assemblyService = new BaseService(tokenResponse, "pluginassemblies");
        pluginService = new BaseService(tokenResponse, "plugintypes");
        stepService = new BaseService(tokenResponse, "sdkmessageprocessingsteps");
        sdkMessageService = new BaseService(tokenResponse, "sdkmessages");
        sdkMessageFilterService = new BaseService(tokenResponse, "sdkmessagefilters");
        secureConfigService = new BaseService(tokenResponse, "sdkmessageprocessingstepsecureconfigs");
        imageService = new BaseService(tokenResponse, "sdkmessageprocessingstepimages");

        const config = {
            Assemblies: []
        };

        for (const assemblyName of assemblyNames) {
            const assemblies = await assemblyService.RetrieveMultiple("", `name eq '${assemblyName}'`, "", "1");

            if (assemblies.length < 1) {
                throw new Error(`Could not retrieve assembly with name: '${assemblyName}'.`);
            }

            const assembly = assemblies[0];

            const assemblyConfig = {
                Id: assembly.pluginassemblyid,
                relativePath: "",
                description: assembly.description,
                version: assembly.version,
                sourcetype: assembly.sourcetype,
                isolationmode: assembly.isolationmode,
                Plugins: []
            }

            const plugins = await pluginService.RetrieveMultiple("", `_pluginassemblyid_value eq '${assembly.pluginassemblyid}'`);

            for (const plugin of plugins) {

                const pluginConfig = {
                    Id: plugin.plugintypeid,
                    typename: plugin.typename,
                    name: plugin.name,
                    friendlyname: plugin.friendlyname,
                    description: plugin.description,
                    Steps: []
                };

                const steps = await stepService.RetrieveMultiple("", `_plugintypeid_value eq '${plugin.plugintypeid}'`);

                for (const step of steps) {

                    const stepConfig = {
                        Id: step.sdkmessageprocessingstepid,
                        name: step.name,
                        message: "",
                        entity: "",
                        filteringattributes: step.filteringattributes,
                        mode: step.mode,
                        stage: step.stage,
                        rank: step.rank,
                        supporteddeployment: step.supporteddeployment,
                        asyncautodelete: step.asyncautodelete,
                        configuration: step.configuration,
                        secureconfiguration: null,
                        Images: []
                    };

                    if (step._sdkmessageid_value) {
                        stepConfig.message = step["_sdkmessageid_value@OData.Community.Display.V1.FormattedValue"];
                    }

                    if (step._sdkmessagefilterid_value) {
                        const sdkMessageFilter = await sdkMessageFilterService.Retrieve(step._sdkmessagefilterid_value, "primaryobjecttypecode");
                        stepConfig.entity = sdkMessageFilter.primaryobjecttypecode;
                    }

                    if (step._sdkmessageprocessingstepsecureconfigid_value) {
                        const secureConfig = await secureConfigService.Retrieve(step._sdkmessageprocessingstepsecureconfigid_value, "sdkmessageprocessingstepsecureconfigid,secureconfig");
                        stepConfig.secureconfiguration = {
                            Id: secureConfig.sdkmessageprocessingstepsecureconfigid,
                            secureconfig: secureConfig.secureconfig
                        };
                    }

                    const images = await imageService.RetrieveMultiple("", `_sdkmessageprocessingstepid_value eq ${step.sdkmessageprocessingstepid}`);

                    for (const image of images) {
                        
                        const imageConfig = {
                            Id: image.sdkmessageprocessingstepimageid,
                            name: image.name,
                            entityalias: image.entityalias,
                            description: image.description,
                            imagetype: image.imagetype,
                            messagepropertyname: image.messagepropertyname,
                            attributes: image.attributes
                        }

                        stepConfig.Images.push(imageConfig);
                    }

                    pluginConfig.Steps.push(stepConfig);
                }

                assemblyConfig.Plugins.push(pluginConfig);
            }

            config.Assemblies.push(assemblyConfig);
        }

        console.info("Writing to file 'configFromD365.json'")

        var configJson = JSON.stringify(config);

        fs.writeFile("./configFromD365.json", configJson);

        console.info("Create Config from D365 - Done.");
    }
    catch (error) {
        console.error(error.message);
    }
}

async function UpsertAssembly(assembly) {

    const assemblyPath = path.resolve(__dirname, assembly.relativePath);
    const assemblyName = path.basename(assemblyPath);

    const assemblyContent = await fs.readFile(assemblyPath, { encoding: 'base64' });
    const upsertMe = {
        description: assembly.description,
        version: assembly.version,
        sourcetype: assembly.sourcetype,
        isolationmode: assembly.isolationmode,
        content: assemblyContent
    };

    let upsertedRecord = null;

    if (assembly.Id) {
        console.info(`Assembly - ${assemblyName} - existing`);
        upsertedRecord = await assemblyService.Update(assembly.Id, upsertMe);
    }
    else {
        console.info(`Assembly - ${assemblyName} - new`);
        upsertedRecord = await assemblyService.Create(upsertMe);
    }

    return upsertedRecord.pluginassemblyid
}

async function UpsertPlugin(plugin, assemblyId) {

    const upsertMe = {
        "pluginassemblyid@odata.bind": `${clientConfig.webApiurl}pluginassemblies(${assemblyId})`,
        typename: plugin.typename,
        name: plugin.name,
        friendlyname: plugin.friendlyname,
        description: plugin.description
    };

    let upsertedRecord = null;

    if (plugin.Id) {
        console.info(`Plugin - ${plugin.typename} - existing`);
        upsertedRecord = await pluginService.Update(plugin.Id, upsertMe);
    }
    else {
        console.info(`Plugin - ${plugin.typename} - new`);
        upsertedRecord = await pluginService.Create(upsertMe);
    }

    return upsertedRecord.plugintypeid;
}

async function UpsertStep(step, pluginName, pluginId) {

    const sdkMessages = await sdkMessageService.RetrieveMultiple("sdkmessageid", `name eq '${step.message}'`, "", "1");

    if (sdkMessages.length < 1) {
        throw new Error(`SDK Message '${step.message}' does not exist.`);
    }

    const sdkMessageId = sdkMessages[0].sdkmessageid;

    let sdkMessageFilterId = null;

    if (step.entity) {
        const sdkMessageFilters = await sdkMessageFilterService.RetrieveMultiple("sdkmessagefilterid", `_sdkmessageid_value eq '${sdkMessageId}' and primaryobjecttypecode eq '${step.entity.toLowerCase()}'`, "", "1");

        if (sdkMessageFilters.length < 1) {
            throw new Error(`Invalid message for entity '${step.message}' or invalid entity name '${step.entity}'.`);
        }

        sdkMessageFilterId = sdkMessageFilters[0].sdkmessagefilterid;
    }

    const upsertMe = {
        "plugintypeid@odata.bind": `${clientConfig.webApiurl}plugintypes(${pluginId})`,
        "sdkmessageid@odata.bind": `${clientConfig.webApiurl}sdkmessages(${sdkMessageId})`,
        name: step.name,
        filteringattributes: step.filteringattributes,
        mode: step.mode,
        stage: step.stage,
        rank: step.rank,
        asyncautodelete: step.asyncautodelete,
        configuration: step.configuration,
        supporteddeployment: step.supporteddeployment,
    };

    if (sdkMessageFilterId) {
        upsertMe["sdkmessagefilterid@odata.bind"] = `${clientConfig.webApiurl}sdkmessagefilters(${sdkMessageFilterId})`;
    }

    let upsertedRecord = null;

    if (step.Id) {
        console.info(`Step - '${pluginName} - ${step.message}'  - existing`);
        upsertedRecord = await stepService.Update(step.Id, upsertMe);
    }
    else {
        console.info(`Step - '${pluginName} - ${step.message}'  - new`);
        upsertedRecord = await stepService.Create(upsertMe);
    }

    return upsertedRecord.sdkmessageprocessingstepid;
}

async function UpsertSecureConfig(secureConfig, stepId) {

    const upsertMe = {
        secureconfig: secureConfig.secureconfig
    };

    let upsertedRecord = null;

    if (secureConfig.Id) {
        console.info(`Secure Config - existing`);
        upsertedRecord = await secureConfigService.Update(secureConfig.Id, upsertMe);
    }
    else {
        console.info(`Secure Config - new`);
        upsertedRecord = await secureConfigService.Create(upsertMe);

        const stepUpdate = {
            "sdkmessageprocessingstepsecureconfigid@odata.bind": `${clientConfig.webApiurl}sdkmessageprocessingstepsecureconfigs(${upsertedRecord.sdkmessageprocessingstepsecureconfigid})`
        }
        await stepService.Update(stepId, stepUpdate);
    }

    return upsertedRecord.sdkmessageprocessingstepsecureconfigid;
}

async function UpsertImage(image, stepId) {

    const upsertMe = {
        "sdkmessageprocessingstepid@odata.bind": `${clientConfig.webApiurl}sdkmessageprocessingsteps(${stepId})`,
        name: image.name,
        entityalias: image.entityalias,
        description: image.description,
        imagetype: image.imagetype,
        attributes: image.attributes,
        messagepropertyname: image.messagepropertyname
    };

    let upsertedRecord = null;

    if (image.Id) {
        console.info(`Image - ${image.name} - existing`);
        upsertedRecord = await imageService.Update(image.Id, upsertMe);
    }
    else {
        console.info(`Image - ${image.name} - new`);
        upsertedRecord = await imageService.Create(upsertMe);
    }

    return upsertedRecord.sdkmessageprocessingstepimageid;
}

EntryPoint();