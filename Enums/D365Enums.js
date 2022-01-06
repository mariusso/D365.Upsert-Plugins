exports.D365Enums = class D365Enums {
    static SourceType = {
        Database: 0
    }

    static IsolationMode = {
        None: 1,
        Sandbox: 2
    }
    
    static Mode = {
        Sync: 0,
        Async: 1
    };

    static Stage = {
        PreValidation: 10,
        PreOperation: 20,
        PostOperation: 40
    };

    static InvocationSource = {
        Parent: 0,
        Child: 1
    };

    static SupportedDeployment = {
        Server: 0,
        Outlook: 1,
        Both: 2
    };

    static ImageType = {
        PreImage: 0,
        PostImage: 1,
        Both: 2
    };
}