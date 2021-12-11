export declare const mockClient: (data?: any) => {
    api: (path: '') => {
        get: () => Promise<any>;
        post: () => Promise<any>;
        patch: () => Promise<any>;
        delete: () => Promise<any>;
    };
};
