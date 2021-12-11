"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.mockClient = void 0;
const mockClient = (data) => {
    return {
        api: (path) => {
            return {
                get: () => Promise.resolve(data),
                post: () => Promise.resolve(data),
                patch: () => Promise.resolve(data),
                delete: () => Promise.resolve(data),
            };
        },
    };
};
exports.mockClient = mockClient;
//# sourceMappingURL=microsoft-graph-client.js.map