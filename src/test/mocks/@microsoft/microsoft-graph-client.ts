export const get = jest.fn()
export const post = jest.fn()
export const patch = jest.fn()
export const del = jest.fn()

export const mockClient = () => {
    return {
        api: function (path: string) {
            return {
                get,
                post,

                patch,
                delete: del,
            }
        },
    }
}
