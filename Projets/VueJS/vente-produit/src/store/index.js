import account from './account'
import product from './product'
vue.use(Vuex)
export default function() {
    const Store = new Vuex.Store({
        modules: {
            account,
            product
        },
        //enable strict mode (adds overhead!)
        //for dev mode only
        strict: process.env.DEV
    })
    return Store
}