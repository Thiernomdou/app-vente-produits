export function setProducts(state, val) {
    //Attribuer une liste de produits aux produits d'état
    state.products = val
}
export function setProduct(state, val) {
    //Affecter un objet avec un produit spécifique à l'état du produit
    state.product = val
}
export function setLoad(state, val) {
    //uploade
    state.uploadingData = val 
}
export function setCart(state, val) {
    /*Attribuer une liste de produits ajoutés dans
    le panier à l'état du panier*/
    state.cart = val
}