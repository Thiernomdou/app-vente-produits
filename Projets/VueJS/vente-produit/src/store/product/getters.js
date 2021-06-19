export function products(state) {
    //retourne l'état qui stocke la liste de produits
    return state.products
}
export function product(state) {
    /*renvoie l'état qui stocke un produit spécifique
    lorsque l'utilisateur veut voir les détails du produit*/
    return state.product
}
export function cart(state) {
    //retourne l'état qui stocke les produits dans le panier
    return state.cart
}