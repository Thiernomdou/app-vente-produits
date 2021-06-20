/* nous allons créer une fonction pour récupérer tous les produits
 avec Axios, une fonction pour récupérer les détails d'un produit, 
 une fonction pour ajouter le produit au panier et 
 une fonction pour retirer le produit du panier*/
import axios from "axios"
import { get } from "core-js/core/dict";

//Action to get products list

    export function getProducts({ commit }) {
        let url = "https://my-json-server.typicode.com/Nelzio/ecommerce-fake-json/products";
        axios.get(url).then((response) => {
            commit("setProducts", response.data);
        }).catch(error => {
            console.log(error);
        })
    }