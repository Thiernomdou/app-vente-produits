import Home from '@/components/Home';
import ShoppingCart from '@/components/ShoppingCart';
import UserSettings from '@/components/UserSettings';
import WishList from '@/components/WishList';

export default [
    // Le chemin racine
    {path: '/', component: Home},
    {path: '/user-settings', component: UserSettings},
    {path: '/wish-list', component: WishList},
    {path: '/shopping-cart', component: ShoppingCart},
]