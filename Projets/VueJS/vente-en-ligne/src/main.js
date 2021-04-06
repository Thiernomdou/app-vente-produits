import '@fortawesome/fontawesome-free/css/all.css';
import '@fortawesome/fontawesome-free/js/all.js';
import 'bootstrap';
import 'bootstrap/dist/css/bootstrap.min.css';

import Vue from 'vue';
import VueRouter from 'vue-router';

import App from './App.vue';
import Routes from './Routes';

Vue.config.productionTip = false
Vue.use(VueRouter)


const router = new VueRouter ({
  routes: Routes,
  mode: 'history'
})

new Vue({
  render: h => h(App),
  router: router
}).$mount('#app')
