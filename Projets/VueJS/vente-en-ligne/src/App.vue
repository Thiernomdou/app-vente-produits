<template>
  <div id="app">
    <header-top></header-top>
    <br><br><br>
    <router-view></router-view>
  </div>

  
</template>

<script>

import Header from './components/Header/Header.vue'
export default {
  name: 'App',
  components: {
    'header-top': Header
  }
}
</script>

<style>

</style>
