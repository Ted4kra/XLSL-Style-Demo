<template>
  <div id="app">
    <XlsxBase />
  </div>
</template>

<script>

import XlsxBase from "@/components/XlsxBase";

export default {
  name: 'App',
  components: {
    XlsxBase
  }
}
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>
