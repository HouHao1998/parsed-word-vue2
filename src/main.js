import Vue from 'vue';
import App from './App.vue';
import vuetify from './plugins/vuetify';
import './assets/styles.css'; // 导入全局样式

Vue.config.productionTip = false;

new Vue({
  vuetify,
  render: h => h(App)
}).$mount('#app');
