import Vue from 'vue'
import App from './App.vue'
import VueRouter from 'vue-router'
// element-ui
import ElementUI from 'element-ui'
import 'element-ui/lib/theme-chalk/index.css'
Vue.use(ElementUI)


Vue.use(VueRouter)
Vue.config.productionTip = false

const routerCfg= [
  {
    path: '/', 
    name: '默认页',
    component:()=>import('./components/Root.vue')
  },{
    path: '/dialog', 
    name: '对话框',
    component:()=>import('./components/Dialog.vue')
  },{
    path: '/taskpane1', 
    name: '任务窗格',
    component:()=>import('./components/TaskPane.vue')
  },{
    path: '/taskpane', 
    name: '页面配置窗口',
    component:()=>import('./components/PageView.vue')
  }
]

new Vue({
  render: h => h(App),
  router: new VueRouter({routes:routerCfg}),
  created: function () {
    //
  }
}).$mount('#app')
