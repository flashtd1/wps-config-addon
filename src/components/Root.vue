<template>
  <div>
    <h1>{{ msg }}</h1>
    <el-button @click="onTask"> 切换页面taskpane </el-button>
    <el-button @click="onConfig"> 切换excle 配置页面 </el-button>
  </div>
</template>

<script>
import ribbon from './ribbon.js'
export default {
  name: 'Root',
  data(){
    return{
      msg: "欢迎来到wps加载项的世界!"
    }
  },
  created(){
    window.ribbon = ribbon   //把ribbon.js中的函数导出，用于在ribbon.xml中配置，从而wps客户端程序能够调用
  },
  methods: {
    onTask() {
      this.$router.push('taskpane')
    },
    onConfig() {
      this.$router.push('taskpane1')
    }
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>

</style>
