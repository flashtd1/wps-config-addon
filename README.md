# et_vue

## Project setup
```
npm install
```

### Compiles and hot-reloads for development
```
npm run serve
```

### Compiles and minifies for production
```
npm run build
```

### Lints and fixes files
```
npm run lint
```

### Customize configuration
See [Configuration Reference](https://cli.vuejs.org/config/).

### wps加载项开发说明
先找到wps.exe所在目录
找到cfgs目录，修改oem.ini文件
在底部加上如下代码
```ini
[Support]
JsApiPlugin=true
JsApiShowWebDebugger=true
```

### 调试
在项目地址执行`npx wpsjs debug`开启调试

### 发布
在项目地址执行`npx wpsjs build`编译代码
按照提示将生成的wps-addon-build目录部署到服务器
剩余步骤下次再写
