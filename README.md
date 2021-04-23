# wps版配置生成工具
> 通过wps表格生成标准json配置，常用于自定义搭建配置界面生成

## 了解
### 效果

### 优势
### 安装体验

## 二次开发
## 项目安装
将项目克隆至本地
```
npm install
```
### 开发
```
npm run serve
```

### 构建
```
npx wpsjs build
```

### 其他配置参考vue
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

## 演示demo
![演示](./wps-table-addon.gif)