<template lang='pug'>
  div
    el-card.header
      .clearfix(slot='header')
        span 表头内容
        el-button(style='float: right; padding: 3px 0', type='text' @click="onAdd(1)") add
      el-row(v-for='item in theader', :key='item' :gutter='2')
        el-col(:span='2')
          div.el-icon-location-outline(@click="onPosition(item.address)" style="margin-top:30px")
        el-col(:span='8')
          div 名称：
            el-input(placeholder='请输入名称', v-model='item.name')
              el-button(slot='append' type="primary"  @click='getName(item)') 获取选中
        el-col(:span='8')
          div 地址(例：'A1')：
            el-input(placeholder='请输入地址', v-model='item.address' )
              template(#append)
                el-button( type="primary"  @click='getAddress(item)') 获取选中
        el-col(:span='4')
          div 数据类型
            el-select(v-model='item.type', placeholder='请选择')
              el-option(v-for='op in options', :key='op.value', :label='op.label', :value='op.value')
        el-col(:span='2')
          div.el-icon-delete(@click="onDelete(item, 1)" style="margin-top:30px")
    el-card.body
      .clearfix(slot='header')
        span 主要内容
        el-button(style='float: right; padding: 3px 0', type='text'  @click="onAdd(2)") add
      el-row(v-for='item in tbody', :key='item' :gutter='2')
        el-col(:span='2')
          div.el-icon-location-outline(@click="onPosition(item.address)" style="margin-top:30px")
        el-col(:span='8')
          div 名称：
            el-input(placeholder='请输入名称', v-model='item.name')
              el-button(slot='append' type="primary"  @click='getName(item)') 获取选中
        el-col(:span='8')
          div 地址(例：'A1')：
            el-input(placeholder='请输入地址', v-model='item.address' )
              template(#append)
                el-button( type="primary"  @click='getAddress(item)') 获取选中
        el-col(:span='4')
          div 数据类型
            el-select(v-model='item.type', placeholder='请选择')
              el-option(v-for='op in options', :key='op.value', :label='op.label', :value='op.value')
        el-col(:span='2')
          div.el-icon-delete(@click="onDelete(item, 2)" style="margin-top:30px")
    el-card.footer
      .clearfix(slot='header')
        span 表尾内容
        el-button(style='float: right; padding: 3px 0', type='text'  @click="onAdd(3)") add
      el-row(v-for='item in tfooter', :key='item' :gutter='2')
        el-col(:span='2')
          div.el-icon-location-outline(@click="onPosition(item.address)" style="margin-top:30px")
        el-col(:span='8')
          div 名称：
            el-input(placeholder='请输入名称', v-model='item.name')
              el-button(slot='append' type="primary"  @click='getName(item)') 获取选中
        el-col(:span='8')
          div 地址(例：'A1')：
            el-input(placeholder='请输入地址', v-model='item.address' )
              template(#append)
                el-button( type="primary"  @click='getAddress(item)') 获取选中
        el-col(:span='4')
          div 数据类型
            el-select(v-model='item.type', placeholder='请选择')
              el-option(v-for='op in options', :key='op.value', :label='op.label', :value='op.value')
        el-col(:span='2')
          div.el-icon-delete(@click="onDelete(item, 3)" style="margin-top:30px")
    el-card.other
      .clearfix(slot='header')
        span 其他内容
      .demo-input-suffix
        | body_header所在行号：
        el-input(placeholder='', v-model='tother.body_start_row')
      .demo-input-suffix
        | footer起始行号：
        el-input(placeholder='', v-model='tother.foot_start_row')
      .demo-input-suffix
        | footer总行数：
        el-input(placeholder='', v-model='tother.foot_row_count')
     
    el-button(type="primary"  @click='onSubmit()') 提交数据
    //- .divItem
    //-   | 这是一个网页，按
    //-   span(style='font-weight: bolder;') "F12"
    //-   | 可以111sdfsdf
    //- .divItem
    //-   | 这个示例展示了wps加载项的相关基础能力，与B/S业务系统的交互，请用浏览器打开：
    //-   span(style='font-weight: bolder;color: slateblue;') {{DemoSpan}}
    //- .divItem
    //-   | 开发文档: 
    //-   span(style='font-weight: bolder;color: slateblue;') https://open.wps.cn/docs/office
    //- hr
    .divItem
      button(style='margin:3px;', @click="onbuttonclick('dockLeft')") 停靠左边
      button(style='margin:3px;', @click="onbuttonclick('dockRight')") 停靠右边
      button(style='margin:3px;', @click="onbuttonclick('hideTaskPane')") 隐藏TaskPane
      //- button(style='margin:3px;', @click="onbuttonclick('addString')") 文档开头添加字符串
      //- button(style='margin:3px;', @click='onDocNameClick()') 取文件名
      //- button(style='margin:3px;', @click='getSheet()') 获取sheet
    //- hr
    //- .divItem
    //-   | 文档文件名为：
    //-   span {{docName}}

</template>

<script>
import axios from 'axios'
import taskPane from './js/taskpane.js'
export default {
  name: 'TaskPane',
  data(){
    return {
      DemoSpan : '',
      docName: '',
      theader:[
        {
          address: 'A1',
          type: 'text',
          name: '当前'
        }
      ],
      tbody:[],
      tfooter:[],
      tother:{
        body_start_row: 6,   // body 数据开始的行
        foot_start_row: 7,   // 表尾 开始的行 
        foot_row_count: 3,   // 表尾总行数  
      },
      options:[
        {
          value: 'text',
          label: '文本'
        },
        {
          value: 'date',
          label: '时间'
        },
        {
          value: 'image',
          label: '图片'
        }
      ]
    }
  },
  methods:{
    onAdd(type) {
      let cell = wps.EtApplication().ActiveCell
      let item = {
        name: cell.Text,
        address: cell.Address(false,false),
        type: 'text',
        col: cell.Column,
        row: cell.Row
      }
      let sheet = wps.EtApplication().ActiveSheet
      if(type == 1) {
          console.log(sheet)
          this.theader.push(item)
      }  else if(type == 2) {
          console.log(sheet)
          this.tbody.push(item)
      } else if(type == 3) {
          console.log(sheet)
          this.tfooter.push(item)
      }
    },
    getName(obj) {
      console.log(obj)
      // let sheet = wps.EtApplication().ActiveSheet
      let cell = wps.EtApplication().ActiveCell
      obj.name = cell.Text
      // console.log(cell )
      // console.log(cell.AddressLocal())
      // console.log(cell.ApplyNames() )
      // console.log(cell.Column)
      // console.log(cell.Row )
      // console.log(cell.Text)
      // console.log(cell.Address() )
    },
    getAddress(obj) {
      console.log(obj)
      let cell = wps.EtApplication().ActiveCell
      obj.address = cell.Address(false,false)
      obj.col = cell.Column,
      obj.Row = cell.Row
    },
    onPosition(address) {
      let sheet = wps.EtApplication().ActiveSheet
      sheet.Range(address).Select()
    },
    onDelete(item, type) {
      if(type == 1) {
        this.theader = this.theader.filter(p=> p!= item)
      } else if(type == 2) {
          this.tbody = this.tbody.filter(p=> p!= item)
      } else if(type == 3) {
          this.tfooter = this.tfooter.filter(p=> p!= item)
      }
    },
    onbuttonclick(id){
      return taskPane.onbuttonclick(id)
    },
    onDocNameClick(){
      this.docName = taskPane.onbuttonclick('getDocName')
    },
    getSheet() {
      let sheet = wps.EtApplication().ActiveWorkbook
      console.log(111, sheet)
      let curSheet = wps.EtApplication().ActiveSheet;
      if (curSheet){
        console.log(2222, curSheet.Rows.length)
        // let cell = curSheet.Cells.Item(1)
        // let {Row: row , Column: col, Text: text}  = cell
        let row = curSheet.UsedRange.Rows.Count
        let col = curSheet.UsedRange.Columns.Count
        let count = curSheet.UsedRange.Count
        console.log(row, col, count)

          for(let i = 1; i <= count; i++){
            if(curSheet.UsedRange.Item(i).Value2){
                console.log(curSheet.UsedRange.Item(i).Value2)
                console.log(curSheet.UsedRange.Item(i).Text)
                console.log(curSheet.UsedRange.Item(i).Address())
            }
        }
        // for(let i = 1; i <= curSheet.Range("A1:D10").Count; i++){
        //     if(Worksheets.Item("Sheet1").Range("A1:D10").Item(i).Value2 < .001 ){
        //         Worksheets.Item("Sheet1").Range("A1:D10").Item(i).Value2 = 0
        //     }
        // }
        // console.log(col, row, text)
        // console.log(cell.Address(), curSheet.Rows.Count)
        // curSheet.Cells.Item(1, 1).Formula="Hello,1123123123 wps加载项!" + curSheet.Cells.Item(1, 1).Formula
      }
    },
    onSubmit() {
      let data = {
        theader: this.theader,
        tfooter: this.tfooter,
        tbody: this.tbody,
        tother: this.tother
      }
      console.log(data)
    }
  },
  mounted() {
      axios.get('/.debugTemp/NotifyDemoUrl').then((res) => {
          this.DemoSpan = res.data;
      });
      let curSheet = wps.EtApplication().ActiveSheet;
      curSheet.Worksheet_SelectionChange = (target)=>{
        console.log(target)
      }
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
.global{
        font-size: 15px;
        min-height: 95%;
    }
.divItem{
    margin-left: 5px;
    margin-bottom: 18px;
    font-size: 15px;
    word-wrap:break-word;
}

.el-header, .el-footer {
    background-color: #B3C0D1;
    color: #333;
    text-align: center;
    line-height: 60px;
  }
  
  .el-aside {
    background-color: #D3DCE6;
    color: #333;
    text-align: center;
    line-height: 200px;
  }
  
  .el-main {
    background-color: #E9EEF3;
    color: #333;
    text-align: center;
    line-height: 160px;
  }
  
  body > .el-container {
    margin-bottom: 40px;
  }
  
  .el-container:nth-child(5) .el-aside,
  .el-container:nth-child(6) .el-aside {
    line-height: 260px;
  }
  
  .el-container:nth-child(7) .el-aside {
    line-height: 320px;
  }
</style>
