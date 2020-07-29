<template lang='pug'>
  div
    el-card.header
      .clearfix(slot='header')
        span 表头内容
        el-button(style='float: right; padding: 3px 5px', type='text' @click="onAddRange(1)") addRangeInRight
        el-button(style='float: right; padding: 3px 5px', type='text' @click="onAddRange(1,false)") addRangeInBottom
        el-button(style='float: right; padding: 3px 5px', type='text' @click="onAdd(1)") addInRight
        el-button(style='float: right; padding: 3px 5px', type='text' @click="onAdd(1,false)") addBottom
        el-button(style='float: right; padding: 3px 5px', type='text' @click="onClean(1)") Clean
      el-row(v-for='(item, index) in theader', :key='index' :gutter='2')
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
              el-option(v-for='(op,index) in options', :key='index', :label='op.label', :value='op.value')
        el-col(:span='2')
          div.el-icon-delete(@click="onDelete(item, 1)" style="margin-top:30px")
    el-card.body
      .clearfix(slot='header')
        span 主要内容
        el-button(style='float: right; padding: 3px 5px', type='text' @click="onAddRange(2)") addRangeInRight
        el-button(style='float: right; padding: 3px 5px', type='text' @click="onAddRange(2,false)") addRangeInBottom
        el-button(style='float: right; padding: 3px 5px', type='text' @click="onAdd(2)") addInRight
        el-button(style='float: right; padding: 3px 5px', type='text' @click="onAdd(2,false)") addBottom
        el-button(style='float: right; padding: 3px 5px', type='text' @click="onClean(2)") Clean
      el-row(v-for='(item,index) in tbody', :key='index' :gutter='2')
        el-col(:span='1')
          div.el-icon-location-outline(@click="onPosition(item.address)" style="margin-top:30px")
        el-col(:span='8')
          div 名称：
            el-input(placeholder='请输入名称', v-model='item.name')
              el-button(slot='append' type="primary"  @click='getName(item)') 获取选中
        el-col(:span='7')
          div 地址(例：'A1')：
            el-input(placeholder='请输入地址', v-model='item.address' )
              template(#append)
                el-button( type="primary"  @click='getAddress(item)') 获取选中
        el-col(:span='3')
          div 数据类型
            el-select(v-model='item.type', placeholder='请选择')
              el-option(v-for='(op, index) in options', :key='index', :label='op.label', :value='op.value')
        el-col(:span='3')
          div 数组结构
          el-checkbox(v-model='item.is_array')
        el-col(:span='2')
          div.el-icon-delete(@click="onDelete(item, 2)" style="margin-top:30px")
    el-card.footer
      .clearfix(slot='header')
        span 表尾内容
        el-button(style='float: right; padding: 3px 5px', type='text' @click="onAddRange(3)") addRangeInRight
        el-button(style='float: right; padding: 3px 5px', type='text' @click="onAddRange(3,false)") addRangeInBottom
        el-button(style='float: right; padding: 3px 5px', type='text' @click="onAdd(3)") addInRight
        el-button(style='float: right; padding: 3px 5px', type='text' @click="onAdd(3,false)") addBottom
        el-button(style='float: right; padding: 3px 5px', type='text' @click="onClean(3)") Clean
      el-row(v-for='(item,index) in tfooter', :key='index' :gutter='2')
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
              el-option(v-for='(op,index) in options', :key='index', :label='op.label', :value='op.value')
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
    el-button(type="primary"  @click='getTest()') 获取sheet数据
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
  data() {
    return {
      DemoSpan : '',
      docName: '',
      theader:[],
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
      ],
      test_body: [
        {
            name:"序号",
            address:"A2",
            type:"text",
            col:1,
            row:2
        },
        {
            name:"产品图示",
            address:"B2",
            type:"image",
            col:2,
            row:2
        },
        {
            name:"商品名称",
            address:"C2",
            type:"text",
            col:3,
            row:2
        },
        {
            name:"型号",
            address:"D2",
            type:"text",
            col:4,
            row:2
        },
        {
            name:"尺寸(W*D*H标准)",
            address:"E2",
            type:"text",
            col:5,
            row:2
        },
        {
            name:"材质",
            address:"F2",
            type:"text",
            col:6,
            row:2
        },
        {
            name:"颜色",
            address:"G2",
            type:"text",
            col:7,
            row:2
        },
        {
            name:"数量",
            address:"H2",
            type:"text",
            col:8,
            row:2
        },
        {
            name:"单位",
            address:"I2",
            type:"text",
            col:9,
            row:2
        },
        {
            name:"价格",
            address:"J2",
            type:"text",
            col:10,
            row:2
        },
        {
            name:"工艺标准和要求",
            address:"K2",
            type:"text",
            col:11,
            row:2
        }
      ]
    }
  },
  methods:{
    onbuttonclick(id){
        return taskPane.onbuttonclick(id)
    },
    onDocNameClick(){
        this.docName = taskPane.onbuttonclick('getDocName')
    },
    onSelectionChange(worksheet, range) {
      console.log(worksheet, range)
      this.selectedCell = range.Address()
      this.selectedCellLocal = range.AddressLocal()
    },
    // 添加一个区域数据
    // isRight 填写的value单元格在右侧相邻 否在在下面相邻
    onAddRange(type, isRight = true) {
      let selection = this.currentSelection()
      console.log(selection.Count)
      for (let index = 1; index <= selection.Count; index++) {
        let cell =  selection.Cells.Item(index)
        if(cell.Text) {
          let nextCell = this.getNextCell(cell, isRight)
          if(type == 2) {
            nextCell = cell
          }
          let item = {
            name: cell.Text,
            address: nextCell.Address(false,false),
            type: 'text',
            col: nextCell.Column,
            row: nextCell.Row
          }
          if(item.name.indexOf('日期') != -1) {
            item.type = 'date'
          } else if(item.name.indexOf('图示') != -1 || item.name.indexOf('图片') != -1 ) {
            item.type = 'image'
          }
          
          if(type == 1) {
            this.theader.push(item)
          }  else if(type == 2) {
            this.tbody.push(item)
          } else if(type == 3) {
            this.tfooter.push(item)
          }
        }
      }
    },
    onAdd(type, isRight = true) {
      let cell = this.currentCell()
      let nextCell = this.getNextCell(cell, isRight)
      if(type == 2) {
        nextCell = cell
      }
      let item = {
        name: cell.Text,
        address: nextCell.Address(false,false),
        type: 'text',
        col: nextCell.Column,
        row: nextCell.Row
      }
      if(item.name.indexOf('日期') != -1) {
        item.type = 'date'
      } else if(item.name.indexOf('图示') != -1 || item.name.indexOf('图片') != -1 ) {
        item.type = 'image'
      }
      if(type == 1) {
          this.theader.push(item)
      }  else if(type == 2) {
          this.tbody.push(item)
      } else if(type == 3) {
          this.tfooter.push(item)
      }
    },
    getName(obj) {
      let cell = this.currentCell()
      obj.name = cell.Text
    },
    getAddress(obj) {
      let cell = this.currentCell()
      obj.address = cell.Address(false,false)
      obj.col = cell.Column,
      obj.Row = cell.Row
    },
    onClean(type) {
      if(type == 1) {
        this.theader = []
      }  else if(type == 2) {
          this.tbody= []
      } else if(type == 3) {
          this.tfooter= []
      }
    },
    onPosition(address) {
      let sheet = this.currentSheet()
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
      let result = {
        theader: this.theader,
        tfooter: this.tfooter,
        tbody: this.tbody,
        tother: this.tother
      }
      var data = JSON.stringify(result) 
      console.log(data)
      let fd = wps.EtApplication().FileDialog(2)
      fd.Filters.Add("文本", "*.json; *.txt;", 2)
      if(fd.Show() == -1){
        let path  = fd.SelectedItems.Item(1)
        wps.FileSystem.writeAsBinaryString (path , data)
      }
    },
    et() {
      return wps.EtApplication();
    },
    currentSheet() {
      return wps.EtApplication().ActiveSheet;
    },
    getNextCell(cell, isRight) {
      if(isRight) {
        if(!cell.MergeCells) {
          return cell.Next
        } else {
          let count = cell.MergeArea.Count
          return cell.MergeArea.Item(count).Next
        }
      } else {
        if(!cell.MergeCells) {
          return this.currentSheet().Cells.Item((cell.Row + 1),cell.Column)
        } else {
          let count = cell.MergeArea.Count
          let endCell = cell.MergeArea.Item(count)
          return this.currentSheet().Cells.Item((endCell.Row + 1),endCell.Column)
        }
      }
    },
    currentCell() {
      return wps.EtApplication().ActiveCell;
    },
    currentSelection() {
      return wps.EtApplication().Selection;
    },
    getTest() {
      let ranges = this.currentSelection()
      console.log(ranges.Columns.Count)
      console.log(ranges.Rows.Count)
      let data = []
      for (let i = 1; i <= ranges.Rows.Count; i++) {
        let  itemData = {}
        let add  = true
        for (let j = 0; j < this.test_body.length; j++) {
          let item = this.test_body[j]
          let cell = ranges.Item(i,item.col)
          if(cell.MergeCells) {
            let first = cell.MergeArea.Address().split(':')[0]
             if(j == 0) {
              itemData = data.find(p=>p.address == first)
              if(itemData) { 
                add = false
              }
              else {
                itemData = {}
              }
            }
            itemData[item.address] = first.Value2
          } else {
            itemData[item.address] = cell.Value2
          }
          if(add) {
            // itemData['address'] = 
            data.push(itemData)
          }
          console.log(cell.MergeArea.Address(false,false))
        }
      }
    }
  },
  mounted() {
      axios.get('/.debugTemp/NotifyDemoUrl').then((res) => {
          this.DemoSpan = res.data;
      });
      taskPane.initSelectionChange(this.onSelectionChange)
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
