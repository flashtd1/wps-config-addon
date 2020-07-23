<template lang='pug'>
  div
    el-button( type="primary" @click="onAdd(3,false)" style="margin-bottom:20px;") 添加选中元素
    el-button( type="primary" @click="layout=[]" style="margin-bottom:20px;") 清空
    el-button( type="primary" @click="onDelete" style="margin-bottom:20px;") 删除选中
    el-button( type="primary" @click="onSubmit" style="margin-bottom:20px;") 生成
    grid-layout(
      :layout.sync="layout"
      :col-num="col_num"
      :row-height="row_height"
      :is-draggable="true"
      :is-resizable="true"
      :is-mirrored="false"
      :vertical-compact="true"
      :margin="margin"
      :use-css-transforms="true"
     )
        grid-item( v-for="item in layout"
          :x="item.x"
          :y="item.y"
          :w="item.w"
          :h="item.h" 
          :i="item.i"
          :key="item.i")
          div
            | 组件： 
            el-select(v-model='item.type', placeholder='请选择' style="width:100px;")
              el-option(v-for='(op,index) in options', :key='index', :label='op.label', :value='op.value')
          div
            | 名称： 
            span {{item.name}}
          div
            | span: 
            | {{item.w}}
          div 
            | key： 
            el-input(v-model="item.key" placeholder="请输入内容" style="width:100px;")
</template>

<script>
 import VueGridLayout from 'vue-grid-layout';
export default {
  components: {
      GridLayout: VueGridLayout.GridLayout,
      GridItem: VueGridLayout.GridItem
    },
  name: 'PageViews',
  data(){
    return {
      col_num:24,
      row_height: 60,
      margin: [10,10],
      span: 24,
      options:[
        {
          value: 'el-input',
          label: '单行文本'
        },
        {
          value: 'el-date-picker',
          label: '时间'
        },
        {
          value: 'Uploader',
          label: '图片'
        },
         {
          value: 'mtext',
          label: '多行文本'
        },
      ],
      layout: [
      //  {"x":0,"y":0,"w":2,"h":2,"i":"0"},
        // {"x":3,"y":0,"w":2,"h":2,"i":"1"},
        // {"x":4,"y":0,"w":2,"h":5,"i":"2"},
        // {"x":6,"y":0,"w":2,"h":3,"i":"3"},
        // {"x":8,"y":0,"w":2,"h":3,"i":"4"},
        // {"x":10,"y":0,"w":2,"h":3,"i":"5"},
        // {"x":0,"y":5,"w":2,"h":5,"i":"6"},
        // {"x":2,"y":5,"w":2,"h":5,"i":"7"},
        // {"x":4,"y":5,"w":2,"h":5,"i":"8"},
        // {"x":6,"y":3,"w":2,"h":4,"i":"9"},
        // {"x":8,"y":4,"w":2,"h":4,"i":"10"},
        // {"x":10,"y":4,"w":2,"h":4,"i":"11"},
        // {"x":0,"y":10,"w":2,"h":5,"i":"12"},
        // {"x":2,"y":10,"w":2,"h":5,"i":"13"},
        // {"x":4,"y":8,"w":2,"h":4,"i":"14"},
        // {"x":6,"y":8,"w":2,"h":4,"i":"15"},
        // {"x":8,"y":10,"w":2,"h":5,"i":"16"},
        // {"x":10,"y":4,"w":2,"h":2,"i":"17"},
        // {"x":0,"y":9,"w":2,"h":3,"i":"18"},
        // {"x":2,"y":6,"w":2,"h":2,"i":"19"}
      ]
    }
  },
  methods:{
    onAdd() {
      let selection = this.currentSelection()
      console.log(selection.Count)
      for (let index = 1; index <= selection.Count; index++) {
        let cell =  selection.Cells.Item(index)
        if(cell.Text) {
          let ys =  this.layout.map(({y})=>{
            return y
          })
          let maxY = Math.max(...ys)
          let temp = {
            x: 0,
            y: maxY,
            w:this.span,
            h:2,
            i: this.layout.length,
            name: cell.Text,
            type: 'el-input',
            key: cell.Address(false,false)
          }
          // if(temp.name.indexOf('日期') != -1) {
          //   temp.type = 'date'
          // } else if(temp.name.indexOf('图示') != -1 || temp.name.indexOf('图片') != -1 ) {
          //   temp.type = 'image'
          // }
          this.layout.push(temp)
        }
      }
      // let cell = this.currentCell()
      // let address = cell.Address(false,false)
      // console.log(address,cell.Value2)
      // let ys =  this.layout.map(({y})=>{
      //   return y
      // })

      // let maxY = Math.max(...ys)
      // console.log(maxY)
      // let temp = {
      //   x: 0,
      //   y: maxY,
      //   w:2,
      //   h:2,
      //   i: this.layout.length,
      //   name: cell.Value2
      // }
      // this.layout.push(temp)
    },
    onDelete() {

    },
    onSubmit() {
      let result = this.layout.map( (item) => {
        let temp = {
          name: item.name,
          component: item.type,
          span: item.w,
          key: item.key
        }
        if(item.type == 'mtext') {
          temp.component = 'el-input'
          temp.attrs = {
            type: 'textarea',
          }
        }
        return temp
      })
      let data = JSON.stringify(result)
      console.log(data) 

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
  },
  mounted() {
   
  }
}
</script>


<style scoped>
/*** EXAMPLE ***/
#content {
    width: 100%;
}

.vue-grid-layout {
    background: #eee;
}

.layoutJSON {
    background: #ddd;
    border: 1px solid black;
    margin-top: 10px;
    padding: 10px;
}

.eventsJSON {
    background: #ddd;
    border: 1px solid black;
    margin-top: 10px;
    padding: 10px;
    height: 100px;
    overflow-y: scroll;
}

.columns {
    -moz-columns: 120px;
    -webkit-columns: 120px;
    columns: 120px;
}



/*.vue-resizable-handle {
    z-index: 5000;
    position: absolute;
    width: 20px;
    height: 20px;
    bottom: 0;
    right: 0;
    background: url('data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBzdGFuZGFsb25lPSJubyI/Pg08IS0tIEdlbmVyYXRvcjogQWRvYmUgRmlyZXdvcmtzIENTNiwgRXhwb3J0IFNWRyBFeHRlbnNpb24gYnkgQWFyb24gQmVhbGwgKGh0dHA6Ly9maXJld29ya3MuYWJlYWxsLmNvbSkgLiBWZXJzaW9uOiAwLjYuMSAgLS0+DTwhRE9DVFlQRSBzdmcgUFVCTElDICItLy9XM0MvL0RURCBTVkcgMS4xLy9FTiIgImh0dHA6Ly93d3cudzMub3JnL0dyYXBoaWNzL1NWRy8xLjEvRFREL3N2ZzExLmR0ZCI+DTxzdmcgaWQ9IlVudGl0bGVkLVBhZ2UlMjAxIiB2aWV3Qm94PSIwIDAgNiA2IiBzdHlsZT0iYmFja2dyb3VuZC1jb2xvcjojZmZmZmZmMDAiIHZlcnNpb249IjEuMSINCXhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgeG1sbnM6eGxpbms9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkveGxpbmsiIHhtbDpzcGFjZT0icHJlc2VydmUiDQl4PSIwcHgiIHk9IjBweCIgd2lkdGg9IjZweCIgaGVpZ2h0PSI2cHgiDT4NCTxnIG9wYWNpdHk9IjAuMzAyIj4NCQk8cGF0aCBkPSJNIDYgNiBMIDAgNiBMIDAgNC4yIEwgNCA0LjIgTCA0LjIgNC4yIEwgNC4yIDAgTCA2IDAgTCA2IDYgTCA2IDYgWiIgZmlsbD0iIzAwMDAwMCIvPg0JPC9nPg08L3N2Zz4=');
    background-position: bottom right;
    padding: 0 3px 3px 0;
    background-repeat: no-repeat;
    background-origin: content-box;
    box-sizing: border-box;
    cursor: se-resize;
}*/

.vue-grid-item:not(.vue-grid-placeholder) {
    background: #ccc;
    border: 1px solid black;
}

.vue-grid-item.resizing {
    opacity: 0.9;
}

.vue-grid-item.static {
    background: #cce;
}

.vue-grid-item .text {
    font-size: 24px;
    text-align: center;
    position: absolute;
    top: 0;
    bottom: 0;
    left: 0;
    right: 0;
    margin: auto;
    height: 100%;
    width: 100%;
}

.vue-grid-item .no-drag {
    height: 100%;
    width: 100%;
}

.vue-grid-item .minMax {
    font-size: 12px;
}

.vue-grid-item .add {
    cursor: pointer;
}

.vue-draggable-handle {
    position: absolute;
    width: 20px;
    height: 20px;
    top: 0;
    left: 0;
    background: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><circle cx='5' cy='5' r='5' fill='#999999'/></svg>") no-repeat;
    background-position: bottom right;
    padding: 0 8px 8px 0;
    background-repeat: no-repeat;
    background-origin: content-box;
    box-sizing: border-box;
    cursor: pointer;
}

</style>