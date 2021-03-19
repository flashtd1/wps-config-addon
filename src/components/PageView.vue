<template lang='pug'>
  div
    el-button( type="primary" @click="setHeader()" style="margin-bottom:20px;") 添加表头
    el-button( type="primary" @click="onAdd()" style="margin-bottom:20px;") 添加选中元素
    el-button( type="primary" @click="layout=[]" style="margin-bottom:20px;") 清空
    el-button( type="primary" @click="onDelete" style="margin-bottom:20px;") 删除选中
    el-button( type="primary" @click="onSubmit" style="margin-bottom:20px;") 生成
    div
      | 表头
      pre {{header}}
    div
      | 结果
      el-input(v-model="result" type="textarea" :row="5")
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
            | 名称： 
            span {{item.name}}
          div 
            span
              | key： 
              el-input(v-model="item.key" size="mini" placeholder="请输入内容" style="width:100px;")
            span
              | 组件
              el-input(v-model="item.component" size="mini" placeholder="请输入内容" style="width:100px;")
          div
            | span: 
            el-input(v-model="item.w" size="mini" placeholder="请输入内容" style="width:100px;")
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
      result: '',
      col_num:24,
      row_height: 60,
      margin: [10,10],
      span: 24,
      header: [],
      layout: [
      //  {"x":0,"y":0,"w":2,"h":2,"i":"0"},
        // {"x":3,"y":0,"w":2,"h":2,"i":"1"},
      ]
    }
  },
  methods:{
    setHeader () {
      this.header = []
      let selection = this.currentSelection()
      for (let index = 0; index < selection.Count; index++) {
        let prop = {
          key: selection.Cells.Item(index + 1).Text,
          column: index
        }
        this.header.push(prop)
      }
    },
    onAdd() {
      let selection = this.currentSelection()
      let row = 0
      let xSum = 0
      let spanColumn = this.header.find(h => h.key == 'span')
      for (let index = 1; index <= selection.Count; index++) {
        let cell =  selection.Cells.Item(index)
        if(cell.Text) {
          // 计算x的位置，如果超出一行大小，换行，x重置
          let span = spanColumn ? parseInt(cell.Offset(0, spanColumn.column).Text) : this.span
          if (xSum + span > this.span) {
            row++
            xSum = 0
          }
          let temp = {
            x: xSum,
            y: row,
            w: span,
            h:2,
            i: this.layout.length
          }
          xSum += span
          let extendProps = this.header.reduce((result, current) => {
            result[current.key] = cell.Offset(0, current.column).Text
            return result
          }, {})
          Object.assign(temp, extendProps)
          this.layout.push(temp)
        }
      }
    },
    onDelete() {

    },
    onSubmit() {
      let result = this.layout.map((item) => {
        let temp = {
          name: item.name,
          component: item.component,
          span: item.span,
          key: item.key,
        }
        // item中带有:的key，需要层级创建对象，在最后一级才进行赋值
        // 先找出所有带：的属性
        let attrProps = Object.keys(item).filter(i => i.includes(':'))
        // 赋值
        // 先构造temp的层级，并在正确的层级上赋值
        for (let i = 0; i < attrProps.length; i++) {
          let levels = attrProps[i].split(':')
          levels.reduce((result, current, index) => {
              if (!result[current]) {
                  if (index == levels.length - 1) {
                    if (item[attrProps[i]]) {
                      result[current] = item[attrProps[i]]
                    }
                  } else {
                    result[current] = {}
                  }
                  return result[current]
              }
              return result[current]
          }, temp)
        }
        // 旧逻辑
        // let attrProps = Object.keys(item).filter(i => i.startsWith('attr_'))
        // attrProps.map((prop) => {
        //   let key = prop.split('attr_')[1]
        //   if (item[prop]) {
        //     Object.assign(temp.attrs, {
        //       [key]: item[prop]
        //     })
        //   }
        // })
        return temp
      })
      let data = JSON.stringify(result)
      console.log(data) 
      this.result = data
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
    }
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