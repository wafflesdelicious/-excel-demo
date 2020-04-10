<template>
  <div class="home">

    <el-form label-width="80px">
      <el-form-item label="URLs:">
        <el-input
          type="textarea"
          :rows="2"
          placeholder="请输入内容"
          v-model="textarea"
          class="input"
          autosize
        ></el-input>
      </el-form-item>
      <el-form-item >
         <el-upload
          class="upload-demo"
          action
          :on-change="handleChange"
          accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
          :auto-upload="false"
          :show-file-list="false"
        >
          <!-- :on-change="handleChange"
          :on-remove="handleRemove"
          :on-exceed="handleExceed"
          :limit="limitUpload" -->
          <!-- <div slot="tip" class="el-upload__tip">只 能 上 传 xlsx / xls 文 件</div> -->
          <el-button size="small" type="primary">点击上传(excel)</el-button>
        </el-upload>
      </el-form-item>
    </el-form>
    <el-table
      :data="arr"
      style="width: 100%"
      ref="table"
      id="table">
      <el-table-column prop="name" label="姓名" width="180"></el-table-column>
      <el-table-column prop="address" label="地址"></el-table-column>
    </el-table>
    <el-button @click="outtable">导出为excel文件</el-button>
  </div>
</template>

<script>
// @ is an alias to /src
// import HelloWorld from '@/components/HelloWorld.vue'
import XLSX from 'xlsx'
import FileSaver from 'file-saver'
export default {
  data () {
    return {
      textarea: '',
      workbook: '',
      arr: []
    }
  },
  methods: {
    handleChange (file, fileList) {
      this.fileTemp = file.raw
      // 判断上传文件是否符合要求
      if (this.fileTemp) {
        if (
          this.fileTemp.type ===
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
          this.fileTemp.type === 'application/vnd.ms-excel'
        ) {
          this.importfxx(this.fileTemp)
        } else {
          this.$message.warning('附件格式错误，请删除后重新上传！')
        }
      } else {
        this.$message.warning('请上传附件！')
      }
    },
    // 导入
    importfxx (obj) {
      // console.log(event.currentTarget.files)
      const _this = this
      this.file = event.currentTarget.files[0]
      var reader = new FileReader()
      reader.onload = function (e) {
        // console.log(e)
        var data = e.target.result
        // 读取excel文件
        var workbook = XLSX.read(data, { type: 'binary' })
        // console.log(workbook)

        var sheetNames = workbook.SheetNames // 工作表名称集合
        var worksheet = workbook.Sheets[sheetNames[0]] // 这里我们只读取第一张sheet
        // console.log(worksheet)

        // XLSX.utils.sheet_to_csv：生成CSV格式
        // XLSX.utils.sheet_to_txt：生成纯文本格式
        // XLSX.utils.sheet_to_html：生成HTML格式
        // XLSX.utils.sheet_to_json：输出JSON格式

        // input文本
        var outdata = XLSX.utils.sheet_to_csv(worksheet)
        outdata = outdata.replace(/,/g, ' ')
        // console.log(outdata)
        _this.textarea = outdata

        // 表格
        var outdata2 = XLSX.utils.sheet_to_json(worksheet)
        console.log(outdata2)
        outdata2.map(v => {
          const obj = {}
          obj.name = v.Name
          obj.address = v.code
          _this.arr.push(obj)
        })
      }
      reader.readAsBinaryString(this.file)
    },
    // 导出
    // aoa_to_sheet: 将一个二维数组转成sheet，会自动处理number、string、boolean、date等类型数据；
    // table_to_sheet: 将一个table dom直接转成sheet，会自动识别colspan和rowspan并将其转成对应的单元格合并；
    // json_to_sheet: 将一个由对象组成的数组转成sheet；
    outtable () {
      var xlsxParam = { raw: true } // 转换成excel时，使用原始的格式
      var elt = document.getElementById('table')
      var sheet = XLSX.utils.table_to_book(elt, xlsxParam)
      console.log(sheet)

      var wbout = XLSX.write(sheet, { bookType: 'xlsx', bookSST: true, type: 'array' }) // 生成同一个excel工作簿
      try {
        // 保存文件
        FileSaver.saveAs(new Blob([wbout], { type: 'application/octet-stream' }), '导出的表格.xlsx')
      } catch (e) {
        if (typeof console !== 'undefined') {
          console.log(e, wbout)
        }
      }
      return wbout
    }
  }
}
</script>
