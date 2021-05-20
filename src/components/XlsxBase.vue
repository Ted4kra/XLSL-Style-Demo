<template>
  <div>
    <el-card>
      <el-table
        size="mini"
        :data="tableData"
      >
        <el-table-column
          prop="shipname"
          label="船名"
        />
        <el-table-column
          prop="vgno"
          label="航次"
        />
        <el-table-column
          prop="blno"
          label="提单号"
        />
        <el-table-column
          prop="ctnno"
          label="箱号"
        />
        <el-table-column
          prop="sealno"
          label="铅封号"
        />
        <el-table-column
          prop="ctntypename"
          label="箱型"
        />
        <el-table-column
          prop="ctnweight"
          label="重量"
        />
        <el-table-column
          prop="carriage"
          label="托运费"
        />
      </el-table>
      <el-button
        style="margin-top: 20px"
        @click="exportExcel"
      >
        导出Excel文件
      </el-button>
    </el-card>
    <el-card style="margin-top: 20px">
      <input
        class="input-file"
        type="file"
        accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
        @change="importData"
      >
      <el-button @click="importAction">
        导入Excel文件
      </el-button>
    </el-card>
  </div>
</template>

<script>
import XLSX from 'XLSX'
import FileSaver from 'file-saver'
export default {
  name: "XlsxBase",
  data() {
    return {
      jsonData: "[{\"shipname\":\"天运河\",\"vgno\":\"175N\",\"blno\":\"PASUQ018843350\",\"ctnno\":\"HCIU2013578\",\"sealno\":\"192490\",\"ctntypename\":\"20GP\",\"ctnweight\":\"28\",\"carriage\":\"200\"},{\"shipname\":\"天运河\",\"vgno\":\"175N\",\"blno\":\"PASUQ018843350\",\"ctnno\":\"HGMU1100660\",\"sealno\":\"002412\",\"ctntypename\":\"20GP\",\"ctnweight\":\"28\",\"carriage\":\"200\"},{\"shipname\":\"天运河\",\"vgno\":\"175N\",\"blno\":\"PASUQ018843350\",\"ctnno\":\"HJLU1224112\",\"sealno\":\"001510\",\"ctntypename\":\"20GP\",\"ctnweight\":\"28\",\"carriage\":\"200\"},{\"shipname\":\"天运河\",\"vgno\":\"175N\",\"blno\":\"PASUQ018843350\",\"ctnno\":\"HJLU1444354\",\"sealno\":\"001509\",\"ctntypename\":\"20GP\",\"ctnweight\":\"28\",\"carriage\":\"200\"},{\"shipname\":\"天运河\",\"vgno\":\"175N\",\"blno\":\"PASUQ018843350\",\"ctnno\":\"TEMU5766268\",\"sealno\":\"192489\",\"ctntypename\":\"20GP\",\"ctnweight\":\"28\",\"carriage\":\"200\"},{\"shipname\":\"昌盛集6\",\"vgno\":\"150S\",\"blno\":\"PASUD018881380\",\"ctnno\":\"TEMU5766458\",\"sealno\":\"007856\",\"ctntypename\":\"20GP\",\"ctnweight\":\"28\",\"carriage\":\"200\"},{\"shipname\":\"昌盛集6\",\"vgno\":\"150S\",\"blno\":\"PASUD018881380\",\"ctnno\":\"TEMU5898325\",\"sealno\":\"007916\",\"ctntypename\":\"20GP\",\"ctnweight\":\"28\",\"carriage\":\"200\"},{\"shipname\":\"昌盛集6\",\"vgno\":\"150S\",\"blno\":\"PASUD018881380\",\"ctnno\":\"YGMU2119056\",\"sealno\":\"000879\",\"ctntypename\":\"20GP\",\"ctnweight\":\"28\",\"carriage\":\"200\"},{\"shipname\":\"昌盛集6\",\"vgno\":\"150S\",\"blno\":\"PASUD018881380\",\"ctnno\":\"YGMU2131365\",\"sealno\":\"002411\",\"ctntypename\":\"20GP\",\"ctnweight\":\"28\",\"carriage\":\"200\"},{\"shipname\":\"昌盛集6\",\"vgno\":\"150S\",\"blno\":\"PASUD018881380\",\"ctnno\":\"YGMU2144017\",\"sealno\":\"000880\",\"ctntypename\":\"20GP\",\"ctnweight\":\"28\",\"carriage\":\"200\"}]",
      tableData: '',
      defaultCellStyle: {
        font: {
          name: '宋体',
          sz: 12,
          color: { auto: 1 }
        },
        border: {},
        alignment: {
          wrapText: 1,
          horizontal: 'center',
          vertical: 'center',
          indent: 0
        }
      },
    }
  },
  created() {
    this.tableData = JSON.parse(this.jsonData)
  },
  methods: {
    importAction() {
      document.querySelector('.input-file').click()
    },
    importData (event) {
      if (!event.currentTarget.files.length) {
        return
      }
      const file = event.currentTarget.files[0]

      const reader = new FileReader()
      reader.readAsBinaryString(file);
      reader.onload = function (e) {
        const data = e.target.result;
        const workBook = XLSX.read(data, {type: 'binary'});
        const options = {
          // 自定义key, 默认是第一行为key
          // header: ["shipname", "vgno", "blno", "ctnno", "sealno", "ctntypename", "ctnweight", "carriage"],
          // 从第二行开始读取 https://github.com/SheetJS/sheetjs/issues/482
          range: 0, // 1,
          // 空白不不读取
          blankrows: false,
          // 默认值，如果是blank，会使用这个默认值
          defval: "",
          // 原始值还是加工过的值
          raw: false
        }
        const excelJson = XLSX.utils.sheet_to_json(workBook.Sheets[workBook.SheetNames[0]], options);
        console.log(excelJson)
      }
    },
    // 导出
    exportExcel() {
      // 表头信息
      let aoa = [
        ['船名',	'航次', '提单号', '箱号','铅封号','箱型', '重量', '托运费']
      ]
      const totalForm = this.tableData.map(item => {
        return [
            item.shipname,
            item.vgno,
            item.blno,
            item.ctnno,
            item.sealno,
            item.ctntypename,
            item.ctnweight,
            item.carriage
        ]
      })
      aoa = aoa.concat(totalForm)
      console.log(aoa)
      // array to sheet
      const sheet = this.sheet_from_array_of_arrays(aoa)
      /// 设置样式
      // 1.设置某个单元格的样式
      sheet['A2'].s = {
        font: {
          name: '宋体',
          sz: 24,
          bold: true,
          color: { rgb: '000000' }
        },
        alignment: {
          horizontal: 'center',
          vertical: 'center',
          wrap_text: 'true'
        }
      }

      // 2.冻结
      sheet["!freeze"] = {
        xSplit: "1",
        ySplit: "1",
        // 坐上角是哪个cell
        topLeftCell: "B2",
        activePane: "bottomRight",
        state: "frozen",
      }
      // 3. 列宽
      /// wch不是像素宽度！！！这个数值具体换算不太清楚，在Excel里拉动的时候会显示
      /// https://docs.sheetjs.com/#column-properties 这里有解释
      sheet['!cols'] = [
        { wch: 10} ,
        { wch: 20 },
        { wch: 30 },
        { wch: 40 },
      ];

      const wbBlob = this.sheet2blob(sheet, 'exportByXLSXStyle')
      // 保存下载
      FileSaver.saveAs(wbBlob, 'exportByXLSXStyle.xlsx')
    },
    // 将一个sheet转成最终的excel文件的blob对象，然后利用URL.createObjectURL下载
    sheet2blob(sheet, sheetName) {
      sheetName = sheetName || 'sheet1'
      const workbook = {
        SheetNames: [sheetName],
        Sheets: {}
      }
      workbook.Sheets[sheetName] = sheet
      // 生成excel的配置项
      const wopts = {
        bookType: 'xlsx', // 要生成的文件类型
        bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
        type: 'binary'
      }

      const wbOut = XLSX.write(workbook, wopts, { defaultCellStyle: this.defaultCellStyle })

      // 字符串转ArrayBuffer
      function s2ab(s) {
        const buf = new ArrayBuffer(s.length)
        const view = new Uint8Array(buf)
        for (let i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF
        return buf
      }

      return new Blob([s2ab(wbOut)], { type: 'application/octet-stream' })
    },
    // 从json转化为sheet
    sheet_from_array_of_arrays(data) {
      const ws = {}
      const range = {
        s: {
          c: 10000000,
          r: 10000000
        },
        e: {
          c: 0,
          r: 0
        }
      }
      for (let R = 0; R !== data.length; ++R) {
        for (let C = 0; C !== data[R].length; ++C) {
          if (range.s.r > R) range.s.r = R
          if (range.s.c > C) range.s.c = C
          if (range.e.r < R) range.e.r = R
          if (range.e.c < C) range.e.c = C
          const cell = {
            v: data[R][C],
            s: this.defaultCellStyle
          }
          if (cell.v == null) continue
          const cell_ref = XLSX.utils.encode_cell({
            c: C,
            r: R
          })

          /* TEST: proper cell types and value handling */
          if (typeof cell.v === 'number') {
            cell.t = 'n'
          } else if (typeof cell.v === 'boolean') {
            cell.t = 'b'
          } else if (cell.v instanceof Date) {
            cell.t = 'n'
            cell.z = XLSX.SSF._table[14]
            cell.v = this.dateNum(cell.v)
          } else {
            cell.t = 's'
          }
          ws[cell_ref] = cell
        }
      }

      /* TEST: proper range */
      if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range)
      return ws
    },
    /* TODO: date1904 logic */
    dateNum(v, date1904) {
      if (date1904) v += 1462
      const epoch = Date.parse(v)
      return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000)
    },
    validValue(value) {
      return value || '/'
    }
  }
}
</script>

<style
    scoped>

</style>
