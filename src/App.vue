<template>
  <div>
    <input
        type="file"
        accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
        @change="handleFileChange"/>

    <button v-if="sheet" @click="download">下载{{sheet.title}}docx</button>
  </div>
</template>

<script>
  import * as XLSX from 'xlsx'
  import {Document, Packer, TextRun} from 'docx'
  import xlsx from './libs/xlsx'
  import { DocumentCreator } from './libs/generator-timu'
  import {saveAs} from 'file-saver'

  export default {
    data () {
      return {
        sheet: null,
        blob: null
      }
    },
    watch: {
    },
    methods: {
      handleFileChange (e) {
        let file = e.target.files && e.target.files[0]
        if (file) {
          this.readExcel(file)
            .then(this.write2Docx)
            .then(({blob, sheet}) => {
              console.log(blob, sheet)
              this.blob = blob
              this.sheet = sheet
            })
            .catch((err) => {
              console.error(err)
              this.blob = null
              this.sheet = null
            })
        }
      },
      readExcel (file) {
        return new Promise((resolve, reject) => {
          xlsx.readWorkbookFromLocalFile(file, wb => {
            // 每一行 转换为 一道题目
            let sheets = []
            wb.SheetNames.forEach(sheetName => {
              let sheet = {
                sheetName: sheetName,
                title: '',
                timuList: []
              }
              let xlsxData = XLSX.utils.sheet_to_json(wb.Sheets[sheetName])
              let allKeyList = xlsx.getHeaderKeyList(wb.Sheets[sheetName])
              xlsxData.forEach((data, index)=> {
                if (index >= 1) {
                  if (typeof data['总序'] === 'string') {

                  } else {
                    sheet.timuList.push(data)
                  }
                } else if (index === 0) {
                  sheet.title = data[allKeyList[0]]
                }
              })
              sheets.push(sheet)
            })
            if (sheets.length >= 1) {
              resolve(sheets[0])
            } else {
              reject()
            }
          })
        })
      },
      write2Docx (sheet) {
        return new Promise((resolve, reject) => {
          // 题目列表转换成word文档
          const documentCreator = new DocumentCreator()
          const doc = documentCreator.create(sheet)
          Packer.toBlob(doc)
            .then((blob) => {
              resolve({
                blob,
                sheet
              })
            })
            .catch(err =>{
              console.error(err)
              reject()
            })
        })
      },
      download () {
        saveAs(this.blob, this.sheet.title + '.docx')
      }
    }
  }
</script>

<style scoped>
header {
  line-height: 1.5;
}

.logo {
  display: block;
  margin: 0 auto 2rem;
}

@media (min-width: 1024px) {
  header {
    display: flex;
    place-items: center;
    padding-right: calc(var(--section-gap) / 2);
  }

  .logo {
    margin: 0 2rem 0 0;
  }

  header .wrapper {
    display: flex;
    place-items: flex-start;
    flex-wrap: wrap;
  }
}
</style>
