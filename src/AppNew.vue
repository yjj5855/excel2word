<template>
  <div>
    <input
        type="file"
        accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
        @change="handleFileChange"/>

    <div v-if="showImagList" v-for="(image,index) in imageList" :key="index">
      <img :src="image.url" style="height: 150px;width: auto;" />
    </div>

    <button v-if="sheet" @click="download('docx')">下载{{sheet.title}}.docx</button>

    <button v-if="sheet" @click="download('ppt')">下载{{sheet.title}}.pptx</button>
  </div>
</template>

<script>
  import Excel from 'exceljs'
  import {Packer} from 'docx'
  import { DocumentCreator } from './libs/generator-timu'
  import {saveAs} from 'file-saver'
  import { PptCreator } from './libs/generator-ppt'

  export default {
    data () {
      return {
        sheet: null,
        blob: null,
        ppt: null,

        imageList: [],
        showImagList: false
      }
    },
    watch: {
    },
    methods: {
      handleFileChange (e) {
        let file = e.target.files && e.target.files[0]
        if (file) {
          this.readExcel(file)
            .then((sheet) => {
              this.sheet = sheet
              this.$nextTick(() => {
                this.showImagList = true
              })
              // Promise.all([
              //   this.write2Docx(sheet),
              //   this.write2Ppt(sheet)
              // ]).then(([blob, ppt]) => {
              //   this.blob = blob
              //   this.ppt = ppt
              // }).catch(err => {
              //   this.blob = null
              //   this.ppt = null
              // })
            })
            .catch((err) => {
              this.sheet = null
            })
        }
      },
      readExcel (file) {
        let self = this
        return new Promise(async (resolve, reject) => {
          const workbook = new Excel.Workbook()
          let data = readFile(file)
          await workbook.xlsx.load(data)
          // console.log(workbook)

          let sheets = []
          workbook.eachSheet(async function(worksheet, sheetId) {
            let sheet = {
              sheetName: worksheet.name,
              title: '',
              timuList: []
            }
            let allKeyList = []
            // 数据
            worksheet.eachRow(function(row, index) {
              if (index === 1) {
                allKeyList = getHeaderKeyList(row)
              }
              if (index === 2) {
                sheet.title = row.values[1]
              } else if (index > 2) {
                let data = formatObj(row, allKeyList)
                if (data) {
                  sheet.timuList.push(data)
                }
              }
            })

            // 图片
            worksheet.getImages().map(async function (image) {
              //  image.range.tl.nativeRow, 'col', image.range.tl.nativeCol, 'imageId', image.imageId
              // console.log('processing image row', image.range.tl.nativeRow, 'col', image.range.tl.nativeCol, 'imageId', image)
              // 把图片添加到timuList对象中
              let row = parseInt(image.range.tl.nativeRow) + 1
              let col = parseInt(image.range.tl.nativeCol) + 1
              let timu = sheet.timuList.find(timu => {
                // console.log(worksheet.getRow(row))
                let obj = formatObj(worksheet.getRow(row), allKeyList)
                return timu['总序'] === obj['总序']
              })
              const img = workbook.model.media.find(m => m.index === image.imageId)
              let url = window.URL.createObjectURL(new Blob([img.buffer]))
              img.url = url
              try {
                let size = await getImageSize(url)
                img.width = size.width
                img.height = size.height
              } catch (err) {
                console.error(err)
              }
              if (timu) {
                if (col <= 10) {
                  timu['题目_image'] = img
                } else {
                  timu['解析_image'] = img
                }
                // self.imageList.push(img)
              } else {
                self.imageList.push(img)
                console.log(`未找到行 ${row} ${col}`, image, img)
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
      },
      write2Docx (sheet) {
        return new Promise((resolve, reject) => {
          // 题目列表转换成word文档
          const documentCreator = new DocumentCreator()
          const doc = documentCreator.create(sheet)
          Packer.toBlob(doc)
            .then((blob) => {
              resolve(blob)
            })
            .catch(err =>{
              console.error(err)
              reject()
            })
        })
      },
      write2Ppt (sheet) {
        return new Promise((resolve, reject) => {
          const pptCreator = new PptCreator()
          const ppt = pptCreator.create(sheet)
          if (ppt) {
            resolve(ppt)
          } else {
            reject()
          }
        })
      },
      download (fileType = 'all') {
        if (fileType === 'all' || fileType === 'docx') {
          this.write2Docx(this.sheet).then(blob => {
            this.blob = blob
            saveAs(this.blob, this.sheet.title + '.docx')
          })
        }
        if (fileType === 'all' || fileType === 'ppt') {
          this.write2Ppt(this.sheet).then(ppt => {
            this.ppt = ppt
            this.ppt.writeFile({ fileName: `${this.sheet.title}.pptx` })
          })
        }
      }
    }
  }

  function readFile (file) {
    return new Promise((resolve, reject) => {
      let reader = new FileReader()
      reader.onload = function(e) {
        let data = e.target.result
        // 读取二进制的excel
        resolve(data)
      }
      reader.readAsArrayBuffer(file)
    })
  }

  function getHeaderKeyList (row) {
    let keyList = []
    keyList = row.values.map((key,index) => {
      return index === 14 ? key + '_1' : key
    })
    return keyList
  }
  function formatObj (row, keyList) {
    let obj = {}
    row.values.forEach((value, index) => {
      obj[keyList[index]] = value
    })

    if (typeof obj['总序'] === 'string' && obj['总序'].length > 10) {
        return null
    } else {
      if (typeof obj['总序'] === 'string') {
        obj['总序'] = Number(obj['总序'].replace(/\s+/g,''))
      }
    }
    return obj
  }
  function getImageSize(url) {
    return new Promise(function (resolve, reject) {
      let image = new Image();
      image.onload = function () {
        resolve({
          width: image.width,
          height: image.height
        });
      };
      image.onerror = function () {
        reject(new Error('error'));
      };
      image.src = url;
      // document.body.append(image)
    });
  }
</script>

<style scoped>
</style>
