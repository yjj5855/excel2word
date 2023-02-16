import pptxgen from 'pptxgenjs'

const topRect = {x: '10%', y: '10%', w: '80%', h: '30%'}
// 1 = 160px
const ipx = 160
const logoRect = {x: '2%', y: 0, w: 261/ipx, h: 87/ipx}

const backgroundColor = 'ffffff'
const textColor = '333333'
const textFontSize = 14
const daanColor = 'ff0000'
export class PptCreator {
  create (sheet) {
    let timuList = sheet.timuList.sort((a, b) => a['总序'] < b['总序'])

    let danxuanList = timuList.filter(item => item['题型'] === '单选')
    let duoxuanList = timuList.filter(item => item['题型'] === '多选')
    const pageLength = timuList.length * 2

    const ppt = new pptxgen()
    ppt.layout = 'LAYOUT_16x9'
    ppt.author = '上海建工';
    ppt.company = '上海建工';
    ppt.revision = '1';
    ppt.subject = sheet.title
    ppt.title = sheet.sheetName

    ppt.defineSlideMaster({
      title: 'MASTER_SLIDE',
      background: { fill: backgroundColor },
      // 配置全局样式
      objects: [
        {
          image: {
            ...logoRect,
            path: '/logo-jiangong.png',
            sizing: {
              type: 'cover'
            }
          }
        }
        // {
        //   text: {
        //     text: 'ajsjf',
        //     options: {
        //       ...topRect,
        //       fontSize: textFontSize,
        //       valign: 'top',
        //       align: 'right',
        //       fontFace: '微软雅黑',
        //       color: textColor
        //     }
        //   }
        // }
      ]
      // slideNumber: { x: 6.5, y: 12, fontSize: 10 }
    })

    danxuanList.map(timu => {
      // 第一页选项
      let slide = ppt.addSlide({masterName: 'MASTER_SLIDE'})
      addTimu(slide, timu)
      addXuanxiang(slide, timu)

      // 第二页 答案和解析
      let slide2 = ppt.addSlide({masterName: 'MASTER_SLIDE'})
      addJieda(slide2, timu)
    })

    duoxuanList.map(timu => {
      // 第一页选项
      let slide = ppt.addSlide({masterName: 'MASTER_SLIDE'})
      addTimu(slide, timu)
      addXuanxiang(slide, timu)

      // 第二页 答案和解析
      let slide2 = ppt.addSlide({masterName: 'MASTER_SLIDE'})
      addJieda(slide2, timu)
    })

    return ppt
  }
}

function addTimu (slide, timu) {
  slide.addText(`${timu['总序']}. ${timu['题目']}`, {
    ...topRect,
    fontSize: textFontSize,
    valign: 'top',
    align: 'left',
    fontFace: '微软雅黑',
    color: textColor
  })
  return slide
}

const bottomRect = {x: '10%', y: '40%', w: '80%', h: '40%'}
function addXuanxiang (slide, timu) {
  const row = [] // 创建数组
  const border = [ // 表格边框
    { type: 'none' },
    { type: 'none' },
    { type: 'none' },
    { type: 'none' }
  ]
  const options = { valign: 'middle', border: border }	//单元格样式配置
  row.push([{ text: `A.${timu['A']}`, options }])
  row.push([{ text: `B.${timu['B']}`, options }])
  row.push([{ text: `C.${timu['C']}`, options }])
  row.push([{ text: `D.${timu['D']}`, options }])

  if (timu['题型'] === '多选') {
    row.push([{ text: `E.${timu['E']}`, options }])
  }
  slide.addTable(row, {
    ...bottomRect,
    // rowH: 0.31, // 单元格默认高度
    valign: 'middle',
    fontSize: textFontSize,
    color: textColor,
    // align: 'center',
    // colW: ['100%']  // 表格每一列宽度
  })

  slide.addText(`【答案】 ${timu['答案']}`, {
    ...bottomRect,
    y: '80%', h: '15%',
    fontSize: textFontSize,
    align: 'left',
    fontFace: '微软雅黑',
    color: daanColor
  })
  return slide
}

function addJieda (slide, timu) {
  slide.addText(`${timu['总序']}. 【解析】 ${timu['调整解析'] || ''}`, {
    ...topRect,
    h: '80%',
    fontSize: textFontSize,
    valign: 'top',
    align: 'left',
    fontFace: '微软雅黑',
    color: textColor
  })
  return slide
}
