import {
  AlignmentType,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  TabStopPosition,
  TabStopType,
  TextRun,
  NumberValueElement,
  Numbering,
  TableBorders,
  WidthType,
  Table,
  TableCell,
  TableRow, BorderStyle, Columns, Column,
  ImageRun
} from "docx";

/**
 *
 八号	5
 七号	5.5
 小六	6.5
 六号	7.5
 小五	9
 五号	10.5
 小四	12
 四号	14
 小三	15
 三号	16
 小二	18
 二号	22
 小一	24
 一号	26
 小初	36
 初号	42
 */

export class DocumentCreator {
  create(sheet) {
    let timuList = sheet.timuList.sort((a, b) => a['总序'] < b['总序'])
    console.log(sheet.timuList)
    let danxuanList = timuList.filter(item => item['题型'] === '单选')
    let duoxuanList = timuList.filter(item => item['题型'] === '多选')
    // console.log(danxuanList, duoxuanList)
    const document = new Document({
      styles: {
        paragraphStyles: [ // 段落样式
          {
            id: "title",
            name: "title",
            run: {
              size: 22 * 2,
              color: "#000000",
              bold: true
            },
            paragraph: { // 段落
              spacing: { // 字间距
                after: 500
              }
            }
          },
          {
            id: "subtitle",
            name: "subtitle",
            run: {
              size: 16 * 2,
              color: "#000000",
              bold: true
            },
            paragraph: { // 段落
              spacing: { // 字间距
                before: 500,
                after: 300
              }
            }
          },
          {
            id: "timu",
            name: "timu",
            run: {
              size: 12 * 2,
              color: "#000000"
            },
            paragraph: { // 段落
              spacing: { // 字间距
                before: 300,
                after: 150
              }
            }
          },
          {
            id: "daan",
            name: "daan",
            run: {
              size: 12 * 2,
              color: "#000000"
            },
            paragraph: { // 段落
              spacing: { // 字间距
                before: 100,
                after: 100
              }
            }
          }
        ]
      },
      numbering: { // 设置项目编号
        config: [
          {
            reference: "my-crazy-numbering",
            levels: [
              {
                level: 1,
                format: "decimal",
                text: "%2.",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: {left: 0, hanging: 360}
                  }
                }
              },
            ]
          },
          {
            reference: "jiexi-numbering",
            levels: [
              {
                level: 1,
                format: "decimal",
                text: "%2.",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: {left: 0, hanging: 360}
                  }
                }
              }
            ]
          }
        ]
      },
      sections: [
        {
          children: [
            new Paragraph({
              text: sheet.title,
              heading: 'title',
              alignment: 'center'
            }),
            new Paragraph({
              text: '单项选择题',
              heading: 'subtitle',
              bold: true,
              alignment: 'center'
            }),
            ...danxuanList.map((timu, index) => {
              return getTimuElement(timu)
            }).reduce((prev, curr) => prev.concat(curr), []),
            new Paragraph({
              text: '多项选择题',
              heading: 'subtitle',
              bold: true,
              alignment: 'center'
            }),
            ...duoxuanList.map((timu, index) => {
              return getTimuElement(timu, 'duo')
            }).reduce((prev, curr) => prev.concat(curr), []),
            new Paragraph({
              text: '参考答案',
              heading: 'subtitle',
              bold: true,
              alignment: 'center'
            }),
            new Paragraph({
              text: '一 单项选择题',
              bold: true,
              heading: 'timu',
            }),
            ...getDaanElement(danxuanList),
            new Paragraph({
              text: '二 多项选择题',
              bold: true,
              heading: 'timu',
            }),
            ...getDaanElement(duoxuanList, 'duo'),
            new Paragraph({
              text: '答案解析',
              heading: 'subtitle',
              bold: true,
              alignment: 'center'
            }),
            new Paragraph({
              text: '一 单项选择题',
              heading: 'timu',
              bold: true
            }),
            ...danxuanList.map((timu, index) => {
              return getJiexiElement(timu)
            }).reduce((prev, curr) => prev.concat(curr), []),
            new Paragraph({
              text: '二 多项选择题',
              bold: true,
              heading: 'timu'
            }),
            ...duoxuanList.map((timu, index) => {
              return getJiexiElement(timu, 'duo')
            }).reduce((prev, curr) => prev.concat(curr), []),
          ]
        }
      ]
    })

    return document
  }
}

function getTimuElement(timu, danOrDuo = 'dan') {
  const arr = []

  // 添加题目
  const p = new Paragraph({
    text: timu['题目'],
    heading: 'timu',
    numbering: {
      reference: 'my-crazy-numbering',
      level: 1
    }
  })
  arr.push(p)

  // 添加题目图片
  if (timu['题目_image']) {
    // console.log(timu['题目_image'])
    const imageRun = new Paragraph({
      children: [
        new ImageRun({
          data: timu['题目_image'].buffer,
          transformation: {
            width: timu['题目_image'].width || 150,
            height: timu['题目_image'].height || 150
          }
        })
      ]
    })
    arr.push(imageRun)
  }

  // 添加选项
  let xuanxiangMax = Math.max(...[
    timu['A'] && String(timu['A']).length || 0,
    timu['B'] && String(timu['B']).length || 0,
    timu['C'] && String(timu['C']).length || 0,
    timu['D'] && String(timu['D']).length || 0,
    timu['E'] && String(timu['E']).length || 0
  ])
  if (xuanxiangMax > 10) {
    // 换行
    arr.push(new Paragraph({
      text: `A.${timu['A']}`,
      heading: 'daan',
    }))
    arr.push(new Paragraph({
      text: `B.${timu['B']}`,
      heading: 'daan',
    }))
    arr.push(new Paragraph({
      text: `C.${timu['C']}`,
      heading: 'daan',
    }))
    arr.push(new Paragraph({
      text: `D.${timu['D']}`,
      heading: 'daan',
    }))
    if (danOrDuo === 'duo') {
      arr.push(new Paragraph({
        text: `E.${timu['E']}`,
        heading: 'daan',
      }))
    }
  } else {
    // 2个选项一排
    let E = []
    if (danOrDuo === 'duo') {
      E = [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  text: `E.${timu['E']}`,
                  heading: 'daan',
                })
              ]
            }),
            new TableCell({
              children: [
                new Paragraph({
                  text: ` `,
                  heading: 'daan',
                })
              ]
            })
          ]
        })
      ]
    }
    arr.push(new Table({
      borders: TableBorders.NONE,
      width: {
        type: WidthType.PERCENTAGE,
        size: 100
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  text: `A.${timu['A']}`,
                  heading: 'daan',
                }),
              ],
            }),
            new TableCell({
              children: [
                new Paragraph({
                  text: `B.${timu['B']}`,
                  heading: 'daan',
                }),
              ],
            })
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  text: `C.${timu['C']}`,
                  heading: 'daan',
                }),
              ]
            }),
            new TableCell({
              children: [
                new Paragraph({
                  text: `D.${timu['D']}`,
                  heading: 'daan',
                })
              ]
            })
          ]
        }),
        ...E
      ],
    }))
  }
  return arr
}

function getDaanElement(list, danOrDuo = 'dan') {
  const arr = []
  const rows = []
  let curRow = null
  let count = danOrDuo === 'dan' ? 10 : 5
  list.map((item, index) => {
    if (index % count === 0) {
      curRow = new TableRow({
        children: [
        ]
      })
    }
    curRow.addCellToIndex(new TableCell({
      children: [
        new Paragraph({
          text: `${item['总序'] <= 9 ? '  ': ''}${item['总序']}. ${item['答案']}`,
          heading: 'timu'
        }),
      ]
    }), index % count)
    if (index % count === count - 1 || index === list.length - 1) {
      rows.push(curRow)
    }
  })
  let table = new Table({
    borders: TableBorders.NONE,
    width: {
      type: WidthType.PERCENTAGE,
      size: 100
    },
    rows: rows
  })
  arr.push(table)
  return arr
}


function getJiexiElement(timu, danOrDuo = 'dan') {
  const arr = []
  const p = new Paragraph({
    text: `【答案】 ${timu['答案']}. ${timu['答案_1']}`,
    heading: 'timu',
    numbering: {
      reference: 'jiexi-numbering',
      level: 1
    }
  })
  arr.push(p)
  const jiexi = new Paragraph({
    text: `【解析】 ${timu['调整解析'] || ''}`,
    heading: 'daan'
  })
  arr.push(jiexi)

  // 添加题目图片
  if (timu['解析_image']) {
    // console.log(timu['题目_image'])
    const imageRun = new Paragraph({
      children: [
        new ImageRun({
          data: timu['解析_image'].buffer,
          transformation: {
            width: timu['解析_image'].width || 150,
            height: timu['解析_image'].height || 150
          }
        })
      ]
    })
    arr.push(imageRun)
  }
  return arr
}
