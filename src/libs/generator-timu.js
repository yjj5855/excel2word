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
  TableRow, BorderStyle, Columns, Column
} from "docx";

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
            id: "subtitle",
            name: "subtitle",
            run: {
              size: 36,
              color: "#000000"
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
              size: 24,
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
              size: 24,
              color: "#000000",
              // margin: {
              //   top: 800,
              //   bottom: 300
              // }
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
              heading: HeadingLevel.TITLE,
              alignment: 'center'
            }),
            new Paragraph({
              text: '单项选择题',
              heading: 'subtitle',
              alignment: 'center'
            }),
            ...danxuanList.map((timu, index) => {
              return getTimuElement(timu)
            }).reduce((prev, curr) => prev.concat(curr), []),
            new Paragraph({
              text: '多项选择题',
              heading: 'subtitle',
              alignment: 'center'
            }),
            ...duoxuanList.map((timu, index) => {
              return getTimuElement(timu, 'duo')
            }).reduce((prev, curr) => prev.concat(curr), []),
            new Paragraph({
              text: '参考答案',
              heading: 'subtitle',
              alignment: 'center'
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: '一 单项选择题',
                  bold: true,
                  heading: 'timu'
                })
              ]
            }),
            ...getDaanElement(danxuanList),
            new Paragraph({
              children: [
                new TextRun({
                  text: '二 多项选择题',
                  bold: true,
                  heading: 'timu'
                })
              ]
            }),
            ...getDaanElement(duoxuanList, 'duo'),
            new Paragraph({
              text: '答案解析',
              heading: 'subtitle',
              alignment: 'center'
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: '一 单项选择题',
                  bold: true,
                  heading: 'timu'
                })
              ]
            }),
            ...danxuanList.map((timu, index) => {
              return getJiexiElement(timu)
            }).reduce((prev, curr) => prev.concat(curr), []),
            new Paragraph({
              children: [
                new TextRun({
                  text: '二 多项选择题',
                  bold: true,
                  heading: 'timu'
                })
              ]
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
  return arr
}
