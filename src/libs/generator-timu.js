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
  Table,
  TableCell,
  TableRow
} from "docx";

export class DocumentCreator {
  create(sheet) {
    let timuList = sheet.timuList.sort((a, b) => a['总序'] < b['总序'])

    let danxuanList = timuList.filter(item => item['题型'] === '单选')
    let duoxuanList = timuList.filter(item => item['题型'] === '多选')
    const document = new Document({
      styles: {
        paragraphStyles: [ // 段落样式
          {
            id: "timu",
            name: "timu",
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
                before: 300,
                after: 150
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
                level: 0,
                format: "upperRoman",
                text: "%1",
                alignment: AlignmentType.START,
                style: {
                  paragraph: {
                    indent: {left: 720, hanging: 260}
                  }
                }
              },
              {
                level: 1,
                format: "decimal",
                text: "%2.",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: {left: 0, hanging: 260}
                  }
                }
              },
              {
                level: 2,
                format: "lowerLetter",
                text: "%3)",
                alignment: AlignmentType.START,
                style: {
                  paragraph: {
                    indent: {left: 2160, hanging: 1700}
                  }
                }
              },
              {
                level: 3,
                format: "upperLetter",
                text: "%4)",
                alignment: AlignmentType.START,
                style: {
                  paragraph: {
                    indent: {left: 2880, hanging: 2420}
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
              alignment: 'center',
              break: 1
            }),
            ...danxuanList.map((timu, index) => {
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
                }))
                arr.push(new Paragraph({
                  text: `B.${timu['B']}`,
                }))
                arr.push(new Paragraph({
                  text: `C.${timu['C']}`,
                }))
                arr.push(new Paragraph({
                  text: `D.${timu['D']}`,
                }))
              } else {
                // 2个选项一排
                arr.push(new Table({
                  // todo 边框要改为透明
                  borders: {top: {color: null}, left: 0, bottom: 0, right: 0},
                  width: {type: 'pct', size: 100},
                  rows: [
                    new TableRow({
                      children: [
                        new TableCell({
                          children: [new Paragraph(`A.${timu['A']}`)],
                        }),
                        new TableCell({
                          children: [new Paragraph(`B.${timu['B']}`)],
                        }),
                      ],
                    }),
                    new TableRow({
                      children: [
                        new TableCell({
                          children: [new Paragraph(`C.${timu['C']}`)],
                        }),
                        new TableCell({
                          children: [new Paragraph(`D.${timu['D']}`)],
                        }),
                      ],
                    })
                  ]
                }))
                // arr.push(new Paragraph({
                //   bidirectional: true,
                //   children: [
                //     new TextRun({
                //       text: `A.${timu['A']}`,
                //       rightToLeft: true,
                //     })
                //   ]
                // }))
                // arr.push(new Paragraph({
                //   bidirectional: true,
                //   children: [
                //     new TextRun({
                //       text: `B.${timu['B']}`,
                //       rightToLeft: true,
                //     })
                //   ]
                // }))
                // arr.push(new Paragraph({
                //   bidirectional: true,
                //   children: [
                //     new TextRun({
                //       text: `C.${timu['C']}`,
                //       rightToLeft: true,
                //     })
                //   ]
                // }))
                // arr.push(new Paragraph({
                //   bidirectional: true,
                //   children: [
                //     new TextRun({
                //       text: `D.${timu['D']}`,
                //       rightToLeft: true,
                //     })
                //   ]
                // }))
              }
              console.log(arr)
              return arr
            }).reduce((prev, curr) => prev.concat(curr), []),
          ]
        }
      ]
    })

    return document
  }
}
