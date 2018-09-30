const fs = require('fs')
const path = require('path')
const Excel = require('xlsx')
const moment = require('moment')
const Yaml = require('js-yaml')

const phonebook = Yaml.safeLoad(fs.readFileSync('./电话本.yaml'))

module.exports = class MergeExcel {
  constructor ({ inputPath, outputPath }) {
    this.inputPath = inputPath
    this.outputPath = outputPath
  }

  go () {
    fs.readdir(this.inputPath, (err, files) => {
      if (err) {
        console.error("无法读取这个目录", err)
        process.exit(1)
      } 
    
      let excels = files.filter(x => isExcelFile(x))
    
      let mergeData = []
      excels.forEach(f => {
        let f1 = path.join(this.inputPath, f)
        let rawData = readExcel(f1)
        let cleanedData = clean(f.split('.')[0], rawData)
        mergeData = mergeData.concat(cleanedData)
      })
  
      // 排序
      let sortedData = mergeData.sort((a, b) => {
        return moment(a.北京时间) - moment(b.北京时间)
      })
  
      // 去重
      let uniqData = {}
      sortedData.forEach(x => {
        uniqData[x.北京时间 + x.发信人] = x
      })
      uniqData = Object.values(uniqData)
  
      // 输出
      outputExcel(uniqData, this.outputPath)
    })
  }
}

const outputExcel = (arr, outputPath) => {
  let wb = Excel.utils.book_new()
  let outputSheet = Excel.utils.json_to_sheet(arr)
  console.log(outputSheet)
  Excel.utils.book_append_sheet(wb, outputSheet, '合并')

  Excel.writeFile(wb, outputPath)
}

const readExcel = (f) => {
  console.log(f)
  let workbook = Excel.readFile(f)

  const sheetNames = workbook.SheetNames
  const sheet = workbook.Sheets[sheetNames[0]]
  let data = Excel.utils.sheet_to_json(sheet)
  return data
}

const isExcelFile = (f) => {
  return f.match('.xls')
}

const clean = (fname, rawData) => {
  // console.log(rawData)

  // 只保留需要的类型
  // text/plain 纯文本 原样保留
  // application/cyfile 文件 取最后的 http://caiyun 网址部分
  let d1 = rawData.filter(x => {
    let arr = ['text/plain', 'application/cyfile']
    // let arr = ['application/cyfile']
    return arr.indexOf(x._ContentType) !== -1
  })

  // 取出需要的列
  // MsgTime       时间
  // PeerUri       接收方
  // _fromUri      发送方
  // _ContentType  类型
  // _content      内容
  // OwnerId
  d1 = d1.map(x => {
    let { MsgTime, _fromUri, PeerUri, _ContentType, _content, OwnerId } = x
    
    let 北京时间 = moment(MsgTime).add(8, 'hour').format('YYYY-MM-DD HH:mm:ss')

    let 消息正文
    if (x._ContentType === 'text/plain') {
      消息正文 = _content
    }
    if (x._ContentType === 'application/cyfile') {
      let m = x._content.match(/.{1}(.+).{2}(http:\/\/.+)/)
      消息正文 = m[2]
    }

    let 发信人 = _fromUri.match(/tel.+(1[0-9]{10})/)[1]
    let 收信人
    try {
      收信人 = PeerUri.match(/tel.+(1[0-9]{10})/)[1]
    } catch (e) {
      收信人 = PeerUri
    }
    if (发信人 === 收信人) {
      收信人 = fname
    }

    收信人 = phonebook[收信人] || 收信人
    发信人 = phonebook[发信人] || 发信人

    return { 北京时间, 发信人, 收信人, 消息类型: _ContentType, 消息正文 }

  })

  return d1
}