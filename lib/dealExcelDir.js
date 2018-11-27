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
    // 判断是目录还是文件
    let isDir = fs.lstatSync(this.inputPath).isDirectory()

    if (isDir) {
      console.log('输入的是目录')
      fs.readdir(this.inputPath, (err, files) => {
        if (err) {
          console.error("无法读取这个目录", err)
          process.exit(1)
        } 

        let excels = files.filter(x => isExcelFile(x))
        let paths = excels.map(x => {
          return path.join(this.inputPath, x)
        })

        this.dealExcels(paths)
      })
    } else {
      console.log('输入的是单个文件')
      let paths = [this.inputPath]
      this.dealExcels(paths)
    }
  }

  dealExcels (paths) {
    let mergeData = []
    paths.forEach(f1 => {
      let rawData = readExcel(f1)
      let fname = f1.split('/').pop().split('.').shift()
      let cleanedData = clean(fname, rawData)
      mergeData = mergeData.concat(cleanedData)
    })

    // 去重
    let uniqData = {}
    mergeData.forEach(x => {
      uniqData[x.北京时间 + x.发信人] = x
    })
    uniqData = Object.values(uniqData)

    // 排序
    let sortedData = uniqData.sort((a, b) => {
      return moment(a.北京时间) - moment(b.北京时间)
    })

    // 分组
    // 如果收信人是 group 直接分组
    // 如果收信人不是 group 按照收发信人分组（私聊）
    let groups = {}
    sortedData.forEach(x => {
      let groupKey = genGroupKey(x)

      groups[groupKey] = groups[groupKey] || []
      groups[groupKey].push(x)
    })
    let groupedData = []
    Object.values(groups).forEach(x => {
      groupedData = groupedData.concat(x).concat([{}])
    })

    // 输出
    outputExcel(groupedData, this.outputPath)
  }
}

const genGroupKey = x => {
  if (x.收信人.includes("group")) {
    return x.收信人
  }

  return [x.发信人, x.收信人].sort().join('-')
}

const outputExcel = (arr, outputPath) => {
  let wb = Excel.utils.book_new()
  let outputSheet = Excel.utils.json_to_sheet(arr)
  // console.log(outputSheet)
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

const dealPhone = (str) => {
  if (str.includes('group')) {
    return str
  }

  try {
    // let phoneRegex = /tel.+(1[0-9]{10})/
    let phoneRegex = /tel.\+(86|852)([0-9]+)/
    let re = str.match(phoneRegex)[2]
    return re
  } catch (e) {
    console.log('处理错误:', str)
    throw e
  }
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
    let 文件名
    if (x._ContentType === 'text/plain') {
      消息正文 = _content
    }
    if (x._ContentType === 'application/cyfile') {
      let content = x._content.split('*').pop()
      let m = content.match(/.{1}(.+).{2}(http:\/\/.+)/)
      文件名 = m[1]
      消息正文 = m[2]
    }

    let 发信人 = dealPhone(_fromUri)
    let 收信人 = dealPhone(PeerUri)
    // try {
    //   收信人 = PeerUri.match(phoneRegex)[1]
    // } catch (e) {
    //   收信人 = PeerUri
    // }
    // if (发信人 === 收信人) {
    //   收信人 = fname
    // }

    收信人 = phonebook[收信人] || 收信人
    发信人 = phonebook[发信人] || 发信人

    return { 北京时间, 发信人, 收信人, 消息类型: _ContentType, 消息正文, 文件名 }

  })

  return d1
}