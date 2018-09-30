const args = require('args')
const dealExcelDir = require('./lib/dealExcelDir')

// CLI args
args
  .option('input', '要处理的目录')
  .option('output', '要合并输出的 excel 文件')

const run = ({ inputPath, outputPath }) => {
  console.log(`开始处理... ${inputPath} --> ${outputPath}`)

  new dealExcelDir({ inputPath, outputPath }).go()
}

const options = args.parse(process.argv)
let inputPath = options.input
let outputPath = options.output

if (inputPath && outputPath) {
  run({ inputPath, outputPath })
} else {
  args.showHelp()
}