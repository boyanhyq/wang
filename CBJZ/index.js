const XLSX =require('xlsx');
const utils  = require('./utils');

const workbook = XLSX.readFile('./demo.xlsx');
// 获取 Excel 中所有表名
const sheetNames = workbook.SheetNames; // 返回 ['sheet1', 'sheet2']
console.log(sheetNames);
// 根据表名获取对应某张表
const codesheet = workbook.Sheets[sheetNames[0]];
const datasheet = workbook.Sheets[sheetNames[1]];

const codes=XLSX.utils.sheet_to_json(codesheet) ;
const datas=XLSX.utils.sheet_to_json(datasheet) ;

console.log(codes);
const finalData = [];

codes.forEach((obj)=>{
  const code = obj['编号'];

  datas.forEach((data)=>{
    const numCode = data['科目名称'];
    if(numCode.indexOf(code)>-1){
      finalData.push({
        ['科目代码']: data['科目代码'],
        ['科目名称']: data['科目名称'].replace(/]/g,'---').replace('[','项目---'),
        ['期末借方余额']: data['期末借方余额'],
      })
    }
})
})


// console.log(finalData)/
const header=['科目代码','科目名称','期末借方余额']
utils.exportJson(header,finalData)