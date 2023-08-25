const fs = require('fs');
const xlsx = require('xlsx');
const { create, element } = require('xmlbuilder2');
const excelToXml = (excelFile, xmlFile) => {
// 读取Excel文件
  const workbook = xlsx.readFile(excelFile); // 替换为您的Excel文件路径
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// 转换Excel数据为JSON
  const jsonData = xlsx.utils.sheet_to_json(worksheet);

// 创建XML文档
  const root = create({version: '1.0', encoding: 'UTF-8'});
  const lightBurnLibrary = root.ele('LightBurnLibrary');

// 将JSON数据转换为XML
  jsonData.forEach((dataEntry, index) => {
    console.log(index, ':', dataEntry)
    const material = lightBurnLibrary.ele('Material', {name: dataEntry['Material name']});

    const entryAttrs = {
      Thickness: dataEntry['Entry Thickness'],
      Desc: dataEntry['Desc'],
      NoThickTitle: dataEntry['NoThickTitle']
    };
    const entry = material.ele('Entry', entryAttrs);

    const cutSettingAttrs = {type: dataEntry['CutSetting type']};
    const cutSetting = entry.ele('CutSetting', cutSettingAttrs);

    cutSetting.ele('index', {Value: dataEntry['index Value']});
    cutSetting.ele('name', {Value: dataEntry['name Value']});
    cutSetting.ele('LinkPath', {Value: dataEntry['LinkPath Value']});
    cutSetting.ele('maxPower', {Value: dataEntry['maxPower Value']});
    cutSetting.ele('maxPower2', {Value: dataEntry['maxPower2 Value']});
    cutSetting.ele('speed', {Value: dataEntry['speed Value']});
    cutSetting.ele('numPasses', {Value: dataEntry['numPasses Value']});
    cutSetting.ele('priority', {Value: dataEntry['priority Value']});
    cutSetting.ele('tabCount', {Value: dataEntry['tabCount Value']});
    cutSetting.ele('tabCountMax', {Value: dataEntry['tabCountMax Value']});
    // 继续添加其他属性...
  });

// 将XML保存到文件
  const xmlString = root.end({prettyPrint: true});
  fs.writeFileSync(xmlFile, xmlString); // 替换为您想要保存的XML文件路径
  console.log('XML转换完成！');
}
const excelFile = process.argv[2]; // the first argument passed in the command line
const xmlFile = process.argv[3]; // the second argument passed in the command line

excelToXml(excelFile, xmlFile);
