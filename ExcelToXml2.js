const fs = require('fs');
const xlsx = require('xlsx');
const { create, element } = require('xmlbuilder2');

// 配置对象，用于映射Excel列名到XML元素和属性
const config = [
  {
    excelColumn: 'Material name',
    xmlElement: 'Material',
    attributes: [
      { name: 'name', excelColumn: 'Material name' }
    ],
    nestedFields: [
      {
        excelColumn: 'Entry Thickness',
        xmlElement: 'Entry',
        attributes: [
          { name: 'Thickness', excelColumn: 'Entry Thickness' },
          { name: 'Desc', excelColumn: 'Desc' },
          { name: 'NoThickTitle', excelColumn: 'NoThickTitle' }
        ],
        nestedFields: [
          {
            excelColumn: 'CutSetting type',
            xmlElement: 'CutSetting',
            attributes: [
              { name: 'type', excelColumn: 'CutSetting type' }
            ],
            fields: [
              { excelColumn: 'index Value', xmlElement: 'index' },
              { excelColumn: 'name Value', xmlElement: 'name' },
              { excelColumn: 'LinkPath Value', xmlElement: 'LinkPath' },
              { excelColumn: 'maxPower Value', xmlElement: 'maxPower' },
              { excelColumn: 'maxPower2 Value', xmlElement: 'maxPower2' },
              { excelColumn: 'speed Value', xmlElement: 'speed' },
              { excelColumn: 'numPasses Value', xmlElement: 'numPasses' },
              { excelColumn: 'priority Value', xmlElement: 'priority' },
              { excelColumn: 'tabCount Value', xmlElement: 'tabCount' },
              { excelColumn: 'tabCountMax Value', xmlElement: 'tabCountMax' },
              { excelColumn: 'ditherMode Value', xmlElement: 'ditherMode' },
              { excelColumn: 'dpi Value', xmlElement: 'dpi' }
            ]
          }
        ]
      }
    ]
  },
  // 添加其他映射...
];

const excelToXml = (excelFile, xmlFile) => {
// 读取Excel文件
  const workbook = xlsx.readFile(excelFile); // 替换为您的Excel文件路径
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// 转换Excel数据为JSON
  const jsonData = xlsx.utils.sheet_to_json(worksheet);



// 创建XML文档
  const root = create({ version: '1.0', encoding: 'UTF-8' });
  const lightBurnLibrary = root.ele('LightBurnLibrary');

// 将JSON数据转换为XML
  jsonData.forEach(dataEntry => {
    const xmlElement = lightBurnLibrary.ele('Material');

    config.forEach(mapping => {
      const nestedElement = xmlElement.ele(mapping.xmlElement);

      mapping.attributes.forEach(attrConfig => {
        nestedElement.att(attrConfig.name, dataEntry[attrConfig.excelColumn]);
      });

      mapping.nestedFields.forEach(nestedMapping => {
        const nestedXmlElement = nestedElement.ele(nestedMapping.xmlElement);

        nestedMapping.attributes.forEach(attrConfig => {
          if (nestedMapping.excelColumn in dataEntry) {
            nestedXmlElement.att(attrConfig.name, dataEntry[nestedMapping.excelColumn]);
          }
        });

        nestedMapping?.fields?.forEach(fieldMapping => {
          if (fieldMapping.excelColumn in dataEntry) {
            const fieldValue = dataEntry[fieldMapping.excelColumn];
            nestedXmlElement.ele(fieldMapping.xmlElement, { Value: fieldValue });
          }
        });
      });
    });
  });

// 将XML保存到文件
  const xmlString = root.end({ prettyPrint: true });
  fs.writeFileSync(xmlFile, xmlString); // 替换为您想要保存的XML文件路径
  console.log('XML转换完成！');

}
const excelFile = process.argv[2]; // the first argument passed in the command line
const xmlFile = process.argv[3]; // the second argument passed in the command line

excelToXml(excelFile, xmlFile);
