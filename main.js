const excelToJson = require("convert-excel-to-json");
const path = require("path");
const fs = require("fs");

const sheetName = "Sheet1"; //待转出的表格名称
const fileDir = path.resolve(__dirname, "./data");
const targetFile = path.resolve(fileDir, "data.xlsx");
const destFile = path.resolve(fileDir, `${new Date().getTime()}.js`);

try {
  const result = excelToJson({
    sourceFile: targetFile,
    header: {
      rows: 2
    },
    sheets: [sheetName],
    columnToKey: {
      A: "id",
      B: "name", // 课程名称
      C: "info", // 课程描述
      D: "subject", // 学科：语文、数学...
      E: "period", // 学段
      F: "grade", // 年级
      G: "classify", // 学科类别：专题、高三语文第一单元...
      H: "type", // 类型：链接或视频
      I: "resource_url", // 资源地址：视频或链接
    }
  });
  const fileStream = Buffer.from(JSON.stringify(result[sheetName]));
  const writeStream = fs.createWriteStream(destFile);
  writeStream.write("var courseData =" + fileStream);
  writeStream.end();
  console.log(`=================`);
  console.log(`转换成功，文件为：${destFile}`);
} catch (err) {
  throw new Error(err);
}
