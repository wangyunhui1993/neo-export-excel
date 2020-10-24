# neo-export-excel

一个用于 excel 导出的强大函数（支持自定义样式、自定义单元格属性、多sheet，表格合并等）

## Installation

You can use this package on the server side as well as the client side.

### [Node.js](http://nodejs.org/):

```
npm install neo-export-excel
```

## Usage

```javascript
import neoExportExcel from "neo-export-excel";
```

## API

```javascript
neoExportExcel({ name: "导出表格的名称", sheets: ObjectArray });
```

eg:

```javascript
var list = [
  { ypsj: "2020-10-10", tbbm: "技术科" },
  { ypsj: "2020-10-11", tbbm: "办公室" },
];
const tHeader = [
  {
    key: "ypsj",
    name: "研判时间",
    style: "",
    format: "",
  },
  { key: "tbbm", name: "填报部门" },
];
var excelInfo = {
  name: "这是表格的名称",
  sheets: [
    {
      name: "这是sheet1",
      content: list,
      tHeader: tHeader,
    },
    {
      name: "这是sheet2",
      content: list,
      tHeader: tHeader,
    },
  ],
};

importToExcel(excelInfo);
```
