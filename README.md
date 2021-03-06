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

eg1:多sheet

```javascript
var list1 = [{
				ypsj: "2020-10-10",
				tbbm: "技术科",
				score: "99.12"
			},
			{
				ypsj: "2020-10-11",
				tbbm: "办公室",
				score: "56.4"
			},
		];
		var list2 = [{
				ypsj: "2020-10-10",
				tbbm: "技术科",
				score: "73.567"
			},
			{
				ypsj: "2020-10-11",
				tbbm: "办公室",
				score: "89"
			},
		];
		const tHeader = [{
				key: "ypsj",
				name: "研判时间",
			},
			{
				key: "tbbm",
				name: "填报部门"
			},
			{
				key: "score",
				name: "代码",
				format: "0.00",  //单元格格式 例："0.00"代表数字保留2位小数，"@"代表文本，"Percent"代表百分比
				style:{"text-align":"right"}  //单元格样式
			},
		];
		var excelInfo = {
			name: "这是表格的名称",
			sheets: [{
					name: "这是sheet1",
					content: list1,
					tHeader: tHeader,
				},
				{
					name: "这是sheet2",
					content: list2,
					tHeader: tHeader,
				},
			],
		};

neoExportExcel(excelInfo);
```


eg2:表格合并

```javascript
var content =
			`<thead>
				<tr>
					<th>研判时间</th>
					<th>填报部门</th>
					<th>代码</th>
				</tr>
			</thead>
			<tbody>
				<tr>
					<td>2020-10-22</td>
					<td>技术科</td>
					<td>123</td>
				</tr>
				<tr>
					<td>2020-10-22</td>
					<td rowspan="2">办公室</td>
					<td>123</td>
				</tr>
				<tr>
					<td>2020-10-22</td>
					<td>123</td>
				</tr>
			</tbody>`
		var excelInfo = {
			name: "这是表格的名称",
			sheets: [{
				name: "这是sheet1",
				content: content,
			}],
		};

neoExportExcel(excelInfo);
```