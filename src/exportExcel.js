var tablesToExcel = (function () {
  var uri = "data:application/vnd.ms-excel;base64,",
    html_start = `<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">`,
    template_ExcelWorksheet = `<x:ExcelWorksheet><x:Name>{SheetName}</x:Name><x:WorksheetSource HRef="sheet{SheetIndex}.htm"/></x:ExcelWorksheet>`,
    template_ListWorksheet = `<o:File HRef="sheet{SheetIndex}.htm"/>`,
    style = `<style type="text/css">
    table th{text-align: center;font-weight: bold;font-size:16px;background-color: #559EC6;color:#fff;height:30px;}
    table td{text-align: center;font-size:16px;}
    </style>`,
    template_HTMLWorksheet =
      `
------=_NextPart_dummy
Content-Location: sheet{SheetIndex}.htm
Content-Type: text/html; charset=utf-8

` +
      html_start +
      `
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link id="Main-File" rel="Main-File" href="../WorkBook.htm">
    <link rel="File-List" href="filelist.xml">
    ${style}
</head>
<body><table  border="1">{SheetContent}</table></body>
</html>`,
    template_WorkBook =
      `MIME-Version: 1.0
X-Document-Type: Workbook
Content-Type: multipart/related; boundary="----=_NextPart_dummy"

------=_NextPart_dummy
Content-Location: WorkBook.htm
Content-Type: text/html; charset=utf-8

` +
      html_start +
      `
<head>
<meta name="Excel Workbook Frameset">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="File-List" href="filelist.xml">
<!--[if gte mso 9]><xml>
 <x:ExcelWorkbook>
    <x:ExcelWorksheets>{ExcelWorksheets}</x:ExcelWorksheets>
    <x:ActiveSheet>0</x:ActiveSheet>
 </x:ExcelWorkbook>
</xml><![endif]-->
</head>
<frameset>
    <frame src="sheet0.htm" name="frSheet">
    <noframes><body><p>This page uses frames, but your browser does not support them.</p></body></noframes>
</frameset>
</html>
{HTMLWorksheets}
Content-Location: filelist.xml
Content-Type: text/xml; charset="utf-8"

<xml xmlns:o="urn:schemas-microsoft-com:office:office">
    <o:MainFile HRef="../WorkBook.htm"/>
    {ListWorksheets}
    <o:File HRef="filelist.xml"/>
</xml>
------=_NextPart_dummy--
`,
    base64 = function (s) {
      return window.btoa(unescape(encodeURIComponent(s)));
    },
    format = function (s, c) {
      return s.replace(/{(\w+)}/g, function (m, p) {
        return c[p];
      });
    };
  return function (tables, filename) {
    var context_WorkBook = {
      ExcelWorksheets: "",
      HTMLWorksheets: "",
      ListWorksheets: "",
    };
    tables.forEach((item, index) => {
      var SheetName = item.name || `sheet${index + 1}`;
      context_WorkBook.ExcelWorksheets += format(template_ExcelWorksheet, {
        SheetIndex: index,
        SheetName: SheetName,
      });
      context_WorkBook.HTMLWorksheets += format(template_HTMLWorksheet, {
        SheetIndex: index,
        SheetContent: item.html,
      });
      context_WorkBook.ListWorksheets += format(template_ListWorksheet, {
        SheetIndex: index,
      });
    });
    console.log(
      "context_WorkBook",
      format(template_WorkBook, context_WorkBook)
    );
    var link = document.createElement("A");
    link.href = uri + base64(format(template_WorkBook, context_WorkBook));
    link.download = filename || "Workbook.xls";
    link.target = "_blank";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };
})();
/* var sheet_1 = `<thead><tr><th>这是sheet1的表头</th></tr><thead><tbody><tr><td>这是sheet1的表体</td></tr><tbody>`
var sheet_2 = `<thead><tr><th>这是sheet2的表头</th></tr><thead><tbody><tr><td>这是sheet2的表体</td></tr><tbody>`
var sheets = [
  { name: "合并详情", html: sheet_1 },
  { name: "发票详情", html: sheet_2 }
];
tablesToExcel(sheets, "导出的excel"); */
function getTable(head, list, tableInfo) {
  //获取完整表格
  let content = "";
  content += getH(head);
  content += "<tbody>";
  list.forEach((item) => {
    content += getB(item, tableInfo);
  });
  content += "</tbody>";
  return content;
  // let ele = document.createElement("table");
  // ele.innerHTML = content;
  // return ele;
}

function getH(list) {
  //拼接表头
  let strH = "<thead><tr>";
  for (var item of list) {
    strH += `<th>${item}</th>`;
  }
  strH += "</tr></thead>";
  return strH;
}

/* 
  用写html的方法导出excel时，excel会自动把一些格式转换一下，有时达不预期的效果，此时可以通过样式进行调整
  mso-number-format:@   文本
  mso-number-format:"0.000"   数字 
  mso-number-format:"mm/dd/yy"    日期 
  mso-number-format:"d\-mmm\-yyyy"   日期
  mso-number-format:Percent   百分比
*/

function getB(list, tableInfo) {
  console.log("list, tableInfo", list, tableInfo);
  //拼接表体
  let strB = "<tr>";
  list.forEach((item, index) => {
    let format = tableInfo[index].format;
    let style = tableInfo[index].style;
    var styleStr = "";
    if (typeof style == "function") {
      let styleFn = style(item, tableInfo);
      style = styleFn;
    }
    if (style instanceof Object) {
      for (var eleKey in style) {
        styleStr += `${eleKey}:${style[eleKey]};`;
      }
    } else {
      style = "";
    }
    if (typeof format == "string") {
      strB += `<td style="mso-number-format:'${format}';${styleStr}">${item}</td>`;
    } else if (typeof format == "function") {
      strB += `<td style="mso-number-format:'${format(
        item,
        tableInfo
      )}';${styleStr}">${item}</td>`;
    } else {
      // eslint-disable-next-line no-useless-escape
      strB += `<td style="mso-number-format:'\@';${styleStr}">${item}</td>`;
    }
  });
  strB += "</tr>";
  return strB;
}

const exportToExcel = function (excelInfo) {
  let excelName = excelInfo.name || "Workbook";
  let sheetList = excelInfo.sheets;
  if (sheetList && sheetList instanceof Array) {
    sheetList.forEach((item) => {
      let content = item.content;
      if (content.nodeType) {
        item.html = content.innerHTML;
      } else if (content instanceof Array) {
        let t_c = [],
          t_e = [];
        for (var H of item.tHeader) {
          t_e.push(H.key);
          t_c.push(H.name || "");
        }
        let data = formatJson(t_e, item.content);
        item.html = getTable(t_c, data, item.tHeader);
      } else if (typeof content == "string") {
        item.html = content;
      }
    });
    tablesToExcel(sheetList, excelName);
  } else {
    console.error("sheets为必传属性");
  }
};
const formatJson = (filterVal, jsonData) => {
  return jsonData.map((v) => filterVal.map((j) => v[j] || ""));
};

/* 

// 单元格内换行：<br style='mso-data-placement:same-cell;'/>
style:单元格样式，支持函数、对象
format：单元格格式，支持函数、对象
content：当为数组是需要和tHeader搭配使用，当为字符串时表示导出的表格元素
const tHeader = [
            { key: "tbbm", name: "填报部门" },
            {
              key: "ypsj",
              name: "研判时间",
              style:'',
              format:'',
            },
            { key: "yxq", name: "有效期" },
            { key: "dtfxsj", name: "动态风险事件" },
            { key: "sjly", name: "事件来源" },
            { key: "jhlb", name: "计划类别" },
            {
              key: "jczdnr",
              name: "检查重点内容",
            },
          ];
          let excelInfo = {
            name: "动态风险事件上报",
            sheets: [
              {
                name: "动态风险事件上报sheet1",
                content: list,
                tHeader: tHeader,
              },
              {
                name: "动态风险事件上报sheet2",
                content: list,
                tHeader: tHeader,
              },
            ],
          };
          importToExcel(excelInfo)

*/
module.exports = exportToExcel;
