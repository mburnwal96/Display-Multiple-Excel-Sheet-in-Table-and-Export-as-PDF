$(document).ready(function(){
        excelSheetParser();
});

function excelSheetParser(){
	/* set up XMLHttpRequest */
	var url = "./airline.xls";
	var oReq = new XMLHttpRequest();
 
	oReq.open("GET", url, true);
	oReq.responseType = "arraybuffer";
 
	oReq.onload = function(e) {
	  var arraybuffer = oReq.response;
 
	  /* convert data to binary string */
	  var data = new Uint8Array(arraybuffer);
	  var arr = new Array();
	  for(var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
  	var bstr = arr.join("");
 
 	 /* Call XLSX */
 	 var workbook = XLSX.read(bstr, {type:"binary"});
	var sheetName = [];
        for(i=0;i<workbook.SheetNames.length;i++){
                sheetName.push(workbook.SheetNames[i]);
        }
	createExcelTab(sheetName);
	for(l=0; l<workbook.SheetNames.length; l++){
 	 	var str=workbook.Sheets[sheetName[l]]["!ref"].split(":");
		 var totalRowColumn = str[1].split("");
		 var column = totalRowColumn[0].charCodeAt(0)-64;
		 var row = rowSize(totalRowColumn);
	 	 var tableHeadRow = $ ('<tr></tr>');
		for(i=0; i<column; i++){
			var excelColumnValue = String.fromCharCode(65 + i)+"1";
			if(workbook.Sheets[sheetName[l]][excelColumnValue] != undefined){
				var tHeadData = $('<th>'+workbook.Sheets[sheetName[l]][excelColumnValue].v+'</th>');
			}
			else{
				var tHeadData = $('<th></th>');
			}
			tableHeadRow.append(tHeadData);
		}
		$("#excelSheetHead"+l).append(tableHeadRow);
		for(j=2; j<=row; j++){
			var tableBodyRow = $('<tr></tr>');
			for(k=0; k<column; k++){
				var excelColumnValue = String.fromCharCode(65 + k)+""+j;
				if(workbook.Sheets[sheetName[l]][excelColumnValue] != undefined){
	                		var tBodyData = $('<td>'+workbook.Sheets[sheetName[l]][excelColumnValue].v+'</td>');
	                		tableBodyRow.append(tBodyData);
				}
				else{
					var tBodyData = $('<td></td>');
	                                tableBodyRow.append(tBodyData);
				}
			}
			$("#excelSheetBody"+l).append(tableBodyRow);
		}
	  } 
    }
	oReq.send();
}
function createExcelTab(tabNameArray){
	  var tabNameUl = $('<ul class="nav nav-tabs"></ul>');
        for(i=0; i<=tabNameArray.length; i++){
            if(i<tabNameArray.length){
                if(i ==0){
                        var tabNameLi = $('<li class = "active"></li>');
                        var divTable = $('<div class = "tab-pane fade in active" style="width:100%"></div>');
                        var table = $('<table class = "table table-bordered"></table>');
                }
                else{
                        var tabNameLi = $('<li></li>');
                        var divTable = $('<div class = "tab-pane fade" style="width:100%"></div>');
                        var table = $('<table class = "table table-bordered"></table>');
                }
                var tabNameA = $('<a></a>');
                tabNameA.attr("data-toggle","tab");
                tabNameA.attr("href","#excelTable"+i);
                tabNameA.attr("id",i);
                tabNameA.text(tabNameArray[i]);
                tabNameLi.append(tabNameA);
                tabNameUl.append(tabNameLi);
                var tableHead = $('<thead class="alert-info"></thead>');
                var tableBody = $('<tbody></tbody>');
                table.attr("id","excelSheet"+i);
                tableHead.attr("id","excelSheetHead"+i);
                tableBody.attr("id","excelSheetBody"+i);
                table.append(tableHead);
                table.append(tableBody);
                divTable.attr("id","excelTable"+i);
                divTable.append(table);
                $("#multipleExcelSheet").append(divTable);
            }
            else{
                var buttonTab = $('<button type="button" class="btn btn-primary" onclick="exportPdf()" style="float: right"><i class="fa fa-file-pdf-o" aria-hidden="true"></i> Export as PDF</button>');
                tabNameUl.append(buttonTab);
            }
        }
        $("#robotApi").append(tabNameUl);
}
function rowSize(rowValue){
	var totalNUmberOfRow = 0;
	for(i=1; i<rowValue.length; i++){
		totalNUmberOfRow = parseInt(rowValue[i])+ (totalNUmberOfRow*10);
	}
	return totalNUmberOfRow;
}

function exportPdf(){
  var tableData = document.getElementById("excelTable"+$('li[class="active"]>a').attr('id'));
  var pdf = new jsPDF('p', 'pt', 'a2');
  var specialElementHandlers = {
    '#editor': function (element, renderer) {
      return true;
    }
  };
  pdf.fromHTML(
    tableData, 
    180, 
    50, 
    {
      'top': 5,
      'elementHandlers': specialElementHandlers
    }, 
    callBack, 
  );
  function callBack (dispose) {
    pdf.output('save',$('li[class="active"]>a').text()+'.pdf');

  }

}


