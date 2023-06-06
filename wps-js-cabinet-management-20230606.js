Attribute Module_Name = "NewMacros"

// eg: "12~15" -> 12,13,14,15
function extractIntFromString(source)
{
	var ss = source.split("~");
	var startNumber = Number.parseInt(ss[0]);	
	var endNumber = Number.parseInt(ss.pop());
	
	var numberList = new Array();
	for(var i = startNumber; i <= endNumber; i++)
	{
		numberList.push(i);
	}	
	
	return numberList;
}


//从周浦机房数据表中，获取机柜编号、机架占用位置
//输入：机柜编号
//输出：机架占用的U号列表
function getOneCabinetInfo(selectedCabinetCode)
{
	var startLine = 2;
	var endLine = 200;	
	
//	机柜列
	var columnCabinet = 16;
//	U位置列
	var columnLocation = 19;	
	
//	已占用的位置
	var usedList = new Array();
	
	for(var i = startLine; i <= endLine; i++)
	{
		 var cabinetCode = new String(Cells.Item(i,columnCabinet));
		 if(cabinetCode.includes(selectedCabinetCode))
		 {
		 	var locationStr = new String(Cells.Item(i,columnLocation));
		 	var numbersLocation =extractIntFromString(locationStr);	
		 	usedList = usedList.concat(numbersLocation); 	
		 }		 			 	
	}
	return usedList;	
}

// 生成一个机柜的占用图表
//输入：机柜编号，生成占表的列数

function showOneCabinetState(selectedCabinetCode,startColumn)
{	
//	从数据表"周浦机房",获取原始unit占用信息
	Sheets.Item("周浦机房信息表").Activate();
	var usedList = getOneCabinetInfo(selectedCabinetCode);
			
//	在数据表"周浦机柜示意图"中,生成机柜占用示意图
	Sheets.Item("周浦机房图表").Activate();	
			
	//	第一行：标题
	Cells.Item(1,startColumn).Formula = selectedCabinetCode;
	//	内容： 从第二行开始
	var lineNumber = 44;
	var unitCode = 44;
	for(var i = 2 ; i <= lineNumber + 1; i++)
	{
//		生成空的unitCode: 44,43,...,1	
		Cells.Item(i,startColumn).Formula = unitCode;
//		判断是否占用
		var current = Number.parseInt(Cells.Item(i,startColumn));
			
		if(usedList.includes(current))
		{
			Range(Cells.Item(i,startColumn),Cells.Item(i,startColumn)).Interior.Color = 65535;			
		}
		unitCode--;
	}
}

//生成多个机柜的占用图
function showMultCabinetState(cabinetList,sheetName)
{
	//	清空数据表:"周浦机房图表"
	Sheets.Item("周浦机房图表").Activate();	
	Range("A1:U45").Select();	
	Selection.Clear();
	
	// 重新生成
	var count = cabinetList.length;
	var startcolum = 1;
	for(var i = 0; i < count; i++ )
	{		
		showOneCabinetState(cabinetList[i],startcolum);			
		startcolum += 1;
	}
	
//	格式化
	Range("A1:AM45").Select();
	Selection.HorizontalAlignment = xlHAlignCenter;
	Selection.Borders.Item(xlEdgeLeft).ColorIndex = xlColorIndexAutomatic;
	Selection.Borders.Item(xlEdgeTop).ColorIndex = xlColorIndexAutomatic;
	Selection.Borders.Item(xlEdgeBottom).ColorIndex = xlColorIndexAutomatic;
	Selection.Borders.Item(xlEdgeRight).ColorIndex = xlColorIndexAutomatic;
	Selection.Borders.Item(xlInsideVertical).ColorIndex = xlColorIndexAutomatic;
	Selection.Borders.Item(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic;	
	
	Range("A1:AM1").Select();
	Selection.Font.Bold = true;
	(obj=>{
		obj.Pattern = xlPatternSolid;
		obj.ThemeColor = 6;
		obj.TintAndShade = 0.800000;
		obj.PatternColorIndex = -4105;
	})(Selection.Interior);	
	
}


/**
 * test Macro
 */
function test()
{
//	test = extractNumberFromString("12~15");

//	getOneCabinetInfo("3-1G01");	

//	showOneCabinetState("3-1G02",52);
	
//	3-1G列
	var cabinetList = ["3-1G01","3-1G02","3-1G03","3-1G04","3-1G05","3-1G06","3-1G07","3-1G08","3-1G09","3-1G10",
//		"3-1G11","3-1G12","3-1G13","3-1G14","3-1G15",
		,"3-1H01","3-1H02","3-1H03","3-1H04","3-1H05","3-1H06","3-1H07","3-1H08","3-1H09","3-1H10",
		,"3-2E08","3-2E09","3-2E10","3-2E11","3-2E12","3-2H13","3-2E14","3-2E15",
		,"3-2F08","3-2F09","3-2F10","3-2F11","3-2F12","3-2F13","3-2F14","3-2F15"];		
	
	showMultCabinetState(cabinetList);	
}


/**
 * test Macro
 */
function test_workbook()
{
	alert(Application.Version)
	alert(Range("A1").Value2)
	alert(ActiveWorkbook.Worksheets.Count)
	
//	workbooks
	alert(ThisWorkbook.Name)
	alert(ActiveWorkbook.Name)
	alert(Workbooks.Item("资产清单2023-05-23.xlsm").FullName)
	alert(Workbooks.Item(1).FullName)
	
	alert(Workbooks.Item("资产清单2023-05-23.xlsm").Path)
	alert(ActiveWorkbook.Name)
	alert(Workbooks.Item("资产清单2023-05-23.xlsm").FullName)
	
//	worksheets
	alert(ActiveSheet.Name)	
	
}

/**
 * test Macro
 */
function test_worksheet()
{
//	worksheets
//	alert(ActiveSheet.Name)	
//	alert(Worksheets(1).Name)
//	alert(Worksheets.Item(2).Name)
//	alert(Worksheets.Item("总表").Name)
	
//	新建工作簿
	var wb =  ThisWorkbook;
	var tmpFile = Workbooks.Add();
	
//	新建工作表，默认位置为第一个工作表
	var tmpSheet =  tmpFile.Sheets.Add();
	tmpSheet.Name = "一日";
	
//	新建工作表，位置在Sheet1的后面
	var tmpSheet =  tmpFile.Sheets.Add(null,Sheets("Sheet1"));
	tmpSheet.Name = "二日";
	
//	删除工作表：Sheet1
	Sheets("Sheet1").Delete()
	
//	保存工作簿
	wb_dir = wb.Path + "\\" + "一月销售数据"
	tmpFile.SaveAs(wb_dir)
	tmpFile.Close()
	
}

/**
 * test Macro
 */
function test_range()
{
//	worksheets
	Range("A1:C3").Select()
	ActiveWorkbook.ActiveSheet.Range("A1:C3").Select();
	Range("A1:C3,D5:G8").Select()
	
	Cells(2,3).Select()
	Cells.Item(2,3).Select()
	
	alert(Range("B2").Value())
	alert(Range("B2").Value2)
	Range("B3").Value2 = 100	
	
}






