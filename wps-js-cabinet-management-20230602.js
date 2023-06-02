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
	var endLine = 181;	
	
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
	var usedList = getOneCabinetInfo(selectedCabinetCode);
		
//	生成机柜占用示意图
	//	第一行：标题
	Cells.Item(1,startColumn).Formula = selectedCabinetCode;
	//	内容： 从第二行开始
	var units = 44;
	var num = 44;
	for(var i = 2 ; i <= units + 1; i++)
	{
//		生成空的占用		
		Cells.Item(i,startColumn).Formula = num;
//		判断是否占用
		var current = Number.parseInt(Cells.Item(i,startColumn));	
		if(usedList.includes(current))
		{
			Range(Cells.Item(i,startColumn),Cells.Item(i,startColumn)).Interior.Color = 65535;			
		}
		num--;
	}
}

//生成多个机柜的占用图
function showMultCabinetState()
{
	var cabinetList = ["3-1G01","3-1G02","3-1G03","3-1G04","3-1G05","3-1G06","3-1G07","3-1G08",
		"3-1G09","3-1G10","3-1G11","3-1G12","3-1G13","3-1G14","3-1G15"];
	var count = cabinetList.length;
	var startcolum = 51;
	for(var i = 0; i < count; i++ )
	{		
		showOneCabinetState(cabinetList[i],startcolum);
		startcolum += 2;
	}
}


/**
 * test Macro
 */
function test()
{
//	test = extractNumberFromString("12~15");

//	getOneCabinetInfo("3-1G01");	

//	showOneCabinetState("3-1G02",52);
	
	showMultCabinetState();

}







