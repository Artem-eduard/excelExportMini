  
var excelExport = {
    config: {
        fileName: "report",
        extension: ".xlsx",
        sheetName: "Sheet1",
        fileFullName: "report.xlsx",
        header: true,
        createEmptyRow: true,
        maxCellWidth: 20
    },
    worksheetObj: {},
    rowCount: 0,
    wsColswidth: [],
    merges: [],
    worksheet: {},
    range: {},
    init: function (options) {
        this.reset();
        if (options) {
            for (var key in this.config) {
                if (options.hasOwnProperty(key)) {
                    this.config[key] = options[key];
                }
            }
        }
        this.config['fileFullName'] = this.config.fileName + this.config.extension;
    },
	
  export: function(JSONData, ReportTitle) {
  var div = document.createElement('div'),
                name,
                xlsxRows = [],
                rows;
            div.style.display = 'none';
            document.body.appendChild(div);
            rows = JSONData;
            
                name = ReportTitle;
				var arr = [];
				var len = JSONData.data.length;
				
			if (JSONData.hasOwnProperty("labels"))
			{
				var arr_row = [];
				for (var key in JSONData.labels)
				{
                
					var cellObj = JSONData.labels[key];
					
						arr_row.push({
						value: cellObj,
						type: "string"
						});
				  
				}
			}
			arr.push(arr_row);
            if (JSONData.hasOwnProperty("data"))
			{
				for (var r in JSONData.data)
				{
					
					var arr_row=[];
					
					for (var key in JSONData.data[r])
					{
						var cellObj = JSONData.data[r][key];
						arr_row.push({
						value: cellObj,
						type: "string"
						});
					}
					arr.push(arr_row);
					
				}
			}

            window.zipcelx({
                filename: name,
                sheet: {
                     data: arr
                }
            });
  }
		}
		