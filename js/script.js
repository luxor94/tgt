let lists
let base

//импорт файла из xlsx в JSON
var ExcelToJSON = function() {

  this.parseExcel = function(file) {
    var reader = new FileReader();

    reader.onload = function(e) {
      var data = e.target.result;
      var workbook = XLSX.read(data, {
        type: 'binary'
      });
      workbook.SheetNames.forEach(function(sheetName) {
        // Here is your object
        var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
        var json_object = JSON.stringify(XL_row_object);
        lists = (JSON.parse(json_object));
        jQuery( '#xlx_json' ).val( json_object );
      })
    };

    reader.onerror = function(ex) {
      console.log(ex);
    };

    reader.readAsBinaryString(file);

  };
};


var ExcelToJSON_Base = function() {

    this.parseExcel = function(file) {
      var reader = new FileReader();
  
      reader.onload = function(e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, {
          type: 'binary'
        });
        workbook.SheetNames.forEach(function(sheetName) {
          // Here is your object
          var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
          var json_object = JSON.stringify(XL_row_object);
          base = (JSON.parse(json_object));
          jQuery( '#xlx_json' ).val( json_object );
        })
      };
  
      reader.onerror = function(ex) {
        console.log(ex);
      };
  
      reader.readAsBinaryString(file);
  
    };
  };

function replacement() {
  for (let i = 0; i < lists.length; i++) {
    for (let j = 0; j < base.length; j++)
    {
      let a = base[j].Наименование
      if (lists[i]['17.Вид номенклатуры'] == a) {
        lists[i]["18.Вес нетто "] = base[j].Нетто;
        lists[i]["19.Вес брутто "] = base[j].Брутто;
        lists[i]["20.ТНВЭД"] = base[j].ТНВЭД;
      }   
    }
  }

  //экспорт в json
  function downloadObjectAsJson(lists, exportName) {
    var dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(lists));
    var downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", exportName + ".json");
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
  }

  var stockList = XLSX.utils.json_to_sheet(lists)
  var wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, stockList, 'list1')
  XLSX.writeFile(wb, 'book.xlsx');
}
