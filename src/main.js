var fileUpload = document.getElementById('fileUpload');
var fileUpload2 = document.getElementById('fileUpload2');

var dataArray = [];

function handleFile(e) {
  var files = e.target.files, f = files[0];
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = new Uint8Array(e.target.result);
    var workbook = XLS.read(data, {type: 'array'});
    var worksheet = workbook.Sheets['PRRAS'];
    for(var i=12; i<33; i++) {
      if (typeof(worksheet['E'+i]) !== 'undefined' && typeof(worksheet['E'+i]['f']) === 'undefined') {
        var obj = {};
        obj[worksheet['C'+i]['v']] = worksheet['E'+i];
        dataArray.push(obj);
      }
    }
    reader.abort();
  };
  reader.readAsArrayBuffer(f);
}
function handleFile2(e) {
  var files = e.target.files, f = files[0];
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = new Uint8Array(e.target.result);
    var workbook = XLSX.read(data, {type: 'array'});
    var worksheet = workbook.Sheets['PRRAS'];
    for(var i=0; i<dataArray.length; i++) {
      var key = parseInt(Object.keys(dataArray[i])[0]);
      if (worksheet['C'+(key+11)]['v'] == key) {
        worksheet['D'+(key+11)] = dataArray[key];
      }
    }
    XLSX.writeFile(workbook, 'out.xls');
    reader.abort();
  };
  reader.readAsArrayBuffer(f);
}

fileUpload.addEventListener('change', handleFile, false);
fileUpload2.addEventListener('change', handleFile2, false);