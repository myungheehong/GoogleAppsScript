// jdbc mysql db conneciton
// mssql connection : jdbc:Microsoft:sqlserver://localhost:1433;databasename=DB명    
// rowdata 시트에서 주소, 포트, 데이터베이스, 사용자, 패스워드 정보 저장
// address = IP 주소, address_port = database 포트
// db = database name, user = database login 
// userpwd = database login password, dbsystem : mysql or mssql select 

function connmysqldb() {   

    var address = SpreadsheetApp.getActive().getSheetByName('rowdata').getRange(5,2).getValue();
    var address_port = SpreadsheetApp.getActive().getSheetByName('rowdata').getRange(5,10).getValue();
    var db = SpreadsheetApp.getActive().getSheetByName('rowdata').getRange(5,8).getValue();
    var user = SpreadsheetApp.getActive().getSheetByName('rowdata').getRange(5,4).getValue();
    var userpwd = SpreadsheetApp.getActive().getSheetByName('rowdata').getRange(5,6).getValue();
    var dbsystem = SpreadsheetApp.getActive().getSheetByName('rowdata').getRange(4,2).getValue();
    var maxcolumn = SpreadsheetApp.getActive().getSheetByName('rowdata').getRange(7,5).getValue();
    var maxrow = SpreadsheetApp.getActive().getSheetByName('rowdata').getRange(7,2).getValue();
    //Logger.log(address + ' ' + address_port + ' ' + db + ' ' + user + ' ' + userpwd + ' ' + dbsystem + ' ' + maxcolumn +' ' + maxrow);
  
    var connQuery = SpreadsheetApp.getActive().getSheetByName('rowdata').getRange(6,2).getValue();
    //Logger.log(connQuery);
  
    var connection = null;
    var result = null;
    var SQLstatement = null;

    //if (address == '') {
    //  Browser.msgBox('IP주소가 입력되지 않았습니다');
    //} 
    
    // 선택된 msyql, mssql 처리에 대한 확인
    if (dbsystem == 'mysql') {
      var dbUrl = 'jdbc:mysql://' + address + ':' + address_port + '/' + db;
    } else if (dbsystem == 'mssql') {
      var dbUrl = 'jdbc:Microsoft:sqlserver://' + address + ':' + address_port + ';' + 'databasename=' + db;
    }
    //Logger.log(dbUrl);  
  
    //database connection function
    var connection = Jdbc.getConnection(dbUrl, user, userpwd);
  
    //database query string
    var SQLstatement = connection.createStatement();
  
    //max reader fetch
    SQLstatement.setMaxRows(maxrow);                
    
    var result = SQLstatement.executeQuery(connQuery);
  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var getcount = result.getMetaData().getColumnCount(); 
    //Logger.log(getcount);

    var row = 0;
    
    // always sheet clear function call
    clearsheet();
   
    var cell = ss.getRange('A09');
    
    for (var i=0; i < getcount; i++){
      cell.offset(row, i).setValue(result.getMetaData().getColumnName(i+1));
    }
      
    //var cell = ss.getRange('A10');
  
    var row = 1;
    while(result.next()) {
      for(var i = 0; i < maxcolumn; i++) {
        cell.offset(row, i).setValue(result.getString(i+1));
      }
      row++;
    }
    result.close();
    SQLstatement.close();
    connection.close();
}