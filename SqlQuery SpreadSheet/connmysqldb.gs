/*
프로그램 목적
- 이 프로그램은 사용자 입력한 서버 정보, 쿼리를 기준으로 데이터를 가져오는 기능입니다. 
- 데이터베이스 연결은 JDBC를 사용합니다.
- mysql connection -> jdbc mysql db conneciton
- mssql connection -> jdbc:Microsoft:sqlserver://localhost:1433;databasename=DB명

기능설명


수정사항
2017.02.22 jdbc 연결에 대한 try catch 처리 기능 추가

*/

function connmysqldb() {   
  
    var error_message = '입력되지 않았습니다 다시 확인해주세요';
  
    // sheet 입력된 값 가져오기
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
  
  if (address == ''){
    Browser.msgBox('ip주소가 ' + error_message);
    return false;
  }
  
  if (address_port == '') {
    Browser.msgBox('포트가 ' + error_message); 
    return false;    
  }
  
  if (user == '') {
    Browser.msgBox('아이디 ' + error_message); 
    return false;    
  }
  
  if (userpwd == '') {
    Browser.msgBox('패스워드 ' + error_message); 
    return false;        
  }
  
  if (db == '') {
    Browser.msgBox('데이터베이스 ' + error_message); 
    return false;
  }
  
    // 선택된 msyql, mssql 선택에 대한 처리
    if (dbsystem == 'mysql') {
      var dbUrl = 'jdbc:mysql://' + address + ':' + address_port + '/' + db;
    } else if (dbsystem == 'mssql') {
      var dbUrl = 'jdbc:Microsoft:sqlserver://' + address + ':' + address_port + ';' + 'databasename=' + db;
    }

  try {  
    //database connection function
    var connection = Jdbc.getConnection(dbUrl, user, userpwd);
  
    //database query string
    var SQLstatement = connection.createStatement();
  
    //max reader fetch
    SQLstatement.setMaxRows(maxrow);                
    
    //Sql Query run
    var result = SQLstatement.executeQuery(connQuery);
  }
  catch (exception){
    Browser.msgBox('데이터베이스 연결에 오류가 발생했습니다. 다시 확인해주세요');
    return false;    
  }
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