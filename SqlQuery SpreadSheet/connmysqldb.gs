/*

���α׷� ����
1. �� ���α׷��� ����� �Է��� ���� ����, ������ �������� �����͸� �������� ����Դϴ�. 
2. �����ͺ��̽� ������ JDBC�� ����մϴ�.
3. mysql connection -> jdbc mysql db conneciton
4. mssql connection -> jdbc:Microsoft:sqlserver://localhost:1433;databasename=DB��

��ɼ���




��������
2017.02.22 jdbc ���ῡ ���� try catch ó�� ��� �߰�
2017.02.27 sheetname �������� ���������� ������ ����
2017.02.27 jdbc ����� ���� �κ��� ���� ������ ���� try catch ����
2017.03.01 server ���� �������� ����� getserverinfo�� ������(��� Ȯ�� ��)
2017.03.02 GetServerInfo() ��ü ������� ����
           GetParameter() �߰�


�̽�����
1. mssql �ý��� ������ ��ȸ �� ���� ����(�����ذ�Ϸ�)


*/

// 2017.03.02 �Լ� ����� ���� ���������� ������ 
var sheet_name = SpreadsheetApp.getActiveSheet().getName();

function conndatabase() {   
    
  // ���� �޽��� ó���� ���� ��� ����  
  var error_message = '�Էµ��� �ʾҽ��ϴ� �ٽ� Ȯ�����ּ���'; //���� �޽��� ó���� ���� ��� 
  var sheet_name = SpreadsheetApp.getActiveSheet().getName();
  //Logger.log(sheet_name)
   
  //var address = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(5,2).getValue();
  //var user = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(5,4).getValue();
  //var userpwd = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(5,6).getValue();
  //var dbsystem = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(4,2).getValue();
  //var address_port = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(5,10).getValue();
  //var db = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(5,8).getValue();

  // 2017.03.01 ������ �������� ���� �Լ��� ������(������ ���뼺) -> ���� ��ü�� ���� ����
  
  var getserver = null;
  var getparameter = null;
  
  var getserver = new GetServerInfo();
  var getparameter = new GetParameter();

  //var address = getserverinfo('address');
  //var address_port = getserverinfo('address_port');
  var address = getserver.address;
  var address_port = getserver.address_port;
  var db = getserver.db;
  var user = getserver.user;
  var userpwd = getserver.userpwd;
  var dbsystem = getserver.dbsystem; 

  var maxcolumn = getparameter.maxcolumn;
  var maxrow = getparameter.maxrow;
  var connQuery = getparameter.connQuery;
  
  //var maxcolumn = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(7,5).getValue();
  //var maxrow = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(7,2).getValue();  
  //var connQuery = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(6,2).getValue();
  //Logger.log(connQuery);
  
  var connection = null;
  var result = null;
  var SQLstatement = null;
  
  if (address == ''){
    Browser.msgBox('ip�ּҰ� ' + error_message);
    return false;
  }
  
  if (address_port == '') {
    Browser.msgBox('��Ʈ�� ' + error_message); 
    return false;    
  }
  
  if (user == '') {
    Browser.msgBox('���̵� ' + error_message); 
    return false;    
  }
  
  if (userpwd == '') {
    Browser.msgBox('�н����� ' + error_message); 
    return false;        
  }
  
  if (db == '') {
    Browser.msgBox('�����ͺ��̽� ' + error_message); 
    return false;
  }
  
  // ���õ� msyql, mssql ���ÿ� ���� ó��
  if (dbsystem == 'mysql') {
    var dbUrl = 'jdbc:mysql://' + address + ':' + address_port + '/' + db;
  } else if (dbsystem == 'mssql') {
    var dbUrl = 'jdbc:sqlserver://' + address + ':' + address_port + ';' + 'databasename=' + db;
  }
    
  try {  
      //database connection function
    var connection = Jdbc.getConnection(dbUrl, user, userpwd);
  }
  catch (exception){
    Browser.msgBox('�����ͺ��̽� ���ῡ ������ �߻��߽��ϴ�. �ٽ� Ȯ�����ּ���');
    return false;    
  }
  
  try {
    //database query string
    var SQLstatement = connection.createStatement();
  
    //max reader fetch
    SQLstatement.setMaxRows(maxrow);                
    
    //Sql Query run
    var result = SQLstatement.executeQuery(connQuery);
  }

  catch (exception){
    Browser.msgBox('�����ͺ��̽� ������ ������ �߻��߽��ϴ�. �ٽ� Ȯ�����ּ���');
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
    
  var row = 1;
  while(result.next()) {
    for(var i = 0; i < getcount; i++) {
      cell.offset(row, i).setValue(result.getString(i+1));
    }
    row++;
  }
  result.close();
  SQLstatement.close();
  connection.close();
}

function GetServerInfo(){ 
  this.address = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(5,2).getValue();
  this.address_port = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(5,10).getValue();
  this.db = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(5,8).getValue(); 
  this.user = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(5,4).getValue();
  this.userpwd = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(5,6).getValue();
  this.dbsystem = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(4,2).getValue();
}

function GetParameter() {
  this.maxcolumn = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(7,5).getValue();
  this.maxrow = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(7,2).getValue();  
  this.connQuery = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(6,2).getValue();
  //var maxcolumn = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(7,5).getValue();
  //var maxrow = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(7,2).getValue();  
  //var connQuery = SpreadsheetApp.getActive().getSheetByName(sheet_name).getRange(6,2).getValue();
}