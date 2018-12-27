
## How to extract database information in a Connect Data Mapping Configuration


<object data="https://github.com/rodnnr/olconnect-dblookupindm/blob/master/SQL_Lookup_in_Connect_DataMapper.pdf" type="application/pdf" width="700px" height="700px">
    <embed src="https://github.com/rodnnr/olconnect-dblookupindm/blob/master/SQL_Lookup_in_Connect_DataMapper.pdf">
        <p>Please download the guide in PDF format: <a href="https://github.com/rodnnr/olconnect-dblookupindm/blob/master/SQL_Lookup_in_Connect_DataMapper.pdf">Download SQL Lookup in Connect DataMapper</a>.</p>
    </embed>
</object>

## For Microsoft SQL

### Open the SQL Connection and Query Database with the Action step

```javascript
var connectionURL = "jdbc:sqlserver://;serverName=localhost\\OLSupport;integratedSecurity=true;databaseName=Northwind";
//var connectionURL = "jdbc:sqlserver://;serverName=localhost\\OLSupport;integratedSecurity=false;databaseName=Northwind";

var custID = data.extract('CustomerID',0).trim();

if(custID){
var sqlConnection = db.connect(connectionURL, "", ""); // for  integratedSecurity=true
//var sqlConnection = db.connect(connectionURL, "sqlUser", "sqlPassword"); // for integratedSecurity=false

var sqlQuery = "SELECT [CompanyName],[ContactName] FROM [dbo].[Customers] where CustomerID=" + "'" + custID + "'";

resultSet = sqlConnection.createStatement().executeQuery(sqlQuery);
resultSet.next();
}
```

### Extract the desired fields from the SQL Query resultSet with the Extract step
```javascript
let companyName;
companyName = resultSet.getRow() ? resultSet.getString("companyName") : "";
```

### Close the resultSet and SQL connection with the Action step
```javascript
resultSet.close();
sqlConnection.close();
```

## For Access, CSV and Excel

**Connecting and querying the database:**

```javascript
//To connect to a Excel database
var connectionURL = "jdbc:odbc:SalesRepsExcel";
var excelConnection = db.connect(connectionURL,"","");

/* To connect to CSV database
var connectionURL = "jdbc:odbc:SalesRepsCSV";
var csvConnection = db.connect(connectionURL,"","");
*/

/* To connect to Access database
var connectionURL = "jdbc:odbc:SalesRepsAccess";
var accessConnection = db.connect(connectionURL,"","");
*/


//Querying Excel database
var repID = record.fields.RepID;

if(repID){
var excelConnection = db.connect(connectionURL,"","");
var excelQuery = "SELECT * FROM [SalesReps$] WHERE RepID=" + "'" + repID + "'";
resultSet = excelConnection.createStatement().executeQuery(excelQuery);
resultSet.next();
}


/* Querying Access database
var repID = record.fields.RepID;

if(repID){
var accessConnection = db.connect(connectionURL,"","");
var accessQuery = "SELECT * FROM SalesReps WHERE RepID=" + "'" + repID + "'";
resultSet = accessConnection.createStatement().executeQuery(accessQuery);
resultSet.next();
}
*/


/*Querying CSV database

var connectionURL = "jdbc:odbc:SalesRepsCSV";
var repID = record.fields.RepID;

if(repID){
var csvConnection = db.connect(connectionURL,"","");
var csvQuery = "SELECT * FROM \"SalesReps.csv\" WHERE RepID=" + "'" + repID + "'";
resultSet = csvConnection.createStatement().executeQuery(csvQuery);
resultSet.next();
}
*/


 

```

**Retrieving the desired fields from the resultSet**

```javascript
let repEmail;
repEmail = resultSet.getRow() ? resultSet.getString("Email") : "";
```

** Closing the resultset and database connection**
```javascript
resultSet.close();
excelConnection.close();

/* 
accessConnection.close()
csvConnection.close();
*/
```
