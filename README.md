
# Database Lookup within Connect Data Mapper

* [Introduction](Introduction)
* [Microsoft SQL Lookup](Microsoft-SQL-Database-Lookup)
  * [Prerequisites for Microsoft SQL](Microsoft-SQL-Database-Lookup#Prerequisites-for-Microsoft-SQL)
  * [SQL Statements with JDBC]( Microsoft-SQL-Database-Lookup#Processing-SQL-Statements-with-JDBC )
  * [Application in Connect Data Mapper]( Microsoft-SQL-Database-Lookup#Application-in-Connect-Data-Mapper )


<object data="https://github.com/rodnnr/olconnect-dblookupindm/blob/master/SQL_Lookup_in_Connect_DataMapper.pdf" type="application/pdf" width="700px" height="700px">
    <embed src="https://github.com/rodnnr/olconnect-dblookupindm/blob/master/SQL_Lookup_in_Connect_DataMapper.pdf">
        <p>Download the guide in PDF format: <a href="https://github.com/rodnnr/olconnect-dblookupindm/blob/master/SQL_Lookup_in_Connect_DataMapper.pdf">Download SQL Lookup in Connect DataMapper</a>.</p>
    </embed>
</object>

## Introduction

Sometimes, the information and data that is required to build a form is scattered across multiple data sources such as Database Management Systems (DBMS) and it is often a challenge to get all the required data in one source. Luckily, the OLConnect Data Mapper can use the JDBC (Java Database Connectivity) driver to connect to these third-party DBMS. JDBC makes it possible to establish a connection with a database, send SQL statements and process the results.
This guide aims to provide the steps that are required to retrieve data from a remote Microsoft SQL Server using the “Action” step in the Data Mapper.

## Microsoft SQL Database Lookup

### Prerequisites for Microsoft SQL
The first part of this guide aims to provide the steps that are required to retrieve data from a remote Microsoft SQL Server using the “Action” step in the Data Mapper
Prerequisites for Microsoft SQL
To ensure a successful connection to Microsoft SQL Server, the below configuration should be made on the machine running your SQL Server instance:

1.	You need to enable mixed mode security when you install Microsoft SQL Server so that you can connect using a user name and password.
2.	The JDBC driver only works with the TCP/IP protocol which is disabled by default on SQL Express. You need to enable the TCP/IP Protocol from the SQL Server Configuration Manager that ships with SQL Express and re-start the SQL Server service. Look under SQL Server Network Configuration ->Protocols for SQLEXPRESS-> TCP/IP->Enable.
3.	Set the TCP Port by right-clicking on “TCP/IP”, then click on “Properties” and clicking on the “IP Addresses” tab. The Default port is 1433.
![TCP Port](https://github.com/rodnnr/oldatamapper-dblookupindm/blob/master/assets/1.png)

4.	Restart the SQL Server Service to apply the changes.
5.	Add TCIP Port 1433 and UDP Port 1434 in your firewall Inbound Rules
![Add TCP 1433 and UDP 1434 to firewall inbound rules](https://github.com/rodnnr/oldatamapper-dblookupindm/blob/master/assets/2.png)
6.	Allow remote connection to the SQL Server. Log into your Microsoft SQL Server instance from SQL Server Management Studio. Right click the server and click on Properties.

    ![SQL Server properties](https://github.com/rodnnr/oldatamapper-dblookupindm/blob/master/assets/3.png)

Navigate to Connections and ensure that Allow remote connections to this server is checked. 
       ![Allow remote connections](https://github.com/rodnnr/oldatamapper-dblookupindm/blob/master/assets/4.png)

### Processing SQL Statements with JDBC

In this example, we are supplied with a CSV data file, which contains information about customers’ orders’ details such as the OrderID, CustomerID and Shipping Address and dates details; but what is missing from the CSV data file is the actual customer’s name (CompanyName) and a contact name (ContactName). This information available in the Customers Table of the Northwind database, which resides on a Microsoft SQL Server.

In general, to process any SQL statement with JDBC, you follow these steps:
*	Establish a connection to the database on the SQL Server.
*	Create a statement.
*	Execute the query.
*	Process the ResultSet object.
*	Close the connection.
   
**Example:** The below script establishes a connection to the Northwind database on a local SQL Server, queries the [Customers] table using the CustomerID field taken from the CSV data file. 

```javascript
let connectionURL = "jdbc:sqlserver://localhost\\SQLEXPRESS:1433;integratedSecurity=false;databaseName=Northwind";
let custID = data.extract('CustomerID',0).trim();

if(custID){
  let sqlConnection = db.connect(connectionURL, "username", "password");
  let sqlQuery = "SELECT [CompanyName],[ContactName] FROM [dbo].[Customers] where CustomerID=" + "'" + custID + "'";

  resultSet = sqlConnection.createStatement().executeQuery(sqlQuery);
  resultSet.next();

  resultSet.close();
  sqlConnection.close();
}
```

 The documentation on how to build the Connection URL is available on the [MSDN website](https://docs.microsoft.com/en-us/sql/connect/jdbc/building-the-connection-url)

### Application in Connect Data Mapper

In this example, we have a CSV data file of customers’ orders, which does not include the actual customers name and contact details. Instead, this information resides on a remote Microsoft SQL Database table. This example will demonstrate how the companyName and contactName fields can be retrieved from the [dbo].[Customers] table of the [Northwind] database.

The first step is to load the CSV in the Data Mapper, then Add Extract Step to perform a standard:
![CustomerID](https://github.com/rodnnr/oldatamapper-dblookupindm/blob/master/assets/5.png) 

Now that the order details have been extracted, we can use the customerID field common between the CSV data file and the [dbo].[Customers] table in a SQL SELECT statement to then retrieve the corresponding customer’s details such as companyName and contactName.
![SQL](https://github.com/rodnnr/oldatamapper-dblookupindm/blob/master/assets/6.png)

For instance, in Microsoft SQL Management Studio, we would normally run the below SELECT statement to retrieve the _CompanyName_ and _ContactName_ of the order where the _CustomerID_ is _‘HANAR’_:
```sql
SELECT CustomerID, CompanyName, ContactName FROM [Northwind].[dbo].[Customers]
WHERE	CustomerID = 'HANAR'
```
![SQL Query](https://github.com/rodnnr/oldatamapper-dblookupindm/blob/master/assets/7.png)
We can execute the same command in the Data Mapper with the following steps:
After the above Extract step, Add Action step to open a connection to the SQL database, then query relevant table and save the result of the query in the resultSet with the following script:

```javascript
let connectionURL = "jdbc:sqlserver://;serverName=localhost\\OLSupport;integratedSecurity=true;databaseName=Northwind";
//let connectionURL = //"jdbc:sqlserver://;serverName=localhost\\OLSupport;integratedSecurity=false;databaseName=Northwind";

var custID = data.extract('CustomerID',0).trim();

if(custID){
  let sqlConnection = db.connect(connectionURL, "", ""); //integratedSecurity=true
  //let sqlConnection = db.connect(connectionURL, "sqlUser", "sqlPassword"); //integratedSecurity=false

  let sqlQuery = "SELECT [CompanyName],[ContactName] FROM [dbo].[Customers] where CustomerID=" + "'" + custID + "'";

  resultSet = sqlConnection.createStatement().executeQuery(sqlQuery);
  resultSet.next();
}
```
![ConnectionURL](https://github.com/rodnnr/oldatamapper-dblookupindm/blob/master/assets/8.png)

Once we have a _resultSet_, we can now extract the relevant _CompanyName_ and _ContactName_ from it. It can be necessary to make sure the resultSet is not empty. After **resultSet.next()**, check the current row with **getRow()**. If it returns 0, then no data was included in the resultSet. If _getRow()>=1_, then you have at least one row.

To do this, simply add an _Extract_ Step. A new field named Field is automatically added in the Data Model. To rename the field, click on the Order and Rename icon in the Field Definition window under the Extract Step properties.

Make sure the Field Definition Mode is set to JavaScript and insert the following JavaScript Ternary expression to extract the _CompanyName_:
```javascript
let companyName;
companyName = resultSet.getRow() ? resultSet.getString("companyName") : "";
```
To extract the _ContactName,_ right-click on the above _Extract_ step, select _“Add a Step”_ and then select _“Add Extract Field”_. Rename the field, set its _Definition Mode_ to JavaScript and insert the following expression:
```javascript
let contactName;
contactName = resultSet.getRow() ? resultSet.getString("ContactName") : "";
```
Note: In case where the _ResultSet_ is empty, instead of returning the empty string, you may choose to output a default value or run another expression.
![Extract](https://github.com/rodnnr/oldatamapper-dblookupindm/blob/master/assets/9.png)

Finally close the ResultSet and the SQL connection as well:
```javascript
resultSet.close();
sqlConnection.close();
```
![close connection](https://github.com/rodnnr/oldatamapper-dblookupindm/blob/master/assets/10.png)



## Microsoft Access, Microsoft Excel and CSV Lookup

### Prerequisites

Since OL Connect is a 64-bit application, we will need the following two prerequisites
*	Microsoft Access Database Engine 2010 for 64-bit Windows
*	64-bit DSN

### Microsoft Access Database Engine 2010 for 64-bit Windows

Download and install Microsoft Access Database Engine 2010 for 64-bit Windows from the following link:
http://www.microsoft.com/en-us/download/details.aspx?displaylang=en&id=13255

Note that launching the installation of a Microsoft Access Database Engine in the usual way, on a machine with an Office installation architecture different from the current one (e.g. 32-bit on 64-bit), may cause the installation to fail. 

![Failed Install of Access Database engine](https://github.com/rodnnr/oldatamapper-dblookupindm/blob/master/assets/12.png)

To have it run properly, you need to launch it from a command line with the “/passive” argument specified:
![install with passive mode](https://github.com/rodnnr/oldatamapper-dblookupindm/blob/master/assets/13.png)

### Create 64-bit Data Source Names (DSN)
To create a 64-bit System DSN:

*	Open the Windows Control Panel and navigate to Administrative Tools > ODBC Data Source(64-bit)
*	Click on the “System DSN” tan and click on Add
*	Select the driver that corresponds to your file type 
	* For Access files, select Microsoft Access Driver (*.mdb, *.accdb)
	* For CSV files, select Microsoft Text Driver (*.txt, *.csv)
	* For Excel files, select Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)
*	Click Finish to select your Access, CSV or Excel file
*	Give your DSN a name. We will call it “SalesRep” for this example
 ![DSNs](https://github.com/rodnnr/oldatamapper-dblookupindm/blob/master/assets/14.png)

### Building JDBC-ODBC Connection URL
The JDBC-ODBC Bridge allows Java applications to use the JDBC API with many existing ODBC drivers.
The general form of the connection URL for Access, Excel and CSV files is:
```javascript
"jdbc:odbc:DSN_Name"
```
 
where:  _DSN_Name_ is the ODB Data Source Name;
Hence, to connect to either of the above database, we can use the following code:
```javascript
//To connect to a Excel database
let connectionURL = "jdbc:odbc:SalesRepsExcel";
let excelConnection = db.connect(connectionURL,"","");

/* To connect to CSV database
let connectionURL = "jdbc:odbc:SalesRepsCSV";
let csvConnection = db.connect(connectionURL,"","");
*/

/* To connect to Access database
let connectionURL = "jdbc:odbc:SalesRepsAccess";
let accessConnection = db.connect(connectionURL,"","");
*/
```
 
### Querying a Microsoft Access Database

Once we have a JDBC connection object, we can then use it to create statements and execute queries, which will return a result set. For a Microsoft Access database, the code is as follow
```javascript
let connectionURL = "jdbc:odbc:SalesRepsAccess";
let repID = record.fields.RepID;

if(repID){
   let accessConnection = db.connect(connectionURL,"","");
   let accessQuery = "SELECT * FROM SalesReps WHERE RepID=" + "'" + repID + "'";
   resultSet = accessConnection.createStatement().executeQuery(accessQuery);
   resultSet.next();
```
 

### Querying a Microsoft Excel Database
Querying a Microsoft Excel database is similar with the procedure for Access, the only difference here is that is that Microsoft Excel Sheet name you are querying from must be followed by the $ sign and enclosed in square brackets:
```javascript
let connectionURL = "jdbc:odbc:SalesRepsExcel";
let repID = record.fields.RepID;

if(repID){
   let excelConnection = db.connect(connectionURL,"","");
   let excelQuery = "SELECT * FROM [SalesReps$] WHERE RepID=" + "'" + repID + "'";
   resultSet = excelConnection.createStatement().executeQuery(excelQuery);
   resultSet.next();
}
```
 

### Querying a CSV file
The code for querying a CSV file is similar with the only difference that the table name to query from is the actual CSV file name and must be enclosed in double quotes:
```javascript
let connectionURL = "jdbc:odbc:SalesRepsCSV";
let repID = record.fields.RepID;

if(repID){
   let csvConnection = db.connect(connectionURL,"","");
   let csvQuery = "SELECT * FROM \"SalesReps.csv\" WHERE RepID=" + "'" + repID + "'";
   resultSet = csvConnection.createStatement().executeQuery(csvQuery);
   resultSet.next();
}
```
 

### Retrieving data from a Result Set
The above query return a result set object. We can then use the methods of the _ResultSet_ object, such as _getString()_, to retrieve the desired data.

For example, use the following code in an_ Extract_ step to retrieve the value of the current Rep’s Email from the Access, Excel or CSV database and assign to a field in the Data Mapper:
```javascript
let repEmail;
repEmail = resultSet.getRow() ? resultSet.getString("Email") : "";
``` 

### Closing JDBC Connections
It is good practice to explicitly close all connections to the database to end each database session rather than leaving the task to the Java’s garbage collection.
To close the above opened connection, you should call close() method as follows:
```javascript
resultSet.close();
accessConnection.close();
//excelConnection.close();
//csvConnection.close();
```
Reference JDBC Basics
https://docs.oracle.com/javase/tutorial/jdbc/basics/


