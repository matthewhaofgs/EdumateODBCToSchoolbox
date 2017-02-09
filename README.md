===

This is now included in https://github.com/matthewhaofgs/OFGS-Account-Sync

Newer changes will be there

===


# ODBCToCSV
Connects to a cloud hosted instance of Edumate, and populate the import templates for schoolbox. 

Requirements:
  *Needs a working ODBC connection to the edumate database. 

   This has been tested with a sequelink ODBC driver
   https://www.progress.com/odbc/sequelink

   Once the driver has been installed, a DSN with the odbc driver needs to be created. 

  *In the root of the application folder, create a text file called config.ini
   config.ini will need the following
   
   connectionstring=  (the connection string for the ODBC connection)
   uploadserver= (FTP details to upload the csv files)
   studentemaildomain= (email domain for students)
  
  
  example config.ini file:
  
uploadserver=schoolbox.example.com;ftpusername;pa$$word;ssh-rsa 2048 00:e2:e7:34:12:81:00:bd:a8:bf:00:9a:00:cc:75:ed
connectionstring=Dsn=EdumateODBC;uid=odbcusername;hst=example.edumate.net;prt=19995;db=saas;pwd=pa$$word1
studentemaildomain=@example.com 
  
 
