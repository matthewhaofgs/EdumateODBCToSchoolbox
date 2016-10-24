# ODBCToCSV
Connects to a cloud hosted instance of Edumate, and populate the import templates for schoolbox. 

Requirements:
  *Needs a working ODBC connection to the edumate database. 

   This has been tested with a sequelink ODBC driver
   https://www.progress.com/odbc/sequelink

   Once the driver has been installed, a DSN with the odbc driver needs to be created. 

  *In the root of the application folder, create a text file called 'connectionString.txt'
   The only line in the text file should be the connection string with the details for your edumate connection in the following format:

   Dsn=EdumateODBC;uid=odbcusername;hst=example.edumate.net;prt=19995;db=saas;pwd=pa$$word1
   (example only)

  *To automatically ftp the csv files to the schoolbox server create a file called 'uploadDetails.txt' in the root of the application folder
   In this file, put the ftp details for the schoolbox server in the format:

   schoolbox.example.com;ftpusername;pa$$word;ssh-rsa 2048 00:e2:e7:34:12:81:00:bd:a8:bf:00:9a:00:cc:75:ed
   (example only)
