# SQLDocumenter

#URL: https://slavasql.blogspot.com/2013/11/stored-procedure-to-document-database.html

#Usage:
1. Just put name of the database you want to document in the third line of the script and run the script.
DECLARE @DBName VARCHAR(128) = '<YOUR DATABASE NAME>'
2. Run the script
3. Copy-paste script's results to .HTML file
4. View that file in any browser (Tested in Chrome/IE/FireFox)
  
#Restriction: Does not work on SQL Server versions earlier than 2005. Not tested in 2017 & 2019.
