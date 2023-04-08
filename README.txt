====> ms-sql branch, for C3 <=======

Imports/updates customers/jobs/specs and syncing costings between the Ink System and C3. 

The InkToC3 (replaces InkToJobs) is written in Visual Basic, originally by Inovex, but now maintained by Sean. 
The source code is on github, mssql brwanch, see https://github.com/Boran/job-InkToJobs/tree/mssql

See also https://boranp.sharepoint.com/:w:/r/sites/allit/_layouts/15/Doc.aspx?sourcedoc=%7BC126E558-F654-40B2-B892-F25AAD48B8DC%7D&file=Ink_inovex%20Notes.docx&action=default&mobileredirect=true

Costings flow 
When a job is created then the job is tagged at being available for dispense. 
At this time the record is created on the costing reports details table for that works order. 
As ink is dispensed and the returns issued/returned then the record for the job is updated. 
At the end of the job the operator needs to select Order Complete. 
When he does this the date the order is completed is filled in on the Costings reports details table. 
Completed costings are transferred to the job system daily (full convert)

Development environment
. Visual Basic 6
. Ms-sql or Mysql ODBC connector
. ODBC??

Installation: The InkToJobs program is installed on the Ink PC: 
copy it to C:\InkToJobs  
install the tool in “MySQL driver 5.1” folder 
click on InkToJobs to start an import, take a few minutes.  
If there is a problem, maybe the DB settings in “Setups” file are wrong. 
In the Control Panel, add a Scheduled task to run C:\ink_to_jobsystem\InkToJobs.exe every morning at 07:00 


Config:
. edit setups.txt
  It is read sequentially, no comments.
. Example, point to DB, and 5378 directory

Database Name : C3_Boran
Driver : sqloledb
Host: bpc3.boran.ie
Password : <PW>
UserName : inkcost
Port : 1433
DispenserDatabasePath : C:\1029921B\1029921B_BE.mdb
Default Print Process : Flexo
Timer Count Value In Seconds : 1800
Days To Look Back : 30
DebugLevel : 0
----
To automate: DebugLevel : 0
Driver : mssqloledb
DebugLevel : 2
DispenserDatabasePath : C:\5378\5378_2007BE.mdbpoint to DB, and 5378 directory
