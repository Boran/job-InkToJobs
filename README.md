InkToJobs
---------

NOTE: the latest code, for MS-SQL (C3) is in the mssql branch. 
The code here is no loonger used.

The program “InkToJobs” is run each day @07:00 on the Ink PC “inovexmillpc” in the mIll that exports pricing to the Jobsystem, and Imports/updates customers/jobs/specs.
The InkToJobs is written in Visual Basic, originally by Inovex, but now maintained by Sean.

The master branch interfaces with the Job System on MYSQL.
See the mssql branch with interfaces to C3 on Microsoft SQL.

Queries:
- Three mysql views v_ink_cust, v_ink_spec, v_ink_job have been created
- /secure/queries/serverside/report_pr runs on milldb each night and reports a list of job inks updated
- To see the most recent ink updates:
```
mysql boranpla
select `date opened`, `date closed`, `works order number` as Job, `design code` as Spec, `design name` as Design, Customer from `ink_costing reports details` order by  `date closed` desc limit 5;
```

Formulations / colours
The ‘fullconvert’ tool has been installed on ‘millink’ to export relevant tables to the jobsystem.
These are visible in the Jobsystem under Config > Lookups > Real Ink Colours, but not yet used

Costings flow
When a job is created then the job is tagged at being available for dispense.
At this time the record is created on the costing reports details table for that works order.
As ink is dispensed and the returns issued/returned then the record for the job is updated.
At the end of the job the operator needs to select Order Complete.
When he does this the date the order is completed is filled in on the Costings reports details table.
Completed costings are transferred to the job system daily.

Installation
------------

The InkToJobs program is installed on the Ink PC, copy it to C:\InkToJobs.
Install also “MySQL driver 5.1” folder.

Click on InkToJobs to start an import, takes a few minutes. 

If there is a problem, maybe the DB settings in “Setups” file is wrong.

In the Control Panel, add a Scheduled task to run C:\ink_to_jobsystem\InkToJobs.exe every morning at 07:00

See the "setups" file for settings.
```
   DebugLevel : 0=normal timed run 1=debug >1=normal no timer
```

There are extensive logs.

Backups:  Install rsyncd
```
cd \rsyncd
cygrunsrv.exe -I rsyncd -e CYGWIN=nontsec -p c:/rsyncd/rsync.exe -a "--config=c:/rsyncd/rsyncd.conf --daemon --no-detach"
cygrunsrv.exe --start rsyncd
```

Then configure backuppc.


Development environment
---------

- On "delphi VM" with WIndows XP
- Visual Basic 6
- Mysql ODBC connector
- edit setups.txt : point to mysqlBD, and 5378 directory


See also the google doc https://docs.google.com/document/d/1pSPYGSR8J53h2mUsSDMyKPb0Tckvz-1NGiHxq1I31-k/edit#

Sean Boran, 2012.

<hr> 
Links to BP git repos: 

 * Production reports https://github.com/Boran/jobcard-prodreports
 * Job System UI: https://github.com/Boran/jobcard-delphi
 * Job UI report sql/layouts: https://github.com/Boran/jobcard-layouts
 * Linux script and server-side job sql: https://github.com/Boran/bpsecure
 * Ink System intwrface: https://github.com/Boran/job-InkToJobs
