VERSION 5.00
Begin VB.Form MainFrm 
   Caption         =   "Ink - Job/C3 synchronisation"
   ClientHeight    =   8925
   ClientLeft      =   1035
   ClientTop       =   645
   ClientWidth     =   12570
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   8925
   ScaleWidth      =   12570
   WindowState     =   2  'Maximized
   Begin VB.CommandButton bbMigrateSpecs 
      BackColor       =   &H00FFFF80&
      Caption         =   "Manually first: migrate custs/specs"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6480
      Width           =   3975
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   10320
      TabIndex        =   10
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtmysqlconnectionstring 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   720
      Width           =   7935
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF80&
      Caption         =   "4. Import Jobs"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6480
      Width           =   5295
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   0
      TabIndex        =   6
      Top             =   6960
      Width           =   12495
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   0
      TabIndex        =   5
      Top             =   4080
      Width           =   12495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF80&
      Caption         =   "3. Import printers/customers/ink types/designs"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
      Caption         =   "1. Connect to C3 System"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Test DB"
      Top             =   480
      Width           =   3615
   End
   Begin VB.CommandButton Customers 
      BackColor       =   &H00FFFF80&
      Caption         =   "2. Import customers"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   5295
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   12495
   End
   Begin VB.Label TextSql 
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   1320
      Width           =   7095
   End
   Begin VB.Label Label1 
      Caption         =   "Counter"
      Height          =   255
      Left            =   9120
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Version C3/ 2020.09.25 -01"
      Height          =   255
      Left            =   9840
      TabIndex        =   9
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "C3 SQL database:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSource As String
Dim AccessConn As New ADODB.Connection
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim IsConnected As Boolean
Dim Answer As Boolean
Dim Answer2 As Boolean

Dim MySQLUserName As String
Dim MySQLPassword As String
Dim MySQLPort As String
Dim MySQLDatabaseName As String
Dim MySQLDriver As String
Dim MySQLHost As String

Dim AccessDBPath As String
Dim PrintProcess As String
Dim db As Database
Dim Timer2Counter, Timer2CounterStart, DaysToLook, DebugLevel As Integer


Public Sub TerminateConnection()
    On Error Resume Next
    
    g_MySQLConn.Close
    Set g_MySQLConn = Nothing
    Set g_MySQLError = Nothing
End Sub


Private Sub bbMigrateSpecs_Click()
' migrate old job customer + spec codes to new C3 codes
Dim Query As String
Dim db, db2 As Database
Dim rstCust, rst2 As DAO.Recordset
Dim Counter As Integer

Set db = OpenDatabase(AccessDBPath)
Set db2 = OpenDatabase(AccessDBPath)
Me.List3.Clear

' 1. add C3 customers
Me.List3.AddItem "-- 1. add C3 customers"
Call Customers_Click

Answer = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)
If Answer = True Then

    ' get all customers with Job and C3 code
    Me.Label1 = "Customer code migration: "
    Me.List3.AddItem Me.Label1
    Query = "SELECT Cust_code, JobSystem_Cust_Code from CV_Ink_Customer"
    CustLog ("------ bbMigrate: " & Query)
    Me.List3.AddItem "-- 2. cust-code on designs"
    rst1.Open Query, g_MySQLConn, adOpenDynamic, adLockOptimistic
    Counter = 0
    Do Until rst1.EOF
        Counter = Counter + 1
        Me.Text1 = Counter
        DoEvents
        
        If Not IsNull(rst1![Cust_code]) Then
            If Not IsNull(rst1![JobSystem_Cust_Code]) Then
                Set rstCust = db.OpenRecordset("SELECT * FROM designs WHERE [Customer] = '" & rst1![JobSystem_Cust_Code] & "'")
                Do Until rstCust.EOF
                    ' Update customer code
                    Me.List3.AddItem "Update job cust " & rst1![JobSystem_Cust_Code]
                    If Not IsNull(rstCust![Customer]) Then
                      Me.List3.AddItem "Update design " & rstCust![Design Code] & " for customer code: " & rst1![JobSystem_Cust_Code] & " changed to " & rst1![Cust_code]
                      rstCust.Edit
                      rstCust![Customer] = rst1![Cust_code]
                      rstCust.Update
                      If DebugLevel = 1 Then
                        CustLog ("bbMigrate, Update design " & rstCust![Design Code] & " for customer code: " & rst1![JobSystem_Cust_Code] & " channged to " & rst1![Cust_code])
                      End If
                    End If
                    rstCust.MoveNext
                Loop
                rstCust.Close
                
            End If
        End If
        rst1.MoveNext
    Loop
    CustLog ("------ bbMigrate: cust done")
    rst1.Close
    
    
    ' Fix live DB: once off
    '           Me.List3.AddItem "--- Prepare to find entries added on 20th sept:"
    '           CustLog ("--- HACK: Prepare to find entries added on 20th sept")
                ' '20/09/2020' and [Date Amended]< '20/09/2020'
                ' WHERE [Date Created] = #2020/09/20#
                ' [Design Code],[Date Created],[Date Amended]
                'Set rstCust = db.OpenRecordset("SELECT top 10 * FROM [Designs] WHERE [Date Created] = #20/09/2020# ")
                'Set rstCust = db.OpenRecordset("SELECT top 10 * FROM [Designs] WHERE [Date Created] < #20/09/2020# ")
     '           Set rstCust = db.OpenRecordset("SELECT FROM [Designs] d WHERE d.[Design Code] like 'A*' and  [Date Created] = #20/09/2020# and [Date Amended]< #20/09/2020#")
     '           Do Until rstCust.EOF
     '               If Not IsNull(rstCust![Design Code]) Then
      '                Me.List3.AddItem "code: " & rstCust![Design Code]
                      ' & Format(rst1![Date Created], "yyyy/mm/dd")
                      'Me.List3.AddItem "code: " & rstCust![Design Code] & " Created " & rst1![Date Created] & ", amended " & rst1![Date Amended]
                      ' delete the record
                      'rstCust.Edit
                      'rstCust.Update
      '                rstCust.Delete
      '              End If
      '              rstCust.MoveNext
      '          Loop
    '          rstCust.Close
    
    Me.Label1 = "Spec nr migration: "
    Me.List3.AddItem "-- 3. get list of specs"
    WriteDesignTraceLog ("-- 3. get list of specs")
    Query = "SELECT SPEC,OldJobSpec,DesignImage from CV_Ink_Articles WHERE OldJobSpec IS NOT null AND SPEC IS NOT null"
    WriteDesignTraceLog ("------ bbMigrate: " & Query)
    Counter = 0
    rst1.Open Query, g_MySQLConn, adOpenDynamic, adLockOptimistic
    Do Until rst1.EOF
        Counter = Counter + 1
        Me.Text1 = Counter
        DoEvents
            
        ' does a design exist with C3 code=Spec already? If yes, skip
        'WriteDesignTraceLog ("SELECT Spec FROM [Design Code] WHERE [Design Code] = '" & rst1![Spec] & "'")
        Set rst2 = db2.OpenRecordset("SELECT [Design Code] FROM [Designs] WHERE [Design Code] = '" & rst1![Spec] & "'")
        If rst2.RecordCount < 1 Then
            ' find designs to rename
            Me.List3.AddItem "-- 4. Find [Designs] with OldJobSpec " & rst1![OldJobSpec]
            WriteDesignTraceLog ("-- 4. Find [Designs] with OldJobSpec " & rst1![OldJobSpec])
            Set rstCust = db.OpenRecordset("SELECT * FROM [Designs] WHERE [Design Code] = '" & rst1![OldJobSpec] & "'")
            Do Until rstCust.EOF
                If Not IsNull(rstCust![Design Code]) Then
                  
                    If DebugLevel = 1 Then
                      WriteDesignTraceLog ("bbMigrate, [Designs] code jobsys: " & rst1![OldJobSpec] & " changed to " & rst1![Spec])
                    End If
                    Me.List3.AddItem "[Designs] OldJobSpec: " & rst1![OldJobSpec] & " change to " & rst1![Spec] & ", image=" & rst1![DesignImage]
                    
                    rstCust.Edit
                    rstCust![Design Code] = rst1![Spec]
                    If Not IsNull(rst1![DesignImage]) Then
                      rstCust![Image Path] = rst1![DesignImage]
                    Else
                      rstCust![Image Path] = "NO DesignImage"
                    End If
                    rstCust.Update
                     
                End If
                rstCust.MoveNext
            Loop
            rstCust.Close
          
        Else
            WriteDesignTraceLog ("-- 4. Skip existing [Designs] " & rst1![Spec])
        End If
        

        Me.List3.AddItem "-- 5. Update [Design Details].."
        WriteDesignTraceLog ("-- 5. Update [Design Details]..")
        Set rstCust = db.OpenRecordset("SELECT * FROM [Design Details] WHERE [Design Code] = '" & rst1![OldJobSpec] & "'")
        Do Until rstCust.EOF
            ' Update code
            If Not IsNull(rstCust![Design Code]) Then
              If DebugLevel = 1 Then
                WriteDesignTraceLog ("bbMigrate, [Design Details] code (jobsys): " & rst1![OldJobSpec] & " changed to " & rst1![Spec])
              End If
              Me.List3.AddItem "[Design Details] code (jobsys): " & rst1![OldJobSpec] & " change to " & rst1![Spec]
              rstCust.Edit
              rstCust![Design Code] = rst1![Spec]
              rstCust.Update
            End If
            rstCust.MoveNext
        Loop
        rstCust.Close
            
        Me.List3.AddItem "-- 6. Update [Works Orders].."
        WriteDesignTraceLog ("-- 6. Update [Works Orders]..")
        Set rstCust = db.OpenRecordset("SELECT * FROM [Works Orders] WHERE [Design Code] = '" & rst1![OldJobSpec] & "'")
        Do Until rstCust.EOF
            ' Update spec code
            If Not IsNull(rstCust![Design Code]) Then
              Me.List3.AddItem "[Works Orders] code (jobsys): " & rst1![OldJobSpec] & " change to " & rst1![Spec]
              If DebugLevel = 1 Then
                WriteDesignTraceLog ("bbMigrate, [Works Orders] code (jobsys): " & rst1![OldJobSpec] & " changed to " & rst1![Spec])
              End If
              rstCust.Edit
              rstCust![Design Code] = rst1![Spec]
              rstCust.Update
            End If
            rstCust.MoveNext
        Loop
        rstCust.Close

        rst1.MoveNext
    Loop
        
    
    WriteDesignTraceLog ("------ bbMigrate: done ")
    Me.List3.AddItem "---- done"
    rst1.Close
    
    Call TerminateConnection
End If
End Sub

Private Sub Command1_Click()
' test DB connection
Answer2 = False
Me.txtmysqlconnectionstring.Text = ""

Answer2 = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)
If Answer2 = True Then
  'MsgBox "Connected to DB Server = "
Else
  MsgBox "Could not connect to the DB server "
End If

End Sub

Private Sub Customers_Click()

Dim MyView As String
Dim MyNewView As String
Dim CustCode As String
Dim JCustCode As String
Dim CustName As String
Dim CustAddress As String
Dim ValueStr As String
Dim HeaderStr As String
Dim Counter As Integer

Me.Label1 = "Customers"
CustLog ("Customers_Click connect to DB " & MySQLDatabaseName)
Answer = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)
CustLog ("Customers_Click Connection answer = " & Answer)
Me.List1.Clear

If Answer = True Then
    MyView = "CV_Ink_Customer"
    HeaderStr = "C3_cust_code" & vbTab & "Job_cust_code" & vbTab & "Customer Name and Address"
    Me.List1.AddItem HeaderStr & "(from " & MyView & ")"
    
    MyNewView = "SELECT Cust_code, JobSystem_Cust_Code, Name, ADDRESS1 from " & MyView
    CustLog ("Customers_Click: " & MyNewView)
    rst1.Open MyNewView, g_MySQLConn, adOpenDynamic, adLockOptimistic
    
    Me.Label1 = "Customers: "
    Do Until rst1.EOF
        Counter = Counter + 1
        Me.Text1 = Counter
        If Not IsNull(rst1![Cust_code]) Then
                CustCode = rst1![Cust_code]
                If Not IsNull(rst1![JobSystem_Cust_Code]) Then
                    JCustCode = rst1![JobSystem_Cust_Code]
                Else
                    JCustCode = ""
                End If
                If Not IsNull(rst1![Name]) Then
                    CustName = rst1![Name]
                Else
                    CustName = "NO CUSTOMER NAME"
                End If
                If Not IsNull(rst1![ADDRESS1]) Then
                    CustAddress = rst1![ADDRESS1]
                Else
                    CustAddress = "NO CUSTOMER ADDRESS"
                End If
                
                CustLog ("Customers_Click: " & CustCode & "," & JCustCode & "," & CustName & "," & CustAddress)
                Me.Text1 = Counter & " " & CustName
                Call UpdateCustomer(CustCode, CustName, JCustCode)
        End If

        ValueStr = CustCode & vbTab & JCustCode & vbTab & CustName & vbTab & vbTab & CustAddress
        Me.List1.AddItem ValueStr
        rst1.MoveNext
    Loop
    
    rst1.Close
    Call TerminateConnection
End If


End Sub


Private Sub Command3_Click()
' import spec
Dim MyView As String
Dim MyNewView As String
Dim Spec As String
Dim CustCode As String
Dim Design As String
Dim Substrate As String
Dim ValueStr As String
Dim HeaderStr As String
Dim InkType As String
Dim PrWidth As String
Dim PrRepeat As String
Dim MyComment As String
Dim DesignImage As String
Dim Printer As String
Dim LastChange As String
Dim LastChangeDate As Date
Dim LastChangeday As String
Dim LastChangemonth As String
Dim LastChangeyear As String
Dim LastChangeTime As String
Dim LastChangeOriginal As String
Dim LastChange24 As String
Dim Counter As Integer
Dim LengthOfString As Integer

Answer = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)
' get printer list, add presses if needed
If Answer = True Then
    MyNewView = "SELECT DISTINCT Printer from CV_Ink_Articles"
    WriteDesignTraceLog ("------ Command3_Click: " & MyNewView)
    rst1.Open MyNewView, g_MySQLConn, adOpenDynamic, adLockOptimistic
    Me.List2.AddItem "CONNECTED"
      
    Me.Label1 = "Printers:"
    Do Until rst1.EOF
        Counter = Counter + 1
        Me.Text1 = Counter
        If Not IsNull(rst1![Printer]) Then
            Printer = rst1![Printer]
        Else
            Printer = "NO PRINTER"
        End If
        
        Me.List2.AddItem "Checking Printer " & Printer
        WriteDesignTraceLog ("Calling AddPress " & Printer)
        Call AddPress(Printer, Printer)
        rst1.MoveNext
    Loop
    Call TerminateConnection
    'WriteDesignTraceLog (CStr(Spec & "," & CustCode & "," & Design & "," & Substrate & "," & PrRepeat & "," & PrWidth & "," & InkType & "," & Printer & "," & LastChangeOriginal & "," & LastChangeday & "/" & LastChangemonth & "/" & LastChangeyear & "," & LastChangeTime & "," & LastChange24))
End If

Call AddCustomers
Call AddInkTypes
Call AddSubstrates
Call AddDesigns
WriteDesignTraceLog ("------ Command3_Click: done ")
Call TerminateConnection

End Sub


Private Sub AddCustomers()

Dim MyView As String
Dim MyNewView As String
Dim Spec As String
Dim CustCode As String
Dim Design As String
Dim Substrate As String
Dim ValueStr As String
Dim HeaderStr As String
Dim InkType As String
Dim PrWidth As String
Dim PrRepeat As String
Dim MyComment As String
Dim DesignImage As String
Dim Printer As String
Dim LastChange As String
Dim LastChangeDate As Date
Dim LastChangeday As String
Dim LastChangemonth As String
Dim LastChangeyear As String
Dim LastChangeTime As String
Dim LastChangeOriginal As String
Dim LastChange24 As String
Dim Counter As Integer


Answer = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)

If Answer = True Then
    'MyView = "`CV_Ink_Articles`"
    MyNewView = "SELECT DISTINCT cust_code from CV_Ink_Articles"
    WriteDesignTraceLog ("------ AddCustomers: " & MyNewView)
    rst1.Open MyNewView, g_MySQLConn, adOpenDynamic, adLockOptimistic
    Me.List2.AddItem "AddCustomers"
    
    Me.Label1 = "Customers:"
    Do Until rst1.EOF
        Counter = Counter + 1
        Me.Text1 = Counter
        DoEvents
        
        If rst1![Cust_code] = "" Then
            CustCode = "NO CUSTOMER CODE"
        ElseIf Not IsNull(rst1![Cust_code]) Then
            CustCode = rst1![Cust_code]
        Else
            CustCode = "NO CUSTOMER CODE"
        End If

        Me.List2.AddItem "Checking Customer: " & CustCode
        'WriteDesignTraceLog ("Calling AddCustomer " & CustCode)
        Call AddCustomer(CustCode, CustCode)
        
        rst1.MoveNext
    Loop
    Call TerminateConnection
End If



End Sub
Private Sub AddInkTypes()

Dim MyView As String
Dim MyNewView As String
Dim Spec As String
Dim CustCode As String
Dim Design As String
Dim Substrate As String
Dim InkCode As String
Dim ValueStr As String
Dim HeaderStr As String
Dim InkType As String
Dim PrWidth As String
Dim PrRepeat As String
Dim MyComment As String
Dim DesignImage As String
Dim Printer As String
Dim LastChange As String
Dim LastChangeDate As Date
Dim LastChangeday As String
Dim LastChangemonth As String
Dim LastChangeyear As String
Dim LastChangeTime As String
Dim LastChangeOriginal As String
Dim LastChange24 As String
Dim Counter As Integer


Answer = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)

If Answer = True Then
    Me.List2.AddItem "AddInkTypes from CV_Ink_Articles"
    MyNewView = "SELECT DISTINCT inktype from CV_Ink_Articles"
    WriteDesignTraceLog ("------ AddInkTypes: " & MyNewView)
    rst1.Open MyNewView, g_MySQLConn, adOpenDynamic, adLockOptimistic
      
    Me.Label1 = "inktypes:"
    Do Until rst1.EOF
        Counter = Counter + 1
        Me.Text1 = Counter
        DoEvents
        
        If Not IsNull(rst1![InkType]) Then
            InkCode = rst1![InkType]
        Else
            InkCode = "NO INK TYPE"
        End If

        Me.List2.AddItem "Checking InkType " & InkCode
        WriteDesignTraceLog ("Calling AddInkType " & InkCode)
        Call AddInkType(InkCode)
        
        rst1.MoveNext
    Loop

    Call TerminateConnection
End If
'Call AddSubstrates
End Sub

Public Function AddSubstrates() As String

Dim MyView As String
Dim MyNewView As String
Dim Spec As String
Dim CustCode As String
Dim Design As String
Dim Substrate As String
Dim ValueStr As String
Dim HeaderStr As String
Dim InkType As String
Dim PrWidth As String
Dim PrRepeat As String
Dim MyComment As String
Dim DesignImage As String
Dim Printer As String
Dim LastChange As String
Dim LastChangeDate As Date
Dim LastChangeday As String
Dim LastChangemonth As String
Dim LastChangeyear As String
Dim LastChangeTime As String
Dim LastChangeOriginal As String
Dim LastChange24 As String
Dim Counter As Integer


Answer = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)
If Answer = True Then
    Me.List2.AddItem "AddSubstrates: Substrate FROM CV_Ink_Articles "
    Me.Label1 = "Substrates:"
    
    MyNewView = "SELECT DISTINCT Substrate FROM CV_Ink_Articles"
    WriteDesignTraceLog ("------ AddSubstrates: " & MyNewView)
    rst1.Open MyNewView, g_MySQLConn, adOpenDynamic, adLockOptimistic
    
    Do Until rst1.EOF
        Counter = Counter + 1
        Me.Text1 = Counter
        DoEvents
        
        If Not IsNull(rst1![Substrate]) Then
            Substrate = rst1![Substrate]
        Else
            Substrate = "NO SUBSTATE CODE"
        End If

        Me.List2.AddItem "Checking Substrate: " & Substrate
        WriteDesignTraceLog ("Calling UpdateSubstrate <" & Substrate & ">")
        Substrate = UpdateSubstrate(Substrate)
        
        rst1.MoveNext
    Loop
     
    Call TerminateConnection
End If

'MsgBox "DONE"
'Call AddDesigns

End Function

Private Sub AddDesigns()

Dim MyView As String
Dim Spec As String
Dim CustCode As String
Dim Design As String
Dim Substrate As String
Dim ValueStr As String
Dim HeaderStr As String
Dim InkType As String
Dim PrWidth As String
Dim PrRepeat As String
Dim MyComment As String
Dim DesignImage As String
Dim Printer As String
Dim LastChange As String
Dim LastChangeDate As Date
Dim LastChangeday As String
Dim LastChangemonth As String
Dim LastChangeyear As String
Dim LastChangeTime As String
Dim LastChangeOriginal As String
Dim LastChange24 As String
Dim Counter As Integer
Dim LengthOfString As Integer


Answer = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)
If Answer = True Then
    'HeaderStr = "Spec" & " | " & "Customer Code" & " | " & "Design"
    HeaderStr = "Spec" & vbTab & vbTab & "Customer Code" & vbTab & vbTab & "Design"
    Me.List2.AddItem HeaderStr
    Me.List2.AddItem ""
    
    MyView = "SELECT * from CV_Ink_Articles"
    WriteDesignTraceLog ("------ AddDesigns: " & MyView)
    rst1.Open MyView, g_MySQLConn, adOpenDynamic, adLockOptimistic
    'Me.List2.AddItem "AddDesigns "
      
    Me.Label1 = "Specs:"
    Do Until rst1.EOF
        Counter = Counter + 1
        Me.Text1 = Counter
        DoEvents
        'MsgBox "Counter = " & Counter
        
        If Not IsNull(rst1![Spec]) Then
            Spec = rst1![Spec]
            
            If rst1![Cust_code] = "" Then
                CustCode = "NO CUSTOMER CODE"
            ElseIf Not IsNull(rst1![Cust_code]) Then
                CustCode = rst1![Cust_code]
            Else
                CustCode = "NO CUSTOMER CODE"
            End If
        
            'If Not IsNull(rst1![Cust_code]) Then
            '   CustCode = rst1![Cust_code]
            'Else
            '   CustCode = "NO CUSTOMER CODE"
            'End If
            
            If Not IsNull(rst1![Design]) Then
                Design = rst1![Design]
            Else
                Design = "NO DESIGN"
            End If
            If Not IsNull(rst1![Substrate]) Then
                Substrate = rst1![Substrate]
            Else
                Substrate = "no SUBSTRATE"
            End If
            If Not IsNull(rst1![PrRepeat]) Then
                PrRepeat = rst1![PrRepeat]
            Else
                PrRepeat = "0"
            End If
            If Not IsNull(rst1![PrWidth]) Then
                PrWidth = rst1![PrWidth]
            Else
                PrWidth = "0"
            End If
            
            If Not IsNull(rst1![InkType]) Then
                InkType = rst1![InkType]
            Else
                InkType = ""
            End If
           
            If Not IsNull(rst1![P_Comment]) Then
                MyComment = rst1![P_Comment]
            Else
                MyComment = " "
            End If
            
            'MsgBox MyComment
            
            If Not IsNull(rst1![DesignImage]) Then
                DesignImage = rst1![DesignImage]
            Else
                DesignImage = "NO DesignImage"
            End If
            
            If Not IsNull(rst1![Printer]) Then
                Printer = rst1![Printer]
            Else
                Printer = "NO PRINTER"
            End If
            
            If Not IsNull(rst1![LastChange]) Then
                LastChangeOriginal = rst1![LastChange]
                LastChange = DateValue(LastChangeOriginal)
                LastChangemonth = Month(LastChange)
                LastChangeday = Day(LastChange)
                LastChangeyear = Year(LastChange)
                LastChangeTime = rst1![LastChange]
                LastChangeTime = TimeValue(LastChangeOriginal)
                LastChange24 = Format(LastChangeOriginal, "hh:mm:ss")
            Else
                LastChange = "NO lastchange"
            End If
            
            If Not IsNumeric(PrWidth) Then
              PrWidth = 0
            End If
            If Not IsNumeric(PrRepeat) Then
              PrRepeat = 0
            End If
            
            WriteDesignTraceLog (CStr(Spec & "," & CustCode & "," & Design & "," & Substrate & "," & PrRepeat & "," & PrWidth & "," & InkType & "," & Printer & "," & LastChangeOriginal & "," & LastChangeday & "/" & LastChangemonth & "/" & LastChangeyear & "," & LastChangeTime & "," & LastChange24))  ' & "," & Printer & "," & LastChange))
            
            Call AddPress(Printer, Printer)
            Call AddCustomer(CustCode, CustCode)
            Substrate = UpdateSubstrate(Substrate)
            'WriteDesignTraceLog ("Substrate now=" & Substrate)
            WriteDesignTraceLog (CStr(Spec & "," & CustCode & "," & Design & "," & Substrate & "," & PrRepeat & "," & PrWidth & "," & InkType & "," & Printer & "," & LastChangeOriginal & "," & LastChangeday & "/" & LastChangemonth & "/" & LastChangeyear & "," & LastChangeTime & "," & LastChange24))  ' & "," & Printer & "," & LastChange))
            'WriteDesignTraceLog "Calling UpdateDesign,"
            Call UpdateDesign(Spec, Design, CustCode, Substrate, Printer, CSng(PrWidth), CSng(PrRepeat), CDate(LastChangeday & "/" & LastChangemonth & "/" & LastChangeyear), LastChange24, MyComment, DesignImage, InkType)

        End If

        'ValueStr = Spec & " | " & CustCode & " | " & Design
        ValueStr = Spec & vbTab & vbTab & CustCode & vbTab & vbTab & Design
        
        Me.List2.AddItem ValueStr
        rst1.MoveNext
    Loop

    rst1.Close
    WriteDesignTraceLog ("------ AddDesigns: done ")
    Call TerminateConnection
End If

End Sub


Private Sub Command4_Click()
' import jobs
Dim MyView As String
Dim job As String
Dim Spec As String
Dim DelDate As String
Dim CustCode As String
Dim Design As String
Dim ValueStr As String
Dim HeaderStr As String
Dim Counter As Integer


WriteJobsTraceLog ("------ Command4_Click connect DB")
Answer = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)
WriteJobsTraceLog ("Connection = " & Answer)

If Answer = True Then
    HeaderStr = "Job" & vbTab & vbTab & "Spec" & vbTab & vbTab & "Del Date" & vbTab & vbTab & "Customer Code" & vbTab & "Design"
    Me.List3.AddItem HeaderStr
    MyView = "select job, Spec, Date, Cust_code, Design  from CV_Ink_PO"
    rst1.Open MyView, g_MySQLConn, adOpenDynamic, adLockOptimistic
      
    Me.Label1 = "Jobs:"
    Do Until rst1.EOF
        Counter = Counter + 1    ' Show the number of jobs being processed
        Me.Text1 = Counter
      
        If Not IsNull(rst1![job]) Then
            job = rst1![job]
          
            If Not IsNull(rst1![Spec]) And (rst1![Spec] > 0) Then
                Spec = rst1![Spec]
            Else
                'Spec = "   "
                ' TODO
                ' Design 99 is a dummy, that exists, to allow a valid db ref
                Spec = "99"
            End If
                    
        End If
        
        WriteJobsTraceLog (job & "," & Spec & "," & rst1![Date] & "," & rst1![Cust_code] & "," & rst1![Design])

        'ValueStr = Spec & " | " & CustCode & " | " & Design
        'ValueStr = job & vbTab & vbTab & Spec & vbTab & vbTab & DelDate & vbTab & vbTab & CustCode & vbTab & vbTab & Design
        ValueStr = job & vbTab & vbTab & Spec & vbTab & vbTab & rst1![Date] & vbTab & vbTab & rst1![Cust_code] & vbTab & vbTab & rst1![Design]
        Me.List3.AddItem ValueStr
        
        If DebugLevel = 1 Then
            WriteJobsTraceLog ("Calling AddJob Job=" & job & " Spec=" & Spec)
        End If
        Call AddJob(job, Spec)
        rst1.MoveNext
    Loop

    rst1.Close
End If
WriteJobsTraceLog ("------ Command4_Click done")
Call TerminateConnection


End Sub



Private Sub Form_Load()

Me.WindowState = 1

'MsgBox "LoadSetups"
Call LoadSetups

'Exit Sub

Call UpdateSubstrate("NO SUBSTRATE")
Call AddCustomer("NO CUSTOMER CODE", "NO CUSTOMER CODE")
'Timer2Counter = 10


'Me.txtDatabaseName.Text = "boranpla"
'Me.txtDriver.Text = "MySQL ODBC 5.1 Driver"
'Me.txtHost.Text = "bpjob"
'Me.txtPassword.Text = "query4inks"
'Me.txtUsername.Text = "inovex"
'Me.txtPort.Text = "3306"

'MsgBox "Form load done"
End Sub

Function CustLog(TraceText As String)
  'Me.List1.AddItem "Error see log for details"
  ' ListFile = App.Path & "\" & "setups.txt"
  Dim logd, logdir, lfile As String
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Dim fso, f
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  logd = App.Path & "\Logs"
  logdir = logd & "\" & "custlog"
  
  If Not fso.folderexists(logd) Then
    MsgBox "Cannot open directory " & logd
  
  Else
   lfile = logdir & Day(Now) & "-" & Month(Now) & "-" & Year(Now) & ".txt"
   If Dir(lfile) = "" Then
        Set f = fso.CreateTextFile(lfile, True)
        f.Close
   End If
   Set f = fso.OpenTextFile(lfile, ForAppending, False)
   f.WriteLine Now & " " & TraceText
   f.Close
  End If

End Function


Function WriteDesignTraceLog(TraceText As String)

  'Me.List1.AddItem "Error see log for details"

  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Dir("Logs\DesignTraceLog " & Day(Now) & " - " & Month(Now) & " - " & Year(Now) & ".txt") = "" Then
        Set f = fso.CreateTextFile("Logs\DesignTraceLog " & Day(Now) & " - " & Month(Now) & " - " & Year(Now) & ".txt", True)
        f.Close
  End If
  
  Set f = fso.OpenTextFile("Logs\DesignTraceLog " & Day(Now) & " - " & Month(Now) & " - " & Year(Now) & ".txt", ForAppending, False)
  f.WriteLine Now & " " & TraceText
  f.Close

End Function

Function WriteJobsTraceLog(TraceText As String)

  'Me.List1.AddItem "Error see log for details"
  
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Dir("Logs\JobsTraceLog " & Day(Now) & " - " & Month(Now) & " - " & Year(Now) & ".txt") = "" Then
        Set f = fso.CreateTextFile("Logs\JobsTraceLog " & Day(Now) & " - " & Month(Now) & " - " & Year(Now) & ".txt", True)
        f.Close
  End If
  
  Set f = fso.OpenTextFile("Logs\JobsTraceLog " & Day(Now) & " - " & Month(Now) & " - " & Year(Now) & ".txt", ForAppending, False)
  f.WriteLine Now & " " & TraceText
  'f.WriteLine TraceText
  
  f.Close

End Function

Private Sub Command5_Click()
' export costings
Dim db As Database
Dim rst As DAO.Recordset
Dim WorksOrderNumberI As String
Dim DesignCodeI As String
Dim DesignNameI As String
Dim DateClosedI As String
Dim DateOpenedI As String
Dim InkWeightI As Single
Dim CustomerI As String
Dim WhiteWeightI As Single
Dim ColoursWeightI As Single
Dim LacquerWeightI As Single
Dim OtherWeightI As Single
Dim ReturnsReturnedI As Single
Dim ReturnsReturnedCostI As Single
Dim ReturnsIssuedI As Single
Dim ReturnsIssuedtoJobCostI As Single
Dim TotalTargetCostI As Single
Dim TargetCostPer1000sqmI As Single
Dim InkCostI As Single
Dim TotalCostI As Single
Dim EstimatedSQMI As Long
Dim SQMI As Long
Dim WhitesCostI As Single
Dim ColoursCostI As Single
Dim LacquerCostI As Single
Dim ActualCost1000I As Single
Dim Counter As Long

WriteExportTraceLog ("Command5_Click Entered Insert data")

Set db = OpenDatabase(AccessDBPath)
Set rst = db.OpenRecordset("SELECT [Costing Reports Details].[Date Closed], [Costing Reports Details].* From [Costing Reports Details] WHERE [Costing Reports Details].[Date Closed] Between Date()-" & DaysToLook & " And Date() ORDER BY [Costing Reports Details].[Date Closed] DESC")
WriteExportTraceLog ("Command5_Click SELECT from 'Costing Reports Details' access")

Me.Label1 = "Costings:"
Me.Text1 = "0"

If rst.RecordCount <> 0 Then
Do Until rst.EOF
    Counter = Counter + 1
    Me.Text1 = "INSERT " & Counter

    WorksOrderNumberI = rst![works order number]
    DesignCodeI = rst![Design Code]
    If Not IsNull(rst![design name]) Then
        DesignNameI = rst![design name]
    Else
        DesignNameI = "NO DESIGN NAME"
    End If
    DateOpenedI = rst![date opened]
    DateClosedI = rst![costing reports details.Date Closed]
    InkWeightI = rst![ink weight]
    If Not IsNull(rst![Customer]) Then
        CustomerI = rst![Customer]
    Else
        CustomerI = "NO CUSTOMER NAME"
    End If
    WhiteWeightI = rst![Whites Weight]
    ColoursWeightI = rst![colours weight]
    LacquerWeightI = rst![Lacquer Weight]
    OtherWeightI = rst![other weight]
    ReturnsReturnedI = rst![returns returned]
    ReturnsReturnedCostI = rst![returns returned cost]
    ReturnsIssuedI = rst![returns issued]
    ReturnsIssuedtoJobCostI = rst![returns issued to job cost]
    TotalTargetCostI = rst![totaltargetcost]
    TargetCostPer1000sqmI = rst![TargetCostPer1000sqm]
    InkCostI = rst![Ink Cost]
    TotalCostI = rst![Total Cost]
    If Not IsNull(rst![Estimated SQM]) Then
        EstimatedSQMI = rst![Estimated SQM]
    Else
        EstimatedSQMI = 0
    End If
    SQMI = rst![SQM]
    WhitesCostI = rst![whites cost]
    ColoursCostI = rst![colours cost]
    LacquerCostI = rst![lacquer cost]
    ActualCost1000I = rst![Actual Cost per 1000sqm ex uplift]
    
    
    DesignNameI = Replace(DesignNameI, "'", "`")
    CustomerI = Replace(CustomerI, "'", "`")
    
    WriteExportTraceLog ("Command5_Click Call InsertData Works Order = " & WorksOrderNumberI)
    Call InsertData(WorksOrderNumberI, DesignCodeI, DesignNameI, InkWeightI, DateOpenedI, CustomerI, WhiteWeightI, ColoursWeightI, LacquerWeightI, OtherWeightI, ReturnsReturnedI, ReturnsReturnedCostI, ReturnsIssuedI, ReturnsIssuedtoJobCostI, TotalTargetCostI, TargetCostPer1000sqmI, InkCostI, TotalCostI, EstimatedSQMI, SQMI, WhitesCostI, ColoursCostI, LacquerCostI, ActualCost1000I, DateClosedI)
    
    If DebugLevel = 1 Then
      WriteExportTraceLog ("Command5_Click Moving to next record")
    End If
    rst.MoveNext
Loop

End If


'Call InsertData("123123123", "987987", "This is another test design", 121.34, "23/11/2011")

End Sub

Private Sub InsertData(worksorder As String, designcode As String, designname As String, inkweight As Single, DateOpened As String, Customer As String, WhiteWeight As Single, ColoursWeight As Single, LacquerWeight As Single, OtherWeight As Single, ReturnsReturned As Single, ReturnsReturnedCost As Single, ReturnsIssued As Single, ReturnsIssuedToJobCost As Single, totaltargetcost As Single, TargetCost1000 As Single, InkCost As Single, TotalCost As Single, EstimatedSQM As Long, SQM As Long, WhitesCost As Single, ColoursCost As Single, LacquerCost As Single, ACper1000 As Single, DateClosed As String)

Dim StrSQL As String
Dim wo As String
Dim dc As String
Dim dn As String
Dim iw As Single
Dim NoRecord As Boolean
Dim rst As New ADODB.Recordset
Dim MyNewDate As String
Dim MyNewDateClosed As String


MyNewDate = Format(DateOpened, "yyyy/mm/dd")
MyNewDateClosed = Format(DateClosed, "yyyy/mm/dd")

WriteExportTraceLog ("InsertData Connecting to SQL Server")
Answer = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)
'Answer = EstablishMySQLConnection("root", "CTRecord", "IBM1", "Boranpla", 3306, "MySQL ODBC 5.1 Driver")


If Answer = True Then
    'WriteExportTraceLog ("InsertData Connection = " & Answer)
    rst.CursorLocation = adUseClient
    
    StrSQL = "SELECT `works order number` FROM `ink_costing reports details` WHERE `works order number` ='" & worksorder & "'"
    rst.Open StrSQL, g_MySQLConn, adOpenKeyset, adLockOptimistic
    
    WriteExportTraceLog ("InsertData Checking Works Order " & worksorder & " Exists on MySQL Server")
    NoRecord = False
    Do Until rst.EOF
        WriteExportTraceLog ("InsertData Works Order " & worksorder & " Exists")
        NoRecord = True
        rst.MoveNext
    Loop

End If

If NoRecord = False Then

    WriteExportTraceLog ("Setting INSERT statement for worksorder = " & worksorder)
    
    StrSQL = "INSERT INTO `ink_costing reports details` (`works order number`,`design code`,`design name`,`ink weight`,`date opened`,`customer`,`whites weight`,`colours weight`,`lacquer weight`,`other weight`,`returns returned`,`returns returned cost`,`returns issued`,`returns issued to job cost`,`totaltargetcost`,`targetcostper1000sqm`,`ink cost`,`total cost`,`estimated sqm`,`sqm`,`whites cost`,`colours cost`,`lacquer cost`,`Actual Cost per 1000sqm ex uplift`,`Date Closed`)" & _
    " VALUES ( '" & worksorder & "' , '" & designcode & "' ,'" & designname & "', '" & inkweight & "', '" & MyNewDate & "', '" & Customer & "', '" & WhiteWeight & "', '" & ColoursWeight & "','" & LacquerWeight & "','" & OtherWeight & "','" & ReturnsReturned & "','" & ReturnsReturnedCost & "','" & ReturnsIssued & "','" & ReturnsIssuedToJobCost & "','" & totaltargetcost & "','" & TargetCost1000 & "','" & InkCost & "','" & TotalCost & "','" & EstimatedSQM & "','" & SQM & "','" & WhitesCost & "','" & ColoursCost & "','" & LacquerCost & "','" & ACper1000 & "','" & MyNewDateClosed & "' )"
    'StrSQL = "INSERT INTO `ink_costing reports details` (`works order number`,`design code`,`design name`,`ink weight`,`date opened`,`customer`,`whites weight`,`colours weight`,`lacquer weight`,`other weight`,`returns returned`,`returns returned cost`,`returns issued`,`returns issued to job cost`,`totaltargetcost`,`targetcostper1000sqm`) VALUES ( '" & worksorder & "' , '" & designcode & "' ,'" & designname & "', '" & inkweight & "', '" & MyNewDate & "', '" & Customer & "', '" & WhiteWeight & "', '" & ColoursWeight & "','" & LacquerWeight & "','" & OtherWeight & "','" & ReturnsReturned & "','" & ReturnsReturnedCost & "','" & ReturnsIssued & "','" & ReturnsIssuedToJobCost & "','" & totaltargetcost & "','" & TargetCost1000 & "')"
    'StrSQL = "INSERT INTO `ink_costing reports details` (`works order number`,`design code`,`design name`,`ink weight`,`date opened`,`customer`,`whites weight`,`colours weight`,`lacquer weight`,`other weight`,`returns returned`,`returns returned cost`,`returns issued`,`returns issued to job cost`,`totaltargetcost`,`targetcostper1000sqm`) VALUES ( '" & worksorder & "' , '" & designcode & "','" & designname & "', '" & inkweight & "', '" & MyNewDate & "','" & Customer & "','" & WhiteWeight & "','" & ColoursWeight & "','" & LacquerWeight & "','" & OtherWeight & "','" & ReturnsReturned & "','" & ReturnsReturnedCost & "','" & ReturnsIssued & "','" & ReturnsIssuedToJobCost & "')"
    
    
    WriteExportTraceLog ("InsertData INTO `ink_costing reports details`")
    rst1.Open StrSQL, g_MySQLConn, adOpenDynamic, adLockOptimistic
    WriteExportTraceLog ("InsertData INSERT statement executed for works order = " & worksorder)
    
    Call TerminateConnection
End If

End Sub

Function WriteExportTraceLog(TraceText As String)

  'Me.List1.AddItem "Error see log for details"
  
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Dir("Logs\ExportTraceLog " & Day(Now) & " - " & Month(Now) & " - " & Year(Now) & ".txt") = "" Then
        Set f = fso.CreateTextFile("Logs\ExportTraceLog " & Day(Now) & " - " & Month(Now) & " - " & Year(Now) & ".txt", True)
        f.Close
  End If
  
  Set f = fso.OpenTextFile("Logs\ExportTraceLog " & Day(Now) & " - " & Month(Now) & " - " & Year(Now) & ".txt", ForAppending, False)
  f.WriteLine Now & " " & TraceText
  
  f.Close

End Function

Function UpdateCustomer(CustomerCode As String, CustomerName As String, JCustCode As String)
Dim db, db2 As Database
Dim rstCust, rst2 As DAO.Recordset

Set db = OpenDatabase(AccessDBPath)
Set db2 = OpenDatabase(AccessDBPath)

Set rstCust = db.OpenRecordset("SELECT * FROM Customers WHERE [Customer code] = '" & CustomerCode & "'")
If rstCust.RecordCount = 0 Then
    ' Code does not exist
    Set rst2 = db2.OpenRecordset("SELECT [Customer Code] FROM [Customers] WHERE [Customer Code] = '" & JCustCode & "'")
    If rst2.RecordCount = 1 Then
        ' a) Rename a Job Cust, or
        If DebugLevel = 1 Then
          CustLog ("UpdateCustomer: rename Jobcust " & JCustCode & " to " & CustomerCode)
        End If
        rst2.Edit
        rst2![Customer Code] = CustomerCode
        rst2![customer name] = CustomerName
        rst2.Update
    Else
        ' b) Add new customer name + code
        CustLog ("UpdateCustomer: code " & CustomerCode & " added - " & CustomerName)
        rstCust.AddNew
        rstCust![Customer Code] = CustomerCode
        rstCust![customer name] = CustomerName
        rstCust.Update
    End If
Else
    ' Code exists: Update customer name, given code
    If DebugLevel = 1 Then
      CustLog ("UpdateCustomer: " & CustomerCode & ": customer name updated")
    End If
    rstCust.Edit
    rstCust![customer name] = CustomerName
    rstCust.Update
End If
rstCust.Close

' deprecated: in fact dont need to do this, just leave old customers in the DB
' further down references will have to be updated though.
'If Not IsNull(JCustCode) Then
'  Set rstCust = db.OpenRecordset("SELECT * FROM Customers WHERE [Customer code] = '" & JCustCode & "'")
'  If rstCust.RecordCount <> 0 Then
'    ' Convert code from JobSystem to C3
'    If DebugLevel = 1 Then
'      CustLog ("UpdateCustomer: Change code from JobSystem: " & "SELECT * FROM Customers WHERE [Customer code] = '" & JCustCode & "'")
'    End If
'    rstCust.Edit
'    rstCust![Customer Code] = CustomerCode
'    rstCust![customer name] = CustomerName
'    rstCust.Update
'  End If
'  rstCust.Close
'End If

db.Close

End Function
Function AddCustomer(CustomerCode As String, CustomerName As String)
Dim db As Database
Dim rstCust As DAO.Recordset

Set db = OpenDatabase(AccessDBPath)
Set rstCust = db.OpenRecordset("SELECT * FROM Customers WHERE [Customer code] = '" & CustomerCode & "'")

If rstCust.RecordCount = 0 Then
    rstCust.AddNew
    rstCust![Customer Code] = CustomerCode
    rstCust![customer name] = CustomerName
    rstCust.Update
End If

rstCust.Close
db.Close

End Function
Function AddInkType(InkTypeCode As String)
Dim db As Database
Dim rstInk As DAO.Recordset

Set db = OpenDatabase(AccessDBPath)
Set rstInk = db.OpenRecordset("SELECT * FROM [Ink Type] WHERE [Ink Type] = '" & InkTypeCode & "'")

If rstInk.RecordCount = 0 Then
    rstInk.AddNew
    rstInk![Ink Type] = InkTypeCode
    rstInk.Update
End If

rstInk.Close
db.Close

End Function
Function AddPress(PressCode As String, PressName As String)
Dim db As Database
Dim qu As String
Dim rstPress As DAO.Recordset

WriteDesignTraceLog ("AccessDB: SELECT * FROM Presses WHERE [Press Number] = '" & PressCode & "'")
Set db = OpenDatabase(AccessDBPath)

Set rstPress = db.OpenRecordset("SELECT * FROM Presses WHERE [Press Number] = '" & PressCode & "'")
If rstPress.RecordCount = 0 Then
    WriteDesignTraceLog ("AddPress: " & PressCode & ", " & PressName)
    rstPress.AddNew
    rstPress![Press Number] = PressCode
    rstPress![Press name] = PressName
    rstPress.Update
End If
rstPress.Close
db.Close

End Function


Function UpdateDesign(designcode As String, designname As String, Customer As String, Substrate As String, Press As String, width As Single, ImpressionLength As Single, DateAmended As String, TimeAmended As String, Comments As String, ImagePath As String, MyInkType As String)

Dim db As Database
Dim rstDesign As DAO.Recordset
Dim q As String

Set db = OpenDatabase(AccessDBPath)
designname = Replace(designname, "'", "`")
designname = Left(designname, 50)

designcode = Replace(designcode, "'", "`")
designcode = Left(designcode, 18)

'Substrate = Replace(Substrate, "'", " ")
'If IsNull(Substrate) Then
'  Substrate = "NO SUBSTATE CODE"
'ElseIf Substrate = "" Then
'  Substrate = "NO SUBSTATE CODE"
'End If

WriteDesignTraceLog ("UpdateDesign: " & designcode & ", substrate=" & Substrate)

q = "SELECT * FROM [Designs] WHERE [Design code] = '" & designcode & "'"
WriteDesignTraceLog (q)
Set rstDesign = db.OpenRecordset(q)
If rstDesign.RecordCount <> 0 Then
    WriteDesignTraceLog ("UpdateDesign: Edit " & designcode)
    
    rstDesign.Edit
    If designname <> "" Then
        rstDesign![design name] = designname
    Else
        rstDesign![design name] = "NO DESIGN NAME"
    End If
    rstDesign![Customer] = Customer
    rstDesign![Substrate] = Substrate
    rstDesign![Printing Press] = Press
    rstDesign![Web Width] = width
    rstDesign![Impression Length] = ImpressionLength
    rstDesign![Date Amended] = CDate(DateAmended)
    rstDesign![Time Amended] = CDate(TimeAmended)
    If MyInkType <> "" Then
        rstDesign![Category] = MyInkType
    End If
    rstDesign![Print Process] = PrintProcess
    rstDesign![Comments] = Comments
    rstDesign![Image Path] = ImagePath
    
    If DebugLevel = 1 Then
        WriteDesignTraceLog ("UpdateDesign/spec " & designcode & " - saving")
    End If
    rstDesign.Update
    
Else
    WriteDesignTraceLog ("UpdateDesign/spec AddNew: " & designcode)
    rstDesign.AddNew
    rstDesign![Design Code] = designcode
    If designname <> "" Then
        rstDesign![design name] = designname
    Else
        rstDesign![design name] = "NO DESIGN NAME"
    End If
    
    rstDesign![Customer] = Customer
    rstDesign![Substrate] = Substrate
    rstDesign![Printing Press] = Press
    rstDesign![Web Width] = width
    rstDesign![Impression Length] = ImpressionLength
    rstDesign![Date Created] = Date
    rstDesign![Date Amended] = CDate(DateAmended)
    rstDesign![Time Amended] = CDate(TimeAmended)
    If MyInkType <> "" Then
        rstDesign![Category] = MyInkType
    End If
    rstDesign![Print Process] = PrintProcess
    rstDesign![Comments] = Comments
    rstDesign![Image Path] = ImagePath
    WriteDesignTraceLog ("UpdateDesign/spec New Spec: update")
    rstDesign.Update
    
    If DebugLevel = 1 Then
        WriteDesignTraceLog ("UpdateDesign " & designcode & " Added")
    End If
End If

rstDesign.Close
db.Close

End Function

Function UpdateSubstrate(SubstrateCode As String)

Dim db As Database
Dim rstSubs As DAO.Recordset

Set db = OpenDatabase(AccessDBPath)
' limit size, remove special chars
SubstrateCode = Left(SubstrateCode, 40)
SubstrateCode = Replace(SubstrateCode, "'", " ")
If IsNull(SubstrateCode) Then
  SubstrateCode = "NO SUBSTATE CODE"
ElseIf SubstrateCode = "" Then
  SubstrateCode = "NO SUBSTATE CODE"
End If
Set rstSubs = db.OpenRecordset("SELECT * FROM [Substrates] WHERE [Substrate code] = '" & SubstrateCode & "'")
If rstSubs.RecordCount <> 0 Then
    'WriteDesignTraceLog ("Substrate exists: " & SubstrateCode)
Else
    WriteDesignTraceLog ("Substrate added: " & SubstrateCode)
    rstSubs.AddNew
    rstSubs![substrate code] = SubstrateCode
    rstSubs![substrate name] = SubstrateCode
    rstSubs.Update
End If

rstSubs.Close
db.Close
UpdateSubstrate = SubstrateCode
End Function

Function AddJob(worksorder As String, designcode As String)

Dim db As Database
Dim rstJobs As DAO.Recordset

Set db = OpenDatabase(AccessDBPath)
Set rstJobs = db.OpenRecordset("SELECT * FROM [Works Orders] WHERE [Works Order Number] = '" & worksorder & "'")
WriteJobsTraceLog ("AddJob: " & "SELECT * FROM [Works Orders] WHERE [Works Order Number] = '" & worksorder & "'")

If rstJobs.RecordCount <> 0 Then
    If DebugLevel = 1 Then
        WriteJobsTraceLog ("AddJob " & worksorder & " Already exists on the ink system")
    End If
Else
    WriteJobsTraceLog ("AddJob: adding new " & worksorder)
    rstJobs.AddNew
    rstJobs![works order number] = worksorder
    rstJobs![Design Code] = designcode
    rstJobs![Date Created] = Date
    rstJobs.Update
End If

rstJobs.Close
db.Close

End Function

Private Sub LoadSetups()
    Dim InputData As String
    Dim Counter As Double
    Dim ListFile As String
    Dim MyPos As Integer
    Dim MyLength As Integer
    Dim fso

    Counter = 0
    DebugLevel = 0
    
    Close #1
    ListFile = App.Path & "\" & "setups.txt"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(ListFile)) Then
    
    Else
        MsgBox ListFile & " is not available!", vbCritical, "InkToJob"
        Exit Sub
    End If

    Open ListFile For Input As #1   ' Open file for input.
    
    ' TODO: Variables are expected in sequential order, horrible.
    ' There must be a general way of reading ini files?
    
    Line Input #1, InputData   ' Read first line of data.
    ' A binary comparison starting at position 1.
    MyPos = InStr(1, InputData, ":", 0)
    MyLength = Len(InputData)
    InputData = Trim(Right(InputData, (MyLength - MyPos)))
    'Me.txtDatabaseName.Text = InputData
    MySQLDatabaseName = InputData

    Line Input #1, InputData   ' Read line of data.
    MyPos = InStr(1, InputData, ":", 0)
    MyLength = Len(InputData)
    InputData = Trim(Right(InputData, (MyLength - MyPos)))
    'Me.TxtDriver.Text = InputData
    MySQLDriver = InputData
    
    Line Input #1, InputData   ' Read line of data.
    MyPos = InStr(1, InputData, ":", 0)
    MyLength = Len(InputData)
    InputData = Trim(Right(InputData, (MyLength - MyPos)))
    'Me.txtHost.Text = InputData
    MySQLHost = InputData
   
    Line Input #1, InputData   ' Read line of data.
    MyPos = InStr(1, InputData, ":", 0)
    MyLength = Len(InputData)
    InputData = Trim(Right(InputData, (MyLength - MyPos)))
    'Me.txtPassword.Text = InputData
    MySQLPassword = InputData
    
    Line Input #1, InputData   ' Read line of data.
    MyPos = InStr(1, InputData, ":", 0)
    MyLength = Len(InputData)
    InputData = Trim(Right(InputData, (MyLength - MyPos)))
    'Me.txtUsername.Text = InputData
    MySQLUserName = InputData
   
    Line Input #1, InputData   ' Read line of data.
    MyPos = InStr(1, InputData, ":", 0)
    MyLength = Len(InputData)
    InputData = Trim(Right(InputData, (MyLength - MyPos)))
    'Me.txtPort.Text = InputData
    MySQLPort = InputData

    Line Input #1, InputData   ' Read line of data.
    MyPos = InStr(1, InputData, ":", 0)
    MyLength = Len(InputData)
    InputData = Trim(Right(InputData, (MyLength - MyPos)))
    'Me.TxtDatabasePath.Text = InputData
    AccessDBPath = InputData

    Line Input #1, InputData   ' Read line of data.
    MyPos = InStr(1, InputData, ":", 0)
    MyLength = Len(InputData)
    InputData = Trim(Right(InputData, (MyLength - MyPos)))
    'Me.TxtDatabasePath.Text = InputData
    PrintProcess = InputData

    Line Input #1, InputData   ' Read line of data.
    MyPos = InStr(1, InputData, ":", 0)
    MyLength = Len(InputData)
    InputData = Trim(Right(InputData, (MyLength - MyPos)))
    'Me.TxtDatabasePath.Text = InputData
    Timer2CounterStart = InputData

    Line Input #1, InputData   ' Read line of data.
    MyPos = InStr(1, InputData, ":", 0)
    MyLength = Len(InputData)
    InputData = Trim(Right(InputData, (MyLength - MyPos)))
    'Me.TxtDatabasePath.Text = InputData
    DaysToLook = InputData
    
    Line Input #1, InputData   ' Read line of data.
    MyPos = InStr(1, InputData, ":", 0)
    MyLength = Len(InputData)
    InputData = Trim(Right(InputData, (MyLength - MyPos)))
    DebugLevel = InputData

    Close #1   ' Close file.

End Sub


Private Sub Form_Terminate()
  'MsgBox "Job/Ink systems sync completed."
End Sub

Private Sub Timer1_Timer()
    ' THIS TIMER IS USED TO CLOSE THE PROGRAM AFTER IMPORT HAS FINISHED.
    
    ' TODO For development, disable:
    Unload MainFrm
    Set MainFrm = Nothing
End

End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    
    ' DEV: comment the following to disable automated execution
    If DebugLevel = 0 Then
        Me.List1.Clear
        Me.List2.Clear
        Me.List3.Clear
        
        Call bbMigrateSpecs_Click
        'Call Customers_Click
        Call Command3_Click
        Call Command4_Click
        'Call Command5_Click
        Me.Timer1.Enabled = True
    End If
    
    'MsgBox "Job/Ink system sync completed."
End Sub
