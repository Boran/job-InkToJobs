VERSION 5.00
Begin VB.Form MainFrm 
   Caption         =   "MySQL Connection"
   ClientHeight    =   9675
   ClientLeft      =   1035
   ClientTop       =   645
   ClientWidth     =   13095
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   9675
   ScaleWidth      =   13095
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   4800
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   1560
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8280
      TabIndex        =   11
      Text            =   "Counter"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Test Button"
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtmysqlconnectionstring 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   4920
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   480
      Width           =   6855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF80&
      Caption         =   "v_ink_job"
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
      Width           =   3615
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
      Height          =   2370
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
      Caption         =   "v_ink_spec"
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
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "SET MYSQL CONNECTION"
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
      Top             =   480
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "v_ink_cust"
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
      Width           =   3615
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
   Begin VB.Label Label8 
      Caption         =   "V2.02"
      Height          =   255
      Left            =   8280
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "MySQL Connection String:"
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
      Left            =   4920
      TabIndex        =   3
      Top             =   120
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
Dim Timer2Counter As Integer
Dim Timer2CounterStart As Integer
Dim DaysToLook As Integer


Public Sub TerminateConnection()
    On Error Resume Next
    
    g_MySQLConn.Close
    Set g_MySQLConn = Nothing
    Set g_MySQLError = Nothing
    
        
End Sub


Private Sub Command1_Click()

Answer2 = False

Me.txtmysqlconnectionstring.Text = ""

'Answer2 = EstablishMySQLConnection(Me.txtUsername.Text, Me.txtPassword.Text, Me.txtHost.Text, Me.txtDatabaseName.Text, Me.txtPort.Text, Me.TxtDriver.Text)
Answer2 = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)


MsgBox "Connected to MySQL Server = " & Answer2

End Sub

Private Sub Command2_Click()

Dim MyView As String
Dim CustCode As String
Dim CustName As String
Dim CustAddress As String
Dim ValueStr As String
Dim HeaderStr As String

WriteTraceLog ("Attempting to connect - Customers")

Answer = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)

WriteTraceLog ("Connection = " & Answer)

If Answer = True Then
    
    HeaderStr = "Customer Code" & vbTab & vbTab & "Customer Name" & vbTab & vbTab & "Customer Address"
    
    Me.List1.AddItem HeaderStr
    
    MyView = "`v_ink_cust`"

    rst1.Open MyView, g_MySQLConn, adOpenDynamic, adLockOptimistic
       
    WriteTraceLog ("Opended v_ink_cust")
       
       
    Do Until rst1.EOF
             
        If Not IsNull(rst1![cust_code]) Then
                CustCode = rst1![cust_code]
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
                WriteTraceLog ("Calling UpdateCustomer - " & CustCode & "," & CustName & "," & CustAddress)
                Call UpdateCustomer(CustCode, CustName)
                
        End If

        ValueStr = CustCode & vbTab & vbTab & CustName & vbTab & vbTab & CustAddress
        
        Me.List1.AddItem ValueStr

        rst1.MoveNext
    Loop

    rst1.Close

    Call TerminateConnection
End If

End Sub


Private Sub Command3_Click()

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


'Answer = EstablishMySQLConnection(Me.txtUsername.Text, Me.txtPassword.Text, Me.txtHost.Text, Me.txtDatabaseName.Text, Me.txtPort.Text, Me.TxtDriver.Text)
Answer = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)

If Answer = True Then
    
    
    MyNewView = "SELECT DISTINCT Printer from v_ink_spec"

    rst1.Open MyNewView, g_MySQLConn, adOpenDynamic, adLockOptimistic
       
    Me.List2.AddItem "CONNECTED"
      
    Do Until rst1.EOF
        Counter = Counter + 1
        Me.Text1 = Counter
        
        If Not IsNull(rst1![Printer]) Then
            Printer = rst1![Printer]
        Else
            Printer = "NO PRINTER"
        End If
        
        Me.List2.AddItem "Checking Printer " & Printer
        
        WriteDesignTraceLog ("Calling Function AddPress " & Printer)
        
        Call AddPress(Printer, Printer)
        
        rst1.MoveNext
    Loop
        
        
    
            
        'WriteDesignTraceLog (CStr(Spec & "," & CustCode & "," & Design & "," & Substrate & "," & PrRepeat & "," & PrWidth & "," & InkType & "," & Printer & "," & LastChangeOriginal & "," & LastChangeday & "/" & LastChangemonth & "/" & LastChangeyear & "," & LastChangeTime & "," & LastChange24))
            

    Call TerminateConnection

End If

Call AddCustomers

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
    
    
    MyView = "`v_ink_spec`"

    MyNewView = "SELECT DISTINCT cust_code from v_ink_spec"


    rst1.Open MyNewView, g_MySQLConn, adOpenDynamic, adLockOptimistic
       
    Me.List2.AddItem "CONNECTED ADD CUSTOMERS"
      
    Do Until rst1.EOF
        Counter = Counter + 1
        Me.Text1 = Counter
        DoEvents
        
        If Not IsNull(rst1![cust_code]) Then
            CustCode = rst1![cust_code]
        Else
            CustCode = "NO CUSTOMER CODE"
        End If

        Me.List2.AddItem "Checking Customer " & CustCode

        WriteDesignTraceLog ("Calling Function AddCustomer " & CustCode)

        Call AddCustomer(CustCode, CustCode)
        
        rst1.MoveNext
    
    Loop

    

    Call TerminateConnection
End If

Call AddInkTypes

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
    
    
    MyView = "`v_ink_spec`"

    MyNewView = "SELECT DISTINCT inktype from v_ink_spec"


    rst1.Open MyNewView, g_MySQLConn, adOpenDynamic, adLockOptimistic
       
    Me.List2.AddItem "CONNECTED ADD INKTYPES"
      
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

        WriteDesignTraceLog ("Calling Function AddInkType " & InkCode)

        Call AddInkType(InkCode)
        
        rst1.MoveNext
    
    Loop

    

    Call TerminateConnection
End If

Call AddSubstrates

End Sub

Private Sub AddSubstrates()

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
    
       
    MyView = "`v_ink_spec`"
    
    MyNewView = "SELECT DISTINCT Substrate FROM v_ink_spec"

    rst1.Open MyNewView, g_MySQLConn, adOpenDynamic, adLockOptimistic
       
    Me.List2.AddItem "CONNECTED ADD SUBSTRATES"
            
    Do Until rst1.EOF
        Counter = Counter + 1
        Me.Text1 = Counter
        DoEvents
        
        If Not IsNull(rst1![Substrate]) Then
            Substrate = rst1![Substrate]
        Else
            Substrate = "NO SUBSTATE CODE"
        End If

        Me.List2.AddItem "Checking Substrate " & Substrate

        WriteDesignTraceLog ("Calling Function UpdateSubstrate " & Substrate)

        Call UpdateSubstrate(Substrate)
        
        rst1.MoveNext
    
    Loop
     
      
      

    Call TerminateConnection
End If

'MsgBox "DONE"
Call AddDesigns

End Sub

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


'Answer = EstablishMySQLConnection(Me.txtUsername.Text, Me.txtPassword.Text, Me.txtHost.Text, Me.txtDatabaseName.Text, Me.txtPort.Text, Me.TxtDriver.Text)
Answer = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)

If Answer = True Then
    
    'HeaderStr = "Spec" & " | " & "Customer Code" & " | " & "Design"
    HeaderStr = "Spec" & vbTab & vbTab & "Customer Code" & vbTab & vbTab & "Design"
    
    Me.List2.AddItem HeaderStr
    
    Me.List2.AddItem ""
    
    MyView = "`v_ink_spec`"

    rst1.Open MyView, g_MySQLConn, adOpenDynamic, adLockOptimistic
       
    Me.List2.AddItem "CONNECTED"
      
    Do Until rst1.EOF
        Counter = Counter + 1
        Me.Text1 = Counter
        DoEvents
    
        'MsgBox "Counter = " & Counter
        
        If Not IsNull(rst1![Spec]) Then
            Spec = rst1![Spec]
            If Not IsNull(rst1![cust_code]) Then
                CustCode = rst1![cust_code]
            Else
                CustCode = "NO CUSTOMER CODE"
            End If
            If Not IsNull(rst1![Design]) Then
                Design = rst1![Design]
            Else
                Design = "NO DESIGN"
            End If
            If Not IsNull(rst1![Substrate]) Then
                Substrate = rst1![Substrate]
            Else
                Substrate = "NO SUBSTRATE"
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
            
            'Call AddPress(Printer, Printer)
            'Call AddCustomer(CustCode, CustCode)
            'Call UpdateSubstrate(Substrate)
            WriteDesignTraceLog "Calling Function Update Design"
            
            Call UpdateDesign(Spec, Design, CustCode, Substrate, Printer, CSng(PrWidth), CSng(PrRepeat), CDate(LastChangeday & "/" & LastChangemonth & "/" & LastChangeyear), LastChange24, MyComment, DesignImage, InkType)

        End If
            

        'ValueStr = Spec & " | " & CustCode & " | " & Design
        ValueStr = Spec & vbTab & vbTab & CustCode & vbTab & vbTab & Design
        
        Me.List2.AddItem ValueStr

        rst1.MoveNext
    Loop

    rst1.Close

    Call TerminateConnection
End If

End Sub


Private Sub Command4_Click()

Dim MyView As String
Dim job As String
Dim Spec As String
Dim DelDate As String
Dim CustCode As String
Dim Design As String
Dim ValueStr As String
Dim HeaderStr As String


WriteJobsTraceLog ("Attempting to connect - Jobs")

Answer = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)

WriteJobsTraceLog ("Connection = " & Answer)

If Answer = True Then
    
    WriteJobsTraceLog ("Connection = " & Answer)
  
    HeaderStr = "Job" & vbTab & vbTab & "Spec" & vbTab & vbTab & "Del Date" & vbTab & vbTab & "Customer Code" & vbTab & vbTab & "Design"
    
    Me.List3.AddItem HeaderStr
        
    MyView = "`v_ink_job`"

    rst1.Open MyView, g_MySQLConn, adOpenDynamic, adLockOptimistic
       
    WriteJobsTraceLog ("Opended v_ink_job")
      
    Do Until rst1.EOF
    
        
        If Not IsNull(rst1![job]) Then
            job = rst1![job]
            
          
            If Not IsNull(rst1![Spec]) Then
                Spec = rst1![Spec]
            Else
                Spec = "   "
            End If
                    
        End If
        
        WriteJobsTraceLog (job & "," & Spec & "," & DelDate & "," & CustCode & "," & Design)

        'ValueStr = Spec & " | " & CustCode & " | " & Design
        ValueStr = job & vbTab & vbTab & Spec & vbTab & vbTab & DelDate & vbTab & vbTab & CustCode & vbTab & vbTab & Design
        
        Me.List3.AddItem ValueStr
        
        WriteJobsTraceLog ("Calling AddJob - " & job & "  -  " & Spec)
        
        Call AddJob(job, Spec)

        rst1.MoveNext
    Loop

    rst1.Close

    Call TerminateConnection
End If


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

MsgBox "Form load done"
End Sub

Function WriteTraceLog(TraceText As String)

  'Me.List1.AddItem "Error see log for details"
  
  Const ForReading = 1, ForWriting = 2, ForAppending = 8, logdir = "Logs\CustomerTraceLog"
  Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  If Not fso.folderexists("Logs") Then
    MsgBox "Cannot open directory " & logdir
  
  Else
   If Dir(logdir & Day(Now) & " - " & Month(Now) & " - " & Year(Now) & ".txt") = "" Then
        Set f = fso.CreateTextFile(logdir & Day(Now) & " - " & Month(Now) & " - " & Year(Now) & ".txt", True)
        f.Close
   End If
   Set f = fso.OpenTextFile(logdir & Day(Now) & " - " & Month(Now) & " - " & Year(Now) & ".txt", ForAppending, False)
   f.WriteLine Now & "," & TraceText
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
  f.WriteLine Now & "," & TraceText
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
  f.WriteLine Now & "," & TraceText
  'f.WriteLine TraceText
  
  f.Close

End Function

Private Sub Command5_Click()
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

WriteExportTraceLog ("Entered Instert data")



Set db = OpenDatabase(AccessDBPath)

Set rst = db.OpenRecordset("SELECT [Costing Reports Details].[Date Closed], [Costing Reports Details].* From [Costing Reports Details] WHERE [Costing Reports Details].[Date Closed] Between Date()-" & DaysToLook & " And Date() ORDER BY [Costing Reports Details].[Date Closed] DESC")

WriteExportTraceLog ("Executed ACCESS SELECT")


If rst.RecordCount <> 0 Then

Do Until rst.EOF

    Counter = Counter + 1
    Me.Text1 = "INSERT " & Counter

    WorksOrderNumberI = rst![works order number]
    DesignCodeI = rst![design code]
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
    
    WriteExportTraceLog ("Call InsertData Works Order = " & WorksOrderNumberI)
    
    
    
    Call InsertData(WorksOrderNumberI, DesignCodeI, DesignNameI, InkWeightI, DateOpenedI, CustomerI, WhiteWeightI, ColoursWeightI, LacquerWeightI, OtherWeightI, ReturnsReturnedI, ReturnsReturnedCostI, ReturnsIssuedI, ReturnsIssuedtoJobCostI, TotalTargetCostI, TargetCostPer1000sqmI, InkCostI, TotalCostI, EstimatedSQMI, SQMI, WhitesCostI, ColoursCostI, LacquerCostI, ActualCost1000I, DateClosedI)
    
    WriteExportTraceLog ("Moving to next record")
    
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

WriteExportTraceLog ("Connecting to SQL Server")


Answer = EstablishMySQLConnection(MySQLUserName, MySQLPassword, MySQLHost, MySQLDatabaseName, MySQLPort, MySQLDriver)
'Answer = EstablishMySQLConnection("root", "CTRecord", "IBM1", "Boranpla", 3306, "MySQL ODBC 5.1 Driver")



If Answer = True Then

WriteExportTraceLog ("Connection = " & Answer)


rst.CursorLocation = adUseClient


StrSQL = "SELECT `works order number` FROM `ink_costing reports details` WHERE `works order number` ='" & worksorder & "'"


rst.Open StrSQL, g_MySQLConn, adOpenKeyset, adLockOptimistic

WriteExportTraceLog ("Checking Works Order " & worksorder & " Exists on MySQL Server")

NoRecord = False
    
Do Until rst.EOF

WriteExportTraceLog ("Works Order " & worksorder & " Exists")

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


WriteExportTraceLog ("Set INSERT statement")

rst1.Open StrSQL, g_MySQLConn, adOpenDynamic, adLockOptimistic

WriteExportTraceLog ("INSERT statement executed for works order = " & worksorder)


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
  f.WriteLine Now & "," & TraceText
  'f.WriteLine TraceText
  
  f.Close

End Function

Function UpdateCustomer(CustomerCode As String, CustomerName As String)
Dim db As Database
Dim rstCust As DAO.Recordset

Set db = OpenDatabase(AccessDBPath)

Set rstCust = db.OpenRecordset("SELECT * FROM Customers WHERE [Customer code] = '" & CustomerCode & "'")

If rstCust.RecordCount = 0 Then

    rstCust.AddNew
    rstCust![Customer Code] = CustomerCode
    rstCust![customer name] = CustomerName
    rstCust.Update
    
    WriteTraceLog ("Customer - " & CustomerCode & " added to table")

Else

    rstCust.Edit
    rstCust![customer name] = CustomerName
    rstCust.Update

    WriteTraceLog ("Customer - " & CustomerCode & " customer name updated")


End If

rstCust.Close
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
Dim rstPress As DAO.Recordset

Set db = OpenDatabase(AccessDBPath)

Set rstPress = db.OpenRecordset("SELECT * FROM Presses WHERE [Press Number] = '" & PressCode & "'")

If rstPress.RecordCount = 0 Then

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

Set db = OpenDatabase(AccessDBPath)

designname = Replace(designname, "'", "`")
designname = Left(designname, 50)
Substrate = Replace(Substrate, "'", " ")

Set rstDesign = db.OpenRecordset("SELECT * FROM [Designs] WHERE [Design code] = '" & designcode & "'")

If rstDesign.RecordCount <> 0 Then
    
    WriteDesignTraceLog ("Editing design " & designcode)
    
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
    rstDesign.Update
    
    WriteDesignTraceLog ("Design - " & designcode & " Updated")
    
   
Else
    
    WriteDesignTraceLog ("Adding New Design " & designcode)
    
    rstDesign.AddNew
    rstDesign![design code] = designcode
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
    rstDesign.Update
    
    WriteDesignTraceLog ("Design - " & designcode & " Added")
    
    
End If

rstDesign.Close
db.Close

End Function
Function UpdateSubstrate(SubstrateCode As String)

Dim db As Database
Dim rstSubs As DAO.Recordset

Set db = OpenDatabase(AccessDBPath)

SubstrateCode = Replace(SubstrateCode, "'", " ")

Set rstSubs = db.OpenRecordset("SELECT * FROM [Substrates] WHERE [Substrate code] = '" & SubstrateCode & "'")

If rstSubs.RecordCount <> 0 Then

    'MsgBox "Substrate Exists"
    
Else

    'MsgBox "Substrate Does Not Exist"
    rstSubs.AddNew
    rstSubs![substrate code] = SubstrateCode
    rstSubs![substrate name] = SubstrateCode
    
    rstSubs.Update
    
End If

rstSubs.Close
db.Close


End Function

Function AddJob(worksorder As String, designcode As String)

Dim db As Database
Dim rstJobs As DAO.Recordset

Set db = OpenDatabase(AccessDBPath)


Set rstJobs = db.OpenRecordset("SELECT * FROM [Works Orders] WHERE [Works Order Number] = '" & worksorder & "'")

If rstJobs.RecordCount <> 0 Then

    WriteJobsTraceLog ("Job - " & worksorder & " Already exists on the database.")
    
Else


    rstJobs.AddNew
    rstJobs![works order number] = worksorder
    rstJobs![design code] = designcode
    rstJobs![Date Created] = Date
    rstJobs.Update
    
    WriteJobsTraceLog ("Job - " & worksorder & " Added to the database.")
    
    
    
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
    
    Close #1
    
    ListFile = App.Path & "\" & "Setups.txt"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(ListFile)) Then
    
    Else
        MsgBox ListFile & " is not available!", vbCritical, "ComputerTel Ltd"
        Exit Sub
    End If

    Open ListFile For Input As #1   ' Open file for input.
    
    ' TODO: Varaibles are expected in sequential order, horrible. There must be a general way of read ini files?
    
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

    Close #1   ' Close file.

End Sub


Private Sub Form_Unload(Cancel As Integer)

End Sub

Private Sub Timer1_Timer()

' THIS TIMER IS USED TO CLOSE THE PROGRAM AFTER IMPORT HAS FINISHED.

Unload MainFrm
Set MainFrm = Nothing

End

End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    Me.List1.Clear
    Me.List2.Clear
    Me.List3.Clear
    Call Command2_Click
    Call Command3_Click
    Call Command4_Click
    Call Command5_Click
    Me.Timer1.Enabled = True
End Sub
