Attribute VB_Name = "Connection"
Option Explicit

'Declare public objects
Public g_MySQLConn As ADODB.Connection
Public g_MySQLError As ADODB.Error

Public Function EstablishMySQLConnection( _
    ByVal strLogin As String, _
    ByVal strPassword As String, _
    ByVal strHost As String, _
    ByVal strDatabase As String, _
    ByVal strPort As String, _
    ByVal strDriver As String) As Boolean

'Setup error handling
On Error GoTo EstablishMySQLConnection_Error

'Declare Local Variables

'__________________ Connection Credentials ____________________

Dim strConnectString As String
Dim strConnectStringDisplay As String
'Const MYSQL_USER_NAME = "root"
'Const MYSQL_PASSWORD = "CTRecord"
'Const MYSQL_HOST = "192.168.1.8"
'Const MYSQL_DATABASE = "boranpla"
'Const MYSQL_PORT = 3306
'Const MyODBCVersion = "MySQL ODBC 5.1 Driver"

Dim MYSQL_USER_NAME As String
Dim MYSQL_PASSWORD As String
Dim MYSQL_HOST As String
Dim MYSQL_DATABASE As String
Dim MYSQL_PORT As String
Dim MYODBCVERSION As String

MYSQL_USER_NAME = strLogin
MYSQL_PASSWORD = strPassword
MYSQL_HOST = strHost
MYSQL_DATABASE = strDatabase
MYSQL_PORT = strPort
MYODBCVERSION = strDriver

'______________________________________________________________

'Set a reference to the ADO connection object
Set g_MySQLConn = New ADODB.Connection

'Build the DSN or DSN-Less part of the connect string

strConnectString = "Provider=MSDASQL;" & _
                    "Driver=" & MYODBCVERSION & ";" & _
                    "Server=" & MYSQL_HOST & ";" & _
                    "Database=" & MYSQL_DATABASE & ";" & _
                    "UID=" & MYSQL_USER_NAME & ";" & _
                    "PWD=" & MYSQL_PASSWORD & ";" & _
                    "Port=" & MYSQL_PORT

strConnectStringDisplay = "Provider=MSDASQL;" & _
                    "Driver=" & MYODBCVERSION & ";" & _
                    "Server=" & MYSQL_HOST & ";" & _
                    "Database=" & MYSQL_DATABASE & ";" & _
                    "UID=" & MYSQL_USER_NAME & ";" & _
                    "Port=" & MYSQL_PORT

MainFrm![txtmysqlconnectionstring].Text = strConnectStringDisplay

'g_MySQLConn.Open strConnectString, strLogin, strPassword
g_MySQLConn.Open strConnectString, MYSQL_USER_NAME, MYSQL_PASSWORD

'Check Conncetion State
If g_MySQLConn.State <> adStateOpen Then
    EstablishMySQLConnection = False
Else
    EstablishMySQLConnection = True
End If


Exit Function


EstablishMySQLConnection_Error:

'Connection failed, display error message
Dim strError
For Each g_MySQLError In g_MySQLConn.Errors
    strError = strError & g_MySQLError.Number & "  :  " & _
        g_MySQLError.Description & vbCrLf & vbCrLf
Next
MsgBox strError, vbCritical + vbOKOnly, "Login Error"

End Function


