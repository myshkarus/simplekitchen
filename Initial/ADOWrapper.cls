VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADOWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Rem /************************************************************************************/
Rem /* FILENAME       :  modshared_ado                                                  */
Rem /* TYPE           :  VBA Module                                                     */
Rem /* DESCRIPTION    :  VBA code which is required for Late Binding ADO DB objects     */
Rem /*                   Found at:                                                      */
Rem /* http://msdn.microsoft.com/en-us/library/windows/desktop/ms678353%28v=vs.85%29.aspx */
Rem /*                                                                                  */
Rem /* AUTHOR         :  Michael D Lueck                                                */
Rem /*                   mlueck@lueckdatasystems.com                                    */
Rem /*                                                                                  */
Rem /* NEEDS          :                                                                 */
Rem /*                                                                                  */
Rem /* USAGE          :                                                                 */
Rem /*                                                                                  */
Rem /* REVISION HISTORY                                                                 */
Rem /*                                                                                  */
Rem /* DATE       REVISED BY DESCRIPTION OF CHANGE                                      */
Rem /* ---------- ---------- -------------------------------------------------------    */
Rem /* 03/01/2012 MDL        Initial Creation                                           */
Rem /************************************************************************************/


Option Explicit
'Option Compare Database

Private Enum ADOCursorLocation
  adUseNone = 1
  adUseServer = 2
  adUseClient = 3
End Enum

Private Enum ADOCursorTypeEnum
  adOpenUnspecified = -1
  adOpenForwardOnly = 0
  adOpenKeyset = 1
  adOpenDynamic = 2
  adOpenStatic = 3
End Enum

Private Enum ADOLockType
  adLockUnspecified = -1
  adLockReadOnly = 1
  adLockPessimistic = 2
  adLockOptimistic = 3
  adLockBatchOptimistic = 4
End Enum


Private cUserID As String
Private cPassword As String
Private SqlQueryString As String
Private connectionString As String

Private cConnection As Object                    'ADODB.Connection
Private cRecordset As Object                     'ADODB.Recordset


Public Function GetRecordset(ByVal sqlQuery As String, Optional ByVal userID As String, Optional ByVal pwd As String) As Object 'ADODB.Recordset
  cUserID = userID
  cPassword = pwd
  
  SqlQueryString = sqlQuery
  cOpenConnection
  cOpenRecordset
  Set GetRecordset = cRecordset
End Function

Private Function cOpenConnection()
  Set cConnection = CreateObject("ADODB.Connection") 'New ADODB.Connection
  Debug.Print connectionString
  cConnection.Open connectionString              ', cUserID, cPassword
End Function

Private Function cCloseConnection()
  cConnection.Close
  Set cConnection = Nothing
End Function

Private Function cOpenRecordset()
  Set cRecordset = CreateObject("ADODB.Recordset") 'New ADODB.Recordset
  cRecordset.CursorLocation = ADOCursorLocation.adUseClient
  
  
  'Set rs = Db.OpenSchema(ADOSchema.adSchemaColumns)
  cRecordset.Open SqlQueryString, cConnection, ADOCursorTypeEnum.adOpenStatic, ADOLockType.adLockReadOnly
End Function

Private Function cCloseRecordset()
  cRecordset.Close
  Set cRecordset = Nothing
End Function

Private Sub Class_Initialize()
  Dim provider As String
  Dim dataSource As String
  Dim initialCatalog As String
  Dim securityInfo As String
  
  On Error GoTo ErrHandler
  
  provider = "Provider=" & SysInfo.provider & ";"
  dataSource = "Data Source =" & DatabasePath & ";"
  initialCatalog = ""
  securityInfo = "Persist Security Info=False"
  
  connectionString = provider & dataSource & securityInfo

 

  '''''  connDB.Open connectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
  '''''  strConn = "Provider=SQLOLEDB.1;" & "Data Source=" & sqlAdr & "; Initial Catalog=" & sqlDb & ";Trusted_connection=yes;"
  '''''
  '''''  Dim strConn As String

  '''''  'Use the SQL Server OLE DB Provider.
  '''''  strConn = "PROVIDER=SQLOLEDB;"
  '''''  'Connect to the Pubs database on the local server.
  '''''  strConn = strConn & "DATA SOURCE=a_sql_server;INITIAL CATALOG=a_database;"
  '''''  'Use an integrated login.
  '''''  strConn = strConn & " INTEGRATED SECURITY=sspi;"
Exit Sub
ErrHandler:
End Sub

Private Sub Class_Terminate()
  If Not (cRecordset Is Nothing) Then Set cRecordset = Nothing
  If Not (cConnection Is Nothing) Then Set cConnection = Nothing
End Sub

