Attribute VB_Name = "Module3"
Public Con100 As ADODB.Connection

Public Sub Conn_2007()
On Error GoTo Error
Dim StrMdbPath, StrConn As String

    StrMdbPath = App.Path & "\Database\" & App.Title & "_DB.mdb"
    StrConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & StrMdbPath & ";Jet OLEDB:Database Password=authentic;"
    Set Con100 = New ADODB.Connection
    Con100.Open StrConn

Exit Sub
Error:
MsgBox "Office 2007 Error", vbInformation
End Sub

Public Function CreateField(ByVal DBConn, strTable, strField, strType As String) As Boolean
Dim Sql As String

    If (strField <> vbNullString) Then
        Sql = "ALTER TABLE " & strTable & " ADD COLUMN " & strField & " " & strType
        DBConn.Execute Sql
        Sql = "UPDATE " & strTable & " Set " & strField & " = 0"
        DBConn.Execute Sql
        CreateField = True
    End If

End Function

Public Function FieldExists(ByVal DBConn, TableName, FieldName As String) As Boolean
Dim Rs As New ADODB.Recordset
Dim FLD As ADODB.Field

Rs.Open TableName, DBConn, adOpenStatic, adLockReadOnly, adCmdTable
For Each FLD In Rs.Fields
    If LCase(FLD.Name) = LCase(FieldName) Then
        FieldExists = True
        Exit For
    End If
Next

End Function


Public Sub Make_Column()
On Error GoTo Error
Dim Sql As String
Dim TableName As String
Dim ColName As String
Dim Row As Long
Dim Col As Long
Dim ColArrey(100) As String

'----------------Table - Model_Set
TableName = "Model_Set"
'  "Printtype"
'    Rs ("IDNo")
'    Rs("LastPartNo") = txtLastPartno.Text
'    Rs("PartNo") = txtPartNo.Text
'    Rs("Darkness") = cboDarkness.ListIndex
'    Rs("RejectionBypass") = Check1.Value
'    Rs("Vendorcode") = txtvendorCode.Text
'    Rs("linecode") = txtlinecode.Text
ColArrey(1) = "DMBypass"
ColArrey(2) = "DM1CurMin"
ColArrey(3) = "DM1CurMax"
ColArrey(4) = "DM2CurMin"
ColArrey(5) = "DM2CurMax"
ColArrey(6) = "DM1VoltMin"
ColArrey(7) = "DM1VoltMax"
ColArrey(8) = "DM2VoltMin"
ColArrey(9) = "DM2VoltMax"
ColArrey(10) = "DMTestCycle"
ColArrey(11) = "PMBypass"
ColArrey(12) = "PM1CurMin"
ColArrey(13) = "PM1CurMax"
ColArrey(14) = "PM1VoltMin"
ColArrey(15) = "PM1VoltMax"
ColArrey(16) = "PMTestCycle"
ColArrey(17) = "BMBypass"
ColArrey(18) = "BM1CurMin"
ColArrey(19) = "BM1CurMax"
ColArrey(20) = "BM2CurMin"
ColArrey(21) = "BM2CurMax"
ColArrey(22) = "BM1VoltMin"
ColArrey(23) = "BM1VoltMax"
ColArrey(24) = "BM2VoltMin"
ColArrey(25) = "BM2VoltMax"
ColArrey(26) = "BMTestCycle"
ColArrey(27) = "HMBypass"
ColArrey(28) = "HM1CurMin"
ColArrey(29) = "HM1CurMax"
ColArrey(30) = "HM1VoltMin"
ColArrey(31) = "HM1VoltMax"
ColArrey(32) = "HMTestCycle"
ColArrey(33) = "LEBypass"
ColArrey(34) = "LEM1CurMin"
ColArrey(35) = "LEM1CurMax"
ColArrey(36) = "LEM1VoltMin"
ColArrey(37) = "LEM1VoltMax"
ColArrey(38) = "LEMTestCycle"
ColArrey(39) = "LIMBypass"
ColArrey(40) = "LIM1CurMin"
ColArrey(41) = "LIM1CurMax"
ColArrey(42) = "LIM1VoltMin"
ColArrey(43) = "LIM1VoltMax"
ColArrey(44) = "LIMTestCycle"
ColArrey(45) = "EKSBypass"
ColArrey(46) = "EKS1CurMin"
ColArrey(47) = "EKS1CurMax"
ColArrey(48) = "EKS1VoltMin"
ColArrey(49) = "EKS1VoltMax"
ColArrey(50) = "EKS2CurMin"
ColArrey(51) = "EKS2CurMax"
ColArrey(52) = "EKS2VoltMin"
ColArrey(53) = "EKS2VoltMax"
ColArrey(54) = "EKSTestCycle"
ColArrey(55) = "ICMinRH"
ColArrey(56) = "ICMaxRH"
ColArrey(57) = "WVMin"
ColArrey(58) = "WVMax"

For Row = 1 To 58
    ColName = ColArrey(Row)
    If FieldExists(Con, TableName, ColName) = False Then
        X = CreateField(Con, TableName, ColName, "varchar(255) DEFAULT 0")
    End If
Next

'----------------Table - Model_Set
TableName = "Model_Report"

ColArrey(1) = "DM1Cur"
ColArrey(2) = "DM2Cur"
ColArrey(3) = "DM1Volt"
ColArrey(4) = "DM2Volt"
ColArrey(5) = "PM1Cur"
ColArrey(6) = "PM1Volt"
ColArrey(7) = "BM1Cur"
ColArrey(8) = "BM2Cur"
ColArrey(9) = "BM1Volt"
ColArrey(10) = "BM2Volt"
ColArrey(11) = "HM1Cur"
ColArrey(12) = "HM1Volt"
ColArrey(13) = "LEM1Cur"
ColArrey(14) = "LEM1Volt"
ColArrey(15) = "EKS1Cur"
ColArrey(16) = "EKS1Volt"
ColArrey(17) = "EKS2Cur"
ColArrey(18) = "EKS2Volt"
ColArrey(19) = "ICLH"
ColArrey(20) = "ICRH"
ColArrey(21) = "DMResult"
ColArrey(22) = "PMResult"
ColArrey(23) = "BMResult"
ColArrey(24) = "HMResult"
ColArrey(25) = "LEMResult"
ColArrey(26) = "EKSResult"

For Row = 1 To 26
    ColName = ColArrey(Row)
    If FieldExists(Con, TableName, ColName) = False Then
        X = CreateField(Con, TableName, ColName, "varchar(255) DEFAULT 0")
    End If
Next





'------------------------------------



'TableName = "user_list"
'ColName = "AccessType"
'If FieldExists(Con100, TableName, ColName) = False Then
'    X = CreateField(Con100, TableName, ColName, "varchar(255) DEFAULT 0")
'End If
'-=========================================

'Sql = "create table Common_Set (ID Counter)"
'Con100.Execute Sql

'TableName = "Common_Set"
'ColName = "SetType"
'If FieldExists(Con100, TableName, ColName) = False Then
'    X = CreateField(Con100, TableName, ColName, "varchar(255) DEFAULT 0")
'End If
'
''Sql = "Update Common_Set Set SetType='CommonSet'"
''Con100.Execute Sql
'
'ColName = "ComPort1"
'If FieldExists(Con100, TableName, ColName) = False Then
'    X = CreateField(Con100, TableName, ColName, "varchar(255) DEFAULT 0")
'End If

'Con100.Close

Exit Sub
Error:
'Con100.Close
End Sub
