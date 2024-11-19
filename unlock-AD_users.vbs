Option Explicit
On Error Resume Next

Dim x

' Windows Event Log'a yazmak için
Dim objShell
Set objShell = CreateObject("WScript.Shell")

' Active Directory'ye bağlan
Dim objRootDSE, strDomainPath
Set objRootDSE = GetObject("LDAP://RootDSE")
If Err.Number <> 0 Then
    objShell.LogEvent 1, "Active Directory'ye baglanma hatasi: " & Err.Description
    x = MsgBox("Active Directory'ye baglanma hatasi: " & Err.Description,64,"Unlock User Accounts")
    WScript.Quit 1
End If
strDomainPath = objRootDSE.Get("defaultNamingContext")

' Log dosyasına ve Event Log'a yazmak için fonksiyon
Sub WriteLog(logMessage)
    On Error Resume Next
    
    ' Her durumda Event Log'a yaz
    objShell.LogEvent 4, logMessage
    
    On Error Goto 0
End Sub

' Son 24 saat içinde kilitlenen hesapları bul ve aç
Dim objConnection, objCommand, objRecordSet
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.CommandText = "<LDAP://" & strDomainPath & ">;(&(objectCategory=person)(objectClass=user)(lockoutTime>=1)(!(lockoutTime=0)));distinguishedName,sAMAccountName;subtree"

Set objRecordSet = objCommand.Execute

' Kilitlenen hesapları aç
Dim objUser
Do Until objRecordSet.EOF
    Set objUser = GetObject("LDAP://" & objRecordSet.Fields("distinguishedName").Value)
    
    ' Hesabın kilitli olup olmadığını kontrol et
    If objUser.IsAccountLocked Then
        ' Hesap kilitliyse kilidi aç
        objUser.IsAccountLocked = False
        objUser.SetInfo
        If Err.Number = 0 Then
            WriteLog "Kullanici hesabi '" & objRecordSet.Fields("sAMAccountName").Value & "' kilidi acildi."
			x = MsgBox(objRecordSet.Fields("sAMAccountName").Value & "' kilidi acildi.",64,"Unlock User Accounts")
        Else
            WriteLog "Kullanici hesabi '" & objRecordSet.Fields("sAMAccountName").Value & "' kilit acma hatasi: " & Err.Description
			x = MsgBox(objRecordSet.Fields("sAMAccountName").Value & "' kilit acma hatasi: " & Err.Description,16,"Unlock User Accounts")
        End If
    Else
        ' Hesap kilitli değilse log'a yaz
        WriteLog "Kullanici hesabi '" & objRecordSet.Fields("sAMAccountName").Value & "' zaten kilitli degil."
    End If
    
    objRecordSet.MoveNext
Loop

' Temizlik
Set objUser = Nothing
Set objRecordSet = Nothing
Set objCommand = Nothing
Set objConnection = Nothing
Set objRootDSE = Nothing
Set objShell = Nothing

On Error Goto 0

'WriteLog "Kilitli hesap kontrolu ve kilit acma islemi tamamlandi."
x = MsgBox("Kilitli hesap kontrolu ve kilit acma islemi tamamlandi.",64,"Kilitli hesap kontrolu")
