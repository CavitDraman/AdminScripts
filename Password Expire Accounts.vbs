On Error Resume Next

' Initialize output string
Dim output
output = ""

' Get domain information
Set rootDSE = GetObject("LDAP://rootDSE")
domainDN = rootDSE.Get("defaultNamingContext")

' Calculate dates for 1 week range
oneWeekAgo = DateAdd("d", -7, Now)

' Create LDAP filter for user accounts
ldapFilter = "(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"

' Set up ADODB connection
Set ado = CreateObject("ADODB.Connection")
ado.Provider = "ADSDSOObject"
ado.Open "Active Directory Provider"

' Execute search
Set objectList = ado.Execute("<LDAP://" & domainDN & ">;" & ldapFilter & ";sAMAccountName,pwdLastSet,msDS-UserPasswordExpiryTimeComputed,distinguishedName;subtree")

' Add header to output
output = output & "Accounts with Passwords Expired in Last Week in " & domainDN & vbCrLf
output = output & "================================================" & vbCrLf
output = output & "Username" & vbTab & vbTab & "Password Last Set" & vbTab & "Password Expiry Date" & vbCrLf
output = output & "------------------------------------------------" & vbCrLf

' Process results
While Not objectList.EOF
    Set user = GetObject(objectList.Fields("ADSPath"))
    
    ' Get user properties
    sAMAccountName = user.Get("sAMAccountName")
    pwdLastSet = user.Get("pwdLastSet")
    
    ' Only process if we have a valid username
    If Not IsNull(sAMAccountName) And Len(sAMAccountName) > 0 Then
        ' Get password expiry time
        user.GetInfoEx Array("msDS-UserPasswordExpiryTimeComputed"), 0
        pwdExpiry = user.Get("msDS-UserPasswordExpiryTimeComputed")
        
        ' Convert pwdLastSet to readable format
        If pwdLastSet = "0" Then
            pwdLastSetStr = "Never Set"
        Else
            pwdLastSetStr = DateAdd("s", CLng(pwdLastSet) / 10000000, #1/1/1601#)
        End If
        
        ' Convert pwdExpiry to readable format
        If pwdExpiry = "0" Then
            pwdExpiryStr = "Must Change"
        ElseIf pwdExpiry = "9223372036854775807" Then
            pwdExpiryStr = "Never Expires"
        Else
            pwdExpiryStr = DateAdd("s", CLng(pwdExpiry) / 10000000, #1/1/1601#)
            ' Show only if password expired within last week
            If pwdExpiryStr < Now And pwdExpiryStr > oneWeekAgo Then
                output = output & sAMAccountName & vbTab & vbTab & pwdLastSetStr & vbTab & pwdExpiryStr & vbCrLf
            End If
        End If
    End If
    
    objectList.MoveNext
Wend

' Cleanup
Set objectList = Nothing
Set ado = Nothing
Set rootDSE = Nothing

' Add completion message
output = output & vbCrLf & "Script completed."

' Display all output at once
WScript.Echo output