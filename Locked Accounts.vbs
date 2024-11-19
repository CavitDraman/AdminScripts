Const ADS_UF_LOCKOUT = 16

ldapFilter = "(&(sAMAccountType=805306368)(lockoutTime>=100))"

Set rootDSE = GetObject("LDAP://rootDSE")
domainDN = rootDSE.Get("defaultNamingContext")

WScript.Echo domainDN & " Locked accounts:"
''WScript.Echo

Set ado = CreateObject("ADODB.Connection")
ado.Provider = "ADSDSOObject"
ado.Open "ADSearch" 
Set objectList = ado.Execute("<LDAP://" & domainDN  & ">;" & ldapFilter & ";ADSPath,distinguishedName;subtree")

While Not objectList.EOF
    Set user = GetObject(objectList.Fields("ADSPath"))

    user.GetInfoEx Array("msDS-User-Account-Control-Computed"), 0
    flags = user.Get("msDS-User-Account-Control-Computed")
    if (flags and ADS_UF_LOCKOUT) then
        WScript.Echo objectList.Fields("distinguishedName")
    End if

    objectList.MoveNext
Wend