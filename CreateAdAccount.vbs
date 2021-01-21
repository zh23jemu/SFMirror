On Error Resume Next
              
Dim strOrg
Dim strDCServer 
Dim strScriptPath
Dim strGetsmtpserver
Dim intGetsmtpserverport
Dim strGetsmtpaccountname
Dim strGetsendemailaddress
Dim strGetsenduserreplyemailaddress
Dim strGetsmtpauthenticate
Dim strGetsendusername
Dim strGetsendpassword
Dim strOracleConnectionString 
Dim strSqlConnectionString
Dim strGetErrorSendto
Dim strGetErrorSendcc
Dim strGetSendto

Set objFileSystem= CreateObject("Scripting.FileSystemObject")
date_log = "log\Create\ToCreateAdAccountFromSF_" & replace(formatdatetime(now(),2),"/","") & "_" & replace(formatdatetime(now(),4) & ":00",":","") & ".log"
Set objLogFile = objFileSystem.OpenTextFile(date_log, 8, True) '8 is for appending


objLogFile.Write now() & " Start" & vbCrLf

If err.number <> 0 Then
  Set objFileSystem = nothing
  wscript.Quit
End If
    

If toGetParameters = false Then
  objLogFile.Write now() & " --Create AD Account Log erro:" & err.description & vbCrLf
  objLogFile.close
  Set objFileSystem = nothing
  wscript.Quit
End If


Set Conn = CreateObject("ADODB.Connection")
Conn.Open strSqlConnectionString
Set ConnMD = CreateObject("ADODB.Connection")
ConnMD.Open strSqlConnectionString


'Set rsMD = CreateObject("ADODB.Recordset")
Set rs = CreateObject("ADODB.Recordset")
Set rsReply = CreateObject("ADODB.Recordset")
'Set rsUpdateHR = CreateObject("ADODB.Recordset")

objLogFile.Write now() & " -- Start Read SF New Employee" & vbCrLf 
     				 	 
strSQL="select pi.EmpID, pi.EFNAME,pi.ELNAME,pi.CNAME as PCNAME,pi.Email,pi.adaccount,pm.Ename AS PMEname,pm.Cname as PMCname,pm.jobLevel,pi.Tel,Displayflag,"
strSQL=strSQL & " di5.DeptCName AS FCNAME,di5.DeptEName AS FENAME,di6.DeptCName AS SCNAME,di6.DeptEName AS SENAME,di7.DeptCName AS TFCNAME,di7.DeptEName AS TENAME,di8.DeptCName AS FOCNAME,di8.DeptEName AS FOENAME,pl.[CName] AS PLCName,pl.[Name] AS PLEName,pl.[Shortname]"
strSQL=strSQL & " from PerInfo pi join PostionMailSetup pm on pi.Position=pm.PositionID"
strSQL=strSQL & " left join DeptInfor di8 on pm.department=di8.DetpID and di8.DeptLevel='08' "
strSQL=strSQL & " left join DeptInfor di7 on (pm.department=di7.DetpID or di8.HLDeptID=di7.DetpID) and di7.DeptLevel='07' "
strSQL=strSQL & " left join DeptInfor di6 on (pm.department=di6.DetpID or di7.HLDeptID=di6.DetpID) and di6.DeptLevel='06' "
strSQL=strSQL & " left join DeptInfor di5 on (pm.department=di5.DetpID or di6.HLDeptID=di5.DetpID) and di5.DeptLevel='05'"
strSQL=strSQL & " left join PlantInfor pl on di5.PlantID=pl.PlantID"
strSQL=strSQL & " where  Country='CHN' and datediff(d,joindate,getdate())=1 and adaccount  like '%.%'"

rs.Open strSQL, Conn, 2, 3, 1

if err.number <> 0 then
     objLogFile.Write now() & "To Create the AD account failed : Read SF datase error" & err.description & vbCrLf
     strSubject = now() & " To Create the AD account failed : Read SF datase error"
     strContent = now() & " To Create the AD account failed : Read SF datase error:<BR>"
     strContent = strContent & err.Description
     MailToOwner strGetsmtpserver, intGetsmtpserverport, strGetsmtpaccountname, strGetsendemailaddress, strGetsenduserreplyemailaddress, strGetsendusername, strGetsendpassword, strGetErrorSendto, strGetErrorSendcc, strSubject, strContent 
     wscript.Quit
end if
strSubject = now() & "--Create Ad account Successful"
strContent = now() & "  Create Ad account for Users:<BR>" 

'rs.MoveFirst
objLogFile.Write now() & " -- Start to Create AD account "& vbCrLf 
While not rs.eof
     ynToCreateMailBox = true 
          strUserJobNo = rs("EmpID").value  
          strUserAccount = LCase(Trim(rs("adaccount").value)) 
          strDisplay = rs("EFNAME").value & " " & rs("ELNAME").value
          'strPlant = UCase(Trim(rs("PLANT")))
          strDesc= rs("PCNAME").value
          StrLan=  rs("Displayflag").value
           
          if ynToCreateMailBox = true then
                  objLogFile.Write now() & " -- Get Ad Account Information : " &  strUserAccount 
                   '" & strOrg & "," & strDCServer & "," & strUserJobNo & "," & strUserAccount & "," & strUserDepartment & "," & strEName & "," & strLastName & "," & strFirstName & "," & ynMailSizeDefaule & "," & intmDBOverHardQuotaLimit & "," & intmDBOverQuotaLimit & "," & intmDBStorageQuota & "," & strExchangeServer & "," & strMailStorge & "," & strMailGroup & "," & strScriptPath & vbCrLf
             if StrLan="857" Then   
               strCreateMailBoxResult = SetADPropertiesAndCreateMailBox(strDCServer,rs("adaccount").value,rs("ELNAME").value,rs("EFNAME").value,rs("PCNAME").value,rs("Tel").value,rs("Email").value,rs("Tel").value,rs("PLEName").value,rs("FENAME").value,rs("PMEname").value,rs("Shortname").value,rs("EmpID").value) 
             Elseif StrLan="858" Then   
              strCreateMailBoxResult = SetADPropertiesAndCreateMailBox(strDCServer,rs("adaccount").value,rs("ELNAME").value,rs("EFNAME").value,rs("PCNAME").value,rs("Tel").value,rs("Email").value,rs("Tel").value,rs("PLCName").value,rs("FCNAME").value,rs("PMCname").value,rs("Shortname").value,rs("EmpID").value) 
             End if                                                           
          else
                 objLogFile.Write now() & " -- Get Ad Account Information false :"  & strUserJobNo & "," & strUserAccount &  vbCrLf
                 strCreateMailBoxResult = "Get Ad Account Information false"
                 
          end if
      If strCreateMailBoxResult = "" Then
              objLogFile.Write now() & " -- To Create the AD Account of " & rs("EmpID") & "  OK ! "  & vbCrLf
              strSQLReply = "insert into [CreateAccountList] ([EmpNo],[AdAccount],[position],[Email],[CreateDate],[Status],[Modifydate],[UpdateSource],[Site])"
              strSQLReply = strSQLReply &" values('"& rs("EmpID").value &"','"& rs("adaccount").value &"','"& rs("jobLevel").value &"','','" & Date & "','A','','SF','"& rs("Shortname").value &"')"            
              objLogFile.Write now() & strSQLReply & vbCrLf
              rsReply.Open strSQLReply , Conn, 1, 1, 1
             
              if err.number <> 0 then 
                   objLogFile.Write  now() & "Update Database " & strUserJobNo & "Error" & err.number & err.description & vbCrLf
                   strSubject = now() & " -- ¦Update Database ("& strUserJobNo & ") Error"
                   strContent = now() & " -- ¦Update Database (" & strUserJobNo & ") Error:<BR>"
                   strContent = strContent & err.Description
                   MailToOwner strGetsmtpserver, intGetsmtpserverport, strGetsmtpaccountname, strGetsendemailaddress, strGetsenduserreplyemailaddress, strGetsendusername, strGetsendpassword, strGetErrorSendto, strGetErrorSendcc, strSubject, strContent
                 err.Clear 
               end if
                            if rsReply.state <> 0 then
                                  rsReply.close
                            End if
           Else
                            objLogFile.Write now() & "-- ?o Create the AD accouont of " & rs("EmpID") & " failed : " & strCreateMailBoxResult & vbCrLf
                            strSubject = now() & " -- ?o Create the AD accouont of " & rs("EmpID").value & " failed : " & strCreateMailBoxResult  
                            strContent = now() & " -- ?o Create the AD accouont of " & rs("EmpID").value & " failed : " & strCreateMailBoxResult& ":<BR>"
                            strContent = strContent & strCreateMailBoxResult
                            MailToOwner strGetsmtpserver, intGetsmtpserverport, strGetsmtpaccountname, strGetsendemailaddress, strGetsenduserreplyemailaddress, strGetsendusername, strGetsendpassword, strGetErrorSendto, strGetErrorSendcc, strSubject, strContent
                            err.Clear 
     end if 
        strContent = strContent & rs("EmpID").value & "    " & rs("adaccount").value & "    " & rs("FCNAME").value & "    " & rs("Shortname").value & "<BR>"
        rs.MoveNext
        err.Clear 
Wend
strContent = strContent & "  Successful"
MailToOwner strGetsmtpserver, intGetsmtpserverport, strGetsmtpaccountname, strGetsendemailaddress, strGetsenduserreplyemailaddress, strGetsendusername, strGetsendpassword, strGetErrorSendto, strGetErrorSendcc, strSubject, strContent
wscript.sleep(600000) 
'SendCreateOKAlert 

objLogFile.Write now() & " End" & vbCrLf
objLogFile.close

Set objFileSystem = nothing
Set Conn = nothing
Set ConnMD = nothing
Set rsMD = nothing
Set rs = nothing
Set rsReply = nothing

wscript.Quit

'Function SetADPropertiesAndCreateMailBox(strOrg,strDCServer,strUserJobNo,strUserAccount,strUserDepartment,strEName,strLastName,strFirstName,ynMailSizeDefaule,intmDBOverHardQuotaLimit,intmDBOverQuotaLimit,intmDBStorageQuota,strExchangeServer,strMailStorge,strMailGroup,strScriptPath,strPlant,strAG)

Function SetADPropertiesAndCreateMailBox(strDCServer,strUserAccount,strLastName,strFirstName,StrChineseName,StrTel,StrEmail,StrMobile,StrCompany,Strdepartment,StrTitle,StrShortname,StrEmpID)
On Error Resume Next

if StrShortname="SZCC" or StrShortname="SZGD" or StrShortname="SZSE" or StrShortname="SZSM" or StrShortname="SZSP" or StrShortname="SZST" then
   strCity="Suzhou"
elseif  StrShortname="BTSE" THEN
   strCity="Baotou"
elseif  StrShortname="CSAS" or StrShortname="CSSM" or StrShortname="CSTG" or StrShortname="CSTL" THEN
   strCity="Changshu"
elseif  StrShortname="ALHK" or StrShortname="HKCS"  THEN
   strCity="Hongkong"
elseif  StrShortname="JXSE"   THEN
   strCity="Jiaxing"
elseif  StrShortname="LYPT" or StrShortname="LYSP"   THEN
  strCity="Luoyang"
elseif  StrShortname="ALSH"   THEN
  strCity="Shanghai"
elseif  StrShortname="YCSE" or StrShortname="YCDF" or StrShortname="YCSM"  THEN
  strCity="Yancheng"
end if
'wscript.echo strCity
 Set objOU = GetObject("LDAP://" & strDCServer & "/" & "OU=Standard Users,OU=Users,OU=Resources,OU="& StrShortname &",OU="& strCity &",OU=China,OU=APAC,DC=CSITEST,DC=COM")
 Set objUser = objOU.Create("user", "cn="& strUserAccount &"")

     objuser.sn = StrChineseName 
     objuser.givenName = strLastName & " " & strFirstName
     objuser.displayName = strLastName & " " & strFirstName & " " & StrChineseName
     objuser.Description = StrChineseName
     objUser.put "physicalDeliveryOfficeName",StrShortname
     objUser.telephoneNumber=StrTel
     'objUser.EmailAddress=StrEmail
     'objUser.wWWHomePage=""

    objUser.Put "sAMAccountName", strUserAccount
    objuser.userPrincipalName = strUserAccount & "@CSITEST.COM"
    
    'objuser.streetAddress = ""
    'objuser.L = ""
    'objuser.st = ""
    'objuser.postalCode=""
    objuser.c = "CN"   

    objuser.pager = StrEmpID
    objuser.mobile = StrMobile
    'objuser.ipphone = ""
   
      
    objUser.Put "title", StrTitle
    objUser.Put "company", StrCompany
    objUser.put "department",Strdepartment
    'objUser.Put "manager", "CN=test test,CN=Users,DC=CSITEST,DC=COM"
        
    
    objuser.SetInfo
    objuser.AccountDisabled = FALSE
    objUser.SetInfo
   ' objuser.SetPassword Password'??????? objuser.SetInfo ??,?????????????ad??
    'Wscript.Echo "??:  " & FullName &vbTab& " ?OU: " & OUname & "   ?????!"
    Set ojbuser = Nothing


SetADPropertiesAndCreateMailBox= ""

If err.number <> 0  Then
    SetADPropertiesAndCreateMailBox= "Create " & strUserJobNoerr & " mailbox : "& err.description
End If

End Function

Public Function MailToOwner(strServer, strPort, strAccountName, strSender, strReplay, strUserName, strPWDg, strTo, strCC, strSubject, strContent)
    
    Set mailCDO = CreateObject("CDO.Message")
    Set cdoConf = CreateObject("CDO.Configuration")
    Set Flds = cdoConf.Fields
    
    On Error Resume Next
    
    Flds("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strServer
    Flds("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = strPort
    Flds("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPort
    Flds("http://schemas.microsoft.com/cdo/configuration/smtpaccountname") = strAccountName
    Flds("http://schemas.microsoft.com/cdo/configuration/sendemailaddress") = strSender
    Flds("http://schemas.microsoft.com/cdo/configuration/senduserreplyemailaddress") = strReplay
    Flds("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
    Flds("http://schemas.microsoft.com/cdo/configuration/sendusername") = strUserName
    Flds("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strPWD
    Flds.Update

    mailCDO.Configuration = cdoConf
    mailCDO.MimeFormatted = True
    mailCDO.MimeFormatted = True
    mailCDO.DSNOptions = cdoDSNSuccess
    mailCDO.To = strTo
    mailCDO.Subject = strSubject
    mailCDO.HTMLBody = strContent
    mailCDO.Send
    Set mailCDO = Nothing
    Set cdoConf = Nothing
    Set Flds = Nothing
    
    MailToOwner = true
    
    If err.number <> 0 Then
        MailToOwner = false
    End If

End Function


Function toGetParameters()
    objLogFile.Write now() & " -- Start get parameters" & vbCrLf
    On Error Resume Next
    Set fso= CreateObject("Scripting.FileSystemObject")
    Set objText = fso.OpenTextFile("MailBox.ini", 1, -1)
    arrIni = Split(objText.ReadAll, vbCrLF)
    
    For i = 0 To UBound(arrIni)
      
        s = Split(arrIni(i), ":=")
        If UBound(s) <> 0 Then
              
              If s(0) = "Org" Then
                strOrg = s(1)
              ElseIf s(0) = "DCServer" Then
                strDCServer = s(1)
              ElseIf s(0) = "ScriptPath" Then
                strScriptPath =  s(1)
          ElseIf s(0) = "smtpserver" Then
                strGetSmtpServer = s(1)
              ElseIf s(0) = "smtpserverport" Then
                intGetSmtpServerport = CInt(s(1))
              ElseIf s(0) = "smtpaccountname" Then
                strGetSmtpaccountname = s(1)
              ElseIf s(0) = "sendemailaddress" Then
                strGetSendEmailAddress = s(1)
              ElseIf s(0) = "senduserreplyemailaddress" Then
                strGetSendUserReplyeMailAddress = s(1)
              ElseIf s(0) = "smtpauthenticate" Then
                strGetsmtpauthenticate = s(1)
              ElseIf s(0) = "sendusername" Then
                strGetsendusername = s(1)
               ElseIf s(0) = "sendErrorto" Then
                strGetErrorSendto = s(1)
              ElseIf s(0) = "sendto" Then
                strGetSendto = s(1)		
              ElseIf s(0) = "sendcc" Then
                strGetErrorSendcc = s(1)	
              ElseIf s(0) = "sendpassword" Then
                strGetsendpassword = s(1)
              ElseIf s(0) = "oracleConnectionString" Then
                strOracleConnectionString = s(1) 
              ElseIf s(0) = "sqlConnectionString" Then
                strSqlConnectionString = s(1)
              End If
        End If
       
    Next
    'TextStream.close
    Set objText = Nothing
    Set fso = Nothing
    toGetParameters = true
    If err.number <> 0 Then 
        toGetParameters = false
    End if
End Function

Function MailContent(strEmpNo)
strMailContent ="Just test"
End Function