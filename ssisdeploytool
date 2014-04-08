title = "SSIS Deploy Tool"         'assing Name of this tool
sqlinstance = getSqlInstance()     'get default sql server instance
currentserver = getCurrentServer() 'get current computer name
'----------------------------
If (InStr(sqlinstance,";") = 0 and len(sqlinstance) > 0) Then     'auto-detected single SQL Server instance
  sqlinstance = InputBox(title &" has detected one SQL Server instance, please verify (should be in SERVER\INSTANCE format):", title, currentserver &"\"& sqlinstance)
ElseIf (InStr(sqlinstance,";") > 0 and len(sqlinstance) > 1) Then 'auto-detected multiple SQL Server instances
  sqlinstance = InputBox(title &" has detected multiple SQL Server instances:"& vbLf &" - "& Replace(sqlinstance,";",vbLf &" - ") & vbLf & _
                         "Please select and input one in SERVER\INSTANCE format:", title, currentserver &"\")
Else                                                              'auto-detect failed to locate SQL Server instance
  sqlinstance = InputBox(title &" hasn't detected SQL Server instance, please input in SERVER\INSTANCE format:", title, currentserver &"\")
End If
'----------------------------
If (len(getSqlVersion(sqlinstance))>0) Then                       'verify input
  Main
Else 
  MsgBox sqlinstance & " is not a valid SQL Server Instance."& vbLf & title &" will exit now.",,title
End If
'++++++++++++++++++++++++++++
Function getCurrentServer()
  Set wshNetwork = WScript.CreateObject( "WScript.Network" )
  getCurrentServer = wshNetwork.ComputerName 
  Set wshNetwork = Nothing
End Function
'++++++++++++++++++++++++++++
Function getSqlInstance()
  on error resume next
  Set oCtx = CreateObject("WbemScripting.SWbemNamedValueSet") 
  oCtx.Add "__ProviderArchitecture", 64 
  oCtx.Add "__RequiredArchitecture", True 
  Set oServices = CreateObject("Wbemscripting.SWbemLocator").ConnectServer("","root\default","","",,,,oCtx) 
  Set Inparams = oServices.Get("StdRegProv").Methods_("EnumKey").Inparameters 
  Inparams.Hdefkey = &h80000002' HKLM 
  Inparams.Ssubkeyname = "SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL" 
  Set Outparams = oServices.Get("StdRegProv").ExecMethod_("EnumValues", Inparams,,oCtx) 
  Set oCtx = Nothing 
  set oServices = Nothing
  If IsArray(Outparams.snames) Then
    for each ins in Outparams.snames
      sInstances = sInstances & ins &";"
    next
    do while right(sInstances,1) = ";"
      sInstances = left(sInstances, len(sInstances)-1) 'remove last ";"
    loop
  Else
    sInstances = ""
  End If
  getSqlInstance = sInstances
End Function
'++++++++++++++++++++++++++++
Function searchFolders(sParentFolder,sFileNamePart)
  on error resume next
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFolder = oFSO.GetFolder(sParentFolder)
  For Each oFile in oFolder.Files
    sToSearch = oFolder.Path &"\"& oFile.Name
    if InStr(len(sToSearch)-len(sFileNamePart),UCase(sToSearch),UCase(sFileNamePart)) then 
      sResult = sResult & sToSearch & ";"
    end if 
  Next
  For Each sSubFolder in oFolder.SubFolders  
    sResult = sResult & searchFolders(sSubFolder,sFileNamePart)
  Next
  do while right(sResult,1) = ";"
    sResult = left(sResult, len(sResult)-1) 'remove last ";"
  loop
  searchFolders = sResult
  set oFSO = Nothing
End Function
'++++++++++++++++++++++++++++
Function execDTUtil(sDTUtilCommand)
  Set oWsh = CreateObject("WScript.Shell")
  strErrorCode = oWsh.Run(sDTUtilCommand,0,True)
  Set oWsh = Nothing
  Select Case strErrorCode
    Case 0 : execDTUtil = 0'"DTUTIL.EXE: The utility executed successfully."
    Case 1 : execDTUtil = 1'"DTUTIL.EXE: The utility failed." 
    Case 4 : execDTUtil = 4'"DTUTIL.EXE: The utility cannot locate the requested package." 
    Case 5 : execDTUtil = 5'"DTUTIL.EXE: The utility cannot load the requested package." 
    Case 6 : execDTUtil = 6'"DTUTIL.EXE: The utility cannot resolve the command line because it contains either syntactic or semantic errors."
    Case Else : execDTUtil = -1'"DTUTIL.EXE: Failed for no reason."
  End Select
End Function
'++++++++++++++++++++++++++++
Function getSqlPaths()
  dim aSQLPaths(), i : i = 0
  Set oWsh = CreateObject("WScript.Shell")
  arrPaths = Split(oWsh.Environment("SYSTEM")("PATH"), ";")
  For Each strPath In arrPaths
    if InStr(strPath, "SQL") then 
      ReDim Preserve aSQLPaths(i)
      aSQLPaths(i) = strPath
      i = i + 1 
    end if
  Next
  getSqlPaths = aSQLPaths
  Set oWsh = Nothing
End Function
'++++++++++++++++++++++++++++
Function getSqlVersion(sSqlServerInstance)
  on error resume next
  Set oCn = CreateObject("ADODB.Connection")
  oCn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=master;Data Source="& sSqlServerInstance 
  set aResult = oCn.execute("select ServerProperty('ProductVersion')")
  getSqlVersion = aResult(0).value
  oCn.Close
  set oCn = Nothing
  set aResult = Nothing
End Function
'============================
Sub Main
  on error resume next
  '====call preparation functions and catch errors====<
  sSqlVersion = getSqlVersion(sqlinstance)
  If Err.Number <> 0 Then
    MsgBox "Error getting SQL Server Instance version: "& vbLf & Err.Description,vbCritical,title
    Err.Clear
    Exit Sub
  End If
  sSqlPaths = getSqlPaths()
  If Err.Number <> 0 Then
    MsgBox "Error getting SQL Server installation paths: "& vbLf & Err.Description,vbCritical,title
    Err.Clear
    Exit Sub
  End If
  for each sSqlPath in getSqlPaths()
    dtutilarr = Split(searchFolders(sSqlPath,"dtutil.exe"),";")
    for each path in dtutilarr
      if InStr(path,Split(sSqlVersion,".")(0)*10) then dtutilpath = path
    next
    'sqlcmdarr = Split(searchFolders(sSqlPath,"sqlcmd.exe"),";")
    'for each path in sqlcmdarr
      'if Instr(path,Split(sSqlVersion,".")(0)*10) then sqlcmdpath = path
    'next
  next
  dtsxarr = Split(searchFolders(".","dtsx"), ";")
  'sqlarr = Split(searchFolders(".","sql"), ";")
  'xmlarr = Split(searchFolders(".","xml"),";")
  '>====end of preparation functions calling====
  '====report findings====<
  sMsg = sMsg & title &" discovered following: "& vbLf & vbLf
  sMsg = sMsg &"Version of provided SQL Server Instance: "& vbLf & vbTab & sSqlVersion & vbLf
  sMsg = sMsg & "DTUtil.exe:"& vbLf & vbTab & dtutilpath & vbLf
  'sMsg = sMsg & "SqlCmd.exe:"& vbLf & vbTab & sqlcmdpath & vbLf
  sMsg = sMsg &"Dtsx files:"& vbLf
  for each path in dtsxarr
    sMsg = sMsg & vbTab & path & vbLf
  next
  'sMsg = sMsg &"Xml files:"& vbLf
  'for each path in xmlarr
    'sMsg = sMsg & vbTab & path & vbLf
  'next
  'sMsg = sMsg &"Sql files:"& vbLf
  'for each path in sqlarr
    'sMsg = sMsg & vbTab & path & vbLf
  'next
  sMsg = sMsg & vbLf & "Are you sure that you want to install the above?"
  sReportReply = MsgBox(sMsg,vbYesNo,title)
  '>====end of findings report====
  'finalDecision = ""
  If sReportReply = vbNo then 'finalDecision = 0 'User replied "No"
    MsgBox "Installation cancelled. Will exit now.",,title
    exit sub
  Else
    If (IsArray(dtsxarr) And Ubound(dtsxarr)=-1) Then
      MsgBox "No *.dtsx to install. Will exit now.",,title
      exit sub
    End If
  End If
  If len(dtutilpath)=0 Then 
    MsgBox "No DTUtil.EXE found. Will exit now.",,title
    exit sub
  End If
  '====begin installing packages====<
  sMsg=""
  For each packagepath in dtsxarr
    filename = Split(packagepath,"\")
    package = Replace(Replace(filename(ubound(filename)),".dtsx",""),"."," ")
    check  =""""& dtutilpath & """ /SQL """& package &""" /EXISTS /SourceServer "& sqlinstance 
    install=""""& dtutilpath & """ /FILE """& packagepath &""" /COPY SQL;"""& package &""" /DestServer "& sqlinstance
    remove =""""& dtutilpath & """ /SQL """& package &""" /DELETE /SourceS "& sqlinstance 
    'msgbox "check:"& check & vblf &"install:"& check & install &"remove:"& remove 'debug
    dtucheck = execDTUtil(check)    '====check if installed====
    If dtucheck = 0 then 
      dturemove = execDTUtil(remove)'====remove if installed====
    End IF
    dtuinstall = execDTUtil(install)'====install new====
    dtucheck = execDTUtil(check)    '====check after installed====
    If dtucheck = 0 then 
      sMsg = sMsg & "SUCCESS!"& vbLf &"Package "& packagepath & vbLf &" was installed." & vbLf
    Else 
      sMsg = sMsg & "FAIL!"& vbLf &"Package " & packagepath & vbLf &" was not installed." & vbLf & vbLf
    End If
  next
  MsgBox sMsg,,title
  '>====end of installing packages====
End Sub
