'---------SCRIPT INFO---------

	' Tested and works Win 2012 r2 (x64) 
	' Script name: Windows-benchmark.vbs
	' Run in CMD with admin privileges
	' Run with cscript to suppress dialogs: cscript.exe /nologo Windows-benchmark.vbs
	' Audit For Windows 2012 r2, Domain Controller or Member Server

	WScript.Echo vbNewLine & vbNewline
	WScript.Echo " #-------------------------------------------------#"
	WScript.Echo " |                                                 |"
	WScript.Echo " |         [ Audit Windows server 2012 r2 ]        |"
	WScript.Echo " |         [ J1Ejota - CIS Benchmark Compliance ]  |"
	WScript.Echo " |         [ Version 1.0.0 ]                       |"
	WScript.Echo " |                                                 |"
	WScript.Echo " |                                                 |"
	WScript.Echo " #-------------------------------------------------#"
	WScript.Echo vbNewLine & vbNewline

	WScript.Echo "     Running scan, please wait ...       "
	
'----------------------------------



'---------Global Variables---------

	strScriptVersion = "0.2"

	Const HKEY_CLASSES_ROOT   = &H80000000
	Const HKEY_CURRENT_USER   = &H80000001
	Const HKEY_LOCAL_MACHINE  = &H80000002
	Const HKEY_USERS          = &H80000003

    strComputer = "." ' Use . for current machine
	
	DQ = Chr(34) ' Insert " in string'

	DC = Null   ' Check if Machine is Domain Controller

	output = ""  ' Declared for getCommandOutput Function
	
    ' Shell Object from execution of cmd commands
	Set WshShell = Wscript.CreateObject("Wscript.Shell")

	' Object that contains Registry 
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
    
    ' Machine name
    StrMachineName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

    ' RSOP
	Set objWmi = GetObject("winmgmts:\\" & strComputer & "\root\rsop\computer")

    ' System information 
    Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem")

    ' Convert system dates
    Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")

    ' Check for DC, Member of DC, or Workgroup
    Set colItems = objWMIService.ExecQuery( "Select * from Win32_ComputerSystem", , 48 )
	For Each objItem in colItems
	    strComputerDomain = objItem.Domain
	    If objItem.PartOfDomain Then
	        Checkdc()
	    Else
	    	WScript.Echo "This machine isn't part of a Domain"
	        wscript.quit
	    End If
	Next
	Set colItems = Nothing


	SET objRoot = GETOBJECT("LDAP://RootDSE")
	SET objDomain = GETOBJECT("LDAP://" & objRoot.GET("defaultNamingContext"))
	SET objHash = CreateObject("Scripting.Dictionary") 
	objHash.Add "DOMAIN_PASSWORD_COMPLEX", &h1 
	objHash.Add "DOMAIN_PASSWORD_STORE_CLEARTEXT", &h16

'----------------------------------



'---------DC CHECK-----------------

FUNCTION Checkdc()

	StrSQL = "Select ADsPath From 'LDAP://" & strComputerDomain & "' Where Name = '" & StrMachineName & "'"

	Set ObjConn = CreateObject("ADODB.Connection")
	ObjConn.Provider = "ADsDSOObject"
	ObjConn.Open "Active Directory Provider"
	Set ObjRS = CreateObject("ADODB.Recordset")
	ObjRS.Open StrSQL, ObjConn
	If Not ObjRS.EOF Then
		ObjRS.MoveFirst
		While Not ObjRS.EOF
			Set ObjThisObject = GetObject(Trim(ObjRS.Fields("ADsPath").Value))
			If StrComp(Trim(ObjThisObject.Class), "COMPUTER", vbTextCompare) = 0 Then
				If Trim(ObjThisObject.primaryGroupID) = 516 Then
					DC = True
				Else
					DC = False
				End If
			End If
			ObjRS.MoveNext
		Wend
	End If

	ObjRS.Close:	Set ObjRS = Nothing
	ObjConn.Close:	Set ObjConn = Nothing
End FUNCTION

'----------------------------------



'---------Get TimeStamp-----------------

	Function LZ(ByVal Number)
	  If Number < 10 Then
	    LZ = "0" & CStr(Number)
	  Else
	    LZ = CStr(Number)
	  End If
	END Function

	Function TimeStamp
	  Dim CurrTime
	  CurrTime = Now()
	  TimeStamp = "_" & CStr(Year(CurrTime)) & _
	     LZ(Month(CurrTime)) & _
	     LZ(Day(CurrTime)) & _
	     LZ(Hour(CurrTime)) & _
	     LZ(Minute(CurrTime)) & _
	     LZ(Second(CurrTime))
	END Function

'----------------------------------


FUNCTION Int8ToSec(BYVAL objInt8)
    ' FUNCTION to convert Integer8 attributes from
    ' 64-bit numbers to seconds.
    DIM lngHigh, lngLow
    lngHigh = objInt8.HighPart
    ' Account for error in IADsLargeInteger property methods.
    lngLow = objInt8.LowPart
    IF lngLow < 0 THEN
        lngHigh = lngHigh + 1
    END IF
    Int8ToSec = -(lngHigh * (2 ^ 32) + lngLow) / (10000000)
END FUNCTION


'Read registry Value--------HKEY--------------PATH--------Key Name--
FUNCTION readRegistry(BYVAL HKEY, BYVAL strKeyPath, BYVAL ValueName)
		
	Const REG_SZ        = 1
	Const REG_EXPAND_SZ = 2
	Const REG_BINARY    = 3
	Const REG_DWORD     = 4
	Const REG_MULTI_SZ  = 7
	Const REG_QWORD     = 11
	
	'Create Key full path
	Clave = HKEY & "\" & strKeyPath & ":" & ValueName
	WScript.Echo Clave
	Call checknameXML(Clave)
	
	'Transform HKEY in a correct value for the commands
	if HKEY = "HKEY_CLASSES_ROOT" then
		HKEY = HKEY_CLASSES_ROOT
	ElseIf HKEY = "HKEY_CURRENT_USER" then
		HKEY = HKEY_CURRENT_USER
	ElseIf HKEY = "HKEY_LOCAL_MACHINE" then
		HKEY = HKEY_LOCAL_MACHINE
	Else 
		HKEY = HKEY_USERS
	End If
	
	'Check if path exits
	If oReg.EnumKey(HKEY, strKeyPath, arrSubKeys) = 0 Then

		oReg.EnumValues HKEY, strKeyPath, arrValueNames, arrValueTypes
		
		'Check if the path has keys'	
		If not varType(arrValueNames) = 1 Then
			
			'Flag to control that key exists'
			Flag = 1
			For I=0 To UBound(arrValueNames)
			
				If Lcase(arrValueNames(I)) = Lcase(ValueName) Then
					
					'Print value of the KEY'

					Select Case arrValueTypes(I)
						Case REG_SZ
							
							oReg.GetStringValue HKEY, strKeyPath, arrValueNames(I), strValue
							Call checkvalueXML(strValue)

						Case REG_EXPAND_SZ
							
							oReg.GetExpandedStringValue HKEY, strKeyPath, arrValueNames(I), strValue
							Call checkvalueXML(strValue)

							If vartype(strValue) = 0 Then
								strValue = "None"
								Call checkvalueXML(strValue)
							End If

						Case REG_BINARY
							
							oReg.GetBinaryValue HKEY, strKeyPath, arrValueNames(I), arrBytes
							strBytes = ""
							For Each uByte in arrBytes
							  strBytes = strBytes & Hex(uByte) & " "
							Next
							
							Call checkvalueXML(strBytes)
							
						Case REG_DWORD
							
							oReg.GetDWORDValue HKEY, strKeyPath, arrValueNames(I), strValue
							Call checkvalueXML(strValue)

						Case REG_QWORD
							
							oReg.GetqWORDValue HKEY, strKeyPath, arrValueNames(I), strValue
							Call checkvalueXML(strValue)
							
						Case REG_MULTI_SZ
							
							oReg.GetMultiStringValue HKEY, strKeyPath, arrValueNames(I), arrValues

							For Each strValue in arrValues
								Call checkvalueXML(strValue)
							Next

							If UBound(arrValues) = -1 Then
								strValue = "None"
								Call checkvalueXML(strValue)
							End If

					End Select


					'Key not found
					Flag = 0
				END IF
			Next

			' Check if Key has been found
			IF Flag = 1 Then
				strvalue = "Not Defined"
				Call checkvalueXML(strValue)
			END IF

		Else

			strvalue = "Not Defined"
			Call checkvalueXML(strValue)

		End If			

	Else

		strvalue = "Not Defined"
		Call checkvalueXML(strValue)

	End If
END FUNCTION


'Gather All Account SIDs on local machine
FUNCTION GetSIDs
    strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys

    'All SIDs are in \ProfileList
    For Each subkey In arrSubKeys
        Set objAccount = objWMIService.Get("Win32_SID.SID='" & subkey & "'")
        If NOT objAccount.ReferencedDomainName = StrMachineName Then
            If NOT objAccount.AccountName = "" Then
                if Instr(subkey, "S-1-5-21") Then
                    results = results & subkey & ","
                End if
            End If
		End If    
    Next

    Result = FilterSIDs(results)
    GetSIDs  = Result
END FUNCTION


'Filter only Established Profiles
FUNCTION FilterSIDs(SID)
    arrResults = Split(SID,",")

    For i = 0 to UBound(arrResults)
        if Len(arrResults(i)) > 10 then 
            strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" & arrResults(i)
            strValueName = "ProfileImagePath"
            oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
            set objFS = CreateObject("Scripting.FileSystemObject")
            Rslt = objFS.FolderExists(strValue)
            if Rslt = True then 
                varSIDs = varSIDS & arrResults(i) & ","
                MountResult = MountNTUser(strValue,arrResults(i))
            end if
        end if
    Next

    if right(varSIDs,1) = "," then varSIDs = left(varSIDs,Len(varSIDs) - 1)
    FilterSIDs = varSIDs
END FUNCTION


'Mount NTUser.dat file(s)
FUNCTION MountNTUser(path,SID)
    Set Exec = WshShell.Exec("%windir%\system32\reg load HKU\" & SID & " " & chr(34) & path & "\ntuser.dat" & chr(34))  
    strOutput = Exec.StdOut.ReadAll
    if instr(1,strOutput,"success") > 1 then
        MountNTUser = "Successful!"
    else
        MountNTUser = "Warning! Failed to Load Hive!"
    end if
END FUNCTION


'Capture command output
FUNCTION getCommandOutput(theCommand)

    Set objCmdExec = WshShell.exec(theCommand)

    Call checknameXML(arrAudit(count))
    
    Do Until objCmdExec.StdOut.AtEndOfStream
    	output = objCmdExec.StdOut.ReadLine()
    Loop
    ' Replace ASCII characters
    output = Replace(output, chr(162), chr(243))
	output = Replace(output, chr(161), chr(237))
	' Delete name of category and black spaces
    output = Trim(Replace(output, arrAudit(count),""))
    ' Delete jump line
    output = Replace(output, vbCr, "")
	output = Replace(output, vbLf, "")

    Call checkvalueXML(output)
    
    count = count + 1
END FUNCTION


'Add Value to ArrAudit '
FUNCTION addtoArray(value)
	ReDim Preserve arrAudit(UBound(arrAudit) + 1)
	arrAudit(UBound(arrAudit)) = value
END FUNCTION


' Check value of User Rigth Assignment
FUNCTION UserRight(regla)
	Set objItems = objWmi.ExecQuery("Select * from RSOP_UserPrivilegeRight Where UserRight=" & DQ & regla & DQ & " AND Precedence=1")

	Call checknameXML(regla)

	If not objItems.count = 0 Then
		For Each objItem In objItems
			If IsNull(objItem.AccountList) Then
				accountlist = "No one,"
			Else
				For Each strAccountList in objItem.AccountList
					accountlist = accountlist & strAccountList & ","
				Next
			End If
			arrList = Split(AccountList, ",")
		Next
		For i=0 To Ubound(arrList) - 1
			Call checkvalueXML(arrList(i))
		Next
	Else
		value = "Not Defined"
		Call checkvalueXML(value)
	End If
END FUNCTION


' Codify and Save XML
Function ParseAndSave(filePath, xmlDoc)
	set xmlWriter = CreateObject("MSXML2.MXXMLWriter")
	set xmlReader = CreateObject("MSXML2.SAXXMLReader")
	Set xmlStream = CreateObject("ADODB.STREAM")
	xmlStream.Open
	xmlStream.Charset = "UTF-8"

	xmlWriter.output = xmlStream
	xmlWriter.indent = True
	xmlWriter.encoding = "UTF-8"

	Set xmlReader.contentHandler = xmlWriter
	Set xmlReader.DTDHandler = xmlWriter
	Set xmlReader.errorHandler = xmlWriter
	xmlReader.putProperty "http://xml.org/sax/properties/lexical-handler", xmlWriter
	xmlReader.putProperty "http://xml.org/sax/properties/declaration-handler", xmlWriter

	xmlReader.parse xmlDoc
	xmlWriter.flush

	xmlStream.SaveToFile filePath, 2

	xmlStream.Close
	Set xmlStream = Nothing
	Set xmlWriter = Nothing
	Set xmlReader = Nothing
END Function
 


Function saveXML

	Set oFSO = CreateObject("Scripting.FileSystemObject")
	sScriptDir = oFSO.GetParentFolderName(WScript.ScriptFullName)

	Call ParseAndSave(sScriptDir & "\" & StrMachineName & TimeStamp & ".xml", xmlDoc)
END Function


'---------Determinate Language---------

	count = 0

	If GetLocale() = 1033 OR GetLocale() = 2057 Then

	 	arrAudit = Array()
	 	'Account Logon'
	 	addtoArray("Credential Validation")
	 	'Account Management'
	 	addtoArray("Application Group Management")
	 	addtoArray("Computer Account Management")
	 	addtoArray("Distribution Group Management")
	 	addtoArray("Other Account Management Events")
	 	addtoArray("Security Group Management")
	 	addtoArray("User Account Management")
	 	'Detailed Tracking'
	 	addtoArray("Process Creation")
	 	'DS Access'
	 	addtoArray("Directory Service Access")
	 	addtoArray("Directory Service Changes")
	 	'Logon/Logoff'
	 	addtoArray("Account Lockout")
	 	addtoArray("Logoff")
	 	addtoArray("Logon")
	 	addtoArray("Other Logon/Logoff Events")
	 	addtoArray("Special Logon")
	 	'Object Access'
	 	addtoArray("Removable Storage")
	 	'Policy Change'
	 	addtoArray("Audit Policy Change")
	 	addtoArray("Authentication Policy Change")
	 	'Privilege Use'
	 	addtoArray("Sensitive Privilege Use")
	 	'System'
	 	addtoArray("IPsec Driver")
	 	addtoArray("Other System Events")
	 	addtoArray("Security State Change")
	 	addtoArray("Security System Extension")
	 	addtoArray("System Integrity")

	ElseIf GetLocale() = 1034  OR GetLocale() = 3082 Then

	 	arrAudit = Array()
	 	'Account Logon'
	 	addtoArray("Validación de credenciales")
	 	'Account Management'	 	
	 	addtoArray("Administración de grupos de aplicaciones")
	 	addtoArray("Administración de cuentas de equipo")
	 	addtoArray("Administración de grupos de distribución")
	 	addtoArray("Otros eventos de administración de cuentas")
	 	addtoArray("Administración de grupos de seguridad")
	 	addtoArray("Administración de cuentas de usuario")
	 	'Detailed Tracking'
	 	addtoArray("Creación del proceso")
	 	'DS Access'
	 	addtoArray("Acceso del servicio de directorio")
	 	addtoArray("Cambios de servicio de directorio")
	 	'Logon/Logoff'
	 	addtoArray("Bloqueo de cuenta")
	 	addtoArray("Cerrar sesión")
	 	addtoArray("Inicio de sesión")
	 	addtoArray("Otros eventos de inicio y cierre de sesión")
		addtoArray("Inicio de sesión especial")
		'Object Access'
	 	addtoArray("Almacenamiento extraíble")
	 	'Policy Change'
	 	addtoArray("Cambio en la directiva de auditoría")
	 	addtoArray("Cambio de la directiva de autenticación")
	 	'Privilege Use'
	 	addtoArray("Uso de privilegio confidencial")
	 	'System'
	 	addtoArray("Controlador IPsec")
	 	addtoArray("Otros eventos de sistema")
	 	addtoArray("Cambio de estado de seguridad")
	 	addtoArray("Extensión del sistema de seguridad")
	 	addtoArray("Integridad del sistema") 	

	else
		
		WScript.Echo "Language error"
		wscript.quit

	End If 

'--------------------------------------



' ====================   XML   =============================

	objCheck = ""
	objFieldValue = ""
	objFieldValues = ""

	Set xmlDoc = CreateObject("Microsoft.XMLDOM")  
	  
	Set objAudit = _
	  xmlDoc.createElement("Audit")  
	xmlDoc.appendChild objAudit  


	'--------- INFO -----------------

		Set objInfo = _
		  xmlDoc.createElement("info")  
		objAudit.appendChild objInfo

		Function infovalue(value,name)
			Set objFieldValue = _
				xmlDoc.createElement(name)
			objFieldValue.Text = value
			objInfo.appendChild objFieldValue
		End Function

		Call infovalue(strScriptVersion, "scriptversion")
		Call infovalue(Now, "executiondate")
		Call infovalue(StrMachineName ,"machinename")
		Call infovalue(strComputerDomain, "domain")

		IF DC Then 
			Call infovalue("True", "DC")
		Else
			Call infovalue("False", "DC")
		End if

		For Each os in SystemSet
		
		    Call infovalue(os.Caption, "system")
		    Call infovalue(os.OperatingSystemSKU, "operatingsystemsku")
		    Call infovalue(os.OSProductSuite, "osproductsuite")
			Call infovalue(os.Version, "version")
		    Call infovalue(os.CodeSet, "codeset")
		    Call infovalue(os.OSLanguage, "language")
			Call infovalue(os.CountryCode, "countrycode")
		    dtmConvertedDate.Value = os.InstallDate
		    dtmInstallDate = dtmConvertedDate.GetVarDate 
			Call infovalue(dtmInstallDate, "installdate")
		Next

	'----------------------------------


	Set objRoot = _
	  xmlDoc.createElement("checks")  
	objAudit.appendChild objRoot  


	Function checkXML(ID)

		Set objCheck = _
		  xmlDoc.createElement("check")  
		objRoot.appendChild objCheck  
		 
		Set xmlns = xmlDoc.createAttribute("id")
		xmlns.text = ID
		objCheck.setAttributeNode xmlns

	END Function

	Function checknameXML(name)

		Set objName = _
			xmlDoc.createElement("name")
		objCheck.appendChild objName

		Set xmlns = xmlDoc.createAttribute("checkname")
		xmlns.text = name
		objName.setAttributeNode xmlns

		Set objFieldValue = _
			xmlDoc.createElement("values")
		objName.appendChild objFieldValue
	END Function

	Function checkvalueXML(valuexml)

		Set objFieldValues = _
			xmlDoc.createElement("value")
			objFieldValues.Text = valuexml
		objFieldValue.appendChild objFieldValues
	END Function


	Function checkvalueusersXML(SID)

		Set xmlns = xmlDoc.createAttribute("sid")
		xmlns.text = SID
		objFieldValues.setAttributeNode xmlns

		Set objAccount = objWMIService.Get("Win32_SID.SID='" & SID & "'")
		username = objAccount.AccountName

		Set xmlns = xmlDoc.createAttribute("username")
		xmlns.text = username
		objFieldValues.setAttributeNode xmlns
	END Function


	Set objIntro = _
	  xmlDoc.createProcessingInstruction _
	  ("xml","version='1.0'")
	xmlDoc.insertBefore _
	  objIntro,xmlDoc.childNodes(0)



	' ************ Users P.99 *******************************

		Dim objusers
		Function checkusers
			Set objusers = _
			  xmlDoc.createElement("users")
			objCheck.appendChild objusers
		End Function

		Function checkuserXML(user)

			Set objFieldValue = _
				xmlDoc.createElement("user")
			objusers.appendChild objFieldValue

			Set xmlns = xmlDoc.createAttribute("username")
			xmlns.text = user
			objFieldValue.setAttributeNode xmlns
		END Function


		Function checkuservalueXML(valuexml, checkname)

			Set objFieldValues = _
				xmlDoc.createElement(checkname)
				objFieldValues.Text = valuexml
			objFieldValue.appendChild objFieldValues

		END Function

	' *******************************************************


' ==========================================================



WScript.Echo vbNewline & " Domain to check benchmark: " & strComputerDomain


WScript.Echo " [*] Account Policies"

	WScript.Echo "     [-] Password Policy"
		Call checkXML("1.1.1")
		Call checknameXML("pwdHistoryLength")

		pwdHistoryLength = objDomain.GET("pwdHistoryLength")
		Call checkvalueXML(pwdHistoryLength)


		Call checkXML("1.1.2")
		Call checknameXML("maxPwdAge")

		maxPwdAge = int(Int8ToSec(objDomain.GET("maxPwdAge")) / 86400)
		Call checkvalueXML(maxPwdAge)


		Call checkXML("1.1.3")
		Call checknameXML("minPwdAge")

		minPwdAge = int(Int8ToSec(objDomain.GET("minPwdAge")) / 86400)
		Call checkvalueXML(minPwdAge)


		Call checkXML("1.1.4")
		Call checknameXML("minPwdLength")

		minPwdLength = objDomain.GET("minPwdLength")
		Call checkvalueXML(minPwdLength)


		Call checkXML("1.1.5")
		Call checknameXML("PwdProperties")

			If &h1 And objDomain.Get("PwdProperties") Then 
				Call checkvalueXML("Enabled")
			Else
				Call checkvalueXML("Disabled")
			End if 

		Call checkXML("1.1.6")
		Call checknameXML("PwdProperties")

			If &h16 And objDomain.Get("PwdProperties") Then 
				Call checkvalueXML("Enabled")
			Else
				Call checkvalueXML("Disabled")
			End if 


	WScript.Echo "     [-] Account Lockout Policy"

		Call checkXML("1.2.1")
		Call checknameXML("lockoutDuration")
		lockoutDuration = Int8ToSec(objDomain.GET("lockoutDuration")) / 60
		Call checkvalueXML(lockoutDuration)

		Call checkXML("1.2.2")
		Call checknameXML("lockoutThreshold")
		lockoutThreshold = objDomain.GET("lockoutThreshold")
		Call checkvalueXML(lockoutThreshold)		

		Call checkXML("1.2.3")
		Call checknameXML("lockoutObservationWindow")
		lockoutObservationWindow = Int8ToSec(objDomain.GET("lockoutObservationWindow")) / 60
		Call checkvalueXML(lockoutObservationWindow)



WScript.Echo " [*] Local Policies"


	WScript.Echo "     [-] User Rights Assignment"

		Call checkXML("2.2.1")
		Call UserRight("SeTrustedCredManAccessPrivilege")

		Call checkXML("2.2.2")
		Call UserRight("SeNetworkLogonRight")

		Call checkXML("2.2.3")
		Call UserRight("SeTcbPrivilege")

		If DC then
			Call checkXML("2.2.4")
			Call UserRight("SeMachineAccountPrivilege")
		End If

		Call checkXML("2.2.5")
		Call UserRight("SeIncreaseQuotaPrivilege")

		Call checkXML("2.2.6")
		Call UserRight("SeInteractiveLogonRight")

		Call checkXML("2.2.7")
		Call UserRight("SeRemoteInteractiveLogonRight")

		Call checkXML("2.2.8")
		Call UserRight("SeBackupPrivilege")

		Call checkXML("2.2.9")
		Call UserRight("SeSystemTimePrivilege")

		Call checkXML("2.2.10")
		Call UserRight("SeTimeZonePrivilege")

		Call checkXML("2.2.11")
		Call UserRight("SeCreatePagefilePrivilege")

		Call checkXML("2.2.12")
		Call UserRight("SeCreateTokenPrivilege")

		Call checkXML("2.2.13")
		Call UserRight("SeCreateGlobalPrivilege")

		Call checkXML("2.2.14")
		Call UserRight("SeCreatePermanentPrivilege")

		Call checkXML("2.2.15")
		Call UserRight("SeCreateSymbolicLinkPrivilege")

		Call checkXML("2.2.16")
		Call UserRight("SeDebugPrivilege")

		Call checkXML("2.2.17") 
		Call UserRight("SeDenyNetworkLogonRight")

		Call checkXML("2.2.18")
		Call UserRight("SeDenyBatchLogonRight")

		Call checkXML("2.2.19")
		Call UserRight("SeDenyServiceLogonRight")

		Call checkXML("2.2.20")
		Call UserRight("SeDenyInteractiveLogonRight")

		Call checkXML("2.2.21")
		Call UserRight("SeDenyRemoteInteractiveLogonRight")
		
		Call checkXML("2.2.22")
		Call UserRight("SeEnableDelegationPrivilege")
		
		Call checkXML("2.2.23")
		Call UserRight("SeRemoteShutdownPrivilege")
		
		Call checkXML("2.2.24")
		Call UserRight("SeAuditPrivilege")
		
		Call checkXML("2.2.25")
		Call UserRight("SeImpersonatePrivilege")
		
		Call checkXML("2.2.26")
		Call UserRight("SeIncreaseBasePriorityPrivilege")
		
		Call checkXML("2.2.27")
		Call UserRight("SeLoadDriverPrivilege")
		
		Call checkXML("2.2.28")
		Call UserRight("SeLockMemoryPrivilege")
		
		If DC then
			Call checkXML("2.2.29")
			Call UserRight("SeBatchLogonRight")
		End If

		Call checkXML("2.2.30")
		Call UserRight("SeSecurityPrivilege")
		
		Call checkXML("2.2.31")
		Call UserRight("SeRelabelPrivilege")
		
		Call checkXML("2.2.32")
		Call UserRight("SeSystemEnvironmentPrivilege")
		
		Call checkXML("2.2.33")
		Call UserRight("SeManageVolumePrivilege")
		
		Call checkXML("2.2.34")
		Call UserRight("SeProfileSingleProcessPrivilege")
		
		Call checkXML("2.2.35")
		Call UserRight("SeSystemProfilePrivilege")
		
		Call checkXML("2.2.36")
		Call UserRight("SeAssignPrimaryTokenPrivilege")
		
		Call checkXML("2.2.37")
		Call UserRight("SeRestorePrivilege")
		
		Call checkXML("2.2.38")
		Call UserRight("SeShutdownPrivilege")

		If DC then
			Call checkXML("2.2.39")
			Call UserRight("SeSyncAgentPrivilege")
		End If

		Call checkXML("2.2.40")
		Call UserRight("SeTakeOwnershipPrivilege")
		

	WScript.Echo "     [+] Security Options"

		WScript.Echo "        [-] Accounts"

			If DC Then
				Set colAccounts = objWMIService.ExecQuery("Select * From Win32_UserAccount")
			else
				Set colAccounts = objWMIService.ExecQuery("Select * From Win32_UserAccount Where domain = '" & StrMachineName & "'")
			End If
				
			For Each objAccount In colAccounts

				If (objAccount.Name = "Administrador" OR objAccount.Name = "Administrator") Then

					Call checkXML("2.3.1.1.1")
					Call checknameXML("Administrator (name) Disabled")
				 		
					If objAccount.Disabled = True Then
				 		Call checkvalueXML("True")
				 	else
				 		Call checkvalueXML("False")
				 	End if
						
				End If

				If Instr(objaccount.SID, "-500") Then

					Call checkXML("2.3.1.1.2")
					Call checknameXML("Administrator (SID) Disabled")

					If objAccount.Disabled = True Then
				 		Call checkvalueXML("True")
				 	else
				 		Call checkvalueXML("False")
				 	End if
			 				
				End If
			Next

			Call checkXML("2.3.1.2")
			Call readRegistry("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "NoConnectedUser")


			For Each objAccount In colAccounts

				If (objAccount.Name = "Invitado" OR objAccount.Name = "Guest") Then

					Call checkXML("2.3.1.3.1")
					Call checknameXML("Guest (name) Disabled")
				 		
					If objAccount.Disabled = True Then
				 		Call checkvalueXML("True")
				 	else
				 		Call checkvalueXML("False")
				 	End if
						
				End If

				If Instr(objaccount.SID, "-501") Then

					Call checkXML("2.3.1.3.2")
					Call checknameXML("Guest (SID) Disabled")

					If objAccount.Disabled = True Then
				 		Call checkvalueXML("True")
				 	else
				 		Call checkvalueXML("False")
				 	End if
			 				
				End If
			Next

			Call checkXML("2.3.1.4")
			Call readRegistry("HKEY_LOCAL_MACHINE", "SYSTEM\CurrentControlSet\Control\Lsa", "LimitBlankPasswordUse")

			Call checkXML("2.3.1.5")
			Call checknameXML("Administrator Renamed")
			
			For Each objAccount In colAccounts
				If Instr(objaccount.SID, "-500") Then
					If (objAccount.Name = "Administrador" OR objAccount.Name = "Administrator") Then
					 	Call checkvalueXML("False")
					else
					 	Call checkvalueXML("True")
					End If
				End if
			Next

			Call checkXML("2.3.1.6")
			Call checknameXML("Guest Renamed")
			
			For Each objAccount In colAccounts
				If Instr(objaccount.SID, "-501") Then
					If (objAccount.Name = "Invitado" OR objAccount.Name = "Guest") Then
					 	Call checkvalueXML("False")
					else
					 	Call checkvalueXML("True")
					End If
				End if
			Next


		WScript.Echo "        [-] Audit"
			Call checkXML("2.3.2.1")
			Call readRegistry("HKEY_LOCAL_MACHINE", "SYSTEM\CurrentControlSet\Control\Lsa", "SCENoApplyLegacyAuditPolicy")

		WScript.Echo "        [-] Audit"
			Call checkXML("9.9.9.9.9.jm")
			Call readRegistry("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Cryptography\CertificateTemplateCache", "Timestamp")


WScript.Echo " [*] Advanced Audit Policy Configuration"


	WScript.Echo "     [-] Account Logon"

		Call checkXML("17.1.1")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)


	WScript.Echo "     [-] Account Management"

		Call checkXML("17.2.1")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)

		Call checkXML("17.2.2")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)

		If DC Then
			Call checkXML("17.2.3")
			Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)
		End if

		Call checkXML("17.2.4")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)

		Call checkXML("17.2.5")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)

		Call checkXML("17.2.6")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)


	WScript.Echo "     [-] Detailed Tracking"

		Call checkXML("17.3.1")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)

	If DC Then
		WScript.Echo "     [-] DS Access"

			Call checkXML("17.4.1")
			Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)

			Call checkXML("17.4.2")
			Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)
	End if

	WScript.Echo "     [-] Logon/Logoff"

		Call checkXML("17.5.1")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)

		Call checkXML("17.5.2")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)

		Call checkXML("17.5.3")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)

		Call checkXML("17.5.4")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)

		Call checkXML("17.5.5")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)


	WScript.Echo "     [-] Object Access"

		Call checkXML("17.6.1")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)


	WScript.Echo "     [-] Policy Change"

		Call checkXML("17.7.1")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)

		Call checkXML("17.7.2")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)


	WScript.Echo "     [-] Privilege Use"

		Call checkXML("17.8.1")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)


	WScript.Echo "     [-] Logon/Logoff"

		Call checkXML("17.9.1")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)

		Call checkXML("17.9.2")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)

		Call checkXML("17.9.3")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)

		Call checkXML("17.9.4")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)

		Call checkXML("17.9.5")
		Call getCommandOutput("auditpol.exe /get /subcategory:" & DQ & arrAudit(count) & DQ)


WScript.Echo " [*] Administrative Templates (User)"

	UserSIDs = Split(GetSIDs,",")

	WScript.Echo "     [-] Control Panel"

			Call checkXML("19.1.3.1")

		    For Each SID In UserSIDs 
		        Call readRegistry("HKEY_USERS", SID & "\SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop", "ScreenSaveActive")
		        Call checkvalueusersXML(SID)
		    Next

			Call checkXML("19.1.3.2")

		    For Each SID In UserSIDs 
		        Call readRegistry("HKEY_USERS", SID & "\SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop", "SCRNSAVE.EXE")
		        Call checkvalueusersXML(SID)
		    Next

			Call checkXML("19.1.3.3")

		    For Each SID In UserSIDs 
		        Call readRegistry("HKEY_USERS", SID & "\SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop", "ScreenSaverIsSecure")
		        Call checkvalueusersXML(SID)
		    Next

			Call checkXML("19.1.3.4")

		    For Each SID In UserSIDs 
		        Call readRegistry("HKEY_USERS", SID & "\SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop", "ScreenSaveTimeOut")
		        Call checkvalueusersXML(SID)
		    Next

			Call checkXML("19.1.3.101")

		    For Each SID In UserSIDs 
		        Call readRegistry("HKEY_USERS", SID & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispScrSavPage")
		        Call checkvalueusersXML(SID)
		    Next		    


	WScript.Echo "     [-] Start Menu and Taskbar"


			Call checkXML("19.5.1.1")

		    For Each SID In UserSIDs 
		        Call readRegistry("HKEY_USERS", SID & "\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\PushNotifications", "NoToastApplicationNotification")
		        Call checkvalueusersXML(SID)
		    Next


	WScript.Echo "     [-] System"

		
			Call checkXML("19.6.5.1.1")

		    For Each SID In UserSIDs 
		        Call readRegistry("HKEY_USERS", SID & "\SOFTWARE\Policies\Microsoft\Assistance\Client\1.0", "NoImplicitFeedback")
		        Call checkvalueusersXML(SID)
		    Next


	WScript.Echo "     [+] Windows Components"

		WScript.Echo "        [-] Attachment Manager"

			Call checkXML("19.7.4.1")

		    For Each SID In UserSIDs 
		        Call readRegistry("HKEY_USERS", SID & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Attachments", "SaveZoneInformation")
		        Call checkvalueusersXML(SID)
		    Next

			Call checkXML("19.7.4.2")

		    For Each SID In UserSIDs 
		        Call readRegistry("HKEY_USERS", SID & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Attachments", "ScanWithAntiVirus")
		        Call checkvalueusersXML(SID)
		    Next

		WScript.Echo "        [-] Network Sharing"

			Call checkXML("19.7.26.1")

		    For Each SID In UserSIDs 
		        Call readRegistry("HKEY_USERS", SID & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoInplaceSharing")
		        Call checkvalueusersXML(SID)
		    Next

		WScript.Echo "        [-] Windows Installer"

			Call checkXML("19.7.39.1")

		    For Each SID In UserSIDs 
		        Call readRegistry("HKEY_USERS", SID & "\SOFTWARE\Policies\Microsoft\Windows\Installer", "AlwaysInstallElevated")
		        Call checkvalueusersXML(SID)
		    Next

		WScript.Echo "        [-] Windows Media Player"

			Call checkXML("19.7.43.2.1")

		    For Each SID In UserSIDs 
		        Call readRegistry("HKEY_USERS", SID & "\SOFTWARE\Policies\Microsoft\WindowsMediaPlayer", "PreventCodecDownload")
		        Call checkvalueusersXML(SID)
		    Next



WScript.Echo " [*] Local Users Audit"


	Call checkXML("99.1")
	Call checkusers


	Const ADS_UF_ACCOUNTDISABLE = &H0002
	Const ADS_UF_LOCKOUT = &H0010
	Const ADS_UF_PASSWD_NOTREQD = &H0020
	Const ADS_UF_PASSWD_CANT_CHANGE = &H0040
	Const ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = &H0080
	Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000
	Const ADS_UF_SMARTCARD_REQUIRED = &H40000
	Const ADS_UF_PASSWORD_EXPIRED = &H800000

	If DC Then
		Set colAccounts = objWMIService.ExecQuery("Select * From Win32_UserAccount Where NOT Name = 'krbtgt'")
	Else
		Set colAccounts = objWMIService.ExecQuery("Select * From Win32_UserAccount Where LocalAccount = True")
	End If

		WScript.Echo "     [-] Obtain Local Users"


	For Each objAccount in colAccounts
	    strUser = objAccount.Name
	    checkuserXML(strUser)
	    accountsinfo(strUser)
	Next

		WScript.Echo "     [-] Check Local Users"

	Function accountsinfo(strUser)

			Set objUser = GetObject("WinNT://" & strComputer & "/" & strUser & ",user")
			flag = objUser.Get("UserFlags")

				If flag AND ADS_UF_ACCOUNTDISABLE Then
				    Call checkuservalueXML("True", "accountdisable")
				Else
				    Call checkuservalueXML("False", "accountdisable")
				End If

				If flag AND ADS_UF_PASSWD_NOTREQD Then
				    Call checkuservalueXML("True", "notrequirepwd")
				Else
				    Call checkuservalueXML("False", "notrequirepwd")
				End If
				 
				If flag AND ADS_PASSWORD_CANT_CHANGE Then
				    Call checkuservalueXML("True", "cantchangepwd")
				Else
				    Call checkuservalueXML("False", "cantchangepwd")
				End If
				 
				If flag AND ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED Then
				    Call checkuservalueXML("True", "pwdencrypted")
				Else
				    Call checkuservalueXML("False", "pwdencrypted")
				End If
				 
				If flag AND ADS_UF_DONT_EXPIRE_PASSWD Then
				    Call checkuservalueXML("True", "pwdexpire")
				Else
				    Call checkuservalueXML("False", "pwdexpire")
				End If
				 
				If flag AND ADS_UF_SMARTCARD_REQUIRED Then
				    Call checkuservalueXML("True", "smartcard")
				Else
				    Call checkuservalueXML("False", "smartcard")
				End If
				 
				If flag AND ADS_UF_PASSWORD_EXPIRED Then
				    Call checkuservalueXML("True", "pwdexpired")
				Else
				    Call checkuservalueXML("False", "pwdexpired")
				End If


			strminPasswordAge = objUser.minPasswordAge / 86400
			Call checkuservalueXML(strminPasswordAge, "minpasswordage")

			strmaxPasswordAge = objUser.maxPasswordAge / 86400
			Call checkuservalueXML(strmaxPasswordAge, "maxpasswordage")

			Call checkuservalueXML(objUser.minPasswordLength, "minpasswordlength")

			Call checkuservalueXML(objUser.passwordHistoryLength, "passwordhistorylength")

			Call checkuservalueXML(objUser.badPasswordAttempts, "badpasswordattempts")

			strautoUnlockInterval = objUser.autoUnlockInterval / 60
			Call checkuservalueXML(strautoUnlockInterval, "autounlockinterval")

			strLockoutObservationInterval = objUser.LockoutObservationInterval / 60
			Call checkuservalueXML(strLockoutObservationInterval, "lockoutobservationinterval")

		    Set colItems = objWMIService.ExecQuery _
		    ("Select * from Win32_NetworkLoginProfile Where Caption = '" & strUser & "'")

			For Each objItem in colItems

			    dtmConvertedDate.Value = objItem.LastLogon
				dtmlastlogon = dtmConvertedDate.GetVarDate
			    Call checkuservalueXML(dtmlastlogon, "lastlogon")
			Next

			If colItems.count = 0 Then
				dtmlastlogon = "Not exists"
				Call checkuservalueXML(dtmlastlogon, "lastlogon")
			End If

	End Function


	Call checkXML("99.2")

    strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys

    'All SIDs are in \ProfileList
    For Each subkey In arrSubKeys
        Set objAccount = objWMIService.Get("Win32_SID.SID='" & subkey & "'")
        If NOT objAccount.AccountName = "" Then
            if Instr(subkey, "S-1-5-21") Then
                results = results & subkey & ","
            End if
        End If
    Next

    arrUsers = Split(results,",")

    For i = 0 to UBound(arrUsers) - 1

    	Set objAccount = objWMIService.Get("Win32_SID.SID='" & arrUsers(i) & "'")

    	Call checknameXML(objAccount.AccountName)
		Call checkvalueXML(arrUsers(i))

    	If NOT objAccount.ReferencedDomainName = StrMachineName Then
    		Set xmlns = xmlDoc.createAttribute("domain")
			xmlns.text = "True"
			objFieldValues.setAttributeNode xmlns
    	else
    		Set xmlns = xmlDoc.createAttribute("domain")
			xmlns.text = "False"
			objFieldValues.setAttributeNode xmlns
        End If
    Next



Call saveXML



WScript.Echo  vbNewLine & " #==================== END AUDIT =======================# " &vbNewLine