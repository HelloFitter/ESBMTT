'该脚本检查XML格式是否正确
Set xdoc = CreateObject("MSXML2.DOMDocument")
esbbasepath = "Z:\WorkSpace\SmartESB\configs\"
Set attrDict = CreateObject("Scripting.Dictionary")
Set wshshell = CreateObject("wscript.shell")
'---------服务治理库地址---------------
serviceurl = "ESBSG/esbsg@159.1.65.153:1521/esbtest"
scriptpath = "SQLScript\"
'=======================
'加载文件中的交易
'=======================
Function LoadTrans()	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set traninfo = fso.OpenTextFile("traninfo.txt",1,False)
	tmpstr = ""
	Do While traninfo.AtEndOfStream <> True
		line = traninfo.ReadLine()
		If Trim(line) <> "" then 
			tmpstr = tmpstr & Replace(line," ",",") & ";"
		End If
	Loop
	LoadTrans = tmpstr
	Set fso = Nothing

End Function

'=======================
'查询对应系统的简称
'=======================
Function queryName(from)
	result = ""
	
	If attrDict.Exists(from) Then
		result = attrDict.Item(from)
	Else
		cmdstr = "sqlplus "& serviceurl &" @" & scriptpath & "GetName.sql" & " " & from
		Set exeRs = wshshell.exec(cmdstr)
		retmsg = exeRs.StdOut.ReadAll()
		If InStr(retmsg,"no rows") = 0 Then
			rets = Split(retmsg,"--------------------------------------------------------------------------------")
			tttt = Split(rets(1),"Disconnected")
			result = Replace(tttt(0),Chr(10),"")
			result = Replace(result,Chr(13),"")
		Else
			MsgBox "未查询到[" & from & "]对应的简称,注意查看!!!"
		End If
		attrDict.Add from,result
	End If
	queryName = result
End Function

Function getOldFile(path,code,filetype)
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set fileDict = CreateObject("Scripting.Dictionary")
	Set oFolder = fso.GetFolder(path)
	best=0
	better=0
	oldFilePath = ""
	For Each otmp In oFolder.Files
		'MsgBox otmp.DateLastModified
		tmp = DateDiff("s",otmp.DateLastModified,Now())

		On Error Resume Next
		fileDict.Add tmp,otmp.Path
		
		'确定是否是需要的文件[根据文件名开头来区分]
		If InStr(otmp.name,filetype) = 1 Then
			'日期小并且包含11为服务码
			If best < tmp And InStr(otmp.name,Left(code,11))  Then
				best = tmp
			Else
				better = tmp
			End If	
		End If

			
	Next

	If best <> 0 Then 
		getOldFile = fileDict.Item(best)
	Else 
		getOldFile = fileDict.Item(better)
	End If 
	Set fileDict = Nothing
	Set fso = Nothing
End Function

Sub Main()
	tranText = LoadTrans()
	If tranText = "" Then
		MsgBox "请检查[ traninfo.txt ]是否已经粘贴交易！",0,"星阳提醒"
		Exit Sub
	Else
		trans = Split(tranText,";")
		For i= 0 To UBound(trans)
			If trans(i) <> "" Then
				traninfo =  Split(trans(i),",")
				code = Trim(traninfo(0))
				from = Trim(traninfo(1))
				dest = Trim(traninfo(2))
				
				enname = queryName(from)

				outservicepath = esbbasepath & "out_conf\metadata\"& dest &"\service_"& code &"_system_" & dest & ".xml"
				xdoc.load(outservicepath)

				'tmp = getOldFile( esbbasepath & "out_conf\metadata\"& dest & "\",code,"service" )
				'MsgBox tmp
				If xdoc.documentElement Is Nothing Then 
					MsgBox code & "=>OUT组包存在问题"
					Exit Sub
				End If

				outchannelepath = esbbasepath & "out_conf\metadata\"& dest &"\channel_" & dest & "_service_" & code & ".xml"
				xdoc.load(outchannelepath)
				If xdoc.documentElement Is Nothing Then 
					MsgBox code & "=>OUT拆包存在问题"
					Exit Sub
				End If

				inservicepath = esbbasepath & "in_conf\metadata\"& enname &"\service_"& code &"_system_" & enname & ".xml"
				xdoc.load(inservicepath)
				If xdoc.documentElement Is Nothing Then 
					MsgBox code & "=>IN组包存在问题"
					Exit Sub
				End If

				inchannelepath = esbbasepath & "in_conf\metadata\"& enname &"\channel_" & enname & "_service_" & code & ".xml"
				xdoc.load(inchannelepath)
				If xdoc.documentElement Is Nothing Then 
					MsgBox code & "=>IN拆包存在问题"
					Exit Sub
				End If

			End If 
		Next

		inserviceIdentify = esbbasepath & "in_conf\serviceIdentify.xml"
		xdoc.load(inserviceIdentify)
		If xdoc.documentElement Is Nothing Then 
			MsgBox "serviceIdentify.xml存在问题"
			Exit Sub
		End If

		insystemIdentify = esbbasepath & "in_conf\systemIdentify.xml"
		xdoc.load(insystemIdentify)
		If xdoc.documentElement Is Nothing Then 
			MsgBox "IN-systemIdentify.xml存在问题"
			Exit Sub
		End If

		outsystemIdentify = esbbasepath & "out_conf\systemIdentify.xml"
		xdoc.load(outsystemIdentify)
		If xdoc.documentElement Is Nothing Then 
			MsgBox "OUT-systemIdentify.xml存在问题"
			Exit Sub
		End If
	End If
	MsgBox "All Config File Is Fine!"
End Sub

Main()

Set xdoc = Nothing
Set wshshell = Nothing
Set attrDict = Nothing