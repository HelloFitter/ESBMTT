'Author:wangxingyang
'Date:20180413
'write at nanjing bank
Set http=CreateObject("MSXML2.XMLHTTP")
Set html = CreateObject("htmlfile")
Set wshshell = CreateObject("wscript.shell")
Set xdoc = CreateObject("MSXML2.DOMDocument")
Set logfso = CreateObject("Scripting.FileSystemObject")
Set tranfso = CreateObject("Scripting.FileSystemObject")
Set scriptfdo = CreateObject("Scripting.FileSystemObject")
Set attrDict = CreateObject("Scripting.Dictionary")
nowdate = Now()
nowdate = Year(nowdate) & Right("0" & Month(nowdate),2)& Right("0" & Day(nowdate),2)& Right("0" & Hour(nowdate),2)& Right("0" & Minute(nowdate),2)&  Right("0" & Second(nowdate),2)
logpath = "Log/Run-[" & nowdate & "].log"
tranpath = "Log/Tran-[" & nowdate & "].log"
scriptpath1 = "Log/Script.log"
Set logfile = logfso.CreateTextFile(logpath,True)
Set tranfile = tranfso.CreateTextFile(tranpath,True)
Set scriptfile = scriptfdo.CreateTextFile(scriptpath1,True)
Dim cookies
'---------�����������̨��ַ----------- [ ע��:��Ҫб�� ]
mainurl = "http://159.1.65.149:18180/"
'---------�ܿ�����ӵ�ַ---------------
linkurl33 = "ESBDATA/esbdata@159.1.33.210:1521/esbprd"
linkurl39 = "ESBDATA/esbdata@159.1.39.130:1521/esbprd"
'---------����������ַ---------------
serviceurl = "ESBSG/esbsg@159.1.65.153:1521/esbtest"
'---------������ʱ�ļ�----------------- [ ע��:��Ҫб�� ]
commpath = "Tmp\"
'---------�ű����Ŀ¼----------------- [ ע��:��Ҫб�� ]
scriptpath = "SQLScript\"
'---------Ԫ���ݴ��Ŀ¼--------------- [ ע��:��Ҫб�� ]
metadatapath = "Metadata\"
'---------svn metadata-----------------
'---------svn·��----------------------
svnpath = "Z:\WorkSpace\SmartESB\configs"
inconfpath = svnpath & "\in_conf\"
outconfpath = svnpath & "\out_conf\"
svndestin = inconfpath & "metadata\"
svndestout = outconfpath & "metadata\"
svnmetadata = svndestin & "metadata.xml"
'---------�������ļ�-----------------
newpath = "C:\Users\srxhx255\Desktop\������������\Metadata\"
'��������ֵ��������ֶ�����
alldata = ""
'�����Ҫ���ӵ����
allNewData = ""
'----------�Ƿ�Ϊ�������ױ�ʶ-----------
'True  ����
'False ��������
isExistFlag = True  
totalNum = 0
'=======================
'�汾������־��ӡ
'=======================
Sub versionLog()
	logfile.WriteLine("------------------" & Now() & "---------------------" & Chr(13) & Chr(13) & Chr(13))
	logfile.WriteLine("=============================================================" & Chr(13))
	logfile.WriteLine("��רTһ " & Now() & "--->Author : wangxingyang")
	logfile.WriteLine("��רTһ " & Now() & "--->Date : 2018/04/13")
	logfile.WriteLine("��רTһ " & Now() & "--->Contact : wangxyao@dcits.com")
	logfile.WriteLine("=============================================================" & Chr(13))
	logfile.WriteLine(Now() & "--->Version : v1.0 At 2018/04/13")
	logfile.WriteLine(Now() & "--->1�� ��齻���Ƿ�Ϊ����")
	logfile.WriteLine(Now() & "--->2�� ������ִֻ�нű�")
	logfile.WriteLine(Now() & "--->3�� ������������ļ�")
	logfile.WriteLine(Now() & "--->4�� �޸Ĺ����ļ��ṹ")
	logfile.WriteLine(Now() & "--->Version : v2.0 At 2018/04/16")
	logfile.WriteLine(Now() & "--->1�� ���ӽ�ѹ�ļ�����")
	logfile.WriteLine(Now() & "--->2�� ���ӱȽ�Ԫ���ݹ���")
	logfile.WriteLine(Now() & "--->3�� ���ӿ��������ÿ⹦��")
	logfile.WriteLine(Now() & "--->Version : v2.1 At 2018/05/04")
	logfile.WriteLine(Now() & "--->1�� �޸ĵ�����BUG[�������г������ƣ����´����ױ���]")
	logfile.WriteLine(Now() & "--->2�� �޸��ļ�ϵͳ����Сд����")
	logfile.WriteLine(Now() & "--->Version : v2.2 At 2018/05/06")
	logfile.WriteLine(Now() & "--->1�� ���ӽ���ִ���ж���־")
	logfile.WriteLine(Now() & "--->2�� ���Ӷ��ļ����д�����")
	logfile.WriteLine(Now() & "--->Version : v2.3 At 2018/05/10")
	logfile.WriteLine(Now() & "--->1�� �޸�Ԫ���ݼ�鲻��������BUG")
	logfile.WriteLine(Now() & "--->2�� �޸��ظ����Ԫ����BUG")
	logfile.WriteLine(Now() & "--->Version : v2.4 At 2018/05/14")
	logfile.WriteLine(Now() & "--->1�� �޸ĵ�Ԫ�����а���Ԫ����ʱ������ӵ�����")
	logfile.WriteLine(Now() & "--->2�� �����޸�Ԫ�����ļ�����")
	logfile.WriteLine(Now() & "--->Version : v2.5 At 2018/05/15")
	logfile.WriteLine(Now() & "--->1�� �����ļ��ȶԹ��ܣ�ʹ��BCompare����Ҫ�ȽϵĹ���")
	logfile.WriteLine(Now() & "--->2�� �������ļ��У������ڵĵ��ù�ϵ������")
	logfile.WriteLine(Now() & "--->Version : v2.6 At 2018/05/17")
	logfile.WriteLine(Now() & "--->1�� �������ݽ��������ж��Ƿ��ܿ�")
	logfile.WriteLine(Now() & "--->2�� ��ӽ��׼�Ʋ�ѯ���棬��һ����������ͬ����ֵ���һ��")
	logfile.WriteLine(Now() & "--->3�� �����������Զ���������ʶ���������ű�")
	logfile.WriteLine(Now() & "--->Version : v2.7 At 2018/05/13")
	logfile.WriteLine(Now() & "--->1�� ��������ģ���滻���鱨�ļ�����")
	logfile.WriteLine(Now() & "--->2�� ����ƴ�Ӹ���������SQL����")
	logfile.WriteLine(Now() & "--->Version : v2.8 At 2018/05/30")
	logfile.WriteLine(Now() & "--->1�� ����ѡ��ִ��ָ������")
	logfile.WriteLine(Now() & "--->2�� �Ż�·���������ٸ���·��")
	logfile.WriteLine(Now() & "--->Version : v2.9 At 2018/06/12")
	logfile.WriteLine(Now() & "--->1�� �޸�Ԫ��������ӵ�BUG")
	logfile.WriteLine(Now() & "--->Version : v3.0 At 2018/11/15")
	logfile.WriteLine(Now() & "--->1�� �޸Ĳ�ѯԪ���ݱȶԶ�ջ���BUG")
	logfile.WriteLine(Now() & "--->2�� ����ճ����ʽУ��")
	logfile.WriteLine("------------------" & Now() & "---------------------" & vbcrlf & vbcrlf)
End Sub
'=======================
'��¼��������ϵͳ
'=======================
Sub login(username,password)
	url = mainurl & "login/"
	http.Open "POST",url,False
	http.setRequestHeader "Content-Type","application/x-www-form-urlencoded; charset=UTF-8"
	http.Send "username=" & username & "&password=" & password
	strHtml = http.responsetext
	cookies = http.getResponseHeader("Set-Cookie")
	logfile.WriteLine(Now() & "--->��¼��������ɹ�")
End Sub
'=======================
'��齻���Ƿ�����������
'=======================
Sub checkTran(code,from,dest)
	logfile.WriteLine(Now() & "--->��ʼ�鿴["& code &"]�Ƿ���ȫ�½���")
	cmdstr = "sqlplus "& linkurl33 &" @" & scriptpath & "CheckTran.sql" & " " & code
	Set exeRs = wshshell.exec(cmdstr)
	retmsg = exeRs.StdOut.ReadAll()
	logfile.WriteLine(Now() & "--->��ѯ�������:")
	logfile.WriteLine(Now() & "--->"& retmsg)
	If InStr(retmsg,"no rows") = 0 Then
		logfile.WriteLine(Now() & "--->����[" & code & "]Ϊ�������ý���!!!")
		fromname = queryName(from)
		tranfile.WriteLine("����[" & code & "] ------> �������ã����ù�ϵΪ: [ " & from & " = " & fromname &" ] -> " & code & " -> [ " & dest & " ]")
		isExistFlag = True
		totalNum = totalNum + 2
	Else
		logfile.WriteLine(Now() & "--->����[" & code & "]Ϊȫ�½���!!!")
		fromname = queryName(from)
		tranfile.WriteLine("����[" & code & "] ------> ȫ�£����ù�ϵΪ: [ " & from & " = " & fromname &" ] -> " & code & " -> [ " & dest & " ]")
		isExistFlag = False
		totalNum = totalNum + 6
	End If
	logfile.WriteLine(Now() & "--->�鿴["& code &"]�Ƿ���ȫ�½��׽���")
End Sub
'=======================
'������ִ�нű�
'=======================
Sub exportAndRunSQLOne(tranId)
	'--------�����ű�--------------
	'�ж��Ƿ��Ǵ���������Ǵ��������������ݿ�ű�
	If isExistFlag Then
		logfile.WriteLine(Now() & "--->����[" & tranId & "]Ϊ�������ý��ף����ܽű���ֱ���˳���")
		Exit Sub
	End If
	logfile.WriteLine(Now() & "--->��ʼ����[" & tranId & "]SQL�ű�")
	type1 = "service"
	operationIds = tranId
	url = mainurl & "esbScriptExport/preview"
	http.Open "POST",url,False
	http.setRequestHeader "Content-Type","application/x-www-form-urlencoded; charset=UTF-8"
	http.setRequestHeader "Cookie",cookies
	http.Send "type=" & type1 & "&operationIds=" & operationIds
	strHtml = http.responsetext
	Set window = html.parentWindow
	window.execScript "var json = " & strHtml, "JScript"
	Set ret = window.json
	sqltmp = ret.result
	'----------�滻Adapter----------
	'---�������
	'------�ɲ��������ڸ��µķ�ʽ
	'------�����õ����滻����
	'----------д�ļ�---------------
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set file = fso.CreateTextFile( scriptpath & operationIds &".sql",True)
	file.WriteLine(sqltmp)
	file.WriteLine("commit;")
	file.WriteLine("exit;")
	file.close()
	logfile.WriteLine(Now() & "--->����[" & tranId & "]SQL�ű����")
	logfile.WriteLine(Now() & "--->�����������£�")
	logfile.WriteLine(Now() & "---> ")
	logfile.WriteLine(sqltmp & "commit;")
	logfile.WriteLine(sqltmp & "exit;")
	'-----------�ܸոձ���������SQL�ļ�---------
	logfile.WriteLine(Now() & "--->��["& linkurl33 &"]��ʼִ�е����ű� " & scriptpath & operationIds &".sql")
	Set exeRs = wshshell.exec("sqlplus " & linkurl33 &" @" & scriptpath & operationIds &".sql")
	retmsg = exeRs.StdOut.ReadAll()
	If InStr(retmsg,"ORA") > 0 Then
		logfile.WriteLine(Now() & " Error->ִ��["& tranId &"]���ڴ�����鿴��־......")
	End If
	logfile.WriteLine(Now() & "--->ƴ�Ӹ�����������俪ʼ..................."& vbcrlf)
	'ƴ�ӷ��񳡾��룬����дSQL����
	sqlText = ""
	scriptfile.WriteLine("---" & tranId)
	If InStr(retmsg,"BIND_PROTOCOLID_REF") > 0 Then
		sqlText = "insert into BINDMAP(SERVICEID, STYPE, LOCATION, VERSION, PROTOCOLID, MAPTYPE) VALUES('" & tranId & "', 'SERVICE', 'local_out', '0', '" & tranId & "Adapter', 'request');" & vbcrlf
	End If
	sqlText = sqlText & "UPDATE SERVICESYSTEMMAP SET ADAPTER = '" & tranId & "Adapter' WHERE SERVICEID = '" & tranId & "';" & vbcrlf
	sqlText = sqlText & "UPDATE BINDMAP SET PROTOCOLID = '" & tranId & "Adapter' WHERE SERVICEID = '" & tranId & "';" & vbcrlf
	scriptfile.WriteLine(sqlText)
	logfile.WriteLine(sqlText)
	logfile.WriteLine(Now() & "--->ƴ�Ӹ���������������..................."& vbcrlf)
	logfile.WriteLine(Now() & "--->ִ�н����")
	logfile.WriteLine(Now() & "-----------------------------------------------------------")
	logfile.WriteLine(retmsg)
	logfile.WriteLine(Now() & "-----------------------------------------------------------")
	logfile.WriteLine(Now() & "--->�����ű�ִ�н��� " & scriptpath & operationIds &".sql")
	'-----------�ܸոձ���������SQL�ļ�---------
	logfile.WriteLine(Now() & "--->��["& linkurl39 &"]��ʼִ�е����ű� " & scriptpath & operationIds &".sql")
	Set exeRs = wshshell.exec("sqlplus " & linkurl39 &" @" & scriptpath & operationIds &".sql")
	retmsg = exeRs.StdOut.ReadAll()
	If InStr(retmsg,"ORA") > 0 Then
		logfile.WriteLine(Now() & " Error->ִ��["& tranId &"]���ڴ�����鿴��־......")
	End If
	logfile.WriteLine(Now() & "--->ִ�н����")
	logfile.WriteLine(Now() & "-----------------------------------------------------------")
	logfile.WriteLine(retmsg)
	logfile.WriteLine(Now() & "-----------------------------------------------------------")
	logfile.WriteLine(Now() & "--->�����ű�ִ�н��� " & scriptpath & operationIds &".sql")
	Set sfile = fso.getfile(scriptpath & operationIds &".sql")
	sfile.delete
	Set fso = Nothing
End Sub
'=======================
'ִ��SQL�ű�
'=======================
Sub RunSqlScript(en_name,ch_name,sqlname)
	If sqlname = "AddChannel" Then
		cmdstr = "sqlplus "& linkurl33 &" @" & scriptpath & sqlname & ".sql" & " " & en_name & " " & en_name & "Connector " & ch_name
	Else
		cmdstr = "sqlplus "& linkurl33 &" @" & scriptpath & sqlname & ".sql" & " " & en_name & " " & en_name & "Adapter " & ch_name
	End If
	
	'logfile.WriteLine(Now() & " RunSqlScript " & cmdstr)

	Set exeRs = wshshell.exec(cmdstr)
	retmsg = exeRs.StdOut.ReadAll()
	If InStr(retmsg,"ORA") > 0 Then
		logfile.WriteLine(Now() & " Error->ִ����������["& en_name &"]���ڴ�����鿴��־......")
	End If
	logfile.WriteLine(Now() & "--->ִ�н����")
	logfile.WriteLine(Now() & "-----------------------------------------------------------")
	logfile.WriteLine(retmsg)
	logfile.WriteLine(Now() & "-----------------------------------------------------------")
	
	If sqlname = "AddChannel" Then
		cmdstr = "sqlplus "& linkurl39 &" @" & scriptpath & sqlname & ".sql" & " " & en_name & " " & en_name & "Connector " & ch_name
	Else
		cmdstr = "sqlplus "& linkurl39 &" @" & scriptpath & sqlname & ".sql" & " " & en_name & " " & en_name & "Adapter " & ch_name
	End If

	Set exeRs = wshshell.exec(cmdstr)
	retmsg = exeRs.StdOut.ReadAll()
	If InStr(retmsg,"ORA") > 0 Then
		logfile.WriteLine(Now() & " Error->ִ����������["& en_name &"]���ڴ�����鿴��־......")
	End If
	logfile.WriteLine(Now() & "--->ִ�н����")
	logfile.WriteLine(Now() & "-----------------------------------------------------------")
	logfile.WriteLine(retmsg)
	logfile.WriteLine(Now() & "-----------------------------------------------------------")
End Sub
'=======================
'�����ű���ִ�У�֧�ֶ��
'=======================
Sub exportAndRunSql(tranIds)
	If InStr(tranIds,",") = 0  Then
		exportAndRunSQLOne(tranIds)
	Else
		myarray = Split(tranIds,",")
		For i= 0 To UBound(myarray)
			exportAndRunSQLOne(myarray(i))
		next
	End If
End Sub
'=======================
'�������׵������ļ�
'=======================
Sub exportConf(from,dest,code)
	'�������ݿ��ѯ�������
	logfile.WriteLine(Now() & "--->׼�����������ļ�")
	logfile.WriteLine(Now() & "--->��ѯ���������ļ��������")
	serviceId = Left(code,11)
	senceId = Right(code,2)
	cmdstr = "sqlplus "& serviceurl &" @" & scriptpath & "GetInfo.sql" & " " & serviceId & " " & senceId & " " & from & " " & dest
	Set exeRs = wshshell.exec(cmdstr)
	retmsg = exeRs.StdOut.ReadAll()
	logfile.WriteLine(Now() & "--->��ѯ���������ļ������������:")
	logfile.WriteLine(Now() & "--->"& retmsg)

	If InStr(retmsg,"no rows") = 0 Then
		logfile.WriteLine(Now() & "--->��ʼ����["& code &"]�����ļ�")
		rets = Split(retmsg,"--------------------------------------------------------------------------------")
		tttt = Split(rets(4),"Disconnected")
		result = tttt(0)
		
		ret = Split(result,Chr(10))

		val1 = Replace(Trim(ret(1)),Chr(13),"")
		val2 = Replace(Trim(ret(3)),Chr(13),"")
		val3 = Replace(Trim(ret(2)),Chr(13),"")
		val4 = Replace(Trim(ret(4)),Chr(13),"")
		
		val1 = Replace(val1,Chr(10),"")
		val2 = Replace(val2,Chr(10),"")
		val3 = Replace(val3,Chr(10),"")
		val4 = Replace(val4,Chr(10),"")

		url = mainurl & "export/exportBatch"
		http.Open "POST",url,False
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded; charset=UTF-8"
		http.setRequestHeader "Accept-Encoding","gzip,deflate,sdch"
		http.setRequestHeader "Cache-Control","max-age=0"
		http.setRequestHeader "Connection","keep-alive"
		http.setRequestHeader "Cookie",cookies
		http.Send "list[0].consumerServiceInvokeId=" & val1 & "&list[0].providerServiceInvokeId=" & val2 & "&list[0].conGeneratorId=" & val3 & "&list[0].proGeneratorId=" & val4
		Set file1 = CreateObject("ADODB.Stream")
		file1.Mode = 3
		file1.Type = 1
		file1.Open()
		file1.write(http.responseBody)
		file1.SaveToFile metadatapath & code & "-" & "metadata.zip",2
		logfile.WriteLine(Now() & "--->����["& code &"]�����ļ����")
		Set file1 = Nothing
	Else
		MsgBox "���ݿ�δ��ѯ��������Ϣ,ע��鿴!!!"
		Exit Sub
	End If 
	'-------------��ѹ�ļ�---------------
	cmdstr = "cmd /c cd %cd%\" & metadatapath  & " & winrar.exe x -ad -y "& code &"-metadata.zip"
	Set exeRs = wshshell.exec(cmdstr)
	'��������ܹ��ȴ�ִ����ɽ�������ܴ���
	retmsg = exeRs.StdOut.ReadAll()
	'-------------�Ƚ�Ԫ����--------------
	'compareMeteData(code)
	'-------------�����ļ���ESB·����-----------
	'copyConfToSvn(code)
End Sub
'=======================
'��ѯ��Ӧϵͳ�ļ��
'=======================
Function queryName(from)
	result = ""
	
	If attrDict.Exists(from) Then
		result = attrDict.Item(from)
		logfile.WriteLine(Now() & "---> �ֵ����ѻ��棬ֱ��ȡ����")
	Else
		logfile.WriteLine(Now() & "---> �ֵ���δ���棬��Ҫ���")
		cmdstr = "sqlplus "& serviceurl &" @" & scriptpath & "GetName.sql" & " " & from
		Set exeRs = wshshell.exec(cmdstr)
		retmsg = exeRs.StdOut.ReadAll()
		If InStr(retmsg,"no rows") = 0 Then
			rets = Split(retmsg,"--------------------------------------------------------------------------------")
			tttt = Split(rets(1),"Disconnected")
			result = Replace(tttt(0),Chr(10),"")
			result = Replace(result,Chr(13),"")
		Else
			MsgBox "δ��ѯ��[" & from & "]��Ӧ�ļ��,ע��鿴!!!"
		End If
		attrDict.Add from,result
	End If
	queryName = result
End Function
'=======================
'��ѯ��Ӧϵͳ������
'=======================
Function queryCHName(from)
	result = ""
	
	If attrDict.Exists(from) Then
		result = attrDict.Item(from)
		logfile.WriteLine(Now() & "--->�ֵ����ѻ��棬ֱ��ȡ����")
	Else
		logfile.WriteLine(Now() & "---> �ֵ���δ���棬��Ҫ���")
		cmdstr = "sqlplus "& serviceurl &" @" & scriptpath & "GetCHName.sql" & " " & from
		logfile.WriteLine(Now() & "---> " & cmdstr)
		Set exeRs = wshshell.exec(cmdstr)
		retmsg = exeRs.StdOut.ReadAll()
		If InStr(retmsg,"no rows") = 0 Then
			rets = Split(retmsg,"--------------------------------------------------------------------------------")
			tttt = Split(rets(1),"Disconnected")
			result = Replace(tttt(0),Chr(10),"")
			result = Replace(result,Chr(13),"")
		Else
			MsgBox "δ��ѯ��[" & from & "]��Ӧ�ļ��,ע��鿴!!!"
		End If
		attrDict.Add from,result
	End If
	queryCHName = result
End Function
'=======================
'�޸ķ���ʶ���ϵͳʶ��
'=======================
Sub modifyXml(from,dest,code)
	logfile.WriteLine(Now() & "--->��ʼ�޸Ĺ����ļ�")

	pathService = commpath & "serviceIdentify.xml"
	pathSystem = commpath & "systemIdentify.xml"
	'�������ݿ��ѯ�������
	logfile.WriteLine(Now() & "--->��ʼ��ѯ[" & from & "]��Ӧ�ļ��")
	serviceId = Left(code,11)
	senceId = Right(code,2)

	result = queryName(from)

	If result = "" Then
		logfile.WriteLine(Now() & "--->û�鵽����أ�ֱ���˳���Ӵ.....")
		Exit Sub
	End If 
	
	logfile.WriteLine(Now() & "--->[" & from & "]===>[" & result & "]")
	logfile.WriteLine(Now() & "--->��ѯ[" & from & "]��Ӧ�ļ�ƽ���")
	
	logfile.WriteLine(Now() & "--->��ʼ�޸�[" & pathService & "]�ļ�")
	xdoc.Load(pathService)
	ReadServiceXml xdoc,result,from,dest,code
	xdoc.Save pathService
	logfile.WriteLine(Now() & "--->�޸�[" & pathService & "]�ļ�����")

	logfile.WriteLine(Now() & "--->��ʼ�޸�[" & pathSystem & "]�ļ�")
	xdoc.Load(pathSystem)
	ReadSystemXml  xdoc,result,dest,code
	xdoc.Save pathSystem
	logfile.WriteLine(Now() & "--->�޸�[" & pathSystem & "]�ļ�����")
	logfile.WriteLine(Now() & "--->�޸Ĺ����ļ�����")
End Sub
'=======================
'��������ʶ���ļ�������½���
'=======================
Sub ReadServiceXml(xdoc,from,ch_name,dest,code)
	'ȡ������channel
	Set nodes = xdoc.documentElement.selectNodes(".//channel")
	For Each node In nodes
		Set Alist = node.Attributes
		Dim node2
		For i = 0 To Alist.Length - 1
			Dim attr
			Set attr = Alist.Item(i)
			If attr.NodeName = "id" And attr.NodeValue = from Then
				'ȡ������ switch
				Set m = node.getElementsByTagName("switch")(0)
				Set childs = m.childNodes

				For Each node1 In node.selectNodes(".//switch")
					For j = 0 To node1.ChildNodes.Length -1
						Set tmp = node1.ChildNodes(j)
						Set node2 = node1
						'�ж��Ƿ���ڸý���
						If tmp.NodeType = 1 And tmp.NodeName = "case" And Trim(tmp.Text) = code Then
							'�еĻ�ֱ���˳�
							logfile.WriteLine(Now() & "--->���ڸõ��ù�ϵ"& from &" -- " & code)
							Exit Sub
						End If 
					Next
				Next
				
				'��ע�ʹ�����ý���
				For n = 1 To childs.length-1
					If childs(n).NodeType = 8 Then
						If Trim(childs(n).NodeValue) = dest Then
							Set nodetmp = xdoc.createElement("case")
							Set nodeattr = xdoc.createAttribute("value")

							If from = "ECIF" Or from = "PPL" Or from = "BIBPS" Then
								nodeattr.text = code
								nodetmp.setAttributeNode nodeattr
								nodetmp.text = code
							Else
								nodeattr.text = "Req" & code
								nodetmp.setAttributeNode nodeattr
								nodetmp.text = code
							End If

							m.insertBefore nodetmp,childs(n+1)

							logfile.WriteLine(Now() & "---> ��ע�ʹ�����ý��� "& from &" --> " & code & " --> " & dest)
							Exit Sub
						End If 
					End If
				Next
				
				'����ע�ͺͽ���
				Set nodetmp = xdoc.createElement("case")
				Set nodeattr = xdoc.createAttribute("value")
				Set nodecomm = xdoc.createComment(dest)

				If from = "ECIF" Or from = "PPL" Or from = "BIBPS"Then
					nodeattr.text = code
					nodetmp.setAttributeNode nodeattr
					nodetmp.text = code
				Else
					nodeattr.text = "Req" & code
					nodetmp.setAttributeNode nodeattr
					nodetmp.text = code
				End If

				m.appendChild(nodecomm)
				m.appendChild(nodetmp)
				Exit Sub
			End If 
		Next
	Next
	'��û�в鵽CHANNEL��Ϊ������������Ҫ��������
	logfile.WriteLine(Now() & "--->**********������Ϊ�������� "& from &" --> " & code & " --> " & dest)
	Set channels = xdoc.documentElement
	logfile.WriteLine(Now() & "---> ���� CHANNEL �ڵ�")
	Set channel = xdoc.createElement("channel")
	Set channelid = xdoc.createAttribute("id")
	Set channeltype = xdoc.createAttribute("type")

	channelid.text = from
	channel.setAttributeNode channelid
	channeltype.text = "dynamic"
	channel.setAttributeNode channeltype
	
	logfile.WriteLine(Now() & "---> ���� Switch �ڵ�")
	Set switch = xdoc.createElement("switch")
	Set switchmode = xdoc.createAttribute("mode")
	Set switchexpression = xdoc.createAttribute("expression")

	switchmode.text = "soap"
	switch.setAttributeNode switchmode
	switchexpression.text = "/soapenv:Envelope/soapenv:Body"
	switch.setAttributeNode switchexpression
	
	logfile.WriteLine(Now() & "---> ���� namespace �ڵ�")
	Set namespace = xdoc.createElement("namespace")
	Set namespacevalue = xdoc.createAttribute("value")

	namespacevalue.text = "soapenv"
	namespace.setAttributeNode namespacevalue
	namespace.text = "http://schemas.xmlsoap.org/soap/envelope/"


	logfile.WriteLine(Now() & "---> ���� ���� �ڵ�")
	Set nodetmp = xdoc.createElement("case")
	Set nodeattr = xdoc.createAttribute("value")
	Set nodecomm = xdoc.createComment(dest)

	If from = "ECIF" Or from = "PPL" Or from = "BIBPS" Then
		nodeattr.text = code
		nodetmp.setAttributeNode nodeattr
		nodetmp.text = code
	Else
		nodeattr.text = "Req" & code
		nodetmp.setAttributeNode nodeattr
		nodetmp.text = code
	End If

	channels.appendChild(channel)
	channel.appendChild(switch)
	switch.appendChild(namespace)
	switch.appendChild(nodecomm)
	switch.appendChild(nodetmp)

	'ִ�� ��������ű� 
	RunSqlScript from,"","AddChannel"
End Sub
'=======================
'��ִ�н��д�ص��ļ�
'=======================
Sub AddLineAfterWord(frompath,topath,word,text)
	'�Ǳ�׼��XML�ļ�������ʹ��XML��������
	logfile.WriteLine(Now() & "--->��ʼ���Ԫ���ݵĲ���")
	If Trim(text) = "" Then
		logfile.WriteLine(Now() & "--->�޲��죬�������")
		Exit Sub 
	End If 
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	pathMetadata4Read = frompath
	pathMetadata4Write = topath
	Set fileMetadata4Read = fso.OpenTextFile(pathMetadata4Read,1,False)
	Set fileMetadata4Write = fso.OpenTextFile(pathMetadata4Write,2,True)
	Do While fileMetadata4Read.AtEndOfStream <> True
		line = fileMetadata4Read.ReadLine()
		If Trim(line) = "</"& word &">" Then 
			fileMetadata4Write.Write(text)
			fileMetadata4Write.WriteLine("</"& word &">")
		Else
			fileMetadata4Write.WriteLine line
		End If
		
	Loop
	fileMetadata4Read.close()
	fileMetadata4Write.close()
	Set fso = Nothing
	logfile.WriteLine(Now() & "--->��������")
	logfile.WriteLine(Now() & "--->")
	logfile.WriteLine(text)
	logfile.WriteLine(Now() & "--->������Ԫ���ݵĲ���")
End Sub
'=======================
'δʹ��
'=======================
Sub notuse()
	Set xdocu = CreateObject("MSXML2.DOMDocument")
	pathMetadata = commpath & "metadata.xml"
	xdocu.Load(pathMetadata)
	
	Set nodetmp1 = xdocu.documentElement
	MsgBox (nodetmp1 Is Nothing)
	Set cNode = xdocu.createElement(nodename)
	
	Set nodeattr1 = xdocu.createAttribute("type")
	nodeattr1.text = nodetype
	cNode.setAttributeNode nodeattr1

	If nodetype = "string" Then 
		Set nodeattr2 = xdocu.createAttribute("length")
		nodeattr2.text = "255"
		cNode.setAttributeNode nodeattr2
	End If 

	nodetmp1.appendChild(cNode)
	xdocu.Save pathMetadata
End Sub 
'=======================
'����ϵͳʶ�������������
'=======================
Sub ReadSystemXml(xdoc,from,dest,code)
	logfile.WriteLine(Now() & from &" --> " & code & " --> " & dest)
	Set nodes = xdoc.documentElement.selectNodes(".//system")
	For Each node In nodes
		Set Alist = node.Attributes
		For i = 0 To Alist.Length - 1
			Dim attr
			Set attr = Alist.Item(i)
			If attr.NodeName = "id" And attr.NodeValue = dest Then
				For Each node1 In node.selectNodes(".//service")
					If node1.Text = code Then
						logfile.WriteLine(Now() & "���ڸý���.......")
						Exit Sub 
					End If
				Next
				'����ý���
				Set cNode = xdoc.createElement("service")
				cNode.Text = code
				node.appendChild(cNode)
				Exit Sub
			End If
		Next
	Next
	'��������ϵͳ����
	logfile.WriteLine(Now() & "--->**********ϵͳ[ " & dest & " ]Ϊ�����ṩ�� "& from &" --> " & code & " --> " & dest)
	Set systems = xdoc.documentElement

	Set system = xdoc.createElement("system")
	Set systemid = xdoc.createAttribute("id")

	systemid.text = dest
	system.setAttributeNode systemid

	Set cNode = xdoc.createElement("service")
	cNode.Text = code
	
	systems.appendChild(system)
	system.appendChild(cNode)

	ch_name=queryCHName(dest)
	'MsgBox (dest & "-->" & ch_name)

	'ִ�� ϵͳ����ű� 
	RunSqlScript dest,ch_name,"AddSystem"
End Sub
'=======================
'����Ԫ����
'=======================
Sub loadMetaData()
	logfile.WriteLine(Now() & "--->��ʼ��������Metadata.xml�ļ�...")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set svnfile = fso.OpenTextFile(svnmetadata,1,False)
	Do While svnfile.AtEndOfStream <> True
		line = svnfile.ReadLine()
		If InStr(line,"type") > 0 Then 
			tmpp = Split(line,"<")(1)
			name = Trim(Split(tmpp," ")(0))
			alldata = alldata & name & ","
		End If
	Loop
	logfile.WriteLine(Now() & "--->��������Metadata.xml���..." )
	Set fso = Nothing
End Sub 
'=======================
'�ȶ�Ԫ����
'=======================
Sub compareMeteData(code)
	'�ļ���ʽ�Ƚ�
	logfile.WriteLine(Now() & "--->��ʼ�Ƚ�[" & code & "]��Ԫ���ݵĲ���")
	Set fso = CreateObject("Scripting.FileSystemObject")
	temp = newpath & code & "-metadata\in_config\servicedefine"

	Set localfile = fso.GetFolder(temp)
	Set outpath = fso.CreateTextFile(newpath & "\out" & code & ".txt",True)	
	
	Set oFiles = localfile.Files
	For Each oFile In oFiles
		tmppath = oFile.Path
		Set file1 = fso.OpenTextFile(tmppath,1,False)
		Do While file1.AtEndOfStream <> True
			line = file1.ReadLine()
			If InStr(line,"metadataid") > 0 Then
				tmpp = Split(line,"metadataid")(1)
				tmpp1 = Trim(Split(tmpp,Chr(34))(1))
				If InStr(line,"type") > 0 Then
					tmpp2 = Split(line,"type")(1)
					tmpp3 = Split(tmpp2,Chr(34))(1)
					isExists = 1
					checkExists alldata,tmpp1,isExists
					'MsgBox len(alldata)   416771
					If (tmpp3 = "array" Or tmpp3 = "sopform") And isExists = 0 Then
						outpath.WriteLine("<" & tmpp1 & " type=" & Chr(34) & "array"& Chr(34) &"/>")
						logfile.WriteLine(Now() & "--->��ʼ���[" & tmpp1 & "]Ԫ����,����Ϊ[array]")
						'����Ѿ���ӵ��ֶΣ������ظ�����
						allNewData = allNewData & "	<" & tmpp1 & " type=" & Chr(34) & "array"& Chr(34) &"/>"& vbcrlf
						alldata = alldata & tmpp1 & ","
					End If 
				Else
					isExists = 1
					checkExists alldata,tmpp1,isExists
					If isExists = 0 Then
						outpath.WriteLine("<" & tmpp1 & " type=" & Chr(34) & "string"& Chr(34) &" length="& Chr(34) & "255" & Chr(34) &"/>")
						logfile.WriteLine(Now() & "--->��ʼ���[" & tmpp1 & "]Ԫ����,����Ϊ[string]")
						allNewData = allNewData & "	<" & tmpp1 & " type=" & Chr(34) & "string"& Chr(34) &" length="& Chr(34) & "255" & Chr(34) &"/>" & vbcrlf
						alldata = alldata & tmpp1 & ","
					End If
				End If

			End If
		Loop
	Next
	Set fso = Nothing
	logfile.WriteLine(Now() & "--->��ɱȽ�[" & code & "]��Ԫ���ݵĲ���")
End Sub
'=======================
'�޸Ĳ�����ļ�
'=======================
Sub modifyPackFile(from,dest,code)
	'�˴���Ҫ���ݸ����ֳ��޸�
	'�޸Ĳ�����ļ��߼���Ҫ�ٴ��ж�
	'isExistFlag = False
	Set fso = CreateObject("Scripting.FileSystemObject")
	logfile.WriteLine(Now() & "--->��ʼ�޸Ĳ�����ļ�[" & code & "]")
	If Not fso.fileExists("Template\" & dest & ".xml") Or isExistFlag Then
		logfile.WriteLine(Now() & "--->û��ģ���ļ�[" & code & "]")
		Exit Sub 
	End If 
	servicepath = newpath & code & "-metadata\out_config\" & dest & "\service_"& code &"_system_" & dest & ".xml"
	servicenewpath = newpath & "service\service_"& code &"_system_" & dest & ".xml"

	channelepath = newpath & code & "-metadata\out_config\" & dest & "\channel_" & dest & "_service_" & code & ".xml"
	channelenewpath = newpath & "service\channel_" & dest & "_service_" & code & ".xml"

	logfile.WriteLine(Now() & "--->��ʼ�޸�����ļ�[" & code & "]")
	xdoc.load(servicepath)
	tmp1 = Replace(xdoc.xml,"s:","")
	tmp2 = Replace(tmp1,"d:","")
	tmp = Replace(tmp2,":","")
	tp = Replace(tmp,"xmlns=" & Chr(34) & "http//esb.dcitsbiz.com/services/"& Left(code,11) & Chr(34),"")
	xdoc.loadXML(tp)
	Set nodes = xdoc.documentElement.selectNodes(".//ReqAppBody")(0)
	addstr = ""
	For Each node In nodes.ChildNodes
		addstr = addstr  & "		" & node.Xml & vbcrlf
	Next
	RepalceFile "Template\" & dest & ".xml",servicenewpath,addstr

	logfile.WriteLine(Now() & "--->��ʼ�޸Ĳ���ļ�[" & code & "]")
	xdoc.load(channelepath)
	tmp1 = Replace(xdoc.xml,"s:","")
	tmp2 = Replace(tmp1,"d:","")
	tmp = Replace(tmp2,":","")
	tp = Replace(tmp,"xmlns=" & Chr(34) & "http//esb.dcitsbiz.com/services/"& Left(code,11) & Chr(34),"")
	xdoc.loadXML(tp)
	Set nodes = xdoc.documentElement.selectNodes(".//RspAppBody")(0)
	addstr = ""
	For Each node In nodes.ChildNodes
		addstr = addstr  & "		" & node.Xml & vbcrlf
	Next
	RepalceFile "Template\" & dest & ".xml",channelenewpath,addstr

	logfile.WriteLine(Now() & "--->����޸Ĳ�����ļ�[" & code & "]")
End Sub
'=======================
'�滻�ļ�
'=======================
Sub RepalceFile(fromfile,tofile,text)
	Set file1 = CreateObject("ADODB.Stream")
	file1.Mode = 3
	file1.Type = 1
	file1.Open()
	file1.LoadFromFile fromfile
	file1.Position = 0
	file1.Type = 2
	file1.Charset = "utf-8"
	tmpfile = file1.ReadText
	file1.Position = 0
	file1.SetEos
	file1.Type = 2
	file1.Charset = "utf-8"
	file1.WriteText Replace(tmpfile,"XXXXXXXXXXXXXXXXXXXX",text)
	file1.SaveToFile tofile,2
	Set file1 = Nothing
End Sub
'=======================
'���Ԫ�����Ƿ����
'=======================
Function checkExists(strBase,strNeed,isExists)
	startPos = InStr(strBase,strNeed)
	If startPos = 0 Then
		isExists = 0
		Exit Function
	Else
		'�˴�����Ҫ��Ԫ�������ֶν�β���������
		isExists = startPos
	End If
	endPos = InStr(startPos,strBase,",")
	newstart = InstrRev(Mid(strBase,1,startPos),",")
	newval =  Mid(strBase,newstart+1,endPos-newstart-1)
	newStrBase = Mid(strBase,endPos)
	'ֻ�г�����ȵ�Ԫ���ݲż������,���ٵݹ�Ĵ���
	While Len(newval) <> Len(strNeed)
		startPos = InStr(newStrBase,strNeed)
		'���µĻ����ַ����в�����Ҫ��Ԫ���ݣ�ֱ���˳�
		If startPos = 0 Then
			isExists = 0
			Exit Function
		End If
		endPos = InStr(startPos,newStrBase,",")
		newstart = InstrRev(Mid(newStrBase,1,startPos),",")
		newval =  Mid(newStrBase,newstart+1,endPos-newstart-1)
		newStrBase = Mid(newStrBase,endPos)
	Wend
	If strNeed <> newval Then
		checkExists newStrBase,strNeed,isExists
	Else
		isExists = newstart+1
		Exit Function
	End If 
End Function
'=======================
'����Ԫ������SVN
'=======================
Sub copyConfToSvn(code,from)
	logfile.WriteLine(Now() & "--->��ʼ�����ļ���SVN[" & code & "]")
	intemp = newpath & code & "-metadata\in_config\"
	outtemp = newpath & code & "-metadata\out_config\"
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	Set oFolder = fso.GetFolder(intemp)
	Set oSubFolder = oFolder.SubFolders

	Set outFolder = fso.GetFolder(outtemp)
	Set outSubFolder = outFolder.SubFolders
	
	'isExistFlag = False

	For Each otmp In oSubFolder
		If isExistFlag Then
			If Trim(LCase(otmp.Name)) <> "servicedefine" Then 
				logfile.WriteLine(Now() & "--->��ʼ�����ļ���-->[" & intemp & otmp.Name & "]")
				fso.CopyFolder intemp & otmp.Name & "*",svndestin
			End If 
		Else
			logfile.WriteLine(Now() & "--->��ʼ�����ļ���-->[" & intemp & otmp.Name & "]")
			fso.CopyFolder intemp & otmp.Name & "*",svndestin
		End If 
	Next 

	For Each itmp In outSubFolder
		If Not isExistFlag Then
			logfile.WriteLine(Now() & "--->��ʼ�����ļ���-->[" & outtemp & "]")
			fso.CopyFolder outtemp & "*",svndestout
		End If 
	Next

	Set fso = Nothing
End Sub
'=======================
'�����ļ��еĽ���
'=======================
Function LoadTrans()
	versionLog()
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set traninfo = fso.OpenTextFile("traninfo.txt",1,False)
	tmpstr = ""

	Do While traninfo.AtEndOfStream <> True
		line = traninfo.ReadLine()
		If Trim(line) <> "" Then
			num = Len(line)- Len(Replace(line," ",""))
			If num <> 2 Then
				Exit Function
			End If 
			tmpstr = tmpstr & Replace(line," ",",") & ";"
		End If
	Loop

	'�����ļ���tmp��
	serviceIdentify = "serviceIdentify.xml"
	systemIdentify = "systemIdentify.xml"
	metadata = "metadata.xml"

	fso.copyfile inconfpath & serviceIdentify,commpath & serviceIdentify
	fso.copyfile inconfpath & systemIdentify,commpath & systemIdentify
	fso.copyfile svnmetadata,commpath & metadata

	LoadTrans = tmpstr

	Set fso = Nothing
End Function
'=======================
' ����SVN
'=======================
Function updateSVN()
	logfile.WriteLine(Now() & "--->��ʼ���� SVN .......")
	cmdstr = "svn update "& svnpath
	Set exeRs = wshshell.exec(cmdstr)
	retmsg = exeRs.StdOut.ReadAll()
	logfile.WriteLine(Now() & "--->���� SVN ��־����:")
	logfile.WriteLine( retmsg)
	logfile.WriteLine(Now() & "--->���� SVN ���.......")
End Function
'=======================
'������
'=======================
Sub Main()
	'tranText = InputBox("��������Ҫ�����Ľ��ף�����[3001300000333,˫��ϵͳ,NCBS;]","����","3001300000333,˫��ϵͳ,NCBS;")
	'==========����SVN������============
	updateSVN
	'==========�����ı��ļ��еĽ���============
	tranText = LoadTrans()
	If tranText = "" Then
		MsgBox "����[ traninfo.txt ]�Ƿ���[ ��ճ������ ] �� [ �ж���ո� ]!",0,"��������"
		Exit Sub
	Else		
		Ans = MsgBox("[��ʼ����]:   ��[ " &  UBound(Split(tranText,";")) & " ]������" & Chr(13) & tranText,VbOKCancel,"��������")
		If Ans = vbCancel Then
			MsgBox "��ȡ���˲�����",0,"��������"
			Exit Sub 
		End If
		'==========��¼��������ƽ̨============
		login "admin","753951"
		'===========��������Ԫ����=============
		loadMetaData()
		'�ָ�ÿ����¼
		trans = Split(tranText,";")
		'������д��ִ�з�������
		tranFlow = InputBox("��������Ҫִ�еĶ���:" & Chr(13) & "    1.  ��齻���Ƿ�Ϊ����       [ 1 ]" & Chr(13) & "    2.  ������ִֻ�нű�         [ 2 ]" & Chr(13) & "    3.  ������������ļ�         [ 3 ]" & Chr(13) & "    4.  �޸Ĺ����ļ��ṹ         [ 4 ]" & Chr(13) & "    5.  �Ƚ�Ԫ���ݲ���           [ 5 ]" & Chr(13) & "    6.  �����ļ���SVN·��        [ 6 ]" & Chr(13) & "    7.  �޸�����ļ�             [ 7 ]" & Chr(13) & "    8.  �޸�Metadata.xml�ļ�     [ 8 ]" & Chr(13) & "    9.  �����ļ��ȶ�             [ 9 ]" & Chr(13) & "    10. ����ȫ��ִ��             [ A ]","����","A")
		tranFlow = UCase(tranFlow)
		If tranFlow = "" Then 
			MsgBox "��ȡ���˲�����"
			Exit Sub
		End If

		For i= 0 To UBound(trans)
			'�ָ����ϸ��Ϣ  -->  3001300000333,˫��ϵͳ,NCBS;
			If trans(i) <> "" Then
				traninfo =  Split(trans(i),",")
				code = Trim(traninfo(0))
				from = Trim(traninfo(1))
				dest = Trim(traninfo(2))
				'=========��齻���Ƿ�Ϊ����===========
				If InStr(tranFlow,"1") > 0 Or InStr(tranFlow,"A") > 0 Then
					checkTran code,from,dest
				End If 
				'==========������ִֻ�нű�============
				If InStr(tranFlow,"2") > 0 Or InStr(tranFlow,"A") > 0 Then
					exportAndRunSQLOne(code)
				End If 
				'==========������������ļ�============
				If InStr(tranFlow,"3") > 0 Or InStr(tranFlow,"A") > 0 Then
					exportConf from,dest,code
				End If
				'==========�޸Ĺ����ļ��ṹ============
				If InStr(tranFlow,"4") > 0 Or InStr(tranFlow,"A") > 0 Then
					modifyXml from,dest,code
				End If 
				'===========�Ƚ�Ԫ���ݲ���=============
				If InStr(tranFlow,"5") > 0 Or InStr(tranFlow,"A") > 0 Then
					compareMeteData code
				End If 
				'========�����ļ���SVN·��=============
				If InStr(tranFlow,"6") > 0 Or InStr(tranFlow,"A") > 0 Then
					copyConfToSvn code,from
				End If 
				'===========�޸�����ļ�===============
				If InStr(tranFlow,"7") > 0 Or InStr(tranFlow,"A") > 0 Then
					modifyPackFile from,dest,code
				End If 
			End If 
		Next
		'�޸�Metadata.xml�ļ�
		If InStr(tranFlow,"8") > 0 Or InStr(tranFlow,"A") > 0 Then
			AddLineAfterWord svnmetadata,commpath & "metadata.xml","metadata",allNewData
		End If 
		'�����ļ��ȶ�
		'CompareAns = MsgBox("��Ҫ�����ļ��ȶ���",VbOKCancel,"��������")
		'If CompareAns <> vbCancel Then
		If InStr(tranFlow,"9") > 0 Or InStr(tranFlow,"A") > 0 Then
			'�ȶ� serviceIdentify.xml
			wshshell.run("BComp.exe " & commpath & "serviceIdentify.xml "& inconfpath &"serviceIdentify.xml")
			'�ȶ� systemIdentify.xml
			wshshell.run("BComp.exe " & commpath & "systemIdentify.xml "& inconfpath &"systemIdentify.xml")
			'�ȶ� metadata.xml
			wshshell.run("BComp.exe " & commpath & "metadata.xml "& inconfpath &"metadata\metadata.xml")
			'in out �ȶ�
			wshshell.run("BComp.exe " & outconfpath & "systemIdentify.xml "& inconfpath &"systemIdentify.xml")
			wshshell.run("BComp.exe " & outconfpath & "metadata\metadata.xml "& inconfpath &"metadata\metadata.xml")
		'End If
		End If
	End If
	logfile.WriteLine(Now() & "--->�����ļ���[" & totalNum & "]��")
	MsgBox "o(��_��)o ������->[ ȫ��������� ] o(��_��)o",0,"��������"
End Sub
Main()
Set http = Nothing
Set html = Nothing
Set wshshell = Nothing
Set xdoc = Nothing
Set logfso = Nothing
Set tranfso = Nothing