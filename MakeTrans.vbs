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
'---------服务治理控制台地址----------- [ 注意:需要斜杠 ]
mainurl = "http://159.1.65.149:18180/"
'---------跑库的连接地址---------------
linkurl33 = "ESBDATA/esbdata@159.1.33.210:1521/esbprd"
linkurl39 = "ESBDATA/esbdata@159.1.39.130:1521/esbprd"
'---------服务治理库地址---------------
serviceurl = "ESBSG/esbsg@159.1.65.153:1521/esbtest"
'---------生成临时文件----------------- [ 注意:需要斜杠 ]
commpath = "Tmp\"
'---------脚本存放目录----------------- [ 注意:需要斜杠 ]
scriptpath = "SQLScript\"
'---------元数据存放目录--------------- [ 注意:需要斜杠 ]
metadatapath = "Metadata\"
'---------svn metadata-----------------
'---------svn路径----------------------
svnpath = "Z:\WorkSpace\SmartESB\configs"
inconfpath = svnpath & "\in_conf\"
outconfpath = svnpath & "\out_conf\"
svndestin = inconfpath & "metadata\"
svndestout = outconfpath & "metadata\"
svnmetadata = svndestin & "metadata.xml"
'---------服务定义文件-----------------
newpath = "C:\Users\srxhx255\Desktop\交易制作工具\Metadata\"
'存放数据字典中所有字段名称
alldata = ""
'存放需要增加的语句
allNewData = ""
'----------是否为存量交易标识-----------
'True  存量
'False 新增复用
isExistFlag = True  
totalNum = 0
'=======================
'版本更新日志打印
'=======================
Sub versionLog()
	logfile.WriteLine("------------------" & Now() & "---------------------" & Chr(13) & Chr(13) & Chr(13))
	logfile.WriteLine("=============================================================" & Chr(13))
	logfile.WriteLine("︻┳═一 " & Now() & "--->Author : wangxingyang")
	logfile.WriteLine("︻┳═一 " & Now() & "--->Date : 2018/04/13")
	logfile.WriteLine("︻┳═一 " & Now() & "--->Contact : wangxyao@dcits.com")
	logfile.WriteLine("=============================================================" & Chr(13))
	logfile.WriteLine(Now() & "--->Version : v1.0 At 2018/04/13")
	logfile.WriteLine(Now() & "--->1、 检查交易是否为新增")
	logfile.WriteLine(Now() & "--->2、 导出并只执行脚本")
	logfile.WriteLine(Now() & "--->3、 导出相关配置文件")
	logfile.WriteLine(Now() & "--->4、 修改公共文件结构")
	logfile.WriteLine(Now() & "--->Version : v2.0 At 2018/04/16")
	logfile.WriteLine(Now() & "--->1、 增加解压文件功能")
	logfile.WriteLine(Now() & "--->2、 增加比较元数据功能")
	logfile.WriteLine(Now() & "--->3、 增加拷贝至配置库功能")
	logfile.WriteLine(Now() & "--->Version : v2.1 At 2018/05/04")
	logfile.WriteLine(Now() & "--->1、 修改弹出框BUG[弹出框有长度限制，导致处理交易报错]")
	logfile.WriteLine(Now() & "--->2、 修改文件系统名称小写问题")
	logfile.WriteLine(Now() & "--->Version : v2.2 At 2018/05/06")
	logfile.WriteLine(Now() & "--->1、 增加交易执行判断日志")
	logfile.WriteLine(Now() & "--->2、 增加读文件进行处理交易")
	logfile.WriteLine(Now() & "--->Version : v2.3 At 2018/05/10")
	logfile.WriteLine(Now() & "--->1、 修改元数据检查不出新增的BUG")
	logfile.WriteLine(Now() & "--->2、 修改重复添加元数据BUG")
	logfile.WriteLine(Now() & "--->Version : v2.4 At 2018/05/14")
	logfile.WriteLine(Now() & "--->1、 修改当元数据中包含元数据时不能添加的问题")
	logfile.WriteLine(Now() & "--->2、 新增修改元数据文件功能")
	logfile.WriteLine(Now() & "--->Version : v2.5 At 2018/05/15")
	logfile.WriteLine(Now() & "--->1、 增加文件比对功能，使用BCompare打开需要比较的工具")
	logfile.WriteLine(Now() & "--->2、 服务定义文件中，不存在的调用关系会新增")
	logfile.WriteLine(Now() & "--->Version : v2.6 At 2018/05/17")
	logfile.WriteLine(Now() & "--->1、 新增根据交易类型判断是否跑库")
	logfile.WriteLine(Now() & "--->2、 添加交易简称查询缓存，在一次运行中相同名称值检查一次")
	logfile.WriteLine(Now() & "--->3、 新增渠道会自动新增服务识别并跑新增脚本")
	logfile.WriteLine(Now() & "--->Version : v2.7 At 2018/05/13")
	logfile.WriteLine(Now() & "--->1、 新增根据模板替换拆组报文件功能")
	logfile.WriteLine(Now() & "--->2、 新增拼接更新适配器SQL功能")
	logfile.WriteLine(Now() & "--->Version : v2.8 At 2018/05/30")
	logfile.WriteLine(Now() & "--->1、 新增选择执行指定流程")
	logfile.WriteLine(Now() & "--->2、 优化路径，尽量少更改路径")
	logfile.WriteLine(Now() & "--->Version : v2.9 At 2018/06/12")
	logfile.WriteLine(Now() & "--->1、 修改元数据少添加的BUG")
	logfile.WriteLine(Now() & "--->Version : v3.0 At 2018/11/15")
	logfile.WriteLine(Now() & "--->1、 修改查询元数据比对堆栈溢出BUG")
	logfile.WriteLine(Now() & "--->2、 增加粘贴格式校验")
	logfile.WriteLine("------------------" & Now() & "---------------------" & vbcrlf & vbcrlf)
End Sub
'=======================
'登录服务治理系统
'=======================
Sub login(username,password)
	url = mainurl & "login/"
	http.Open "POST",url,False
	http.setRequestHeader "Content-Type","application/x-www-form-urlencoded; charset=UTF-8"
	http.Send "username=" & username & "&password=" & password
	strHtml = http.responsetext
	cookies = http.getResponseHeader("Set-Cookie")
	logfile.WriteLine(Now() & "--->登录服务治理成功")
End Sub
'=======================
'检查交易是否是新增复用
'=======================
Sub checkTran(code,from,dest)
	logfile.WriteLine(Now() & "--->开始查看["& code &"]是否是全新交易")
	cmdstr = "sqlplus "& linkurl33 &" @" & scriptpath & "CheckTran.sql" & " " & code
	Set exeRs = wshshell.exec(cmdstr)
	retmsg = exeRs.StdOut.ReadAll()
	logfile.WriteLine(Now() & "--->查询结果如下:")
	logfile.WriteLine(Now() & "--->"& retmsg)
	If InStr(retmsg,"no rows") = 0 Then
		logfile.WriteLine(Now() & "--->交易[" & code & "]为新增复用交易!!!")
		fromname = queryName(from)
		tranfile.WriteLine("交易[" & code & "] ------> 新增复用，调用关系为: [ " & from & " = " & fromname &" ] -> " & code & " -> [ " & dest & " ]")
		isExistFlag = True
		totalNum = totalNum + 2
	Else
		logfile.WriteLine(Now() & "--->交易[" & code & "]为全新交易!!!")
		fromname = queryName(from)
		tranfile.WriteLine("交易[" & code & "] ------> 全新，调用关系为: [ " & from & " = " & fromname &" ] -> " & code & " -> [ " & dest & " ]")
		isExistFlag = False
		totalNum = totalNum + 6
	End If
	logfile.WriteLine(Now() & "--->查看["& code &"]是否是全新交易结束")
End Sub
'=======================
'导出并执行脚本
'=======================
Sub exportAndRunSQLOne(tranId)
	'--------导出脚本--------------
	'判断是否是存量，如果是存量，不用跑数据库脚本
	If isExistFlag Then
		logfile.WriteLine(Now() & "--->交易[" & tranId & "]为新增复用交易，不跑脚本。直接退出！")
		Exit Sub
	End If
	logfile.WriteLine(Now() & "--->开始导出[" & tranId & "]SQL脚本")
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
	'----------替换Adapter----------
	'---后期添加
	'------可采用先跑在更新的方式
	'------或者拿到后替换更新
	'----------写文件---------------
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set file = fso.CreateTextFile( scriptpath & operationIds &".sql",True)
	file.WriteLine(sqltmp)
	file.WriteLine("commit;")
	file.WriteLine("exit;")
	file.close()
	logfile.WriteLine(Now() & "--->导出[" & tranId & "]SQL脚本完成")
	logfile.WriteLine(Now() & "--->导出内容如下：")
	logfile.WriteLine(Now() & "---> ")
	logfile.WriteLine(sqltmp & "commit;")
	logfile.WriteLine(sqltmp & "exit;")
	'-----------跑刚刚保存下来的SQL文件---------
	logfile.WriteLine(Now() & "--->在["& linkurl33 &"]开始执行导出脚本 " & scriptpath & operationIds &".sql")
	Set exeRs = wshshell.exec("sqlplus " & linkurl33 &" @" & scriptpath & operationIds &".sql")
	retmsg = exeRs.StdOut.ReadAll()
	If InStr(retmsg,"ORA") > 0 Then
		logfile.WriteLine(Now() & " Error->执行["& tranId &"]存在错误，请查看日志......")
	End If
	logfile.WriteLine(Now() & "--->拼接更新适配器语句开始..................."& vbcrlf)
	'拼接服务场景码，用于写SQL更新
	sqlText = ""
	scriptfile.WriteLine("---" & tranId)
	If InStr(retmsg,"BIND_PROTOCOLID_REF") > 0 Then
		sqlText = "insert into BINDMAP(SERVICEID, STYPE, LOCATION, VERSION, PROTOCOLID, MAPTYPE) VALUES('" & tranId & "', 'SERVICE', 'local_out', '0', '" & tranId & "Adapter', 'request');" & vbcrlf
	End If
	sqlText = sqlText & "UPDATE SERVICESYSTEMMAP SET ADAPTER = '" & tranId & "Adapter' WHERE SERVICEID = '" & tranId & "';" & vbcrlf
	sqlText = sqlText & "UPDATE BINDMAP SET PROTOCOLID = '" & tranId & "Adapter' WHERE SERVICEID = '" & tranId & "';" & vbcrlf
	scriptfile.WriteLine(sqlText)
	logfile.WriteLine(sqlText)
	logfile.WriteLine(Now() & "--->拼接更新适配器语句结束..................."& vbcrlf)
	logfile.WriteLine(Now() & "--->执行结果：")
	logfile.WriteLine(Now() & "-----------------------------------------------------------")
	logfile.WriteLine(retmsg)
	logfile.WriteLine(Now() & "-----------------------------------------------------------")
	logfile.WriteLine(Now() & "--->导出脚本执行结束 " & scriptpath & operationIds &".sql")
	'-----------跑刚刚保存下来的SQL文件---------
	logfile.WriteLine(Now() & "--->在["& linkurl39 &"]开始执行导出脚本 " & scriptpath & operationIds &".sql")
	Set exeRs = wshshell.exec("sqlplus " & linkurl39 &" @" & scriptpath & operationIds &".sql")
	retmsg = exeRs.StdOut.ReadAll()
	If InStr(retmsg,"ORA") > 0 Then
		logfile.WriteLine(Now() & " Error->执行["& tranId &"]存在错误，请查看日志......")
	End If
	logfile.WriteLine(Now() & "--->执行结果：")
	logfile.WriteLine(Now() & "-----------------------------------------------------------")
	logfile.WriteLine(retmsg)
	logfile.WriteLine(Now() & "-----------------------------------------------------------")
	logfile.WriteLine(Now() & "--->导出脚本执行结束 " & scriptpath & operationIds &".sql")
	Set sfile = fso.getfile(scriptpath & operationIds &".sql")
	sfile.delete
	Set fso = Nothing
End Sub
'=======================
'执行SQL脚本
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
		logfile.WriteLine(Now() & " Error->执行新增渠道["& en_name &"]存在错误，请查看日志......")
	End If
	logfile.WriteLine(Now() & "--->执行结果：")
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
		logfile.WriteLine(Now() & " Error->执行新增渠道["& en_name &"]存在错误，请查看日志......")
	End If
	logfile.WriteLine(Now() & "--->执行结果：")
	logfile.WriteLine(Now() & "-----------------------------------------------------------")
	logfile.WriteLine(retmsg)
	logfile.WriteLine(Now() & "-----------------------------------------------------------")
End Sub
'=======================
'导出脚本并执行，支持多个
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
'导出交易的配置文件
'=======================
Sub exportConf(from,dest,code)
	'连接数据库查询相关数据
	logfile.WriteLine(Now() & "--->准备导出配置文件")
	logfile.WriteLine(Now() & "--->查询导出配置文件所需参数")
	serviceId = Left(code,11)
	senceId = Right(code,2)
	cmdstr = "sqlplus "& serviceurl &" @" & scriptpath & "GetInfo.sql" & " " & serviceId & " " & senceId & " " & from & " " & dest
	Set exeRs = wshshell.exec(cmdstr)
	retmsg = exeRs.StdOut.ReadAll()
	logfile.WriteLine(Now() & "--->查询导出配置文件所需参数如下:")
	logfile.WriteLine(Now() & "--->"& retmsg)

	If InStr(retmsg,"no rows") = 0 Then
		logfile.WriteLine(Now() & "--->开始导出["& code &"]配置文件")
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
		logfile.WriteLine(Now() & "--->导出["& code &"]配置文件完成")
		Set file1 = Nothing
	Else
		MsgBox "数据库未查询到配置信息,注意查看!!!"
		Exit Sub
	End If 
	'-------------解压文件---------------
	cmdstr = "cmd /c cd %cd%\" & metadatapath  & " & winrar.exe x -ad -y "& code &"-metadata.zip"
	Set exeRs = wshshell.exec(cmdstr)
	'添加这句才能够等待执行完成结果，才能串行
	retmsg = exeRs.StdOut.ReadAll()
	'-------------比较元数据--------------
	'compareMeteData(code)
	'-------------拷贝文件至ESB路径下-----------
	'copyConfToSvn(code)
End Sub
'=======================
'查询对应系统的简称
'=======================
Function queryName(from)
	result = ""
	
	If attrDict.Exists(from) Then
		result = attrDict.Item(from)
		logfile.WriteLine(Now() & "---> 字典中已缓存，直接取缓存")
	Else
		logfile.WriteLine(Now() & "---> 字典中未缓存，需要查库")
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
'=======================
'查询对应系统的名称
'=======================
Function queryCHName(from)
	result = ""
	
	If attrDict.Exists(from) Then
		result = attrDict.Item(from)
		logfile.WriteLine(Now() & "--->字典中已缓存，直接取缓存")
	Else
		logfile.WriteLine(Now() & "---> 字典中未缓存，需要查库")
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
			MsgBox "未查询到[" & from & "]对应的简称,注意查看!!!"
		End If
		attrDict.Add from,result
	End If
	queryCHName = result
End Function
'=======================
'修改服务识别和系统识别
'=======================
Sub modifyXml(from,dest,code)
	logfile.WriteLine(Now() & "--->开始修改公共文件")

	pathService = commpath & "serviceIdentify.xml"
	pathSystem = commpath & "systemIdentify.xml"
	'连接数据库查询相关数据
	logfile.WriteLine(Now() & "--->开始查询[" & from & "]对应的简称")
	serviceId = Left(code,11)
	senceId = Right(code,2)

	result = queryName(from)

	If result = "" Then
		logfile.WriteLine(Now() & "--->没查到简称呢，直接退出了哟.....")
		Exit Sub
	End If 
	
	logfile.WriteLine(Now() & "--->[" & from & "]===>[" & result & "]")
	logfile.WriteLine(Now() & "--->查询[" & from & "]对应的简称结束")
	
	logfile.WriteLine(Now() & "--->开始修改[" & pathService & "]文件")
	xdoc.Load(pathService)
	ReadServiceXml xdoc,result,from,dest,code
	xdoc.Save pathService
	logfile.WriteLine(Now() & "--->修改[" & pathService & "]文件结束")

	logfile.WriteLine(Now() & "--->开始修改[" & pathSystem & "]文件")
	xdoc.Load(pathSystem)
	ReadSystemXml  xdoc,result,dest,code
	xdoc.Save pathSystem
	logfile.WriteLine(Now() & "--->修改[" & pathSystem & "]文件结束")
	logfile.WriteLine(Now() & "--->修改公共文件结束")
End Sub
'=======================
'解析服务识别文件并添加新交易
'=======================
Sub ReadServiceXml(xdoc,from,ch_name,dest,code)
	'取出所有channel
	Set nodes = xdoc.documentElement.selectNodes(".//channel")
	For Each node In nodes
		Set Alist = node.Attributes
		Dim node2
		For i = 0 To Alist.Length - 1
			Dim attr
			Set attr = Alist.Item(i)
			If attr.NodeName = "id" And attr.NodeValue = from Then
				'取出所有 switch
				Set m = node.getElementsByTagName("switch")(0)
				Set childs = m.childNodes

				For Each node1 In node.selectNodes(".//switch")
					For j = 0 To node1.ChildNodes.Length -1
						Set tmp = node1.ChildNodes(j)
						Set node2 = node1
						'判断是否存在该交易
						If tmp.NodeType = 1 And tmp.NodeName = "case" And Trim(tmp.Text) = code Then
							'有的话直接退出
							logfile.WriteLine(Now() & "--->存在该调用关系"& from &" -- " & code)
							Exit Sub
						End If 
					Next
				Next
				
				'在注释处插入该交易
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

							logfile.WriteLine(Now() & "---> 在注释处插入该交易 "& from &" --> " & code & " --> " & dest)
							Exit Sub
						End If 
					End If
				Next
				
				'插入注释和交易
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
	'若没有查到CHANNEL则为新增渠道，需要新增渠道
	logfile.WriteLine(Now() & "--->**********此渠道为新增渠道 "& from &" --> " & code & " --> " & dest)
	Set channels = xdoc.documentElement
	logfile.WriteLine(Now() & "---> 新增 CHANNEL 节点")
	Set channel = xdoc.createElement("channel")
	Set channelid = xdoc.createAttribute("id")
	Set channeltype = xdoc.createAttribute("type")

	channelid.text = from
	channel.setAttributeNode channelid
	channeltype.text = "dynamic"
	channel.setAttributeNode channeltype
	
	logfile.WriteLine(Now() & "---> 新增 Switch 节点")
	Set switch = xdoc.createElement("switch")
	Set switchmode = xdoc.createAttribute("mode")
	Set switchexpression = xdoc.createAttribute("expression")

	switchmode.text = "soap"
	switch.setAttributeNode switchmode
	switchexpression.text = "/soapenv:Envelope/soapenv:Body"
	switch.setAttributeNode switchexpression
	
	logfile.WriteLine(Now() & "---> 新增 namespace 节点")
	Set namespace = xdoc.createElement("namespace")
	Set namespacevalue = xdoc.createAttribute("value")

	namespacevalue.text = "soapenv"
	namespace.setAttributeNode namespacevalue
	namespace.text = "http://schemas.xmlsoap.org/soap/envelope/"


	logfile.WriteLine(Now() & "---> 新增 交易 节点")
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

	'执行 渠道插入脚本 
	RunSqlScript from,"","AddChannel"
End Sub
'=======================
'将执行结果写回到文件
'=======================
Sub AddLineAfterWord(frompath,topath,word,text)
	'非标准的XML文件，不能使用XML函数处理
	logfile.WriteLine(Now() & "--->开始添加元数据的差异")
	If Trim(text) = "" Then
		logfile.WriteLine(Now() & "--->无差异，无需操作")
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
	logfile.WriteLine(Now() & "--->内容如下")
	logfile.WriteLine(Now() & "--->")
	logfile.WriteLine(text)
	logfile.WriteLine(Now() & "--->完成添加元数据的差异")
End Sub
'=======================
'未使用
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
'解析系统识别并添加新增交易
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
						logfile.WriteLine(Now() & "存在该交易.......")
						Exit Sub 
					End If
				Next
				'插入该交易
				Set cNode = xdoc.createElement("service")
				cNode.Text = code
				node.appendChild(cNode)
				Exit Sub
			End If
		Next
	Next
	'若是新增系统问题
	logfile.WriteLine(Now() & "--->**********系统[ " & dest & " ]为新增提供方 "& from &" --> " & code & " --> " & dest)
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

	'执行 系统插入脚本 
	RunSqlScript dest,ch_name,"AddSystem"
End Sub
'=======================
'加载元数据
'=======================
Sub loadMetaData()
	logfile.WriteLine(Now() & "--->开始加载最新Metadata.xml文件...")
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
	logfile.WriteLine(Now() & "--->加载最新Metadata.xml完成..." )
	Set fso = Nothing
End Sub 
'=======================
'比对元数据
'=======================
Sub compareMeteData(code)
	'文件方式比较
	logfile.WriteLine(Now() & "--->开始比较[" & code & "]与元数据的差异")
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
						logfile.WriteLine(Now() & "--->开始添加[" & tmpp1 & "]元数据,类型为[array]")
						'添加已经添加的字段，以免重复新增
						allNewData = allNewData & "	<" & tmpp1 & " type=" & Chr(34) & "array"& Chr(34) &"/>"& vbcrlf
						alldata = alldata & tmpp1 & ","
					End If 
				Else
					isExists = 1
					checkExists alldata,tmpp1,isExists
					If isExists = 0 Then
						outpath.WriteLine("<" & tmpp1 & " type=" & Chr(34) & "string"& Chr(34) &" length="& Chr(34) & "255" & Chr(34) &"/>")
						logfile.WriteLine(Now() & "--->开始添加[" & tmpp1 & "]元数据,类型为[string]")
						allNewData = allNewData & "	<" & tmpp1 & " type=" & Chr(34) & "string"& Chr(34) &" length="& Chr(34) & "255" & Chr(34) &"/>" & vbcrlf
						alldata = alldata & tmpp1 & ","
					End If
				End If

			End If
		Loop
	Next
	Set fso = Nothing
	logfile.WriteLine(Now() & "--->完成比较[" & code & "]与元数据的差异")
End Sub
'=======================
'修改拆组包文件
'=======================
Sub modifyPackFile(from,dest,code)
	'此处需要根据各个现场修改
	'修改拆组包文件逻辑需要再次判断
	'isExistFlag = False
	Set fso = CreateObject("Scripting.FileSystemObject")
	logfile.WriteLine(Now() & "--->开始修改拆组包文件[" & code & "]")
	If Not fso.fileExists("Template\" & dest & ".xml") Or isExistFlag Then
		logfile.WriteLine(Now() & "--->没有模版文件[" & code & "]")
		Exit Sub 
	End If 
	servicepath = newpath & code & "-metadata\out_config\" & dest & "\service_"& code &"_system_" & dest & ".xml"
	servicenewpath = newpath & "service\service_"& code &"_system_" & dest & ".xml"

	channelepath = newpath & code & "-metadata\out_config\" & dest & "\channel_" & dest & "_service_" & code & ".xml"
	channelenewpath = newpath & "service\channel_" & dest & "_service_" & code & ".xml"

	logfile.WriteLine(Now() & "--->开始修改组包文件[" & code & "]")
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

	logfile.WriteLine(Now() & "--->开始修改拆包文件[" & code & "]")
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

	logfile.WriteLine(Now() & "--->完成修改拆组包文件[" & code & "]")
End Sub
'=======================
'替换文件
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
'检查元数据是否存在
'=======================
Function checkExists(strBase,strNeed,isExists)
	startPos = InStr(strBase,strNeed)
	If startPos = 0 Then
		isExists = 0
		Exit Function
	Else
		'此处若需要的元数据在字段结尾会出现问题
		isExists = startPos
	End If
	endPos = InStr(startPos,strBase,",")
	newstart = InstrRev(Mid(strBase,1,startPos),",")
	newval =  Mid(strBase,newstart+1,endPos-newstart-1)
	newStrBase = Mid(strBase,endPos)
	'只有长度相等的元数据才继续检查,减少递归的次数
	While Len(newval) <> Len(strNeed)
		startPos = InStr(newStrBase,strNeed)
		'若新的基础字符串中不含需要的元数据，直接退出
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
'拷贝元数据至SVN
'=======================
Sub copyConfToSvn(code,from)
	logfile.WriteLine(Now() & "--->开始拷贝文件至SVN[" & code & "]")
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
				logfile.WriteLine(Now() & "--->开始拷贝文件夹-->[" & intemp & otmp.Name & "]")
				fso.CopyFolder intemp & otmp.Name & "*",svndestin
			End If 
		Else
			logfile.WriteLine(Now() & "--->开始拷贝文件夹-->[" & intemp & otmp.Name & "]")
			fso.CopyFolder intemp & otmp.Name & "*",svndestin
		End If 
	Next 

	For Each itmp In outSubFolder
		If Not isExistFlag Then
			logfile.WriteLine(Now() & "--->开始拷贝文件夹-->[" & outtemp & "]")
			fso.CopyFolder outtemp & "*",svndestout
		End If 
	Next

	Set fso = Nothing
End Sub
'=======================
'加载文件中的交易
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

	'拷贝文件至tmp下
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
' 更新SVN
'=======================
Function updateSVN()
	logfile.WriteLine(Now() & "--->开始更新 SVN .......")
	cmdstr = "svn update "& svnpath
	Set exeRs = wshshell.exec(cmdstr)
	retmsg = exeRs.StdOut.ReadAll()
	logfile.WriteLine(Now() & "--->更新 SVN 日志如下:")
	logfile.WriteLine( retmsg)
	logfile.WriteLine(Now() & "--->更新 SVN 完成.......")
End Function
'=======================
'主方法
'=======================
Sub Main()
	'tranText = InputBox("请输入需要开发的交易，形如[3001300000333,双屏系统,NCBS;]","提醒","3001300000333,双屏系统,NCBS;")
	'==========更新SVN至最新============
	updateSVN
	'==========加载文本文件中的交易============
	tranText = LoadTrans()
	If tranText = "" Then
		MsgBox "请检查[ traninfo.txt ]是否已[ 经粘贴交易 ] 或 [ 有多余空格 ]!",0,"星阳提醒"
		Exit Sub
	Else		
		Ans = MsgBox("[开始处理]:   共[ " &  UBound(Split(tranText,";")) & " ]个交易" & Chr(13) & tranText,VbOKCancel,"星阳提醒")
		If Ans = vbCancel Then
			MsgBox "你取消了操作！",0,"星阳提醒"
			Exit Sub 
		End If
		'==========登录服务治理平台============
		login "admin","753951"
		'===========加载最新元数据=============
		loadMetaData()
		'分割每条记录
		trans = Split(tranText,";")
		'根据填写的执行方法进行
		tranFlow = InputBox("请输入需要执行的动作:" & Chr(13) & "    1.  检查交易是否为新增       [ 1 ]" & Chr(13) & "    2.  导出并只执行脚本         [ 2 ]" & Chr(13) & "    3.  导出相关配置文件         [ 3 ]" & Chr(13) & "    4.  修改公共文件结构         [ 4 ]" & Chr(13) & "    5.  比较元数据差异           [ 5 ]" & Chr(13) & "    6.  拷贝文件至SVN路径        [ 6 ]" & Chr(13) & "    7.  修改组包文件             [ 7 ]" & Chr(13) & "    8.  修改Metadata.xml文件     [ 8 ]" & Chr(13) & "    9.  运行文件比对             [ 9 ]" & Chr(13) & "    10. 流程全部执行             [ A ]","提醒","A")
		tranFlow = UCase(tranFlow)
		If tranFlow = "" Then 
			MsgBox "你取消了操作！"
			Exit Sub
		End If

		For i= 0 To UBound(trans)
			'分割交易明细信息  -->  3001300000333,双屏系统,NCBS;
			If trans(i) <> "" Then
				traninfo =  Split(trans(i),",")
				code = Trim(traninfo(0))
				from = Trim(traninfo(1))
				dest = Trim(traninfo(2))
				'=========检查交易是否为新增===========
				If InStr(tranFlow,"1") > 0 Or InStr(tranFlow,"A") > 0 Then
					checkTran code,from,dest
				End If 
				'==========导出并只执行脚本============
				If InStr(tranFlow,"2") > 0 Or InStr(tranFlow,"A") > 0 Then
					exportAndRunSQLOne(code)
				End If 
				'==========导出相关配置文件============
				If InStr(tranFlow,"3") > 0 Or InStr(tranFlow,"A") > 0 Then
					exportConf from,dest,code
				End If
				'==========修改公共文件结构============
				If InStr(tranFlow,"4") > 0 Or InStr(tranFlow,"A") > 0 Then
					modifyXml from,dest,code
				End If 
				'===========比较元数据差异=============
				If InStr(tranFlow,"5") > 0 Or InStr(tranFlow,"A") > 0 Then
					compareMeteData code
				End If 
				'========拷贝文件至SVN路径=============
				If InStr(tranFlow,"6") > 0 Or InStr(tranFlow,"A") > 0 Then
					copyConfToSvn code,from
				End If 
				'===========修改组包文件===============
				If InStr(tranFlow,"7") > 0 Or InStr(tranFlow,"A") > 0 Then
					modifyPackFile from,dest,code
				End If 
			End If 
		Next
		'修改Metadata.xml文件
		If InStr(tranFlow,"8") > 0 Or InStr(tranFlow,"A") > 0 Then
			AddLineAfterWord svnmetadata,commpath & "metadata.xml","metadata",allNewData
		End If 
		'运行文件比对
		'CompareAns = MsgBox("需要进行文件比对吗？",VbOKCancel,"星阳提醒")
		'If CompareAns <> vbCancel Then
		If InStr(tranFlow,"9") > 0 Or InStr(tranFlow,"A") > 0 Then
			'比对 serviceIdentify.xml
			wshshell.run("BComp.exe " & commpath & "serviceIdentify.xml "& inconfpath &"serviceIdentify.xml")
			'比对 systemIdentify.xml
			wshshell.run("BComp.exe " & commpath & "systemIdentify.xml "& inconfpath &"systemIdentify.xml")
			'比对 metadata.xml
			wshshell.run("BComp.exe " & commpath & "metadata.xml "& inconfpath &"metadata\metadata.xml")
			'in out 比对
			wshshell.run("BComp.exe " & outconfpath & "systemIdentify.xml "& inconfpath &"systemIdentify.xml")
			wshshell.run("BComp.exe " & outconfpath & "metadata\metadata.xml "& inconfpath &"metadata\metadata.xml")
		'End If
		End If
	End If
	logfile.WriteLine(Now() & "--->配置文件共[" & totalNum & "]个")
	MsgBox "o(∩_∩)o 啦啦啦->[ 全部处理完成 ] o(∩_∩)o",0,"星阳提醒"
End Sub
Main()
Set http = Nothing
Set html = Nothing
Set wshshell = Nothing
Set xdoc = Nothing
Set logfso = Nothing
Set tranfso = Nothing