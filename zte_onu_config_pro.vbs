# $language = "VBScript"
# $interface = "1.0"
'本脚本适用于secureCRT 7.1 以上版本，其它低版本没有进行测试
'本脚本适用于自动配置中兴C320和C300 OLT下的ONU的自动配置
'本脚本的作者是 山东广电网络有限公司济宁分公司的 姬广超
'Email:	newchoose@163.com	欢迎交流学习

dim g_fso
set g_fso = CreateObject("Scripting.FileSystemObject")
Const ForWriting = 2 
Const ForAppending = 8

'delayTime 为扫描中兴C320 OLT的周期时间
dim delayTime
delayTime = 1000 * 5

dim objCurrentTab
dim strPrompt

Set re = New RegExp
re.Global = True  
re.Pattern = "(\d/\s\d+/\d+)\s+(\d+)\s+"

Sub main()

	dim olt_name
	olt_name = ""
	
	set objCurrentTab = crt.GetScriptTab
	
	'获取外层vlan信息
	'打开当前目录下的zte_olt_info.csv文件，并将文件内容全部读取到strFileData里面
	set objStream = g_fso.openTextFile(".\zte_olt_info.csv", 1, false)
	strFileData = objStream.ReadAll
	objStream.close
	'Name,IP,UserName,Password,Svlan2,Svlan1
	'中心机房C320,172.16.90.200,zte,zte,2064,1064
	'中心机房C300,172.16.90.201,zte,zte,2540,1540
	'李营站,172.16.90.202,zte,zte,2619,1619
	'接庄站,172.16.90.203,zte,zte,2539,1539
	'二十里铺站,172.16.90.204,zte,zte,2699,1699
	'石桥站,172.16.90.205,zte,zte,2679,1679
	'长沟站,172.16.90.206,zte,zte,2659,1659
	
	vLines = split(strFileData, vbcrlf)
	'分析读取到的每一行，并且忽略首行
	for i = 1 to UBound(vLines)
		oltInfo = split(vLines(i), ",")
		if inStr(objCurrentTab.Caption, oltInfo(1)) then
			'获取配置文件中外层vlan的信息
			cvlan2 = oltInfo(4)
			cvlan1 = oltInfo(5)
			'获取OLT名称和ip信息
			olt_name = oltInfo(0)
			olt_ip = oltInfo(1)
			username = oltInfo(2)
			passwd = oltInfo(3)
			exit for
		end if
	next
	
	if olt_name = "" then
		msgbox "没有找到当前ZTE-OLT的配置信息！"
		exit sub
	end if
	
	'重置一下连接，输入密码
	objCurrentTab.Session.Disconnect
	objCurrentTab.session.Connect()
	'输入用户名和密码，并确认是否进入#号模式，并写日志
	inputPasswd olt_ip, username, passwd
	'获取当前界面的命令提示符
	strPrompt = getstrPrompt(objCurrentTab)
	
	do	
		'如果连接断开则不断的尝试重新连接
		On Error Resume Next
	
		objCurrentTab.Screen.Synchronous = True
		objCurrentTab.Screen.Send "show onu unauthentication" & vbcr
		objCurrentTab.Screen.waitForString vbcr
		strResult = crt.Screen.ReadString(strPrompt)
		objCurrentTab.Screen.Synchronous = false
		
		if Instr(strResult, "40529: No related information") then 
			'do something
		else
			'使用正则表达式获取onu序号和mac地址
			re. = "epon-onu_(\d/\d+/\d+):.+([0-9a-f]{4}\.[0-9a-f]{4}\.[0-9a-f]{4})"
			If re.Test(strResults) <> True Then
				MsgBox "异常错误！"
				crt.quit
			Else
				Set matches = re.Execute(strResults)
				For Each match In matches
					epon_num = match.SubMatches(0)
					onu_mac = match.SubMatches(1)
				Next
			End If
			
			'下面的代码的功能是通过截取命令结果分析出该 Pon 口下最后一个ONU的编号
			objCurrentTab.Screen.Synchronous = True
			objCurrentTab.Screen.Send "show running-config interface epon-olt_" & epon_num & vbcr
			objCurrentTab.Screen.waitForString vbcr
			strResult = crt.Screen.ReadString(strPrompt)
			objCurrentTab.Screen.Synchronous = false
			
			'下面的代码功能是对返回的结果进行分行
			strLines = Split(strResult, vbcrlf)
			'如果当前ONU为该PON口下的第一个设备
			if UBound(strLines) - LBound(strLines) = 6 then
				last_num = 0
			'否则获取当前该PON口下的最后一个设备的编号
			else
				'获取最后一行的数组索引
				lastIndex = UBound(strLines) - 3
				'strLines(lastIndex) = "onu 2 type ZTE-F400 mac c4a3.66c7.a8ae ip-cfg static"
				str1 = Split(strLines(lastIndex), " ")
				'str1(1) = "13"
				last_num = str1(1)
			end if
			
			'配置ONU的所有必要参数都获取到啦，接下来就调用一个Sub就OK啦
			'为了应付raisecom类型的ONU上线，强制ONU的类型为ZTE-F400
			if instr(onu_type, "ZTE") then
				config_ZTE_ONU epon_num, onu_type, onu_mac, last_num + 1, cvlan2, cvlan1
			else
				config_ZTE_ONU epon_num, "ZTE-F400", onu_mac, last_num + 1, cvlan2, cvlan1
			end if


			'写日志是必须的啊，这才显的专业!
			Set logStream = g_fso.OpenTextFile(".\zte_onu_config_log.txt", 8, True)
			logStream.WriteLine Now & ", ZTE-OLT: " & olt_name & "(" & olt_ip & ")  add an ONU : " & _
					epon_num & ":" & last_num + 1 & ", type: " & onu_type & ", mac: " & onu_mac 
			logStream.close
			
		end if
		
		nError = Err.Number 
		strErr = Err.Description 
		' Restore normal error handling so that any other errors in our 
		' script are not masked/ignored 
		On Error Goto 0 
		
		'发现错误进行写日志，并且尝试重新连接远程OLT
		If nError <> 0 Then
			Set logStream = g_fso.OpenTextFile(".\zte_onu_config_log.txt", 8, True)
			logStream.writeLine Now & ", An Error happened on ZTE-OLT: " & olt_name & "(" & olt_ip & ") . Error: " & strErr
			objCurrentTab.Session.Disconnect
			logStream.writeLine Now & ", The session for ZTE-OLT: " & olt_name & "(" & olt_ip & ") was disconnected. Trying reConnect..."
			logStream.close
			objCurrentTab.session.Connect()
			'输入用户名和密码，并确认是否进入#号模式
			inputPasswd olt_ip, username, passwd
		end if
		
		'给定时间内休息一会
		crt.sleep delayTime
		
	loop 
	
End Sub

'该过程的作用是：输入OLT的用户名和密码（用户名和密码根据当地自己情况）, 并确认是否进入#号模式
Sub inputPasswd(olt_ip, username, passwd)

	
	Set objCurrentTab = crt.GetScriptTab
	objCurrentTab.Screen.Synchronous = True
	
	objCurrentTab.Screen.WaitForString "sername:"
	objCurrentTab.Screen.Send username & chr(13)
	objCurrentTab.Screen.WaitForString "assword:"
	objCurrentTab.Screen.Send passwd & chr(13)
	
	'判断是否进入#号模式
	if objCurrentTab.Screen.WaitForString("#", 3) <> true then
		msgbox "没有进入#号模式，请检查用户名和密码相关信息！程序执行失败！"
		crt.Quit
	end if
	
	'打开日志文件, 如果没有则新建该文件
	Set logStream = g_fso.OpenTextFile(".\zte_onu_config_log.txt", 8, True)
	logStream.writeLine Now & ", The Script has been running at ZTE-OLT: " & olt_name & "(" & olt_ip & ")"
	logStream.close
	objCurrentTab.Screen.Synchronous = false
End Sub

'该函数的功能是获取给定界面的命令提示符
Function getstrPrompt(objCurrentTab)

	objCurrentTab.activate
	
	if objCurrentTab.Session.Connected = True  then
		
			objCurrentTab.Screen.Send vbcrlf
			objCurrentTab.Screen.WaitForString vbcr

			Do 
			' Attempt to detect the command prompt heuristically... 
				Do 
					bCursorMoved = objCurrentTab.Screen.WaitForCursor(1)
				Loop Until bCursorMoved = False
			' Once the cursor has stopped moving for about a second, we'll 
			' assume it's safe to start interacting with the remote system. 
			' Get the shell prompt so that we can know what to look for when 
			' determining if the command is completed. Won't work if the prompt 
			' is dynamic (e.g., changes according to current working folder, etc.) 
				nRow = objCurrentTab.Screen.CurrentRow 
				strPrompt = objCurrentTab.screen.Get(nRow, 0, nRow, objCurrentTab.Screen.CurrentColumn - 1)
				' Loop until we actually see a line of text appear (the 
				' timeout for WaitForCursor above might not be enough 
				' for slower-responding hosts. 
				strPrompt = Trim(strPrompt)
				If strPrompt <> "" Then Exit Do
			Loop 
		
			getstrPrompt = strPrompt
		
	end if

End Function


'该过程的功能是用给定的参数配置ONU
'epon_num：为指定ONU所在的PON口的序号
'onu_type：为指定ONU的型号
'onu_mac： 为指定ONU的mac地址
'onu_num： 为指定ONU的序号
'cvlan2：  为互联网的外层vlan
'cvlan1:   为点播的外层vlan
Sub config_ZTE_ONU(epon_num, onu_type, onu_mac, onu_num, cvlan2, cvlan1)
	
	'根据ONU的序号自动计算出互联网和点播的内层Vlan
	onu_vlan2 = 2000 + onu_num
	onu_vlan1 = 1000 + onu_num
	
	objCurrentTab.Screen.Synchronous = True
	
	objCurrentTab.Screen.Send "con t" & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "interface epon-olt_" & epon_num & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "onu " & onu_num & " type "  & onu_type & " mac " & onu_mac & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "exit" & vbcr
	objCurrentTab.Screen.WaitForString "#"
	
	objCurrentTab.Screen.Send "interface epon-onu_" & epon_num & ":" & onu_num & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "service-port 1 vport 1 user-vlan " & onu_vlan2 & " to " & onu_vlan2 & " svlan " & cvlan2 & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "service-port 2 vport 1 user-vlan " & onu_vlan1 & " to " & onu_vlan1 & " svlan " & cvlan1 & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "exit" & vbcr
	objCurrentTab.Screen.WaitForString "#"
	
	objCurrentTab.Screen.Send "pon-onu-mng epon-onu_" & epon_num & ":" & onu_num & vbcr
	objCurrentTab.Screen.WaitForString "#"
	
	'因为中兴C320下发配置慢的原因，这里需要延迟
	'如果输入命令太快会提示：%Code 40796: This operation is forbidden, because down config is in process!

	crt.sleep 1000 * 25
	objCurrentTab.Screen.Send "auto-config" & vbcr
	crt.Screen.WaitForString("#")

	'do
		'crt.sleep 1000
		'objCurrentTab.Screen.Send "auto-config" & vbcr
		'crt.Screen.WaitForString("#")
		'nRow = objCurrentTab.Screen.CurrentRow - 1
		'cmdResult = objCurrentTab.screen.Get(nRow, 0, nRow, objCurrentTab.Screen.columns)
		'msgbox cmdResult
	'loop while Instr(cmdResult, "down config is in process!") 
	
	objCurrentTab.Screen.Send "vlan port eth_0/1 mode tag vlan " &  onu_vlan2 & " priority 0"  & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "vlan port eth_0/2 mode tag vlan " &  onu_vlan1 & " priority 0"  & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "vlan port eth_0/3 mode tag vlan " &  onu_vlan2 & " priority 0"  & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "vlan port eth_0/4 mode tag vlan " &  onu_vlan2 & " priority 0"  & vbcr
	objCurrentTab.Screen.WaitForString "#"
	
	objCurrentTab.Screen.Send "save" & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "exit" & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "exit" & vbcr
	objCurrentTab.Screen.WaitForString "#"
	objCurrentTab.Screen.Send "write" & vbcr
	objCurrentTab.Screen.WaitForString "#"
	
	objCurrentTab.Screen.Synchronous = false
	
End Sub



