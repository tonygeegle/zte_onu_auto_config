# $language = "VBScript"
# $interface = "1.0"
'���ű�������secureCRT 7.1 ���ϰ汾�������Ͱ汾û�н��в���
'���ű��������Զ���������C320��C300 OLT�µ�ONU���Զ�����
'���ű��������� ɽ������������޹�˾�����ֹ�˾�� ���㳬
'Email:	newchoose@163.com	��ӭ����ѧϰ

dim g_fso
set g_fso = CreateObject("Scripting.FileSystemObject")
Const ForWriting = 2 
Const ForAppending = 8

'delayTime Ϊɨ������C320 OLT������ʱ��
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
	
	'��ȡ���vlan��Ϣ
	'�򿪵�ǰĿ¼�µ�zte_olt_info.csv�ļ��������ļ�����ȫ����ȡ��strFileData����
	set objStream = g_fso.openTextFile(".\zte_olt_info.csv", 1, false)
	strFileData = objStream.ReadAll
	objStream.close
	'Name,IP,UserName,Password,Svlan2,Svlan1
	'���Ļ���C320,172.16.90.200,zte,zte,2064,1064
	'���Ļ���C300,172.16.90.201,zte,zte,2540,1540
	'��Ӫվ,172.16.90.202,zte,zte,2619,1619
	'��ׯվ,172.16.90.203,zte,zte,2539,1539
	'��ʮ����վ,172.16.90.204,zte,zte,2699,1699
	'ʯ��վ,172.16.90.205,zte,zte,2679,1679
	'����վ,172.16.90.206,zte,zte,2659,1659
	
	vLines = split(strFileData, vbcrlf)
	'������ȡ����ÿһ�У����Һ�������
	for i = 1 to UBound(vLines)
		oltInfo = split(vLines(i), ",")
		if inStr(objCurrentTab.Caption, oltInfo(1)) then
			'��ȡ�����ļ������vlan����Ϣ
			cvlan2 = oltInfo(4)
			cvlan1 = oltInfo(5)
			'��ȡOLT���ƺ�ip��Ϣ
			olt_name = oltInfo(0)
			olt_ip = oltInfo(1)
			username = oltInfo(2)
			passwd = oltInfo(3)
			exit for
		end if
	next
	
	if olt_name = "" then
		msgbox "û���ҵ���ǰZTE-OLT��������Ϣ��"
		exit sub
	end if
	
	'����һ�����ӣ���������
	objCurrentTab.Session.Disconnect
	objCurrentTab.session.Connect()
	'�����û��������룬��ȷ���Ƿ����#��ģʽ����д��־
	inputPasswd olt_ip, username, passwd
	'��ȡ��ǰ�����������ʾ��
	strPrompt = getstrPrompt(objCurrentTab)
	
	do	
		'������ӶϿ��򲻶ϵĳ�����������
		On Error Resume Next
	
		objCurrentTab.Screen.Synchronous = True
		objCurrentTab.Screen.Send "show onu unauthentication" & vbcr
		objCurrentTab.Screen.waitForString vbcr
		strResult = crt.Screen.ReadString(strPrompt)
		objCurrentTab.Screen.Synchronous = false
		
		if Instr(strResult, "40529: No related information") then 
			'do something
		else
			'ʹ��������ʽ��ȡonu��ź�mac��ַ
			re. = "epon-onu_(\d/\d+/\d+):.+([0-9a-f]{4}\.[0-9a-f]{4}\.[0-9a-f]{4})"
			If re.Test(strResults) <> True Then
				MsgBox "�쳣����"
				crt.quit
			Else
				Set matches = re.Execute(strResults)
				For Each match In matches
					epon_num = match.SubMatches(0)
					onu_mac = match.SubMatches(1)
				Next
			End If
			
			'����Ĵ���Ĺ�����ͨ����ȡ�������������� Pon �������һ��ONU�ı��
			objCurrentTab.Screen.Synchronous = True
			objCurrentTab.Screen.Send "show running-config interface epon-olt_" & epon_num & vbcr
			objCurrentTab.Screen.waitForString vbcr
			strResult = crt.Screen.ReadString(strPrompt)
			objCurrentTab.Screen.Synchronous = false
			
			'����Ĵ��빦���ǶԷ��صĽ�����з���
			strLines = Split(strResult, vbcrlf)
			'�����ǰONUΪ��PON���µĵ�һ���豸
			if UBound(strLines) - LBound(strLines) = 6 then
				last_num = 0
			'�����ȡ��ǰ��PON���µ����һ���豸�ı��
			else
				'��ȡ���һ�е���������
				lastIndex = UBound(strLines) - 3
				'strLines(lastIndex) = "onu 2 type ZTE-F400 mac c4a3.66c7.a8ae ip-cfg static"
				str1 = Split(strLines(lastIndex), " ")
				'str1(1) = "13"
				last_num = str1(1)
			end if
			
			'����ONU�����б�Ҫ��������ȡ�������������͵���һ��Sub��OK��
			'Ϊ��Ӧ��raisecom���͵�ONU���ߣ�ǿ��ONU������ΪZTE-F400
			if instr(onu_type, "ZTE") then
				config_ZTE_ONU epon_num, onu_type, onu_mac, last_num + 1, cvlan2, cvlan1
			else
				config_ZTE_ONU epon_num, "ZTE-F400", onu_mac, last_num + 1, cvlan2, cvlan1
			end if


			'д��־�Ǳ���İ�������Ե�רҵ!
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
		
		'���ִ������д��־�����ҳ�����������Զ��OLT
		If nError <> 0 Then
			Set logStream = g_fso.OpenTextFile(".\zte_onu_config_log.txt", 8, True)
			logStream.writeLine Now & ", An Error happened on ZTE-OLT: " & olt_name & "(" & olt_ip & ") . Error: " & strErr
			objCurrentTab.Session.Disconnect
			logStream.writeLine Now & ", The session for ZTE-OLT: " & olt_name & "(" & olt_ip & ") was disconnected. Trying reConnect..."
			logStream.close
			objCurrentTab.session.Connect()
			'�����û��������룬��ȷ���Ƿ����#��ģʽ
			inputPasswd olt_ip, username, passwd
		end if
		
		'����ʱ������Ϣһ��
		crt.sleep delayTime
		
	loop 
	
End Sub

'�ù��̵������ǣ�����OLT���û��������루�û�����������ݵ����Լ������, ��ȷ���Ƿ����#��ģʽ
Sub inputPasswd(olt_ip, username, passwd)

	
	Set objCurrentTab = crt.GetScriptTab
	objCurrentTab.Screen.Synchronous = True
	
	objCurrentTab.Screen.WaitForString "sername:"
	objCurrentTab.Screen.Send username & chr(13)
	objCurrentTab.Screen.WaitForString "assword:"
	objCurrentTab.Screen.Send passwd & chr(13)
	
	'�ж��Ƿ����#��ģʽ
	if objCurrentTab.Screen.WaitForString("#", 3) <> true then
		msgbox "û�н���#��ģʽ�������û��������������Ϣ������ִ��ʧ�ܣ�"
		crt.Quit
	end if
	
	'����־�ļ�, ���û�����½����ļ�
	Set logStream = g_fso.OpenTextFile(".\zte_onu_config_log.txt", 8, True)
	logStream.writeLine Now & ", The Script has been running at ZTE-OLT: " & olt_name & "(" & olt_ip & ")"
	logStream.close
	objCurrentTab.Screen.Synchronous = false
End Sub

'�ú����Ĺ����ǻ�ȡ���������������ʾ��
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


'�ù��̵Ĺ������ø����Ĳ�������ONU
'epon_num��Ϊָ��ONU���ڵ�PON�ڵ����
'onu_type��Ϊָ��ONU���ͺ�
'onu_mac�� Ϊָ��ONU��mac��ַ
'onu_num�� Ϊָ��ONU�����
'cvlan2��  Ϊ�����������vlan
'cvlan1:   Ϊ�㲥�����vlan
Sub config_ZTE_ONU(epon_num, onu_type, onu_mac, onu_num, cvlan2, cvlan1)
	
	'����ONU������Զ�������������͵㲥���ڲ�Vlan
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
	
	'��Ϊ����C320�·���������ԭ��������Ҫ�ӳ�
	'�����������̫�����ʾ��%Code 40796: This operation is forbidden, because down config is in process!

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



