; 坐标常量
_winLocalX := 0 * (A_ScreenDPI / 96) ;微信窗口X坐标
_winLocalY := 0 * (A_ScreenDPI / 96) ;微信窗口Y坐标
_winWidth := 880 * (A_ScreenDPI / 96) ;微信窗口宽
_winHeight := 700 * (A_ScreenDPI / 96) ;微信窗口高
_contactLocalX := 26 * (A_ScreenDPI / 96) ;微信联系人按钮X坐标
_contactLocalY := 145 * (A_ScreenDPI / 96) ;微信联系人按钮Y坐标
_contactDetailUpLeftX := 370 * (A_ScreenDPI / 96) ;微信联系人详情左上角X坐标
_contactDetailUpLeftY := 60 * (A_ScreenDPI / 96) ;微信联系人详情左上角Y坐标
_contactDetailLowRightX := 750 * (A_ScreenDPI / 96) ;微信联系人详情右下角X坐标
_contactDetailLowRightY := 440 * (A_ScreenDPI / 96) ;微信联系人详情右下角Y坐标
_contactRemarkX := 488 * (A_ScreenDPI / 96) ;微信联系人备注X坐标
_contactRemarkY := 240 * (A_ScreenDPI / 96) ;微信联系人备注X坐标
_contactListX := 150 * (A_ScreenDPI / 96) ;微信联系人列表X坐标
_contactListY := 300 * (A_ScreenDPI / 96) ;微信联系人列表Y坐标
_msExcelComObject := "Excel.Application" ;Ms Excel com object
_ahkWeChatAHKClassName := "ahk_class WeChatMainWndForPC" ;微信窗口名
_projectUrl := "https://github.com/XgHao/WeChat-Contact" ;项目地址

; 自定义异常-excel导出异常
class ExcelExportError extends Error
{}

; 自定义异常-微信窗口异常
class WeChatWinError extends Error
{}

; 自定义异常-用户终止操作
class UserTerminateOperationError extends Error
{}

; 设置优先级高，鼠标速度最快
ProcessSetPriority "High"
SetKeyDelay 0

;提示
msgbox("1.win+c开始导出`r`n2.win+ecs停止", "说明", "OK")

; ESC 退出
#esc::ExitApp

; 找到微信联系信息
#c::
{	
	try 
	{
		; 检测微信窗口是否打开
		CheckWeChatWin()

		; 获取微信联系人数
		contactCount := GetContactCount()

		; 定位微信窗口
		SetWeChatWin()

		; 创建文件路径
		path := CreateFilePath()

		; 保存到Excel
		SaveContactToXlsx(path, contactCount)
	}
	catch ExcelExportError as excelErr ;Excel导出异常
	{
		; Excel导出失败，将导出为csv
		SaveContactToCsv(path, contactCount)
	}
	catch WeChatWinError as weChatWinErr ;微信窗口异常
	{
		Return
	}
	catch UserTerminateOperationError as UserErr ;用户终止操作
	{
		Return
	}
	catch as e
	{
		; 错误信息
		errorMsg := Type(e) " thrown!`n`nwhat: " e.what "`nfile: " e.file
                . "`nline: " e.line "`nmessage: " e.message "`nextra: " e.extra

		; 反馈信息	
        Result := MsgBox("发生了错误，是否反馈？`n`n错误信息如下:`n" errorMsg, "出错了", "YesNo Icon!")
        if Result = "Yes"
            run _projectUrl "/issues/new"
		Return
	}

	; 打开文件路径
	Run "explore " path

	StarMyGitHub()
}

; 检测微信窗口是否打开
CheckWeChatWin()
{
	if !WinExist(_ahkWeChatAHKClassName)
	{
		MsgBox("未检测到微信窗口，请打开微信，按下【win+c】重新运行", "未找到微信", "OK Icon?")
		Throw WeChatWinError("Wechat Windows Is Not Active")
	}
}

; 定位微信窗口
SetWeChatWin()
{
	CheckWeChatWin()
	winactivate ; 激活微信窗口
	winmove _winLocalX,_winLocalY,_winWidth,_winHeight
	Sleep 100

	; 切换到联系人选项卡
	click _contactLocalX, _contactLocalY
	Sleep 100

	; 定位到最后一个微信好友
	click Format("{1} {2} Middle", _contactListX, _contactListY)
	Sleep 100
	send "{End}"
}

; 创建文件路径
CreateFilePath()
{
	path := StrReplace(Format("{1}\微信好友记录", A_WorkingDir), "\\", "\") ;路径
	if !DirExist(path) ;不存在该路径则创建
		DirCreate path
	Return path
}

;设置Excel标题
SetTitle(&objExcel)
{
	objExcel.Cells(1,1).Value := "昵称"
	objExcel.Cells(1,2).Value := "签名"
	objExcel.Cells(1,3).Value := "地区"
	objExcel.Cells(1,4).Value := "微信号"
	objExcel.Cells(1,5).Value := "来源"
	objExcel.Cells(1,6).Value := "备注"
}

;获取联系人详情
GetContactDetail()
{
	Loop 3	;循环3次，若3次还没有复制到内容视为失败
	{
		mouseclickdrag "L", _contactDetailLowRightX, _contactDetailLowRightY, _contactDetailUpLeftX, _contactDetailUpLeftY ;选中联系人信息详情
		A_Clipboard := "" ;清空剪切板
		Sleep 50
		sendinput "^c"	;复制
		if ClipWait(0.1, 1)	;超时抛出异常
			Break
	}
	
	contactMap := Map()	;新建Map对象，存放联系人信息
	wechatDetailArr := StrSplit(A_Clipboard, "`n") ;通过换行符进行分割

	len := wechatDetailArr.Length
	if (len = 5)
    {
        contactMap[1] := wechatDetailArr[1] ;昵称
        contactMap[2] := wechatDetailArr[2] ;签名
        contactMap[3] := wechatDetailArr[3] ;地区
        contactMap[4] := wechatDetailArr[4] ;微信ID
        contactMap[5] := wechatDetailArr[5] ;来源
    }
    else if (len = 4) ;没有签名的情况
    {
        contactMap[1] := wechatDetailArr[1]
        contactMap[2] := ""
        contactMap[3] := wechatDetailArr[2]
        contactMap[4] := wechatDetailArr[3]
        contactMap[5] := wechatDetailArr[4]
    }
	else ;其余数组长度都当做失败
	{
		MoveToNextContact()	;移动到下个联系人
		Throw Error("The copy text is error.")	;复制文字失败
	}

	A_Clipboard := "" ;清空剪切板
	click _contactRemarkX, _contactRemarkY
	sleep 20
	sendinput "^a^c" ;全选
	ClipWait(0.1, 1) ;等待0.1s超时视为无备注
	contactMap[6] := A_Clipboard ;备注
	MoveToNextContact()	;移动到下个联系人

	Return contactMap
}

;移动到下一个联系人
MoveToNextContact()
{
	click Format("{1} {2} Middle", _contactListX, _contactListY)
	send "{Up}"
	Sleep 50 ;缓存时间等待下个联系人加载
}

;输入联系人数
GetContactCount()
{
	contactCountValue := 0
	Loop
	{
		inputBoxResult := InputBox("请输入你的微信好友数，可以在联系人界面中通讯录管理中查看，`n`n初次运行请输入较小的值来测试结果`n`nTips: 在导出期间尽可能避免其他操作`n", "请输入读取联系人的次数")
		contactCountValue := inputBoxResult.Value
		result := inputBoxResult.Result
		if (result = "Cancel")
			Throw UserTerminateOperationError("GetContactCount User Cancel")	;终止操作
		Try
		{
			if (contactCountValue > 0)
				Break
			Else
				throw Error()
		}
		catch
		{
			MsgBox("无效的数字，请重新输入", "错误", "OK Icon!")
		}
	}

	Return contactCountValue
}

; 保存联系人到Xlsx
SaveContactToXlsx(path, contactCount)
{
	objExcel := ""
	Try 
	{
		objExcel := ComObject(_msExcelComObject) ;创建Excel-COM对象
		objExcel.Workbooks.Add ;创建工作簿
		SetTitle(&objExcel) ;设置Excel标题

		errorCount := 0 ;失败次数，当连续3次失败后，视为导出完成
		row := 2 ;第一行为标题，所以这里从第二行开始
		loop
		{
			try
			{
				for key, value in GetContactDetail()
					objExcel.Cells(row, key).Value := value
				row++ ;插入Excel成功，自增1
				errorCount := 0 ;成功后失败次数归零，重新计数
			}
			catch ; TODO 获取联系人信息失败
			{
				if	(++errorCount > 3) ;自增失败次数
					Break ;联系失败3次，视为导出成功
			}

			; 当行数大于微信联系人时，导出完毕跳出
			if (row - 1 > contactCount)
				Break
		}	
		
		fileName := Format("{1}\{2}.xlsx", path, A_Now) ;文件名
		objExcel.ActiveWorkbook.SaveAs(fileName) ;保存文件
	}
	catch as err
	{
		throw ExcelExportError(err.message)
	}
	Finally
	{
		try
			objExcel.Quit
	}
}

;保存联系人为Csv
SaveContactToCsv(path, contactCount)
{
	fileName := Format("{1}\{2}.csv", path, A_Now) ;文件名
	csvFile := FileOpen(fileName, "w")

	; 标题
	csvFile.WriteLine("昵称,签名,地区,微信号,来源,备注")
	errorCount := 0 ;失败次数，当连续3次失败后，视为导出完成
	loop
	{
		try
		{
			csvFile.WriteLine(FormatContactCsv(GetContactDetail()))
			contactCount-- ;写入成功，自减1
			errorCount := 0 ;成功后失败次数归零，重新计数
		}
		catch ; TODO 获取联系人信息失败
		{
			if	(++errorCount > 3) ;自增失败次数
				Break ;联系失败3次，视为导出成功
		}

		; 当行数大于微信联系人时，导出完毕跳出
		if (contactCount <= 0)
			Break
	}	
	
	csvFile.Close()
}

FormatContactCsv(map)
{
	return Format("{1},{2},{3},{4},{5}", map[1], map[2], map[3], map[4], map[5])
}

StarMyGitHub()
{
	Result := MsgBox("导出成功！可在打开的文件夹中查看`n`n制作不易，若对你有帮忙请赏一个Star⭐吧", "导出成功", "YesNo Iconi")
	if Result = "Yes"
		run _projectUrl
}