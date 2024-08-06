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
_ahkWeChatAHKClassName := "ahk_class WeChatMainWndForPC" ;微信窗口名
_projectUrl := "https://github.com/XgHao/WeChat-Contact" ;项目地址
;_msExcelComObject := "Excel.Application" ;Ms Excel com object

; 自定义异常-复制文字异常
class CopyError extends Error
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

		; 保存为Csv
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
	click _contactListX, _contactListY
	; click Format("{1} {2} Middle", _contactListX, _contactListY)
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

;输入联系人数
GetContactCount()
{
	contactCountValue := 0
	Loop
	{
		inputBoxResult := InputBox("请输入你的微信好友数，可以在联系人界面中通讯录管理中查看`n`n初次运行请输入较小的值来测试结果`n`nTips: 在导出期间尽可能避免其他操作`n`nTips: 由于企业微信用户的存在，实际的好友数要大于通讯录管理界面的好友数，所以你在输入时需要加上若干值，比如通讯录管理界面中显示有240个好友，那么你可以输入260`n", "请输入读取联系人的次数", "W400 H400")
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

;#region 新版本获取方式 [version >= 3.7.0]

;获取联系人详情
GetContactDetail()
{
	totalWaitTime := 2000 ; 总等待时间2s
	interval := 50 ; 每次等待50ms

	waitTime := 0

	Loop {
		A_Clipboard := "" ; 清空剪贴板
		Sleep 10 ; 缓冲时间
		sendinput "^a" ; 全选
		Sleep 20 ; 缓冲时间
		sendinput "^c" ; 复制

		if ClipWait(interval / 1000, 0) ; ClipWait参数为秒，所以Interval需要除以1000
		{
			break ; 如果剪贴板更新，则退出循环
		}

		waitTime += interval
		if (waitTime >= totalWaitTime) {
			; 移动到上一个
			send "{Up}"
			Sleep 50 ; 缓存时间等待下个联系人加载
			Throw CopyError("copy error")
		}
	}

	data := A_Clipboard
	contactMap := Map()	;新建Map对象，存放联系人信息
	wechatDetailArr := StrSplit(data, "`n") ;通过换行符进行分割
	len := wechatDetailArr.Length

	switch len {
		case 8: ;wechatDetailArr[7]是描述
		{
			contactMap[1] := wechatDetailArr[1] ;备注
			contactMap[2] := wechatDetailArr[2] ;昵称
			contactMap[3] := wechatDetailArr[3] ;微信ID
			contactMap[4] := wechatDetailArr[4] ;地区
			contactMap[5] := wechatDetailArr[6] ;标签
			contactMap[6] := wechatDetailArr[8] ;个性签名
		}
		case 7:
		{
			contactMap[1] := wechatDetailArr[1] ;备注
			contactMap[2] := wechatDetailArr[2] ;昵称
			contactMap[3] := wechatDetailArr[3] ;微信ID
			contactMap[4] := wechatDetailArr[4] ;地区
			contactMap[5] := wechatDetailArr[6] ;标签
			contactMap[6] := wechatDetailArr[7] ;个性签名
		}
		case 6:
		{
			; 有备注，但是没有地区信息，最后两个可以是【地区，标签，描述，个性签名】中任意两个，默认给最大可能性 【标签，个性签名】
			if (wechatDetailArr[1] = wechatDetailArr[4])
			{
				contactMap[1] := wechatDetailArr[1] ;备注
				contactMap[2] := wechatDetailArr[2] ;昵称
				contactMap[3] := wechatDetailArr[3] ;微信ID
				contactMap[4] := "-" ;地区
				contactMap[5] := wechatDetailArr[5] ;标签
				contactMap[6] := wechatDetailArr[6] ;个性签名
			}

			; 有备注，地区信息，最后两个可以是【地区，标签，描述，个性签名】中任意一个
			else if (wechatDetailArr[1] = wechatDetailArr[5])
			{
				contactMap[1] := wechatDetailArr[1] ;备注
				contactMap[2] := wechatDetailArr[2] ;昵称
				contactMap[3] := wechatDetailArr[3] ;微信ID
				contactMap[4] := wechatDetailArr[4] ;地区
				if (StrLen(wechatDetailArr[6]) >= 5)
				{
					contactMap[5] := "-" ;标签
					contactMap[6] := wechatDetailArr[6] ;个性签名
				}
				else
				{
					contactMap[5] := wechatDetailArr[6] ;标签
					contactMap[6] := "-" ;个性签名
				}
			}

			; 没有备注，wechatDetailArr[5]是描述
			else
			{
				contactMap[1] := "-" ;备注
				contactMap[2] := wechatDetailArr[1] ;昵称
				contactMap[3] := wechatDetailArr[2] ;微信ID
				contactMap[4] := wechatDetailArr[3] ;地区
				contactMap[5] := wechatDetailArr[4] ;标签
				contactMap[6] := wechatDetailArr[6] ;个性签名
			}
		}
		case 5:
		{
			; 有备注，但是没有地区信息，最后一个可能是标签或个性签名，这里根据长度判断
			if (wechatDetailArr[1] = wechatDetailArr[4])
			{
				contactMap[1] := wechatDetailArr[1] ;备注
				contactMap[2] := wechatDetailArr[2] ;昵称
				contactMap[3] := wechatDetailArr[3] ;微信ID
				contactMap[4] := "-" ;地区
				if (StrLen(wechatDetailArr[5]) >= 5)
				{
					contactMap[5] := "-" ;标签
					contactMap[6] := wechatDetailArr[5] ;个性签名
				}
				else
				{
					contactMap[5] := wechatDetailArr[5] ;标签
					contactMap[6] := "-" ;个性签名
				}
			}

			; 有备注，地区信息
			else if (wechatDetailArr[1] = wechatDetailArr[5])
			{
				contactMap[1] := wechatDetailArr[1] ;备注
				contactMap[2] := wechatDetailArr[2] ;昵称
				contactMap[3] := wechatDetailArr[3] ;微信ID
				contactMap[4] := wechatDetailArr[4] ;地区
				contactMap[5] := "-" ;标签
				contactMap[6] := "-" ;个性签名
			}

			; 没有备注，能确定的只有昵称ID，后续可以是【地区，标签，描述，个性签名】中任意三个，默认给最大可能性 【地区，标签，个性签名】
			else
			{
				contactMap[1] := "-" ;备注
				contactMap[2] := wechatDetailArr[1] ;昵称
				contactMap[3] := wechatDetailArr[2] ;微信ID
				contactMap[4] := wechatDetailArr[3] ;地区
				contactMap[5] := wechatDetailArr[4] ;标签
				contactMap[6] := wechatDetailArr[5] ;个性签名
			}
		}
		case 4:	; 备注 + 昵称微信ID  或者 昵称
		{
			; 有备注，但是没有地区信息
			if (wechatDetailArr[1] = wechatDetailArr[4])
			{
				contactMap[1] := wechatDetailArr[1] ;备注
				contactMap[2] := wechatDetailArr[2] ;昵称
				contactMap[3] := wechatDetailArr[3] ;微信ID
				contactMap[4] := "-" ;地区
				contactMap[5] := "-" ;标签
				contactMap[6] := "-" ;个性签名
			}

			; 无备注，只能确定昵称及ID，后续可以是【地区，标注，描述，个性签名】中任意连续两个，默认给最大可能性 【地区，个性签名】
			if (wechatDetailArr[1] != wechatDetailArr[4])
			{
				contactMap[1] := "-" ;备注
				contactMap[2] := wechatDetailArr[1] ;昵称
				contactMap[3] := wechatDetailArr[2] ;微信ID
				contactMap[4] := wechatDetailArr[3] ;地区
				contactMap[5] := "-" ;标签
				contactMap[6] := wechatDetailArr[4] ;个性签名
			}
		}
		case 3:	; 昵称，微信ID，地区（无法区分企业微信）
		{
			contactMap[1] := "-" ;备注
			contactMap[2] := wechatDetailArr[1] ;昵称
			contactMap[3] := wechatDetailArr[2] ;微信ID
			contactMap[4] := wechatDetailArr[3] ;地区
			contactMap[5] := "-" ;标签或描述
			contactMap[6] := "-" ;个性签名
		}
		case 2:	; 仅有昵称及微信ID
		{
			contactMap[1] := "-" ;备注
			contactMap[2] := wechatDetailArr[1] ;昵称
			contactMap[3] := wechatDetailArr[2] ;微信ID
			contactMap[4] := "-" ;地区
			contactMap[5] := "-" ;标签
			contactMap[6] := "-" ;个性签名
		}
	}

	; 元数据
	contactMap[7] := data

	; 移动到上一个
	send "{Up}"
	
	Sleep Random(20, 100) ;缓存时间等待下个联系人加载

	Return contactMap
}

;保存联系人为Csv
SaveContactToCsv(path, contactCount)
{
	fileName := Format("{1}\{2}.csv", path, A_Now) ;文件名
	csvFile := FileOpen(fileName, "w", "UTF-8")

	; 标题
	csvFile.WriteLine("备注,昵称,微信号,地区,标签,个性签名,元数据")
	errorCount := 0 ;失败次数，当连续3次失败后，视为导出完成
	loop
	{
		try
		{
			csvFile.WriteLine(FormatContactCsv(GetContactDetail()))
			contactCount-- ;写入成功，自减1
			errorCount := 0 ;成功后失败次数归零，重新计数
		}
		catch as e
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
	return Format("{1},{2},{3},{4},{5},{6},{7}", FormatCsvItem(map[1]), FormatCsvItem(map[2]), FormatCsvItem(map[3]), FormatCsvItem(map[4]), FormatCsvItem(map[5]), FormatCsvItem(map[6]), FormatCsvItem(StrReplace(map[7], "`n", "        ")))
}

FormatCsvItem(item)
{
	if (StrLen(item) = 1)
	{
		return item
	}

	if (SubStr(item,1,1) = "+" || SubStr(item,1,1) = "-" || SubStr(item,1,1) = "=" )
	{
		return "'" . item
	}
	
	return item
}

;#endregion

StarMyGitHub()
{
	Result := MsgBox("导出成功！可在打开的文件夹中查看`n`n制作不易，若对你有帮忙请赏一个Star⭐吧", "导出成功", "YesNo Iconi")
	if Result = "Yes"
		run _projectUrl
}