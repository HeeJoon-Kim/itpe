#NoTrayIcon

scnt := 0
fcnt := 0

Gui, font,bold s12
Gui, Add, Text, y20 vSuccessVar, Success: %scnt%
Gui, Add, Text, x180 y20 vFailVar, Fail: %fcnt%
Gui, font
Gui, Add, Button, x110 y60 w80 h30 gSelectFile, 파일 열기
;Gui, Show ,x200 y200 w300 h100, License

Gui, Show, x200 y100 w300 h100, License

return

GuiClose:
	ExitApp
	return

SelectFile:
	LF := "`n" ; newline
	whr := ComObjCreate("WinHttp.WinHttpRequest.5.1") ; request object
	
	FileEncoding, UTF-8 ; set file encoding
	
	; select target file
	FileSelectFile, SelectedFile, 3, , Open a file, Text Documents (*.txt; *.csv)
	if (SelectedFile != "") 
	{
		url := "https://www.gov.kr/mw/NisCertificateConfirmExecute.do"
		SplitPath, SelectedFile,, SourceFilePath, SourceFileExt, SourceFileNoExt
		DestFile := SourceFilePath "\" SourceFileNoExt "_result." SourceFileExt
		
		if FileExist(DestFile)
		{
			MsgBox, 4,, 결과 파일이 존재합니다. 지우고 새로 만드시겠습니까?`n '아니로'를 선택하시면 내용이 추가 됩니다.`n`nFILE: %DestFile%
			IfMsgBox, Yes
				FileDelete, %DestFile%
		}
		
		; file read line by line
		Loop, read, %SelectedFile%, %DestFile%
		{
			output := A_LoopReadLine
			arr := StrSplit(A_LoopReadLine, ",")
			
			try
			{
				whr.open("POST", url, true)
				
				whr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8")
				whr.send("reqtInstCode=B490007&reqtUserName=" . arr[1] . "&ctftNo=" . arr[2])
				whr.WaitForResponse()
				data := whr.ResponseText()
			} catch e {
				MsgBox, % e.message
				return
			}
			
			checkResult := SubStr(data, -2, 1)
			
			If (checkResult == "Y") {
				scnt += 1
				GuiControl,, SuccessVar, Success: %scnt%
			} else {
				fcnt += 1
				GuiControl,, FailVar, Fail: %fcnt%
			}
			
			output := output . "," . checkResult . LF
			FileAppend, % output ; output file
		}
		
		MsgBox, , Success, 완료 되었습니다.
	}
		
		;MsgBox, The user selected the following:`n%SelectedFile%
	
	return

GetCheckResult(name, sn) {
	out := ""
	return out
}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; - - - - - - - - - - - - CreateFormData.ahk - - - - - - - - - - - - - - - - - - - - -
/*
	Version 6 feb 2020
	From .: https://gist.github.com/tmplinshi/8428a280bba58d25ef0b
	By .: tmplinshi
	CreateFormData - Creates "multipart/form-data" for http post
	https://www.autohotkey.com/boards/viewtopic.php?t=7647
	Usage: CreateFormData(ByRef retData, ByRef retHeader, objParam)
		retData   - (out) Data used for HTTP POST.
		retHeader - (out) Content-Type header used for HTTP POST.
		objParam  - (in)  An object defines the form parameters.
		            To specify files, use array as the value. Example:
		                objParam := { "key1": "value1"
		                            , "upload[]": ["1.png", "2.png"] }
	Requirements: BinArr.ahk -- https://gist.github.com/tmplinshi/a97d9a99b9aa5a65fd20
	Version    : 1.30 / 2019-01-13 - The file parameters are now placed at the end of the retData
	             1.20 / 2016-06-17 - Added CreateFormData_WinInet(), which can be used for VxE's HTTPRequest().
	             1.10 / 2015-06-23 - Fixed a bug
	             1.00 / 2015-05-14
*/

; Used for WinHttp.WinHttpRequest.5.1, Msxml2.XMLHTTP ...
CreateFormData(ByRef retData, ByRef retHeader, objParam) {
	New CreateFormData(retData, retHeader, objParam)
}

; Used for WinInet
CreateFormData_WinInet(ByRef retData, ByRef retHeader, objParam) {
	New CreateFormData(safeArr, retHeader, objParam)

	size := safeArr.MaxIndex() + 1
	VarSetCapacity(retData, size, 1)
	DllCall("oleaut32\SafeArrayAccessData", "ptr", ComObjValue(safeArr), "ptr*", pdata)
	DllCall("RtlMoveMemory", "ptr", &retData, "ptr", pdata, "ptr", size)
	DllCall("oleaut32\SafeArrayUnaccessData", "ptr", ComObjValue(safeArr))
}

Class CreateFormData {
	__New(ByRef retData, ByRef retHeader, objParam) {

		CRLF := "`r`n"

		Boundary := this.RandomBoundary()
		BoundaryLine := "------------------------------" . Boundary

		; Loop input paramters
		binArrs := []
		fileArrs := []
		For k, v in objParam
		{	
			If IsObject(v) {
				For i, FileName in v
				{
					str := BoundaryLine . CRLF
					     . "Content-Disposition: form-data; name=""" . k . """; filename=""" . FileName . """" . CRLF
					     . "Content-Type: " . this.MimeType(FileName) . CRLF . CRLF
					fileArrs.Push( BinArr_FromString(str) )
					fileArrs.Push( BinArr_FromFile(FileName) )
					fileArrs.Push( BinArr_FromString(CRLF) )
				}
			} Else {
				MsgBox, %k% %v%
				str := BoundaryLine . CRLF
				     . "Content-Disposition: form-data; name=""" . k """" . CRLF . CRLF
				     . v . CRLF
				binArrs.Push( BinArr_FromString(str) )
			}
		}

		binArrs.push( fileArrs* )

		str := BoundaryLine . "--" . CRLF
		binArrs.Push( BinArr_FromString(str) )

		retData := BinArr_Join(binArrs*)
		retHeader := "multipart/form-data; boundary=----------------------------" . Boundary
	}

	RandomBoundary() {
		str := "0|1|2|3|4|5|6|7|8|9|a|b|c|d|e|f|g|h|i|j|k|l|m|n|o|p|q|r|s|t|u|v|w|x|y|z"
		Sort, str, D| Random
		str := StrReplace(str, "|")
		Return SubStr(str, 1, 12)
	}

	MimeType(FileName) {
		n := FileOpen(FileName, "r").ReadUInt()
		Return (n        = 0x474E5089) ? "image/png"
		     : (n        = 0x38464947) ? "image/gif"
		     : (n&0xFFFF = 0x4D42    ) ? "image/bmp"
		     : (n&0xFFFF = 0xD8FF    ) ? "image/jpeg"
		     : (n&0xFFFF = 0x4949    ) ? "image/tiff"
		     : (n&0xFFFF = 0x4D4D    ) ? "image/tiff"
		     : "application/octet-stream"
	}

}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; - - - - - - - - - - - - BinArr.ahk  - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Update: 2015-6-4 - Added BinArr_ToFile()
; https://gist.github.com/tmplinshi/a97d9a99b9aa5a65fd20
; By tmplinshi
BinArr_FromString(str) {
	oADO := ComObjCreate("ADODB.Stream")

	oADO.Type := 2 ; adTypeText
	oADO.Mode := 3 ; adModeReadWrite
	oADO.Open
	oADO.Charset := "UTF-8"
	oADO.WriteText(str)

	oADO.Position := 0
	oADO.Type := 1 ; adTypeBinary
	oADO.Position := 3 ; Skip UTF-8 BOM
	return oADO.Read, oADO.Close
}

BinArr_FromFile(FileName) {
	oADO := ComObjCreate("ADODB.Stream")

	oADO.Type := 1 ; adTypeBinary
	oADO.Open
	oADO.LoadFromFile(FileName)
	return oADO.Read, oADO.Close
}

BinArr_Join(Arrays*) {
	oADO := ComObjCreate("ADODB.Stream")

	oADO.Type := 1 ; adTypeBinary
	oADO.Mode := 3 ; adModeReadWrite
	oADO.Open
	For i, arr in Arrays
		oADO.Write(arr)
	oADO.Position := 0
	return oADO.Read, oADO.Close
}

BinArr_ToString(BinArr, Encoding := "UTF-8") {
	oADO := ComObjCreate("ADODB.Stream")

	oADO.Type := 1 ; adTypeBinary
	oADO.Mode := 3 ; adModeReadWrite
	oADO.Open
	oADO.Write(BinArr)

	oADO.Position := 0
	oADO.Type := 2 ; adTypeText
	oADO.Charset  := Encoding 
	return oADO.ReadText, oADO.Close
}

BinArr_ToFile(BinArr, FileName) {
	oADO := ComObjCreate("ADODB.Stream")

	oADO.Type := 1 ; adTypeBinary
	oADO.Open
	oADO.Write(BinArr)
	oADO.SaveToFile(FileName, 2)
	oADO.Close
}