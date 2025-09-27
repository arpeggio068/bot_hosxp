#RequireAdmin
#Include <WinAPI.au3>
#include "OpenCV-Match_UDF_Mod.au3"
#include <WindowsConstants.au3>

Opt("MouseCoordMode", 1)
Opt("WinTitleMatchMode", 2)
Global $sHosXp_Title = "BMS-HOSxP XE 4.0"

Func TestToClick($picPath = @ScriptDir&"\Match\dad_name.png", $x1 = 0, $y1 = 370, $x2 = 51, $y2 = 400, $tol = 0.75)
 Local Const $kScreen = 1.25 ;constant for screen size 125% = 1.25  for screen size 100% = 1
 Local Const $kPosition = 5  ;constant for screen size 125% position. Set $kPosition = 0 if screen size = 100%
 Local $Match = _MatchPicture("", $picPath, $x1*$kScreen, $y1*$kScreen, $x2*$kScreen, $y2*$kScreen, $tol)
 $realX = 0
 $realY = 0
 If $kPosition > 0 Then
	$realX = $Match[4] - ($Match[4]/$kPosition)
    $realY = $Match[5] - ($Match[5]/$kPosition)
 Else
	$realX = $Match[4]
    $realY = $Match[5]
 EndIf
 ;_ArrayDisplay($Match1)
 If $Match[0] > 0 Then
	MouseMove($realX,$realY,10)
	;MouseClick($MOUSE_CLICK_LEFT, $realX, $realY, 1, 2)
 EndIf

EndFunc

Func FindToClick($picPath = @ScriptDir&"\Match\dad_name.png", $x1 = 0, $y1 = 370, $x2 = 51, $y2 = 400, $tol = 0.75)
 Local Const $kScreen = 1.25 ;constant for screen size 125% = 1.25  for screen size 100% = 1
 Local Const $kPosition = 5  ;constant for screen size 125% position. Set $kPosition = 0 if screen size = 100%
 Local $Match = _MatchPicture("", $picPath, $x1*$kScreen, $y1*$kScreen, $x2*$kScreen, $y2*$kScreen, $tol)
 $bClick = False
 $realX = 0
 $realY = 0
 If $kPosition > 0 Then
	$realX = $Match[4] - ($Match[4]/$kPosition)
    $realY = $Match[5] - ($Match[5]/$kPosition)
 Else
	$realX = $Match[4]
    $realY = $Match[5]
 EndIf
 ;_ArrayDisplay($Match1)
 If $Match[0] > 0 Then
	;MouseMove($realX,$realY,10)
	MouseClick($MOUSE_CLICK_LEFT, $realX, $realY, 1, 2)
	$bClick = True
	;Sleep(1000)
 EndIf
 Return $bClick
 EndFunc

 Func FindToClickRt($picPath = @ScriptDir&"\Match\dad_name.png", $x1 = 0, $y1 = 370, $x2 = 51, $y2 = 400, $tol = 0.75)
 Local Const $kScreen = 1.25 ;constant for screen size 125% = 1.25  for screen size 100% = 1
 Local Const $kPosition = 5  ;constant for screen size 125% position. Set $kPosition = 0 if screen size = 100%
 Local $Match = _MatchPicture("", $picPath, $x1*$kScreen, $y1*$kScreen, $x2*$kScreen, $y2*$kScreen, $tol)
 $realX = 0
 $realY = 0
 If $kPosition > 0 Then
	$realX = $Match[4] - ($Match[4]/$kPosition)
    $realY = $Match[5] - ($Match[5]/$kPosition)
 Else
	$realX = $Match[4]
    $realY = $Match[5]
 EndIf
 ;_ArrayDisplay($Match1)
 If $Match[0] > 0 Then
	;MouseMove($realX,$realY,10)
	MouseClick( $MOUSE_CLICK_RIGHT , $realX, $realY, 1, 2)
	;Sleep(1000)
 EndIf

EndFunc

Func FindToCon($picPath = @ScriptDir&"\Match\dad_name.png", $x1 = 0, $y1 = 370, $x2 = 51, $y2 = 400, $tol = 0.75)
 Local $kScreen = 1.25 ;constant for screen size 125% = 1.25  for screen size 100% = 1
 Local $Match = _MatchPicture("", $picPath, $x1*$kScreen, $y1*$kScreen, $x2*$kScreen, $y2*$kScreen, $tol)
 ;_ArrayDisplay($Match1)
 If $Match[0] > 0 Then
	 ;ConsoleWrite("result = match")
	 Return True
 Else
	 ;ConsoleWrite("result = not match")
	 Return False
 EndIf
EndFunc

Func SetHnBox($hWndXp)
  ; เลือกเฉพาะ TcxGroupBox Instance 1
    Local $hCtrl = ControlGetHandle($hWndXp, "", "[CLASS:TcxGroupBox; INSTANCE:1]") ;ด้านขวาของ list คนไข้ เส้นแนวตั้งสีฟ้าข้างๆ ไม่ใช่ตาราง grid
    Local $aPos = ControlGetPos($hWndXp, "", $hCtrl)
    ;MsgBox(0, "HN Box Pos", "x=" & $aPos[0] & ", y=" & $aPos[1] & ", w=" & $aPos[2] & ", h=" & $aPos[3])
    If $aPos[0] <> 0 Or $aPos[1] <> 125 Or $aPos[2] <> 400 Or $aPos[3] <> 676 Then
        ControlMove($hWndXp, "", $hCtrl, 0, 125, 400, 676)
		Sleep(1000)
    EndIf
EndFunc

Func ActHosXp1()
   Local $hWndXp = WinGetHandle($sHosXp_Title)
   Sleep(1000)
   WinActivate($hWndXp)
   Sleep(1000)
   ;SetHnBox($hWndXp)
   ;Sleep(1000)
   ;SetZeroPrice($hWndXp)
EndFunc

Func ActHosXp2()
    Local $hWndXp = WinGetHandle($sHosXp_Title)
    Sleep(200)
    WinSetState($hWndXp, "", @SW_RESTORE)
	Sleep(200)
	WinSetState($hWndXp, "", @SW_MAXIMIZE)
	Sleep(200)
    DllCall("user32.dll", "int", "SetForegroundWindow", "hwnd", $hWndXp)
    Sleep(500)
	;SetHnBox($hWndXp)
    ;Sleep(1000)
    ;SetZeroPrice($hWndXp)
EndFunc


;Global $sDtMenu = "\Match\xp_dental_menu.png"  ;ไม่ใช้
Global $sFoundVisit = "\Match\xp_hn_found_visit.png"   ;รูปชายหญิง
Global $aFvFirstPos = [0, 81,75,169] ;start loop position
Global $aFvPos = [267, 101, 470, 199]  ;end loop position
Global $sFiLock =  "\Match\xp_finance_lock.png"  ;ภาพกล่องข้อความ finance lock
Global $aFiLockPos = [589, 354, 955, 470]
;Global $sClosePt = "\Match\xp_close_pt.png"  ;ไม่ใช้
Global $sLoadPtSuccess2 = "\Match\xp_load_pt_success2.png" ;รูปฟันหน้าบนด้าน Li ล่างสุด
Global $aLoadPtPos = [641, 626,1118, 794]
;Global $sDentalCare =  "\Match\xp_dental_care.png" ;ไม่ใช้
Global $sF1 =   "\Match\xp_f1_click.png"  ;ปุ่มซักประวัติ
Global $aF1Pos = [ 6, 492,128, 557]
Global $sCC =   "\Match\xp_cc.png"  ;ข้อความ CC
Global $aCcPos = [253, 365, 436, 485]
;Global $sAddCC =   "\Match\xp_add_cc.png"  ;ไม่ใช้
Global $sRec_allergy_iden =   "\Match\rec_allergy_iden.png"  ;select box การแพ้ยา
Global $aRecAllergyPos = [313, 230, 547, 310]
Global $sF4 =   "\Match\xp_f4_click.png"  ;ปุ่มหัตถการ
Global $aF4Pos = [6, 526, 131, 596]
Global $sAddItem =   "\Match\xp_add_item.png"  ;ปุ่มเครืองหมายบวกเพิ่มหัตถการ
Global $aAddItemBtnPos = [170, 118, 546, 241]
Global $sPratomItem =   "\Match\xp_item_found_pratom.png"  ;ข้อความนักเรียน ประถม
Global $sFluorideItem =   "\Match\xp_item_found_fluoride.png"  ;ข้อความ ฟลูออไรด์วาร์นิช
Global $aItemTxtPos = [241, 252, 719, 461]  ;item txt position
Global $sSaveItem =   "\Match\xp_item_save.png"   ;ปุ่มติ๊กถูกในวงกลมสีเขียว สำหรับบันทึกหัตถการ
Global $aSaveItemBtnPos = [1061, 702, 1332, 807]  ; save item btn position
Global $sTask =   "\Match\xp_task_on_item.png"   ;เมนู Task สีเขียว
Global $aTaskPos = [325, 366, 629, 525]
Global $sOpdPrice1 =   "\Match\opd_price1.png"   ;ข้อความ ค่าบริการผู้ป่วยนอก พื้นสีฟ้า
Global $sOpdPrice2 =   "\Match\opd_price2.png"  ;ข้อความ ค่าบริการผู้ป่วยนอก ไม่มีพื้น
Global $aOpdPricePos = [354, 572, 950, 771]
Global $sEditPrice =   "\Match\xp_edit_price.png"   ; ข้อความ Inv Setting
Global $aEditPricePos = [209, 633, 1500, 812]
;Global $sFinalSave =   "\Match\xp_final_save1.png"  ;ไม่ใช้
Global $sFinalSave2 =   "\Match\xp_final_save2.png"  ;ภาพติ๊กถูกในวงกลมสีเขียวตอนบันทึกครั้งสุดท้าย
Global $aFinalS2Pos = [420, 275, 761, 499]
Global $sLastOK =   "\Match\xp_lastOK.png"
Global $aLastOkPos = [311, 18, 1068, 600]

Func SetZeroPrice($hWndXp)
		if FindToCon(@ScriptDir&$sOpdPrice1, 354, 572, 950, 771, 0.75) Then
			FindToClickRt(@ScriptDir&$sOpdPrice1,354, 572, 950, 771, 0.75)
			Sleep(1200)
		Else
			FindToClickRt(@ScriptDir&$sOpdPrice2, 354, 572, 950, 771, 0.75)
			Sleep(1200)
		EndIf
        Sleep(500)
     TestToClick(@ScriptDir&$sEditPrice,209, 633, 1500, 812, 0.75)   ; ข้อความ Inv Setting
	 Sleep(1000)
	 ControlSend($hWndXp, "","" , "q")
	 Sleep(1000)
EndFunc

_OpenCV_Startup()
ActHosXp2()
Sleep(1000)
;TestToClick(@ScriptDir&$sFoundVisit, $aFvPos[0], $aFvPos[1], $aFvPos[2], $aFvPos[3], 0.75)   ;รูปชายหญิงที่ตำแหน่งค่อนมาตรงกลาง
;TestToClick(@ScriptDir&$sFoundVisit, $aFvFirstPos[0], $aFvFirstPos[1], $aFvFirstPos[2], $aFvFirstPos[3], 0.8)  ;รูปชายหญิงที่อยู่ตำแหน่งซ้ายสุดของจอ
;TestToClick(@ScriptDir&$sFiLock, $aFiLockPos[0], $aFiLockPos[1], $aFiLockPos[2], $aFiLockPos[3], 0.75) ;ภาพกล่องข้อความ finance lock
;TestToClick(@ScriptDir&$sLoadPtSuccess2, $aLoadPtPos[0], $aLoadPtPos[1], $aLoadPtPos[2], $aLoadPtPos[3], 0.80) ;รูปฟันหน้าบนด้าน Li ล่างสุด
;TestToClick(@ScriptDir&$sF1, $aF1Pos[0], $aF1Pos[1], $aF1Pos[2], $aF1Pos[3], 0.80)  ;ปุ่มซักประวัติ
;TestToClick(@ScriptDir&$sCC, $aCcPos[0], $aCcPos[1], $aCcPos[2], $aCcPos[3], 0.80)  ;ข้อความ CC
;TestToClick(@ScriptDir&$sRec_allergy_iden, $aRecAllergyPos[0], $aRecAllergyPos[1], $aRecAllergyPos[2], $aRecAllergyPos[3], 0.75)  ;select box การแพ้ยา
;TestToClick(@ScriptDir&$sF4, $aF4Pos[0], $aF4Pos[1], $aF4Pos[2], $aF4Pos[3], 0.75)  ;ปุ่มหัตถการ
TestToClick(@ScriptDir&$sAddItem, $aAddItemBtnPos[0], $aAddItemBtnPos[1], $aAddItemBtnPos[2], $aAddItemBtnPos[3], 0.80)   ;ปุ่มเพิ่มหัตถการ
;TestToClick(@ScriptDir&$sPratomItem, $aItemTxtPos[0],  $aItemTxtPos[1],  $aItemTxtPos[2],  $aItemTxtPos[3], 0.75) ;ข้อความนักเรียน ประถม
;TestToClick(@ScriptDir&$sFluorideItem, $aItemTxtPos[0],  $aItemTxtPos[1],  $aItemTxtPos[2],  $aItemTxtPos[3], 0.75) ;ข้อความ ฟลูออไรด์วาร์นิช
;TestToClick(@ScriptDir&$sSaveItem, $aSaveItemBtnPos[0], $aSaveItemBtnPos[1], $aSaveItemBtnPos[2], $aSaveItemBtnPos[3], 0.75) ;ปุ่มติ๊กถูกในวงกลมสีเขียว สำหรับบันทึกหัตถการ
;TestToClick(@ScriptDir&$sTask, $aTaskPos[0], $aTaskPos[1], $aTaskPos[2], $aTaskPos[3], 0.75)  ;เมนู Task สีเขียว
;TestToClick(@ScriptDir&$sOpdPrice1, $aOpdPricePos[0], $aOpdPricePos[1], $aOpdPricePos[2], $aOpdPricePos[3], 0.75) ;ข้อความ ค่าบริการผู้ป่วยนอก พื้นสีฟ้า
;TestToClick(@ScriptDir&$sOpdPrice2, $aOpdPricePos[0], $aOpdPricePos[1], $aOpdPricePos[2], $aOpdPricePos[3], 0.75) ;ข้อความ ค่าบริการผู้ป่วยนอก ไม่มีพื้น
;TestToClick(@ScriptDir&$sEditPrice, $aEditPricePos[0], $aEditPricePos[1], $aEditPricePos[2], $aEditPricePos[3], 0.75)  ;ข้อความ Inv Setting  ,change
;TestToClick(@ScriptDir&$sFinalSave2, $aFinalS2Pos[0], $aFinalS2Pos[1], $aFinalS2Pos[2], $aFinalS2Pos[3], 0.75)  ;ภาพติ๊กถูกในวงกลมสีเขียวตอนบันทึกครั้งสุดท้าย
;TestToClick(@ScriptDir&$sLastOK, $aLastOkPos[0], $aLastOkPos[1], $aLastOkPos[2], $aLastOkPos[3], 0.75)


_OpenCV_Shutdown()
Exit(0)