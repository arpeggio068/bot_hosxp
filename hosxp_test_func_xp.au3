#RequireAdmin
#include <Excel.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#Include <WinAPI.au3>
#Include <SendMessage.au3>
#include <FileConstants.au3>
#include <Array.au3>
#include <StringConstants.au3>
#include "OpenCV-Match_UDF_Mod.au3"
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Inet.au3>
#include <CppKeySend.au3>

Opt("MouseCoordMode", 1)
Opt("WinTitleMatchMode", 2)

Global $sHosXp_Title = "BMS-HOSxP XE 4.0"

Func SetHnBox($hWndXp)
    ; เลือกเฉพาะ TcxGroupBox Instance 3
    Local $hCtrl = ControlGetHandle($hWndXp, "", "[CLASS:TcxGroupBox; INSTANCE:1]")

    ; เอาตำแหน่งปัจจุบันออกมา
    Local $aPos = ControlGetPos($hWndXp, "", $hCtrl)
    ;MsgBox(0, "HN Box Pos", "x=" & $aPos[0] & ", y=" & $aPos[1] & ", w=" & $aPos[2] & ", h=" & $aPos[3])

    ; ถ้าไม่ตรงตำแหน่ง ให้ย้ายใหม่
    If $aPos[0] <> 0 Or $aPos[1] <> 125 Or $aPos[2] <> 400 Or $aPos[3] <> 676 Then
        ControlMove($hWndXp, "", $hCtrl, 0, 125, 400, 676)
    EndIf
EndFunc


Func ActHosXp()
    Local $hWndXp = WinGetHandle($sHosXp_Title)
    WinActivate($hWndXp)
    Sleep(500)
    SetHnBox($hWndXp)
EndFunc

ActHosXp()

Exit(0)