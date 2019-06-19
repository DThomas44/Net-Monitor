/*
    Net Monitor
    Author: Daniel Thomas (Elesar/Dinenon)
    Date: 6/17/2019

    This script will monitor a specified network interface & log the avg and peak
    bandwidth use over a specified period of time.

    To Do:
        Complete actual logging implementation
        Implement selecting an interface in the settings GUI
*/
;<=====  System Settings  =====================================================>
#SingleInstance Force
#NoEnv
SetBatchLines, -1

;<=====  Settings  ============================================================>
settings := loadXML(A_ScriptDir . "\inc\settings.xml")
tdoc := loadXML(A_ScriptDir . "\inc\transform.xslt")

if (!settings.selectSingleNode("/netMonitor/settings/interface").text
    && !settings.selectSingleNode("/netMonitor/settings/autoInterface").text)
    settingsGUI(settings)

;<=====  Variables  ===========================================================>
downRate := 0
upRate := 0
maxDown := 0
maxUp := 0
totalDown := 0
totalUp := 0

;<=====  Timers  ==============================================================>
SetTimer, getNetData, 1000
SetTimer, logNetUse, % (settings.selectSingleNode("/netMonitor/settings/logInterval").text * 60000)
SetTimer, checkTime, 6000

;<=====  Setup NET Object  ====================================================>
if settings.selectSingleNode("/netMonitor/settings/autoInterface").text
{
    NET := new XNET(True)
}
else
{
    ;NET := new XNET(False)
    MsgBox, % "Manual interface selection not yet supported."
    ExitApp
}

;<=====  Start Log File  ======================================================>
logFile := LogStart(settings)
Return

;<=====  Functions  ===========================================================>
checkTime(){
    global
    if ((A_Hour == 00) && (A_Min == 00))
    {
        logStop(logFile)
        logFile := logStart(settings)
    }
}

getNetData(){
    Global
    ;Get current net stats
    NET.Update()
    downRate := NET.RxBPS
    upRate := NET.TxBPS

    ;Update max rates
    if (downRate > maxDown)
        maxDown := downRate
    if (upRate > maxUp)
        maxUp := upRate

    ;Update totals
    totalDown += downRate
    totalUp += upRate

    return
}

loadXML(file){
    xmlFile := fileOpen(file, "r")
    xml := xmlFile.read()
    xmlFile.Close()

    doc := ComObjCreate("MSXML2.DOMdocument.6.0")
    doc.async := false
    if !doc.loadXML(xml)
    {
        MsgBox, % "Could not load" . file . "!"
    }

    return doc
}

log(fileName, text){
    FormatTime, TimeStamp, A_Now, [dd/MMM/yyyy HH:mm:ss]
    logFile := fileOpen(fileName, "a")
    logFile.Write(TimeStamp . " " . text . "`r`n")
    logFile.Close()
    return 1
}

logNetUse(){
    Global
    FormatTime, TimeStamp, A_Now, [dd/MMM/yyyy HH:mm:ss]
    file := fileOpen(logFile, "a")
    file.Write(TimeStamp . " Max Down:" . (maxDown/125000) . "Mbps - Max Up:" . (maxUp/125000)
        . "Mbps - Total Down: " . sizeFormat(totalDown) . " - Total Up: " . sizeFormat(totalUp) . "`r`n")
    file.Close()

    ;Clear counters for next logging cycle
    maxDown := 0
    maxUp := 0

    return 1
}

logStart(settings){
    FormatTime, TimeStamp, A_Now, dd-MMM-yyyy_HH-mm-ss
    IfNotExist, % A_ScriptDir . "\Logs\"
        FileCreateDir, % A_ScriptDir . "\Logs"
    fileName := A_ScriptDir . "\Logs\" . TimeStamp . ".txt"
    Try {
        logFile := fileOpen(fileName, "w")
    }
    catch e {
        MsgBox, Failed to open file for logging!`n%A_LastError%
    }
    FormatTime, TimeStamp, A_Now, [dd/MMM/yyyy HH:mm:ss]
    logFile.Write(TimeStamp . " Logging started. Logging every "
        . settings.selectSingleNode("/netMonitor/settings/logInterval").text
        . " minutes using default interface."
        ;. settings.selectSingleNode("/netMonitor/settings/interface").text
        . "`r`n")
    logFile.Close()
    return fileName
}

logStop(fileName){
    FormatTime, TimeStamp, A_Now, [dd/MMM/yyyy HH:mm:ss]
    logFile := fileOpen(fileName, "a")
    logFile.Write(TimeStamp . " Logging finished.`r`n")
    logFile.Close()
    return 1
}

saveSettings(settings, tdoc){
    try {
        resultDoc := ComObjCreate("MSXML2.DOMdocument.6.0")
        settings.transformNodeToObject(tdoc, resultDoc)
        resultDoc.save(A_ScriptDir . "\inc\settings.xml")
        ObjRelease(Object(resultDoc))
    } catch e {
        MsgBox, % "Could not save settings.`n" . e . "`n" . e.description
        return 0
    }
    return 1
}

settingsCancel(){
    Global
    Gui, Settings:Hide
    Return
}

settingsGUI(settings){
    Global
    Gui, Settings:New
    Gui, Settings:Add, Text, x5 y5, Log Interval (Minutes):
    Gui, Settings:Add, Edit, x+5 yp-2 w50
    Gui, Settings:Add, UpDown, vLogInterval Range0-60, % settings.selectSingleNode("/netMonitor/settings/logInterval").text
    Gui, Settings:Add, CheckBox, x5 y+10 w200 vAutoInterface Disabled, Automatic interface selection
    GuiControl,, AutoInterface, % settings.selectSingleNode("/netMonitor/settings/autoInterface").text
    Gui, Settings:Add, Text, x5 y+10, Interface:
    Gui, Settings:Add, Edit, x+5 yp-2 vInterface Disabled, % settings.selectSingleNode("/netMonitor/settings/interface").text
    Gui, Settings:Add, Button, x5 y+10 w75 gSettingsCancel, Cancel
    Gui, Settings:Add, Button, x+10 yp w75 gSettingsOK, OK
    Gui, Show
    Return
}

settingsOK(){
    Global
    Gui, Settings:Submit
    node := settings.selectSingleNode("/netMonitor/settings/logInterval")
    node.text := LogInterval
    SaveSettings(settings, tdoc)
    Return
}

sizeFormat(bytes,round:=0){
    size:=0
    if (bytes >= 1000000000000)
    {     
        size :=Round(bytes/1000000000000,2) " TB"
    }
    else if (bytes >= 1000000000)
    {
        size :=Round(bytes/1000000000,2) " GB"
    }
    else if (bytes >= 1000000)
    {
        if (round)
        {
            size :=Round(bytes/1000000,0) " MB"
        }
        else   
            size :=Round(bytes/1000000,1) " MB"                 
    }
    else if (bytes >= 1000)
    {
        size :=Round(bytes/1000) " kB"
    }
    else if (!bytes || bytes == 0)
    {
        size :=0 " kB"
    }
    else
    {
            size := bytes " B"               
    } 
    return size
}

;<=====  Includes  ============================================================>
#Include %A_ScriptDir%\inc
#Include XNET.ahk
#Include Common Functions.ahk
