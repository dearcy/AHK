/*
-------------------------------------------------------
   Create short-cut keys for common command/tools
   Author: kang
   Version: 0.1
-------------------------------------------------------
*/
#SingleInstance force

#Include D:\Dropbox\Work\AutoHotKey\AutoHotkeyProgram\Lib\MI.ahk
#Include D:\Dropbox\Work\AutoHotKey\AutoHotkeyProgram\Lib\WinServ.ahk
; set this script startup automaticly when system start up
{
   SplitPath, A_Scriptname, , , , OutNameNoExt 
   LinkFile=%A_StartupCommon%\%OutNameNoExt%.lnk
   IfNotExist, %LinkFile% 
     FileCreateShortcut, %A_ScriptFullPath%, %LinkFile% 
   SetWorkingDir, %A_ScriptDir%
}


;Favorite menu
{
    ;¡ý Define main menu items ¡ý
    MenuList=
    (
      &Lunch        (win+w)       >          Run D:\Dropbox\Work\Lunch\Lunch.xlsm
      &Security     (win+s)       >          Run D:\Dropbox\Personal\security.xlsx
      &Calc         (win+c)       >          Run C:\windows\system32\calc.exe
      &Paint        (win+p)       >          Run C:\windows\system32\mspaint.exe
      &NotePad      (win+n)       >          Run notepad.exe
      &QQ           (win+q)       >          Run D:\Program Files (x86)\Tencent\QQ\QQProtect\Bin\QQProtect.exe
      &TTPlayer     (win+m)       >          Run D:\Program Files (x86)\TTPlayer\ttplayer.exe
      &To-Do                      >          Run D:\Dropbox\To-do.txt
      &gascfg.xml                 >          Run D:\ironman\config\service\gascfg.xml
      &Work                       >          ProjectDoc             > Run D:\Dropbox\Work\BrowserAgentDoc\Project Guidance.docx
      &Folder                     >          IronmanSrc             > Run F:\ironman\trunk\smallie\src\gas
      &Folder                     >          FF24Bin                > Run F:\ff24\trunk\gomez-build\ff-dist\dist\bin
      &Folder                     >          IEAgentSrc             > Run F:\ironman\trunk\smallie\src\app
      &Folder                     >          IronmanInstall         > Run D:\ironman    
      &Folder                     >          Dropbox                > Run D:\Dropbox
      &Folder                     >          issues                 > Run D:\Dropbox\Work\issue
      &Tools                      >          EventViewer            > Run C:\windows\system32\eventvwr.exe
      &Tools                      >          Services               > Run C:\windows\system32\services.msc
      &Tools                      >          RemoteComputer         > Run D:\Dropbox\Work\RemoteComputer.rtsx
      &Tools                      >          DebugViewer            > Run D:\Dropbox\Tools\SysinternalsSuite\Dbgview.exe
      &Tools                      >          GomezRecorder          > Run D:\Program Files (x86)\Gomez\Gomez Recorder\GomezRecorder.exe -csmode -debug
      &Tools                      >          WinSCP                 > Run D:\Dropbox\Tools\WinScp517\WinSCP.exe
    )
    ;¡ü Define main menu items ¡ü
      
    ;¡ý Add menu items from MenuList above ¡ý
    Loop, parse, MenuList, `n
    {
      StringSplit, MenuLevel,A_LoopField, >    
        Menulevel1 = %MenuLevel1%	;Remove leading and trailing space thus avoid duplicate menu
        Menulevel2 = %MenuLevel2%	;Remove leading and trailing space thus avoid duplicate menu
      if (Menulevel3 = "") {
        Menu, MyMenu, Add, %MenuLevel1%,Action        
        Command := varize(MenuLevel1)   ;Remove illegal characters        
        %Command% = %MenuLevel2%        
        ;MsgBox %MenuLevel2%
        if( InStr(MenuLevel2, ".exe") AND InStr(MenuLevel2, "Run")) {
         IconPath := trim(SubStr(MenuLevel2, InStr(MenuLevel2, "Run") + 3, InStr(MenuLevel2, ".exe") + 1 - InStr(MenuLevel2, "Run")))         
        } Else If( InStr(MenuLevel2, "Run"))
        {
           Temp := trim(SubStr(MenuLevel2, InStr(MenuLevel2, "Run") + 3))
           SplitPath, Temp,,,FileExtension
           if( FileExtension = "Lnk")
           FileGetShortcut, %Temp%, IconPath
           Else
           IconPath := AssocQueryApp(FileExtension)
        }
        ;MsgBox A_LoopField
        MI_SetMenuItemIcon("MyMenu", A_LoopField, IconPath, 1, 16)
      } Else if(MenuLevel4 = "" ){    
        Menu, %MenuLevel1%, Add, %MenuLevel2%, Action
        Menu, MyMenu, Add, %MenuLevel1%, :%MenuLevel1%    
        Command := varize(MenuLevel2)   ;Remove illegal characters
        %Command% = %MenuLevel3%
        if( InStr(MenuLevel3, ".exe") AND InStr(MenuLevel3, "Run") )
        IconPath := trim(SubStr(MenuLevel3, InStr(MenuLevel3, "Run") + 3, InStr(MenuLevel3, ".exe") + 1 - InStr(MenuLevel3, "Run")))
        Else If( InStr(MenuLevel3, "Run"))
        {
        Temp := trim(SubStr(MenuLevel3, InStr(MenuLevel3, "Run") + 3))
        SplitPath, Temp,,,FileExtension
        if( FileExtension = "Lnk")
        FileGetShortcut, %Temp%, IconPath
        Else
        IconPath := AssocQueryApp(FileExtension)
        }
        ;MsgBox %A_LoopField%
        ; set folder icon for parent menu item
        MI_SetMenuItemIcon("MyMenu", A_LoopField, "C:\Windows\System32\shell32.dll", 4, 16)
        if (IconPath = ""){
            ; set folder icon for no icon path item
            MI_SetMenuItemIcon(MenuLevel1, A_LoopField, "C:\Windows\System32\shell32.dll", 4, 16)
        } else {
            MI_SetMenuItemIcon(MenuLevel1, A_LoopField, IconPath, 1, 16)
        }
      }
      else
        {
        Menu, %MenuLevel2%, Add, %MenuLevel3%, Action
        Menu, %MenuLevel1%, Add, %MenuLevel2%, :%MenuLevel2%
        Menu, MyMenu, Add, %MenuLevel1%, :%MenuLevel1%
        Command := varize(MenuLevel3)   ;Remove illegal characters
        %Command% = %MenuLevel4%
        if( InStr(MenuLevel4, ".exe") AND InStr(MenuLevel4, "Run") )
        {
        IconPath := trim(SubStr(MenuLevel4, InStr(MenuLevel4, "Run") + 3, InStr(MenuLevel4, ".exe") + 1 - InStr(MenuLevel4, "Run")))
        MI_SetMenuItemIcon(MenuLevel2, A_Index, IconPath, 1, 16)
        }
        VarSetCapacity(MenuLevel4,0)	;Reset Menulevel4 to blank
        }
    }
    ;¡ü Process MenuList above ¡ü    
    ;~ f_Hotkey = #MButton
    ;~ Hotkey, %f_Hotkey%, DisplayMenu
    ;~ DisplayMenu:
    MButton:: Send {MButton}
    MButton & RButton::
    {
        CoordMode, Menu, Screen
        ;X := A_ScreenWidth/2.3
        ;Y := A_ScreenHeight/3.3
        CoordMode, Mouse , Screen
        MouseGetPos, X , Y
        Menu, MyMenu, Show, %X%, %Y%
        ;send {MButton}
        Return
     }
    ;exitapp
    Action:
       selected := varize(A_ThisMenuItem) ;Remove illegal characters
       exe = % %selected% ; exe is this menu next item
       if( InStr(exe, ".xlsx")) {      
         xlsx := Trim(SubStr(exe,4))      
         Run,  %xlsx%
         return
       }
       
       DynaRun("#notrayicon `n"exe)
    ;exitapp
    
    ;Function Library:
    AssocQueryApp(ext) {
        RegRead, type, HKCU, Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.%Ext%, Application
        If !ErrorLevel { ;Current user has overridden default setting
            RegRead, act, HKCU, Software\Classes\Applications\%type%\shell
            If ErrorLevel
                act = open
            RegRead, cmd, HKCU, Software\Classes\Applications\%type%\shell\%act%\command
            }
        Else {           ;Default setting
            RegRead, type, HKCR, .%Ext%
            RegRead, act , HKCR, %type%\shell
            If ErrorLevel
                act = open
        RegRead, cmd , HKCR, %type%\shell\%act%\command
        EXEPosition := InStr(cmd,".exe",false,0,1)
        exepath := Trim(SubStr(cmd,1,EXEPosition+3),"""")
    }
     Return, exepath
    }


    DynaRun(TempScript, pipename="")
    {      
      static _:="uint",@:="Ptr"
      If pipename =
      name := "AHK" A_TickCount
      Else
      name := pipename
      __PIPE_GA_ := DllCall("CreateNamedPipe","str","\\.\pipe\" name,_,2,_,0,_,255,_,0,_,0,@,0,@,0)
      __PIPE_ := DllCall("CreateNamedPipe","str","\\.\pipe\" name,_,2,_,0,_,255,_,0,_,0,@,0,@,0)
      if (__PIPE_=-1 or __PIPE_GA_=-1)
        Return 0
      Run, %A_AhkPath% "\\.\pipe\%name%",,UseErrorLevel HIDE, PID
      If ErrorLevel
      MsgBox, 262144, ERROR,% "Could not open file:`n" __AHK_EXE_ """\\.\pipe\" name """"
      DllCall("ConnectNamedPipe",@,__PIPE_GA_,@,0)
      DllCall("CloseHandle",@,__PIPE_GA_)
      DllCall("ConnectNamedPipe",@,__PIPE_,@,0)
      script := (A_IsUnicode ? chr(0xfeff) : (chr(239) . chr(187) . chr(191))) TempScript      
      if !DllCall("WriteFile",@,__PIPE_,"str",script,_,(StrLen(script)+1)*(A_IsUnicode ? 2 : 1),_ "*",0,@,0)
        Return A_LastError
      DllCall("CloseHandle",@,__PIPE_)
      Return PID
    }

    varize(var, autofix = true) {
      Loop, Parse, var
      {
        c = %A_LoopField%
        x := Asc(c)
        If (x > 47 and x < 58) or (x > 64 and x < 91) or (x > 96 and x < 123)
        or c = "#" or c = "_" or c = "@" or c = "$" or c = "?" or c = "[" or c = "]"
        IfEqual, autofix, 1, SetEnv, nv, %nv%%c%
        Else er++
      } If StrLen(var) > 254
      IfEqual, autofix, 1, StringLeft, var, var, 254
      Else er++
      IfEqual, autofix, 1, Return, nv
      Else Return, er
    }
    return
}

;System Shortcut
{
   #SingleInstance force
   ;set environment in windows( Win + PrintScrin )
   #PrintScreen::Run D:\Dropbox\Work\autoScripts\SetEnvironment.bat
   ;********************************************
   #w::Run D:\Dropbox\Work\Lunch\Lunch.xlsm                                                                     ;( Win + w )
   #c::Run calc.exe                                                                                             ;( Win + c )
   #p::Run mspaint.exe                                                                                          ;( Win + p ) 
   #n::Run notepad.exe                                                                                          ;( Win + n )
   #q::Run D:\Program Files (x86)\Tencent\QQ\QQProtect\Bin\QQProtect.exe                                        ;( Win + q )   
   #y::Run D:\Program Files\YouDao\Dict\YodaoDict.exe
   #m::Run D:\Program Files (x86)\TTPlayer\ttplayer.exe
   ; For ff24
   #b::Run F:\ff24\trunk\gomez-build\ff-dist\dist\bin
   #z::Run F:\ff24\trunk\gomez-build\ff-dist\dist\bin\storage\FF\log\agt-0-ga.log
   #a::Run C:\mozilla-build\start-msvc10.bat
   #s::Run D:\Dropbox\Personal\security.xlsx
   ;Volume Down
   !WheelDown::Send {Volume_Down 2}
   ;Volume Up
   !WheelUp::Send {Volume_Up 2}
   ;toggle mute
   !MButton::Send {Volume_Mute}
   !F1::
      Send {F4}
      Sleep, 500
      Send {F5}
      Sleep, 500
      Send {F6}
      return
   ;only use for window
   ; #IfWinActive ahk_class CabinetWClass
   {
      CapsLock & k::send {Up}
      CapsLock & j::send {Down}
      CapsLock & h::send {Left}
      CapsLock & l::send {Right}         
   }      
   ;opacity
   {
      ; changing window transparencies
      #WheelUp::  ; Increments transparency up by 3.375% (with wrap-around)
          DetectHiddenWindows, on
          WinGet, curtrans, Transparent, A
          if ! curtrans
              curtrans = 255
          newtrans := curtrans + 8
          if newtrans > 0
          {
              WinSet, Transparent, %newtrans%, A
          }
          else
          {
              WinSet, Transparent, OFF, A
              WinSet, Transparent, 255, A
          }
      return

      #WheelDown::  ; Increments transparency down by 3.375% (with wrap-around)
          DetectHiddenWindows, on
          WinGet, curtrans, Transparent, A
          if ! curtrans
              curtrans = 255
          newtrans := curtrans - 8
          if newtrans > 0
          {
              WinSet, Transparent, %newtrans%, A
          }
          ;else
          ;{
          ;    WinSet, Transparent, 255, A
          ;    WinSet, Transparent, OFF, A
          ;}
      return

      #o::  ; Reset Transparency Settings
          WinSet, Transparent, 255, A
          WinSet, Transparent, OFF, A
      return

      #g::  ; Press Win+G to show the current settings of the window under the mouse.
          MouseGetPos,,, MouseWin
          WinGet, Transparent, Transparent, ahk_id %MouseWin%
          ToolTip Translucency:`t%Transparent%`n
          Sleep 2000
          ToolTip
      return
   }
}


;Project Shortcut
{ 
   #SingleInstance force
   ;Run Script and close agent (~)
   SC029::
   { 
      ;get default script name
      Loop, HKEY_CLASSES_ROOT, *\shell\Set as Default script\DefaultScript, 1, 1
      {
         RegRead defaultScript
         SplitPath , defaultScript , scriptname
      }      
      ;MsgBox , 1 , , "run %defaultScript%'?"
      ;IfMsgBox OK
      RunWait , %comspec% /c nc localhost 8700 < "%defaultScript%" > "%defaultScript%.xml" , , hide
      ;Run, %comspec% , ,min
      ;RunWait , %comspec% /c nc localhost 8700 < "%defaultScript%" > "%defaultScript%.xml"
      Process , Exist , panzilla.exe
         Process , Close , panzilla.exe
      Process , Exist , GomezIEAgent.exe
         Process , Close , GomezIEAgent.exe
      return
   }
   #4::
   { 
      ;get default script name
      Loop, HKEY_CLASSES_ROOT, *\shell\Set as Default script\DefaultScript, 1, 1
      {
         RegRead defaultScript
         SplitPath , defaultScript , scriptname
      }      
      ;MsgBox , 1 , , "run %defaultScript%'?"
      ;IfMsgBox OK
      RunWait , %comspec% /c nc localhost 8800 < "%defaultScript%" > "%defaultScript%.xml" , , hide
      ;Run, %comspec% , ,min
      ;RunWait , %comspec% /c nc localhost 8700 < "%defaultScript%" > "%defaultScript%.xml"
      Process , Exist , panzilla.exe
         Process , Close , panzilla.exe
      Process , Exist , GomezIEAgent.exe
         Process , Close , GomezIEAgent.exe
      return
   }
   ^SC029::
   { 
      ;get default script name
      Loop, HKEY_CLASSES_ROOT, *\shell\Set as Default script\DefaultScript, 1, 1
      {
         RegRead defaultScript
         SplitPath , defaultScript , scriptname
      }      
      MsgBox , 1 , , "Send %defaultScript%'to 7878?"
      IfMsgBox OK
         RunWait , %comspec% /c nc localhost 7878 < "%defaultScript%" > "%defaultScript%_ironman.xml" , , hide
         MsgBox "%defaultScript% is finished"
      return
   }
   !SC029::
   {
      Loop, HKEY_CLASSES_ROOT, *\shell\Set as Default script\DefaultScript, 1, 1
      {
         RegRead defaultScript
         SplitPath , defaultScript , scriptname
      }      
      MsgBox , 4160 , Default Script , The default script is: %defaultScript%" , 3      
      return
   }
   
   !8::
   { 
      RunWait %comspec% /c nc localhost 8700 < D:\Dropbox\Work\issue\SimpleGSL\www.google.com.gsl > C:\Users\u124016\Desktop\google.xml  , , hide
      Process , Exist , panzilla.exe
         Process , Close , panzilla.exe
      Process , Exist , GomezIEAgent.exe
         Process , Close , GomezIEAgent.exe
      return
   }
   
   ;Delete old log file and Run Panzilla(Alt + 1)   
   !1::
   {
      MsgBox , 4129 , , "Run Panzilla ?" , 2
      IfMsgBox Cancel
         return
      EnvGet,PanzillWorkFolder,K_Panzilla_Bin
      ;MsgBox %PanzillWorkFolder%
      SetWorkingDir, %PanzillWorkFolder%\gomez-tools
      IfExist %PanzillWorkFolder%\storage\FF\log\agt-0-ga.log
         FileDelete %PanzillWorkFolder%\storage\FF\log\agt-0-ga.log
      RunWait %PanzillWorkFolder%\gomez-tools\panzilla-launch.cmd
      return
   }

   ;Delete log file and Run IE Agent
   !2::
   {
      MsgBox , 4129 , , "Run IEAgent ?" , 2
      IfMsgBox Cancel
         return
      EnvGet , IEWorkFolder , K_IE_Bin
      SetWorkingDir, %IEWorkFolder%\debug
      IfExist %IEWorkFolder%\debug\agent.log
         FileDelete %IEWorkFolder%\debug\agent.log
      RunWait GomezIEAgent.exe -vw 1 -p 8700 -cp 8701
      return      
   }
   
   ;Start mozilla build window
   !3::Run C:\mozilla-build\start-msvc10.bat
   
   ;Delete old log file and Run Panzilla 13(Alt + 1)   
   !4::
   {
      MsgBox , 4129 , , "Run Panzilla 13?" , 2
      IfMsgBox Cancel
         return
      EnvGet,PanzillWorkFolder,K_Panzilla_Bin
      ;MsgBox %PanzillWorkFolder%
      SetWorkingDir, F:\FxAgent\gomez-build\ff-dist\dist\bin\gomez-tools
      IfExist F:\FxAgent\gomez-build\ff-dist\dist\bin\storage\FF\log\agt-0-ga.log
         FileDelete F:\FxAgent\gomez-build\ff-dist\dist\bin\storage\FF\log\agt-0-ga.log
      RunWait F:\FxAgent\gomez-build\ff-dist\dist\bin\gomez-tools\panzilla-launch.cmd
      return
   }
}
#i::
{ 
   Run, %comspec% /c "taskkill /F /IM imsvc.exe /T"
   Run, %comspec% /c "taskkill /F /IM wintestxue.exe /T"
   Run, %comspec% /c "taskkill /F /IM wintestkang.exe /T"
   ;WinServ("Ironman", False) 
   return
}
;~ #r::
;~ {
   ;~ WinServ("Ironman", True) 
   ;~ return
;~ }
;Temp Shortcut
#t::
{ 
   
   ;~ Process , Exist , imsvc.exe
   ;~ If (ErrorLevel != 0) {
      ;~ Run, %comspec% /c "taskkill /F /IM imsvc.exe /T"
   ;~ }
   ;~ Process , Exist , "wintestxue.exe"
   ;~ NewPID = %ErrorLevel%
   ;~ If (NewPID = 0) {
      ;~ Run, %comspec% /c "taskkill /F /IM wintestxue.exe /T"
   ;~ }
   ;~ Process , Exist , "wintestkang.exe"
   ;~ If (ErrorLevel != 0) {
      ;~ Process , Close , "wintestkang.exe"
   ;~ }
   ;RunWait,sc stop "Ironman" ;Stop AdobeARM service.    
   FileDelete , D:\Ironman\service\ironman_service.log
   FileCopy , F:\ironman\trunk\smallie\src\gas\Release\imsvc.exe, D:\Ironman\service\Release, 1
   ;RunWait,sc start "Ironman" ;Stop AdobeARM service.
   
   ;WinServ("Ironman", True) 
   return
}
