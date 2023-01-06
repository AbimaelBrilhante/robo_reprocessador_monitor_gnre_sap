If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "sp01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSTRIP_BL1/tabpSCR1/ssub%_SUBSCREEN_BL1:RSPOSP01NR:0100/ctxtS_RQCRED-LOW").text = "01012023"
session.findById("wnd[0]/usr/tabsTABSTRIP_BL1/tabpSCR1/ssub%_SUBSCREEN_BL1:RSPOSP01NR:0100/ctxtS_RQCRED-HIGH").text = "01012023"
session.findById("wnd[0]/usr/tabsTABSTRIP_BL1/tabpSCR1/ssub%_SUBSCREEN_BL1:RSPOSP01NR:0100/ctxtS_RQCRED-HIGH").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_BL1/tabpSCR1/ssub%_SUBSCREEN_BL1:RSPOSP01NR:0100/ctxtS_RQCRED-HIGH").caretPosition = 8
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/chk[1,3]").setFocus
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[0]/mbar/menu[0]/menu[2]/menu[1]").select
