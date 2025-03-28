/*
Origin VBS Script:

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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nfpl9"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtFKKL1-GPART").text = "12345"
session.findById("wnd[0]/usr/cmbFKKL1-LSTYP").key = "OPEN"
session.findById("wnd[0]/usr/cmbFKKL1-LSTYP").setFocus
session.findById("wnd[0]/tbar[0]/btn[0]").press

And how this would be written as Rust:
*/

use sap_scripting::*;

fn main() -> crate::Result<()> {
    // Initialise the environment.
    let com_instance = SAPComInstance::new().expect("Couldn't get COM instance");
    let wrapper = com_instance
        .sap_wrapper()
        .expect("Couldn't get SAP wrapper");
    let engine = wrapper
        .scripting_engine()
        .expect("Couldn't get GuiApplication instance");

    let connection: GuiConnection = sap_scripting::GuiApplicationExt::children(&engine)?
        .element_at(0)?
        .downcast()
        .expect("expected connection, but got something else!");
    eprintln!("Got connection");
    let session: GuiSession = sap_scripting::GuiConnectionExt::children(&connection)?
        .element_at(0)?
        .downcast()
        .expect("expected session, but got something else!");

    let wnd: GuiMainWindow = session
        .find_by_id("wnd[0]".to_owned())?
        .downcast()
        .expect("no window!");
    wnd.maximize().unwrap();

    let tbox_comp: GuiOkCodeField = session
        .find_by_id("wnd[0]/tbar[0]/okcd".to_owned())?
        .downcast()
        .expect("no ok code field!");
    tbox_comp.set_text("/nfpl9".to_owned()).unwrap();
    wnd.send_v_key(0).unwrap();

    Ok(())
}
