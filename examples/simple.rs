/*
Example VBS script:

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

*/

use sap_scripting::*;

fn main() -> Result<()> {
    // Initialise the environment.
    let com_instance = SAPComInstance::new()?;
    eprintln!("Got COM instance");
    let wrapper = com_instance.sap_wrapper()?;
    eprintln!("Got wrapper");
    let engine = wrapper.scripting_engine()?;
    eprintln!("Got scripting engine");
    let connection = engine.get_connection(0)?;
    eprintln!("Got connection");
    let session = connection.get_session()?;
    eprintln!("Got session");
    let tbox_comp = session.find_by_id("wnd[0]/tbar[0]/okcd")?;
    println!("Tbox comp type: {}", tbox_comp.kind()?);

    Ok(())
}