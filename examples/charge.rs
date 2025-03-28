use sap_scripting::*;

/// Demonstrate the start of the process of creating a charge
fn main() -> crate::Result<()> {
    let com_instance = SAPComInstance::new().expect("Couldn't get COM instance");
    let wrapper = com_instance
        .sap_wrapper()
        .expect("Couldn't get SAP wrapper");
    let engine = wrapper
        .scripting_engine()
        .expect("Couldn't get GuiApplication instance");

    let connection: GuiConnection = sap_scripting::GuiApplicationExt::children(&engine)?
        .element_at(0)?
        .cast()
        .expect("expected connection, but got something else!");
    eprintln!("Got connection");
    let session: GuiSession = sap_scripting::GuiConnectionExt::children(&connection)?
        .element_at(0)?
        .cast()
        .expect("expected session, but got something else!");

    let wnd: GuiMainWindow = session
        .find_by_id("wnd[0]".to_owned())?
        .cast()
        .expect("no window!");
    wnd.maximize().unwrap();
    session.start_transaction("fpe1".to_string())?;

    let ctxt: GuiCTextField = session
        .find_by_id("wnd[0]/usr/ctxtFKKKO-BLART".to_string())?
        .cast()
        .expect("expected doc type ctextfield");
    ctxt.set_text("P1".to_string())?;

    let ctxt: GuiCTextField = session
        .find_by_id("wnd[0]/usr/ctxtFKKKO-WAERS".to_string())?
        .cast()
        .expect("expected currency ctextfield");
    ctxt.set_text("GBP".to_string())?;

    let txt: GuiTextField = session
        .find_by_id("wnd[0]/usr/txtFKKKO-XBLNR".to_string())?
        .cast()
        .expect("expected reference textfield");
    txt.set_text("XA12345678".to_string())?;

    Ok(())
}
