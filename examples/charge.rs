use sap_scripting::*;

/// Demonstrate the start of the process of creating a charge
fn main() -> crate::Result<()> {
    let com_instance = SAPComInstance::new().expect("Couldn't get COM instance");
    let wrapper = com_instance.sap_wrapper().expect("Couldn't get SAP wrapper");
    let engine = wrapper.scripting_engine().expect("Couldn't get GuiApplication instance");

    let connection = match sap_scripting::GuiApplication_Impl::children(&engine)?.element_at(0)? {
        SAPComponent::GuiConnection(conn) => conn,
        _ => panic!("expected connection, but got something else!"),
    };
    eprintln!("Got connection");
    let session = match sap_scripting::GuiConnection_Impl::children(&connection)?.element_at(0)? {
        SAPComponent::GuiSession(session) => session,
        _ => panic!("expected session, but got something else!"),
    };

    if let SAPComponent::GuiMainWindow(wnd) = session.find_by_id("wnd[0]".to_owned())? {
        wnd.maximize().unwrap();
        session.start_transaction("fpe1".to_string())?;

        match session.find_by_id("wnd[0]/usr/ctxtFKKKO-BLART".to_string())? {
            SAPComponent::GuiCTextField(ctxt) => ctxt.set_text("P1".to_string())?,
            _ => panic!("expected doc type ctextfield")
        }
        match session.find_by_id("wnd[0]/usr/ctxtFKKKO-WAERS".to_string())? {
            SAPComponent::GuiCTextField(ctxt) => ctxt.set_text("GBP".to_string())?,
            _ => panic!("expected currency ctextfield")
        }
        match session.find_by_id("wnd[0]/usr/txtFKKKO-XBLNR".to_string())? {
            SAPComponent::GuiTextField(txt) => txt.set_text("XA12345678".to_string())?,
            _ => panic!("expected reference textfield")
        }
    } else {
        panic!("no window!");
    }

    Ok(())
}
