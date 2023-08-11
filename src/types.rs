use windows::{core::*, Win32::System::Com::*};

use crate::idispatch_ext::IDispatchExt;
use crate::traits::*;
use crate::variant_ext::VariantExt;

macro_rules! type_struct {
    ($(#[$attr:meta])* => $name: ident: $($trait: ident,)*) => {
        $(#[$attr])*
        pub struct $name {
            pub(crate) inner: IDispatch,
        }

        impl HasDispatch for $name {
            fn get_idispatch(&self) -> &IDispatch {
                &self.inner
            }
        }
        $(impl $trait for $name {})*
    };
}

macro_rules! forward_func {
    ($(#[$attr:meta])* => $snake_name: ident, $name: expr) => {
        $(#[$attr])*
        pub fn $snake_name(&self) -> crate::Result<()> {
            let _ = self.inner.call($name, vec![])?;
            Ok(())
        }
    };
}

macro_rules! forward_func_1_arg {
    ($(#[$attr:meta])* => $snake_name: ident, $name: expr, $arg_name: ident, $arg_ty: ty, $arg_transformer: ident) => {
        $(#[$attr])*
        pub fn $snake_name(&self, $arg_name: $arg_ty) -> crate::Result<()> {
            let _ = self.inner.call($name, vec![VARIANT::$arg_transformer($arg_name)])?;
            Ok(())
        }
    };
}

macro_rules! get_property {
    ($(#[$attr:meta])* => $snake_name: ident, $name: expr, $kind: ty, $transformer: ident) => {
        $(#[$attr])*
        pub fn $snake_name(&self) -> crate::Result<$kind> {
            Ok(self.inner.get($name)?.$transformer()?)
        }
    };
}

macro_rules! set_property {
    ($(#[$attr:meta])* => $snake_name: ident, $name: expr, $kind: ty, $transformer: ident) => {
        $(#[$attr])*
        pub fn $snake_name(&self, value: $kind) -> crate::Result<()> {
            let _ = self.inner.set($name, VARIANT::$transformer(value))?;
            Ok(())
        }
    };
}

/// A wrapper over the SAP scripting engine, equivalent to CSapROTWrapper.
pub struct SAPWrapper {
    inner: IDispatch,
}

impl SAPWrapper {
    pub(crate) fn new() -> crate::Result<Self> {
        unsafe {
            let clsid: GUID = CLSIDFromProgID(w!("SapROTWr.SapROTWrapper"))?;
            let p_clsid: *const GUID = &clsid;
            log::debug!("CSapROTWrapper CLSID: {:?}", clsid);

            let dispatch: IDispatch =
                CoCreateInstance(p_clsid, None, CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER)?;
            Ok(SAPWrapper { inner: dispatch })
        }
    }

    /// Get the Scripting Engine object from this wrapper.
    pub fn scripting_engine(&self) -> crate::Result<GuiApplication> {
        log::debug!("Getting UI ROT entry...");
        let result = self
            .inner
            .call("GetROTEntry", vec![VARIANT::from_str("SAPGUI")])?;

        let sap_gui = result.to_idispatch()?;

        log::debug!("Getting scripting engine.");
        let scripting_engine = sap_gui.call("GetScriptingEngine", vec![])?;

        Ok(GuiApplication {
            inner: scripting_engine.to_idispatch()?.clone(),
        })
    }
}

/// SAPComponent wraps around all of the GuiComponent types, allowing a more Rust-y way of exporing components.
pub enum SAPComponent {
    /// The GuiAbapEditor object represents the new ABAP editor control available as of SAP_BASIS release
    /// 6.20 (see also SAP Note 930742). GuiAbapEditor extends GuiShell.
    GuiAbapEditor(GuiVComponent), //(GuiAbapEditor),
    /// The GuiApoGrid is an object specifically created for SAP SCM applications. It implements a planning
    /// board, which is similar to a GuiGridView control. GuiApoGrid extends GuiShell.
    GuiApoGrid(GuiVComponent), //(GuiApoGrid),
    /// The GuiApplication represents the process in which all SAP GUI activity takes place. If the scripting
    /// component is accessed by attaching to an SAP Logon process, then GuiApplication will represent SAP
    /// Logon. GuiApplication is a creatable class. However, there must be only one component of this type
    /// in any process. GuiApplication extends the GuiContainer Object.
    GuiApplication(GuiApplication),
    /// The GuiBarChart is a powerful tool to display and modify time scale diagrams.
    GuiBarChart(GuiVComponent), //(GuiBarChart),
    /// A GuiBox is a simple frame with a name (also called a "Group Box"). The items inside the frame are not
    /// children of the box. The type prefix is "box".
    GuiBox(GuiVComponent), //(GuiBox),
    /// GuiButton represents all push buttons that are on dynpros, the toolbar or in table controls. GuiButton
    /// extends the GuiVComponent Object. The type prefix is btn, the name property is the fieldname taken
    /// from the SAP data dictionary There is one exception: For tabstrip buttons, it is the button id set in
    /// screen painter that is taken from the SAP data dictionary.
    GuiButton(GuiButton),
    /// The calendar control can be used to select single dates or periods of time. GuiCalendar extends the
    /// GuiShell Object.
    GuiCalendar(GuiVComponent), //(GuiCalendar),
    /// The GuiChart object is of a very technical nature. It should only be used for recording and playback, as
    /// most of the parameters cannot be determined in any other way.
    GuiChart(GuiVComponent), //(GuiChart),
    /// GuiCheckBox extends the GuiVComponent Object. The type prefix is chk, the name is the fieldname taken
    /// from the SAP data dictionary.
    GuiCheckBox(GuiCheckBox),
    /// GuiColorSelector displays a set of colors for selection. It extends the GuiShell Object.
    GuiColorSelector(GuiVComponent), //(GuiColorSelector),
    /// The GuiComboBox looks somewhat similar to GuiCTextField, but has a completely different implementation.
    /// While pressing the combo box button of a GuiCTextField will open a new dynpro or control in which a
    /// selection can be made, GuiComboBox retrieves all possible choices on initialization from the server, so
    /// the selection is done solely on the client. GuiComboBox extends the GuiVComponent Object. The type prefix
    /// is cmb, the name is the fieldname taken from the SAP data dictionary. GuiComboBox inherits from the
    /// GuiVComponent Object.
    GuiComboBox(GuiComboBox),
    ///
    GuiComboBoxControl(GuiVComponent), //(GuiComboBoxControl),
    /// Members of the Entries collection of a GuiComboBox are of type GuiComBoxEntry.
    GuiComboBoxEntry(GuiVComponent), //(GuiComboBoxEntry),
    /// GuiComponent is the base class for most classes in the Scripting API. It was designed to allow generic
    /// programming, meaning you can work with objects without knowing their exact type.
    GuiComponent(GuiComponent),
    /// This interface resembles GuiVContainer. The only difference is that it is not intended for visual objects
    /// but rather administrative objects such as connections or sessions. Objects exposing this interface will
    /// therefore support GuiComponent but not GuiVComponent. GuiContainer extends the GuiComponent Object.
    GuiContainer(GuiVComponent), //(GuiContainer),
    /// A GuiContainerShell is a wrapper for a set of the GuiShell Object. GuiContainerShell extends the GuiVContainer
    /// Object. The type prefix is shellcont, the name is the last part of the id, shellcont[n].
    GuiContainerShell(GuiVComponent), //(GuiContainerShell),
    /// If the cursor is set into a text field of type GuiCTextField a combo box button is displayed to the right of
    /// the text field. Pressing this button is equivalent to pressing the F4 key. The button is not represented in
    /// the scripting object model as a separate object; it is considered to be part of the text field.
    ///
    /// There are no other differences between GuiTextField and GuiCTextField. GuiCTextField extends the GuiTextField.
    /// The type prefix is ctxt, the name is the Fieldname taken from the SAP data dictionary.
    GuiCTextField(GuiCTextField),
    /// The GuiCustomControl is a wrapper object that is used to place ActiveX controls onto dynpro screens. While
    /// GuiCustomControl is a dynpro element itself, its children are of GuiContainerShell type, which is a container
    /// for controls. GuiCustomControl extends the GuiVContainer Object. The type prefix is cntl, the name is the
    /// fieldname taken from the SAP data dictionary.
    GuiCustomControl(GuiVComponent), //(GuiCustomControl),
    /// The GuiDialogShell is an external window that is used as a container for other shells, for example a toolbar.
    /// GuiDialogShell extends the GuiVContainer Object. The type prefix is shellcont, the name is the last part of
    /// the id, shellcont[n].
    GuiDialogShell(GuiVComponent), //(GuiDialogShell),
    /// The GuiEAIViewer2D control is used to view 2-dimensional graphic images in the SAP system. The user can carry
    /// out redlining over the loaded image. The scripting wrapper for this control records all user actions during
    /// the redlining process and reproduces the same actions when the recorded script is replayed.
    GuiEAIViewer2D(GuiVComponent), //(GuiEAIViewer2D),
    /// The GuiEAIViewer3D control is used to view 3-dimensional graphic images in the SAP system.
    GuiEAIViewer3D(GuiVComponent), //(GuiEAIViewer3D),
    /// A GuiFrameWindow is a high level visual object in the runtime hierarchy. It can be either the main window or
    /// a modal popup window. See the GuiMainWindow and GuiModalWindow sections for examples. GuiFrameWindow itself
    /// is an abstract interface. GuiFrameWindow extends the GuiVContainer Object. The type prefix is wnd, the name
    /// is wnd plus the window number in square brackets.
    GuiFrameWindow(GuiFrameWindow), //(GuiFrameWindow),
    /// The GuiGosShell is only available in New Visual Design mode. GuiGOSShell extends the GuiVContainer Object.
    /// The type prefix is shellcont, the name is the last part of the id, shellcont[n].
    GuiGOSShell(GuiVComponent), //(GuiGOSShell),
    /// For the graphic adapter control only basic members from GuiShell are available. Recording and playback is
    /// not possible.
    GuiGraphAdapt(GuiVComponent), //(GuiGraphAdapt),
    /// The grid view is similar to the dynpro table control, but significantly more powerful. GuiGridView extends
    /// the GuiShell Object.
    GuiGridView(GuiVComponent), //(GuiGridView),
    /// The GuiHTMLViewer is used to display an HTML document inside SAP GUI. GuiHTMLViewer extends the GuiShell
    /// Object.
    GuiHTMLViewer(GuiVComponent), //(GuiHTMLViewer),
    ///
    GuiInputFieldControl(GuiVComponent), //(GuiInputFieldControl),
    /// GuiLabel extends the GuiVComponent Object. The type prefix is lbl, the name is the fieldname taken from the
    /// SAP data dictionary.
    GuiLabel(GuiVComponent), //(GuiLabel),
    /// This window represents the main window of an SAP GUI session.
    GuiMainWindow(GuiMainWindow), //(GuiMainWindow),
    /// For the map control only basic members from GuiShell are available. Recording and playback is not possible.
    GuiMap(GuiVComponent), //(GuiMap),
    /// A GuiMenu may have other GuiMenu objects as children. GuiMenu extends the GuiVContainer Object. The type prefix
    /// is menu, the name is the text of the menu item. If the item does not have a text, which is the case for
    /// separators, then the name is the last part of the id, menu[n].
    GuiMenu(GuiMenu),
    /// Only the main window has a menubar. The children of the menubar are menus. GuiMenubar extends the GuiVContainer
    /// Object. The type prefix and name are mbar.
    GuiMenubar(GuiMenubar),
    /// A GuiModalWindow is a dialog pop-up.
    GuiModalWindow(GuiVComponent), //(GuiModalWindow),
    /// The GuiNetChart is a powerful tool to display and modify entity relationship diagrams. It is of a very technical
    /// nature and should only be used for recording and playback, as most of the parameters cannot be determined in
    /// any other way.
    GuiNetChart(GuiVComponent), //(GuiNetChart),
    ///
    GuiOfficeIntegration(GuiVComponent), //(GuiOfficeIntegration),
    /// The GuiOkCodeField is placed on the upper toolbar of the main window. It is a combo box into which commands can
    /// be entered. Setting the text of GuiOkCodeField will not execute the command until server communication is
    /// started, for example by emulating the Enter key (VKey 0). GuiOkCodeField extends the GuiVComponent Object. The
    /// type prefix is okcd, the name is empty.
    GuiOkCodeField(GuiOkCodeField),
    /// There are some differences between GuiTextField and GuiPasswordField:
    ///
    /// - The Text and DisplayedText properties cannot be read for a password field. The returned text is always empty.
    /// During recording the password is also not saved in the recorded script.
    /// - The properties HistoryCurEntry, HistoryCurIndex, HistoryIsActive and HistoryList are not supported, because
    /// password fields do not offer an input history
    /// - The property IsListElement is not supported, because password fields cannot be placed on ABAP lists
    GuiPasswordField(GuiVComponent), //(GuiPasswordField),
    /// The picture control displays a picture on an SAP GUI screen. GuiPicture extends the GuiShell Object.
    GuiPicture(GuiVComponent), //(GuiPicture),
    /// GuiRadioButton extends the GuiVComponent Object. The type prefix is rad, the name is the fieldname taken from the
    /// SAP data dictionary.
    GuiRadioButton(GuiRadioButton),
    /// For the SAP chart control only basic members from GuiShell are available. Recording and playback is not possible.
    GuiSapChart(GuiVComponent), //(GuiSapChart),
    /// The GuiScrollbar class is a utility class used for example in GuiScrollContainer or GuiTableControl.
    GuiScrollbar(GuiVComponent), //(GuiScrollbar),
    /// This container represents scrollable subscreens. A subscreen may be scrollable without actually having a scrollbar,
    /// because the existence of a scrollbar depends on the amount of data displayed and the size of the GuiUserArea.
    /// GuiScrollContainer extend sthe GuiVContainer Object. The type prefix is ssub, the name is generated from the data
    /// dictionary settings.
    GuiScrollContainer(GuiVComponent), //(GuiScrollContainer),
    /// GuiShell is an abstract object whose interface is supported by all the controls. GuiShell extends the GuiVContainer
    /// Object. The type prefix is shell, the name is the last part of the id, shell[n].
    GuiShell(GuiVComponent), //(GuiShell),
    /// This container represents non-scrollable subscreens. It does not have any functionality apart from to the inherited
    /// interfaces. GuiSimpleContainer extends the GuiVContainer Object. The type prefix is sub, the name is is generated
    /// from the data dictionary settings.
    GuiSimpleContainer(GuiVComponent), //(GuiSimpleContainer),
    /// GuiSplit extends the GuiShell Object.
    GuiSplit(GuiVComponent), //(GuiSplit),
    /// The GuiSplitterContainer represents the dynpro splitter element, which was introduced in the Web Application Server
    /// ABAP in NetWeaver 7.1. The dynpro splitter element is similar to the activeX based splitter control, but it is a
    /// plain dynpro element.
    GuiSplitterContainer(GuiVComponent), //(GuiSplitterContainer),
    /// For the stage control only basic members from GuiShell are available. Recording and playback is not possible.
    GuiStage(GuiVComponent), //(GuiStage),
    /// GuiStatusbar represents the message displaying part of the status bar on the bottom of the SAP GUI window. It does
    /// not include the system and login information displayed in the rightmost area of the status bar as these are available
    /// from the GuiSessionInfo object. GuiStatusbar extends the GuiVComponent Object. The type prefix is sbar.
    GuiStatusbar(GuiVComponent), //(GuiStatusbar),
    /// The parent of the GuiStatusPane objects is the status bar (see also GuiStatusbar Object). The GuiStatusPane objects
    /// reflect the individual areas of the status bar, for example "pane[0]" refers to the section of the status bar where
    /// the messages are displayed. See also GuiStatusbar Object. The first pane of the GuiStatusBar (pane[0]) can have a
    /// child of type GuiStatusBarLink, if a service request link is displayed.
    GuiStatusPane(GuiVComponent), //(GuiStatusPane),
    /// The GuiTab objects are the children of a GuiTabStrip object. GuiTab extends the GuiVContainer Object. The type prefix
    /// is tabp, the name is the id of the tab’s button taken from SAP data dictionary.
    GuiTab(GuiVComponent), //(GuiTab),
    /// The table control is a standard dynpro element, in contrast to the GuiCtrlGridView, which looks similar. GuiTableControl
    /// extends the GuiVContainer Object. The type prefix is tbl, the name is the fieldname taken from the SAP data dictionary.
    GuiTableControl(GuiVComponent), //(GuiTableControl),
    /// A tab strip is a container whose children are of type GuiTab. GuiTabStrip extends the GuiVContainer Object. The type
    /// prefix is tabs, the name is the fieldname taken from the SAP data dictionary.
    GuiTabStrip(GuiVComponent), //(GuiTabStrip),
    /// The TextEdit control is a multiline edit control offering a number of possible benefits. With regard to scripting,
    /// the possibility of protecting text parts against editing by the user is especially useful. GuiTextedit extends the
    /// GuiShell Object.
    GuiTextedit(GuiVComponent), //(GuiTextedit),
    /// GuiTextField extends the GuiVComponent Object. The type prefix is txt, the name is the fieldname taken from the
    /// SAP data dictionary.
    GuiTextField(GuiTextField),
    /// The titlebar is only displayed and exposed as a separate object in New Visual Design mode. GuiTitlebar extends the
    /// GuiVContainer Object. The type prefix and name of GuiTitlebar are titl.
    GuiTitlebar(GuiVComponent), //(GuiTitlebar),
    /// Every GuiFrameWindow has a GuiToolbar. The GuiMainWindow has two toolbars unless the second has been turned off by
    /// the ABAP application. In classical SAP GUI themes, the upper toolbar is called “system toolbar” or “GUI toolbar” ,
    /// while the second toolbar is called “application toolbar”. In SAP GUI themes as of Belize and in integration scenarios
    /// (like embedded into SAP Business Client), only a single toolbar (“merged toolbar") is displayed. Additionally, a footer
    /// also containing buttons originally coming from the system or application toolbar may be displayed.
    GuiToolbar(GuiVComponent), //(GuiToolbar),
    /// A Tree view.
    GuiTree(GuiVComponent), //(GuiTree),
    /// The GuiUserArea comprises the area between the toolbar and status bar for windows of GuiMainWindow type and the area
    /// between the titlebar and toolbar for modal windows, and may also be limited by docker controls. The standard dynpro
    /// elements can be found only in this area, with the exception of buttons, which are also found in the toolbars.
    GuiUserArea(GuiVComponent), //(GuiUserArea),
    /// The GuiVComponent interface is exposed by all visual objects, such as windows, buttons or text fields. Like GuiComponent,
    /// it is an abstract interface. Any object supporting the GuiVComponent interface also exposes the GuiComponent interface.
    /// GuiVComponent extends the GuiComponent Object.
    GuiVComponent(GuiVComponent),
    /// An object exposes the GuiVContainer interface if it is both visible and can have children. It will then also expose
    /// GuiComponent and GuiVComponent. Examples of this interface are windows and subscreens, toolbars or controls having
    /// children, such as the splitter control. GuiVContainer extends the GuiContainer Object and the GuiVComponent Object.
    GuiVContainer(GuiVComponent), //(GuiVContainer),
    /// GuiVHViewSwitch represents the “View Switch” object that was introduced with the Belize theme in SAP GUI. The View Switch
    /// is placed in the header area of the SAP GUI main window and can be used to select different views within an application.
    /// Many screens can be displayed in different ways (for example, as a tree or list). To switch from one view to another in
    /// a comfortable way, these screens may make use of the View Switch:
    GuiVHViewSwitch(GuiVComponent), //(GuiVHViewSwitch),
}

impl From<GuiComponent> for SAPComponent {
    fn from(value: GuiComponent) -> Self {
        if let Ok(kind) = value.kind() {
            log::debug!("Converting component {kind} to SAPComponent.");
            match kind.as_str() {
                "GuiAbapEditor" => SAPComponent::GuiAbapEditor(value.into_vcomponent_unchecked()),
                "GuiApoGrid" => SAPComponent::GuiApoGrid(value.into_vcomponent_unchecked()),
                "GuiApplication" => {
                    SAPComponent::GuiApplication(GuiApplication { inner: value.inner })
                }
                "GuiBarChart" => SAPComponent::GuiBarChart(value.into_vcomponent_unchecked()),
                "GuiBox" => SAPComponent::GuiBox(value.into_vcomponent_unchecked()),
                "GuiButton" => SAPComponent::GuiButton(GuiButton { inner: value.inner }),
                "GuiCalendar" => SAPComponent::GuiCalendar(value.into_vcomponent_unchecked()),
                "GuiChart" => SAPComponent::GuiChart(value.into_vcomponent_unchecked()),
                "GuiCheckBox" => SAPComponent::GuiCheckBox(GuiCheckBox { inner: value.inner }),
                "GuiColorSelector" => {
                    SAPComponent::GuiColorSelector(value.into_vcomponent_unchecked())
                }
                "GuiComboBox" => SAPComponent::GuiComboBox(GuiComboBox { inner: value.inner }),
                "GuiComboBoxControl" => {
                    SAPComponent::GuiComboBoxControl(value.into_vcomponent_unchecked())
                }
                "GuiComboBoxEntry" => {
                    SAPComponent::GuiComboBoxEntry(value.into_vcomponent_unchecked())
                }
                "GuiComponent" => SAPComponent::GuiComponent(value.into()),
                "GuiContainer" => SAPComponent::GuiContainer(value.into_vcomponent_unchecked()),
                "GuiContainerShell" => {
                    SAPComponent::GuiContainerShell(value.into_vcomponent_unchecked())
                }
                "GuiCTextField" => SAPComponent::GuiCTextField(GuiCTextField { inner: value.inner }),
                "GuiCustomControl" => {
                    SAPComponent::GuiCustomControl(value.into_vcomponent_unchecked())
                }
                "GuiDialogShell" => SAPComponent::GuiDialogShell(value.into_vcomponent_unchecked()),
                "GuiEAIViewer2D" => SAPComponent::GuiEAIViewer2D(value.into_vcomponent_unchecked()),
                "GuiEAIViewer3D" => SAPComponent::GuiEAIViewer3D(value.into_vcomponent_unchecked()),
                "GuiFrameWindow" => {
                    SAPComponent::GuiFrameWindow(GuiFrameWindow { inner: value.inner })
                }
                "GuiGOSShell" => SAPComponent::GuiGOSShell(value.into_vcomponent_unchecked()),
                "GuiGraphAdapt" => SAPComponent::GuiGraphAdapt(value.into_vcomponent_unchecked()),
                "GuiGridView" => SAPComponent::GuiGridView(value.into_vcomponent_unchecked()),
                "GuiHTMLViewer" => SAPComponent::GuiHTMLViewer(value.into_vcomponent_unchecked()),
                "GuiInputFieldControl" => {
                    SAPComponent::GuiInputFieldControl(value.into_vcomponent_unchecked())
                }
                "GuiLabel" => SAPComponent::GuiLabel(value.into_vcomponent_unchecked()),
                "GuiMainWindow" => SAPComponent::GuiMainWindow(GuiMainWindow { inner: value.inner, }),
                "GuiMap" => SAPComponent::GuiMap(value.into_vcomponent_unchecked()),
                "GuiMenu" => SAPComponent::GuiMenu(GuiMenu { inner: value.inner }),
                "GuiMenubar" => SAPComponent::GuiMenubar(GuiMenubar { inner: value.inner }),
                "GuiModalWindow" => SAPComponent::GuiModalWindow(value.into_vcomponent_unchecked()),
                "GuiNetChart" => SAPComponent::GuiNetChart(value.into_vcomponent_unchecked()),
                "GuiOfficeIntegration" => {
                    SAPComponent::GuiOfficeIntegration(value.into_vcomponent_unchecked())
                }
                "GuiOkCodeField" => {
                    SAPComponent::GuiOkCodeField(GuiOkCodeField { inner: value.inner })
                }
                "GuiPasswordField" => {
                    SAPComponent::GuiPasswordField(value.into_vcomponent_unchecked())
                }
                "GuiPicture" => SAPComponent::GuiPicture(value.into_vcomponent_unchecked()),
                "GuiRadioButton" => SAPComponent::GuiRadioButton(GuiRadioButton { inner: value.inner }),
                "GuiSapChart" => SAPComponent::GuiSapChart(value.into_vcomponent_unchecked()),
                "GuiScrollbar" => SAPComponent::GuiScrollbar(value.into_vcomponent_unchecked()),
                "GuiScrollContainer" => {
                    SAPComponent::GuiScrollContainer(value.into_vcomponent_unchecked())
                }
                "GuiShell" => SAPComponent::GuiShell(value.into_vcomponent_unchecked()),
                "GuiSimpleContainer" => {
                    SAPComponent::GuiSimpleContainer(value.into_vcomponent_unchecked())
                }
                "GuiSplit" => SAPComponent::GuiSplit(value.into_vcomponent_unchecked()),
                "GuiSplitterContainer" => {
                    SAPComponent::GuiSplitterContainer(value.into_vcomponent_unchecked())
                }
                "GuiStage" => SAPComponent::GuiStage(value.into_vcomponent_unchecked()),
                "GuiStatusbar" => SAPComponent::GuiStatusbar(value.into_vcomponent_unchecked()),
                "GuiStatusPane" => SAPComponent::GuiStatusPane(value.into_vcomponent_unchecked()),
                "GuiTab" => SAPComponent::GuiTab(value.into_vcomponent_unchecked()),
                "GuiTableControl" => {
                    SAPComponent::GuiTableControl(value.into_vcomponent_unchecked())
                }
                "GuiTabStrip" => SAPComponent::GuiTabStrip(value.into_vcomponent_unchecked()),
                "GuiTextedit" => SAPComponent::GuiTextedit(value.into_vcomponent_unchecked()),
                "GuiTextField" => SAPComponent::GuiTextField(GuiTextField { inner: value.inner }),
                "GuiTitlebar" => SAPComponent::GuiTitlebar(value.into_vcomponent_unchecked()),
                "GuiToolbar" => SAPComponent::GuiToolbar(value.into_vcomponent_unchecked()),
                "GuiTree" => SAPComponent::GuiTree(value.into_vcomponent_unchecked()),
                "GuiUserArea" => SAPComponent::GuiUserArea(value.into_vcomponent_unchecked()),
                "GuiVComponent" => SAPComponent::GuiVComponent(value.into_vcomponent_unchecked()),
                "GuiVContainer" => SAPComponent::GuiVContainer(value.into_vcomponent_unchecked()),
                "GuiVHViewSwitch" => {
                    SAPComponent::GuiVHViewSwitch(value.into_vcomponent_unchecked())
                }
                _ => SAPComponent::GuiComponent(value),
            }
        } else {
            SAPComponent::GuiComponent(value)
        }
    }
}

type_struct! {
    /// The GuiApplication represents the process in which all SAP GUI activity takes place. If the scripting
    /// component is accessed by attaching to an SAP Logon process, then GuiApplication will represent SAP
    /// Logon. GuiApplication is a creatable class. However, there must be only one component of this type in
    /// any process. GuiApplication extends the GuiContainer Object.
    => GuiApplication: GuiComponentMethods, GuiContainerMethods,
}

impl GuiApplication {
    /// Start a new connection with the provided connection string.
    pub fn start_connection<S>(&self, connection_string: S) -> crate::Result<GuiConnection>
    where
        S: AsRef<str>,
    {
        Ok(GuiConnection {
            inner: self
                .inner
                .call(
                    "OpenConnectionByConnectionString",
                    vec![VARIANT::from_str(connection_string.as_ref())],
                )?
                .to_idispatch()?
                .clone(),
        })
    }

    /// Get the nth session from the scripting engine.
    pub fn get_connection(&self, n: i32) -> crate::Result<GuiConnection> {
        Ok(GuiConnection {
            inner: self
                .inner
                .get_named("Children", VARIANT::from_i32(n))?
                .to_idispatch()?
                .clone(),
        })
    }
}

type_struct! {
    /// A GuiConnection represents the connection between SAP GUI and an application server. Connections can be
    /// opened from SAP Logon or from GuiApplication’s openConnection and openConnectionByConnectionString
    /// methods. GuiConnection extends the GuiContainer Object. The type prefix for GuiConnection is con, the
    /// name is con plus the connection number in square brackets.
    => GuiConnection: GuiContainerMethods, GuiComponentMethods,
}

impl GuiConnection {
    /// This method closes a connection along with all its sessions.
    pub fn close_connection(&self) -> crate::Result<()> {
        self.inner.call("CloseConnection", vec![])?;
        Ok(())
    }

    /// A session can be closed by calling this method of the connection. Closing the last session of a connection will close the connection, too.
    ///
    /// The parameter "Id" must contain the id of the session to close (like "/app/con[0]/ses[0]").
    pub fn close_session<S>(&self, id: S) -> crate::Result<()>
    where
        S: AsRef<str>,
    {
        self.inner
            .call("CloseSession", vec![VARIANT::from_str(id.as_ref())])?;
        Ok(())
    }
}

type_struct! {
    /// The GuiSession provides the context in which a user performs a certain task such as working with
    /// a transaction. It is therefore the access point for applications, which record a user’s actions
    /// regarding a specific task or play back those actions. GuiSession extends GuiContainer. The type
    /// prefix is ses, the name is ses plus the session number in square brackets.
    => GuiSession: GuiComponentMethods, GuiContainerMethods,
}

type_struct! {
    /// GuiButton represents all push buttons that are on dynpros, the toolbar or in table controls.
    /// GuiButton extends the GuiVComponent Object. The type prefix is btn, the name property is the
    /// fieldname taken from the SAP data dictionary There is one exception: For tabstrip buttons, it is
    /// the button id set in screen painter that is taken from the SAP data dictionary.
    => GuiButton: GuiBoxMethods, GuiComponentMethods, GuiVComponentMethods,
}

impl GuiButton {
    forward_func! {
        /// This emulates manually pressing a button. Pressing a button will always cause server
        /// communication to occur, rendering all references to elements below the window level
        /// invalid.
        => press, "Press"
    }
    get_property! {
        /// This property is True if the button is displayed emphasized (in Fiori Visual Themes:
        /// The leftmost button in the footer and buttons configured as
        /// "Fiori Usage D Display<->Change").
        => emphasized, "Emphasized", bool, to_bool
    }
    get_property! {
        /// Left label of the GuiButton. The label is assigned in the Screen Painter, using the flag
        /// 'assign left'.
        => left_label, "LeftLabel", SAPComponent, to_sap_component
    }
    get_property! {
        /// Right label of the GuiButton. This property is set in Screen Painter using the 'assign
        /// right' flag.
        => right_label, "RightLabel", SAPComponent, to_sap_component
    }
}

type_struct! {
    /// This interface resembles GuiVContainer. The only difference is that it is not intended for visual
    /// objects but rather administrative objects such as connections or sessions. Objects exposing this
    /// interface will therefore support GuiComponent but not GuiVComponent. GuiContainer extends the
    /// GuiComponent Object.
    => GuiComponent: GuiComponentMethods, GuiContainerMethods,
}

impl GuiComponent {
    fn into_vcomponent_unchecked(self) -> GuiVComponent {
        GuiVComponent { inner: self.inner }
    }
}

type_struct! {
    /// A GuiFrameWindow is a high level visual object in the runtime hierarchy. It can be either the main
    /// window or a modal popup window. See the GuiMainWindow and GuiModalWindow sections for examples.
    /// GuiFrameWindow itself is an abstract interface. GuiFrameWindow extends the GuiVContainer Object.
    /// The type prefix is wnd, the name is wnd plus the window number in square brackets.
    => GuiFrameWindow: GuiComponentMethods, GuiVComponentMethods, GuiContainerMethods, GuiVContainerMethods, GuiFrameWindowMethods,
}

type_struct! {
    /// This window represents the main window of an SAP GUI session. GuiMainWindow extends the
    /// GuiFrameWindow Object.
    => GuiMainWindow: GuiComponentMethods, GuiVComponentMethods, GuiContainerMethods, GuiVContainerMethods, GuiFrameWindowMethods,
}

impl GuiMainWindow {
    get_property! {
        /// This property it True if the application toolbar, the lower toolbar within SAP GUI,
        /// is visible. Setting this property to False will hide the application toolbar.
        => buttonbar_visible, "ButtonbarVisible", bool, to_bool
    }
    set_property! {
        /// This property it True if the application toolbar, the lower toolbar within SAP GUI,
        /// is visible. Setting this property to False will hide the application toolbar.
        => set_buttonbar_visible, "ButtonbarVisible", bool, from_bool
    }
    get_property! {
        /// This property it True if the status bar at the bottom of the SAP GUI window is
        /// visible. Setting this property to False will hide the status bar. When the status
        /// bar is hidden, messages will be displayed in a popup instead.
        => statusbar_visible, "StatusbarVisible", bool, to_bool
    }
    set_property! {
        /// This property it True if the status bar at the bottom of the SAP GUI window is
        /// visible. Setting this property to False will hide the status bar. When the status
        /// bar is hidden, messages will be displayed in a popup instead.
        => set_statusbar_visible, "StatusbarVisible", bool, from_bool
    }
    get_property! {
        /// This property it True if the title bar is visible. Setting this property to False
        /// will hide the title bar. The title bar is only available in New Visual Design, not
        /// in Classic Design.
        => titlebar_visible, "TitlebarVisible", bool, to_bool
    }
    set_property! {
        /// This property it True if the title bar is visible. Setting this property to False
        /// will hide the title bar. The title bar is only available in New Visual Design, not
        /// in Classic Design.
        => set_titlebar_visible, "TitlebarVisible", bool, from_bool
    }
    get_property! {
        /// This property it True if the system toolbar, the upper toolbar within SAP GUI, is
        /// visible. Setting this property to False will hide the system toolbar.
        => toolbar_visible, "ToolbarVisible", bool, to_bool
    }
    set_property! {
        /// This property it True if the system toolbar, the upper toolbar within SAP GUI, is
        /// visible. Setting this property to False will hide the system toolbar.
        => set_toolbar_visible, "ToolbarVisible", bool, from_bool
    }
}

type_struct! {
    /// The GuiOkCodeField is placed on the upper toolbar of the main window. It is a combo box into which
    /// commands can be entered. Setting the text of GuiOkCodeField will not execute the command until
    /// server communication is started, for example by emulating the Enter key (VKey 0). GuiOkCodeField
    /// extends the GuiVComponent Object. The type prefix is okcd, the name is empty.
    => GuiOkCodeField: GuiComponentMethods, GuiContainerMethods, GuiVComponentMethods,
}

impl GuiOkCodeField {
    get_property! {
        /// In SAP GUI designs newer than Classic design the GuiOkCodeField can be collapsed using the arrow
        /// button to the right of it. In SAP GUI for Windows the GuiOkCodeField may also be collapsed via a
        /// setting in the Windows registry.
        ///
        /// This property contains False is the GuiOkCodeField is collapsed.
        => opened, "Opened", bool, to_bool
    }
}

type_struct! {
    /// The GuiVComponent interface is exposed by all visual objects, such as windows, buttons or text fields.
    /// Like GuiComponent, it is an abstract interface. Any object supporting the GuiVComponent interface
    /// also exposes the GuiComponent interface. GuiVComponent extends the GuiComponent Object.
    => GuiVComponent: GuiComponentMethods, GuiContainerMethods, GuiVComponentMethods,
}

type_struct! {
    /// GuiCheckBox extends the GuiVComponent Object. The type prefix is chk, the name is the fieldname taken
    /// from the SAP data dictionary.
    => GuiCheckBox: GuiComponentMethods, GuiVComponentMethods,
}

type_struct! {
    /// The GuiComboBox looks somewhat similar to GuiCTextField, but has a completely different implementation.
    /// While pressing the combo box button of a GuiCTextField will open a new dynpro or control in which a
    /// selection can be made, GuiComboBox retrieves all possible choices on initialization from the server,
    /// so the selection is done solely on the client. GuiComboBox extends the GuiVComponent Object. The type
    /// prefix is cmb, the name is the fieldname taken from the SAP data dictionary. GuiComboBox inherits from
    /// the GuiVComponent Object.
    => GuiComboBox: GuiComponentMethods, GuiVComponentMethods,
}

type_struct! {
    /// If the cursor is set into a text field of type GuiCTextField a combo box button is displayed to the right
    /// of the text field. Pressing this button is equivalent to pressing the F4 key. The button is not
    /// represented in the scripting object model as a separate object; it is considered to be part of the text
    /// field.
    ///
    /// There are no other differences between GuiTextField and GuiCTextField. GuiCTextField extends the
    /// GuiTextField. The type prefix is ctxt, the name is the Fieldname taken from the SAP data dictionary.
    => GuiCTextField: GuiTextFieldMethods, GuiVComponentMethods, GuiComponentMethods,
}

type_struct! {
    /// A GuiMenu may have other GuiMenu objects as children. GuiMenu extends the GuiVContainer Object. The type
    /// prefix is menu, the name is the text of the menu item. If the item does not have a text, which is the case
    /// for separators, then the name is the last part of the id, menu[n].
    => GuiMenu: GuiVComponentMethods, GuiVContainerMethods, GuiContainerMethods, GuiComponentMethods,
}

impl GuiMenu {
    forward_func! {
        /// Select the menu.
        => select, "Select"
    }
}

type_struct! {
    /// Only the main window has a menubar. The children of the menubar are menus. GuiMenubar extends the
    /// GuiVContainer Object. The type prefix and name are mbar.
    => GuiMenubar: GuiVComponentMethods, GuiVContainerMethods, GuiContainerMethods, GuiComponentMethods,
}

type_struct! {
    /// GuiRadioButton extends the GuiVComponent Object. The type prefix is rad, the name is the fieldname taken from the SAP data dictionary.
    => GuiRadioButton: GuiVComponentMethods, GuiComponentMethods,
}

impl GuiRadioButton {
    forward_func! {
        /// Selecting a radio button automatically deselects all the other buttons within that group.
        /// This may cause a server roundtrip, depending on the definition of the button in the screen
        /// painter.
        => select, "Select"
    }

    // TODO properties
}

type_struct! {
    /// GuiTextField extends the GuiVComponent Object. The type prefix is txt, the name is the fieldname taken from the SAP data dictionary.
    => GuiTextField: GuiVComponentMethods, GuiComponentMethods, GuiTextFieldMethods,
}
