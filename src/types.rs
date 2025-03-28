use com_shim::{com_shim, IDispatchExt, VariantTypeExt};
use windows::{core::*, Win32::System::Com::*, Win32::System::Variant::*};

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
            .call("GetROTEntry", vec![VARIANT::variant_from("SAPGUI".to_string())])?;

        let sap_gui: &IDispatch = result.variant_into()?;

        log::debug!("Getting scripting engine.");
        let scripting_engine = sap_gui.call("GetScriptingEngine", vec![])?;

        Ok(GuiApplication {
            inner: <com_shim::VARIANT as VariantTypeExt<'_, &IDispatch>>::variant_into(&scripting_engine)?.clone(),
        })
    }
}

/// SAPComponent wraps around all of the GuiComponent types, allowing a more Rust-y way of exporing components.
pub enum SAPComponent {
    /// The GuiApplication represents the process in which all SAP GUI activity takes place. If the scripting
    /// component is accessed by attaching to an SAP Logon process, then GuiApplication will represent SAP
    /// Logon. GuiApplication is a creatable class. However, there must be only one component of this type
    /// in any process. GuiApplication extends the GuiContainer Object.
    GuiApplication(GuiApplication),
    /// The GuiBarChart is a powerful tool to display and modify time scale diagrams.
    GuiBarChart(GuiBarChart),
    /// A GuiBox is a simple frame with a name (also called a "Group Box"). The items inside the frame are not
    /// children of the box. The type prefix is "box".
    GuiBox(GuiBox),
    /// GuiButton represents all push buttons that are on dynpros, the toolbar or in table controls. GuiButton
    /// extends the GuiVComponent Object. The type prefix is btn, the name property is the fieldname taken
    /// from the SAP data dictionary There is one exception: For tabstrip buttons, it is the button id set in
    /// screen painter that is taken from the SAP data dictionary.
    GuiButton(GuiButton),
    /// The calendar control can be used to select single dates or periods of time. GuiCalendar extends the
    /// GuiShell Object.
    GuiCalendar(GuiCalendar),
    /// The GuiChart object is of a very technical nature. It should only be used for recording and playback, as
    /// most of the parameters cannot be determined in any other way.
    GuiChart(GuiChart),
    /// GuiCheckBox extends the GuiVComponent Object. The type prefix is chk, the name is the fieldname taken
    /// from the SAP data dictionary.
    GuiCheckBox(GuiCheckBox),
    /// GuiColorSelector displays a set of colors for selection. It extends the GuiShell Object.
    GuiColorSelector(GuiColorSelector),
    /// The GuiComboBox looks somewhat similar to GuiCTextField, but has a completely different implementation.
    /// While pressing the combo box button of a GuiCTextField will open a new dynpro or control in which a
    /// selection can be made, GuiComboBox retrieves all possible choices on initialization from the server, so
    /// the selection is done solely on the client. GuiComboBox extends the GuiVComponent Object. The type prefix
    /// is cmb, the name is the fieldname taken from the SAP data dictionary. GuiComboBox inherits from the
    /// GuiVComponent Object.
    GuiComboBox(GuiComboBox),
    ///
    GuiComboBoxControl(GuiComboBoxControl),
    /// Members of the Entries collection of a GuiComboBox are of type GuiComBoxEntry.
    GuiComboBoxEntry(GuiComboBoxEntry),
    /// GuiComponent is the base class for most classes in the Scripting API. It was designed to allow generic
    /// programming, meaning you can work with objects without knowing their exact type.
    GuiComponent(GuiComponent),
    /// A GuiConnection represents the connection between SAP GUI and an application server. Connections can be opened
    /// from SAP Logon or from GuiApplication’s openConnection and openConnectionByConnectionString methods.
    /// GuiConnection extends the GuiContainer Object. The type prefix for GuiConnection is con, the name is con
    /// plus the connection number in square brackets.
    GuiConnection(GuiConnection),
    /// This interface resembles GuiVContainer. The only difference is that it is not intended for visual objects
    /// but rather administrative objects such as connections or sessions. Objects exposing this interface will
    /// therefore support GuiComponent but not GuiVComponent. GuiContainer extends the GuiComponent Object.
    GuiContainer(GuiContainer),
    /// A GuiContainerShell is a wrapper for a set of the GuiShell Object. GuiContainerShell extends the GuiVContainer
    /// Object. The type prefix is shellcont, the name is the last part of the id, shellcont\[n\].
    GuiContainerShell(GuiContainerShell),
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
    GuiCustomControl(GuiCustomControl),
    /// The GuiDialogShell is an external window that is used as a container for other shells, for example a toolbar.
    /// GuiDialogShell extends the GuiVContainer Object. The type prefix is shellcont, the name is the last part of
    /// the id, shellcont\[n\].
    GuiDialogShell(GuiDialogShell),
    /// The GuiEAIViewer2D control is used to view 2-dimensional graphic images in the SAP system. The user can carry
    /// out redlining over the loaded image. The scripting wrapper for this control records all user actions during
    /// the redlining process and reproduces the same actions when the recorded script is replayed.
    GuiEAIViewer2D(GuiEAIViewer2D),
    /// The GuiEAIViewer3D control is used to view 3-dimensional graphic images in the SAP system.
    GuiEAIViewer3D(GuiEAIViewer3D),
    /// A GuiFrameWindow is a high level visual object in the runtime hierarchy. It can be either the main window or
    /// a modal popup window. See the GuiMainWindow and GuiModalWindow sections for examples. GuiFrameWindow itself
    /// is an abstract interface. GuiFrameWindow extends the GuiVContainer Object. The type prefix is wnd, the name
    /// is wnd plus the window number in square brackets.
    GuiFrameWindow(GuiFrameWindow),
    /// The GuiGosShell is only available in New Visual Design mode. GuiGOSShell extends the GuiVContainer Object.
    /// The type prefix is shellcont, the name is the last part of the id, shellcont\[n\].
    GuiGOSShell(GuiGOSShell),
    /// For the graphic adapter control only basic members from GuiShell are available. Recording and playback is
    /// not possible.
    GuiGraphAdapt(GuiGraphAdapt),
    /// The grid view is similar to the dynpro table control, but significantly more powerful. GuiGridView extends
    /// the GuiShell Object.
    GuiGridView(GuiGridView),
    /// The GuiHTMLViewer is used to display an HTML document inside SAP GUI. GuiHTMLViewer extends the GuiShell
    /// Object.
    GuiHTMLViewer(GuiHTMLViewer),
    ///
    GuiInputFieldControl(GuiInputFieldControl),
    /// GuiLabel extends the GuiVComponent Object. The type prefix is lbl, the name is the fieldname taken from the
    /// SAP data dictionary.
    GuiLabel(GuiLabel),
    /// This window represents the main window of an SAP GUI session.
    GuiMainWindow(GuiMainWindow),
    /// For the map control only basic members from GuiShell are available. Recording and playback is not possible.
    GuiMap(GuiMap),
    /// A GuiMenu may have other GuiMenu objects as children. GuiMenu extends the GuiVContainer Object. The type prefix
    /// is menu, the name is the text of the menu item. If the item does not have a text, which is the case for
    /// separators, then the name is the last part of the id, menu\[n\].
    GuiMenu(GuiMenu),
    /// Only the main window has a menubar. The children of the menubar are menus. GuiMenubar extends the GuiVContainer
    /// Object. The type prefix and name are mbar.
    GuiMenubar(GuiMenubar),
    /// A GuiModalWindow is a dialog pop-up.
    GuiModalWindow(GuiModalWindow),
    /// The GuiNetChart is a powerful tool to display and modify entity relationship diagrams. It is of a very technical
    /// nature and should only be used for recording and playback, as most of the parameters cannot be determined in
    /// any other way.
    GuiNetChart(GuiNetChart),
    ///
    GuiOfficeIntegration(GuiOfficeIntegration),
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
    GuiPasswordField(GuiPasswordField),
    /// The picture control displays a picture on an SAP GUI screen. GuiPicture extends the GuiShell Object.
    GuiPicture(GuiPicture),
    /// GuiRadioButton extends the GuiVComponent Object. The type prefix is rad, the name is the fieldname taken from the
    /// SAP data dictionary.
    GuiRadioButton(GuiRadioButton),
    /// For the SAP chart control only basic members from GuiShell are available. Recording and playback is not possible.
    GuiSapChart(GuiSapChart),
    /// The GuiScrollbar class is a utility class used for example in GuiScrollContainer or GuiTableControl.
    GuiScrollbar(GuiScrollbar),
    /// This container represents scrollable subscreens. A subscreen may be scrollable without actually having a scrollbar,
    /// because the existence of a scrollbar depends on the amount of data displayed and the size of the GuiUserArea.
    /// GuiScrollContainer extend sthe GuiVContainer Object. The type prefix is ssub, the name is generated from the data
    /// dictionary settings.
    GuiScrollContainer(GuiScrollContainer),
    /// GuiSession is self-contained in that ids within the context of a session remain valid independently of other connections
    /// or sessions being open at the same time. Usually an external application will first determine with which session to
    /// interact. Once that is clear, the application will work more or less exclusively on that session. Traversing the object
    /// hierarchy from the GuiApplication to the user interface elements, it is the session among whose children the highest
    /// level visible objects can be found. In contrast to objects like buttons or text fields, the session remains valid until
    /// the corresponding main window has been closed, whereas buttons, for example, are destroyed during each server
    /// communication.
    GuiSession(GuiSession),
    /// GuiShell is an abstract object whose interface is supported by all the controls. GuiShell extends the GuiVContainer
    /// Object. The type prefix is shell, the name is the last part of the id, shell\[n\].
    GuiShell(GuiShell),
    /// This container represents non-scrollable subscreens. It does not have any functionality apart from to the inherited
    /// interfaces. GuiSimpleContainer extends the GuiVContainer Object. The type prefix is sub, the name is is generated
    /// from the data dictionary settings.
    GuiSimpleContainer(GuiSimpleContainer),
    /// GuiSplit extends the GuiShell Object.
    GuiSplit(GuiSplit),
    /// The GuiSplitterContainer represents the dynpro splitter element, which was introduced in the Web Application Server
    /// ABAP in NetWeaver 7.1. The dynpro splitter element is similar to the activeX based splitter control, but it is a
    /// plain dynpro element.
    GuiSplitterContainer(GuiSplitterContainer),
    /// For the stage control only basic members from GuiShell are available. Recording and playback is not possible.
    GuiStage(GuiStage),
    /// GuiStatusbar represents the message displaying part of the status bar on the bottom of the SAP GUI window. It does
    /// not include the system and login information displayed in the rightmost area of the status bar as these are available
    /// from the GuiSessionInfo object. GuiStatusbar extends the GuiVComponent Object. The type prefix is sbar.
    GuiStatusbar(GuiStatusbar),
    /// The parent of the GuiStatusPane objects is the status bar (see also GuiStatusbar Object). The GuiStatusPane objects
    /// reflect the individual areas of the status bar, for example "pane\[0\]" refers to the section of the status bar where
    /// the messages are displayed. See also GuiStatusbar Object. The first pane of the GuiStatusBar (pane\[0\]) can have a
    /// child of type GuiStatusBarLink, if a service request link is displayed.
    GuiStatusPane(GuiStatusPane),
    /// The GuiTab objects are the children of a GuiTabStrip object. GuiTab extends the GuiVContainer Object. The type prefix
    /// is tabp, the name is the id of the tab’s button taken from SAP data dictionary.
    GuiTab(GuiTab),
    /// The table control is a standard dynpro element, in contrast to the GuiCtrlGridView, which looks similar. GuiTableControl
    /// extends the GuiVContainer Object. The type prefix is tbl, the name is the fieldname taken from the SAP data dictionary.
    GuiTableControl(GuiTableControl),
    /// A tab strip is a container whose children are of type GuiTab. GuiTabStrip extends the GuiVContainer Object. The type
    /// prefix is tabs, the name is the fieldname taken from the SAP data dictionary.
    GuiTabStrip(GuiTabStrip),
    /// The TextEdit control is a multiline edit control offering a number of possible benefits. With regard to scripting,
    /// the possibility of protecting text parts against editing by the user is especially useful. GuiTextedit extends the
    /// GuiShell Object.
    GuiTextedit(GuiTextedit),
    /// GuiTextField extends the GuiVComponent Object. The type prefix is txt, the name is the fieldname taken from the
    /// SAP data dictionary.
    GuiTextField(GuiTextField),
    /// The titlebar is only displayed and exposed as a separate object in New Visual Design mode. GuiTitlebar extends the
    /// GuiVContainer Object. The type prefix and name of GuiTitlebar are titl.
    GuiTitlebar(GuiTitlebar),
    /// Every GuiFrameWindow has a GuiToolbar. The GuiMainWindow has two toolbars unless the second has been turned off by
    /// the ABAP application. In classical SAP GUI themes, the upper toolbar is called “system toolbar” or “GUI toolbar” ,
    /// while the second toolbar is called “application toolbar”. In SAP GUI themes as of Belize and in integration scenarios
    /// (like embedded into SAP Business Client), only a single toolbar (“merged toolbar") is displayed. Additionally, a footer
    /// also containing buttons originally coming from the system or application toolbar may be displayed.
    GuiToolbar(GuiToolbar),
    /// A Tree view.
    GuiTree(GuiTree),
    /// The GuiUserArea comprises the area between the toolbar and status bar for windows of GuiMainWindow type and the area
    /// between the titlebar and toolbar for modal windows, and may also be limited by docker controls. The standard dynpro
    /// elements can be found only in this area, with the exception of buttons, which are also found in the toolbars.
    GuiUserArea(GuiUserArea),
    /// The GuiVComponent interface is exposed by all visual objects, such as windows, buttons or text fields. Like GuiComponent,
    /// it is an abstract interface. Any object supporting the GuiVComponent interface also exposes the GuiComponent interface.
    /// GuiVComponent extends the GuiComponent Object.
    GuiVComponent(GuiVComponent),
    /// An object exposes the GuiVContainer interface if it is both visible and can have children. It will then also expose
    /// GuiComponent and GuiVComponent. Examples of this interface are windows and subscreens, toolbars or controls having
    /// children, such as the splitter control. GuiVContainer extends the GuiContainer Object and the GuiVComponent Object.
    GuiVContainer(GuiVContainer),
    /// GuiVHViewSwitch represents the “View Switch” object that was introduced with the Belize theme in SAP GUI. The View Switch
    /// is placed in the header area of the SAP GUI main window and can be used to select different views within an application.
    /// Many screens can be displayed in different ways (for example, as a tree or list). To switch from one view to another in
    /// a comfortable way, these screens may make use of the View Switch:
    GuiVHViewSwitch(GuiVHViewSwitch),
}

impl From<IDispatch> for SAPComponent {
    fn from(value: IDispatch) -> Self {
        let value = GuiComponent { inner: value };
        if let Ok(mut kind) = value.r_type() {
            log::debug!("Converting component {kind} to SAPComponent.");
            if kind.as_str() == "GuiShell" {
                log::debug!("Kind is shell, checking subkind.");
                if let Ok(sub_kind) = (GuiShell { inner: value.inner.clone() }).sub_type() {
                    // use subkind if a GuiShell
                    log::debug!("Subkind is {sub_kind}");
                    kind = sub_kind;
                }
            }
            match kind.as_str() {
                // ! Types that extend from GuiShell are not prefixed with `Gui` as they use SubType.
                "GuiApplication" => {
                    SAPComponent::GuiApplication(GuiApplication { inner: value.inner })
                }
                "BarChart" => SAPComponent::GuiBarChart(GuiBarChart { inner: value.inner }),
                "GuiBox" => SAPComponent::GuiBox(GuiBox { inner: value.inner }),
                "GuiButton" => SAPComponent::GuiButton(GuiButton { inner: value.inner }),
                "Calendar" => SAPComponent::GuiCalendar(GuiCalendar { inner: value.inner }),
                "Chart" => SAPComponent::GuiChart(GuiChart { inner: value.inner }),
                "GuiCheckBox" => SAPComponent::GuiCheckBox(GuiCheckBox { inner: value.inner }),
                "ColorSelector" => {
                    SAPComponent::GuiColorSelector(GuiColorSelector { inner: value.inner })
                }
                "GuiComboBox" => SAPComponent::GuiComboBox(GuiComboBox { inner: value.inner }),
                "ComboBoxControl" => {
                    SAPComponent::GuiComboBoxControl(GuiComboBoxControl { inner: value.inner })
                }
                "GuiComboBoxEntry" => {
                    SAPComponent::GuiComboBoxEntry(GuiComboBoxEntry { inner: value.inner })
                }
                "GuiComponent" => SAPComponent::GuiComponent(value),
                "GuiConnection" => {
                    SAPComponent::GuiConnection(GuiConnection { inner: value.inner })
                }
                "GuiContainer" => SAPComponent::GuiContainer(GuiContainer { inner: value.inner }),
                "ContainerShell" => {
                    SAPComponent::GuiContainerShell(GuiContainerShell { inner: value.inner })
                }
                "GuiCTextField" => {
                    SAPComponent::GuiCTextField(GuiCTextField { inner: value.inner })
                }
                "GuiCustomControl" => {
                    SAPComponent::GuiCustomControl(GuiCustomControl { inner: value.inner })
                }
                "GuiDialogShell" => {
                    SAPComponent::GuiDialogShell(GuiDialogShell { inner: value.inner })
                }
                "EAIViewer2D" => {
                    SAPComponent::GuiEAIViewer2D(GuiEAIViewer2D { inner: value.inner })
                }
                "EAIViewer3D" => {
                    SAPComponent::GuiEAIViewer3D(GuiEAIViewer3D { inner: value.inner })
                }
                "GuiFrameWindow" => {
                    SAPComponent::GuiFrameWindow(GuiFrameWindow { inner: value.inner })
                }
                "GuiGOSShell" => SAPComponent::GuiGOSShell(GuiGOSShell { inner: value.inner }),
                "GraphAdapt" => {
                    SAPComponent::GuiGraphAdapt(GuiGraphAdapt { inner: value.inner })
                }
                "GridView" => SAPComponent::GuiGridView(GuiGridView { inner: value.inner }),
                "HTMLViewer" => {
                    SAPComponent::GuiHTMLViewer(GuiHTMLViewer { inner: value.inner })
                }
                "InputFieldControl" => {
                    SAPComponent::GuiInputFieldControl(GuiInputFieldControl { inner: value.inner })
                }
                "GuiLabel" => SAPComponent::GuiLabel(GuiLabel { inner: value.inner }),
                "GuiMainWindow" => {
                    SAPComponent::GuiMainWindow(GuiMainWindow { inner: value.inner })
                }
                "Map" => SAPComponent::GuiMap(GuiMap { inner: value.inner }),
                "GuiMenu" => SAPComponent::GuiMenu(GuiMenu { inner: value.inner }),
                "GuiMenubar" => SAPComponent::GuiMenubar(GuiMenubar { inner: value.inner }),
                "GuiModalWindow" => {
                    SAPComponent::GuiModalWindow(GuiModalWindow { inner: value.inner })
                }
                "NetChart" => SAPComponent::GuiNetChart(GuiNetChart { inner: value.inner }),
                "OfficeIntegration" => {
                    SAPComponent::GuiOfficeIntegration(GuiOfficeIntegration { inner: value.inner })
                }
                "GuiOkCodeField" => {
                    SAPComponent::GuiOkCodeField(GuiOkCodeField { inner: value.inner })
                }
                "GuiPasswordField" => {
                    SAPComponent::GuiPasswordField(GuiPasswordField { inner: value.inner })
                }
                "Picture" => SAPComponent::GuiPicture(GuiPicture { inner: value.inner }),
                "GuiRadioButton" => {
                    SAPComponent::GuiRadioButton(GuiRadioButton { inner: value.inner })
                }
                "SapChart" => SAPComponent::GuiSapChart(GuiSapChart { inner: value.inner }),
                "GuiScrollbar" => SAPComponent::GuiScrollbar(GuiScrollbar { inner: value.inner }),
                "GuiScrollContainer" => {
                    SAPComponent::GuiScrollContainer(GuiScrollContainer { inner: value.inner })
                }
                "GuiSession" => SAPComponent::GuiSession(GuiSession { inner: value.inner }),
                "GuiShell" => SAPComponent::GuiShell(GuiShell { inner: value.inner }),
                "GuiSimpleContainer" => {
                    SAPComponent::GuiSimpleContainer(GuiSimpleContainer { inner: value.inner })
                }
                "Split" => SAPComponent::GuiSplit(GuiSplit { inner: value.inner }),
                "SplitterContainer" => {
                    SAPComponent::GuiSplitterContainer(GuiSplitterContainer { inner: value.inner })
                }
                "Stage" => SAPComponent::GuiStage(GuiStage { inner: value.inner }),
                "GuiStatusbar" => SAPComponent::GuiStatusbar(GuiStatusbar { inner: value.inner }),
                "GuiStatusPane" => {
                    SAPComponent::GuiStatusPane(GuiStatusPane { inner: value.inner })
                }
                "GuiTab" => SAPComponent::GuiTab(GuiTab { inner: value.inner }),
                "GuiTableControl" => {
                    SAPComponent::GuiTableControl(GuiTableControl { inner: value.inner })
                }
                "GuiTabStrip" => SAPComponent::GuiTabStrip(GuiTabStrip { inner: value.inner }),
                "Textedit" => SAPComponent::GuiTextedit(GuiTextedit { inner: value.inner }),
                "GuiTextField" => SAPComponent::GuiTextField(GuiTextField { inner: value.inner }),
                "GuiTitlebar" => SAPComponent::GuiTitlebar(GuiTitlebar { inner: value.inner }),
                "GuiToolbar" => SAPComponent::GuiToolbar(GuiToolbar { inner: value.inner }),
                "Tree" => SAPComponent::GuiTree(GuiTree { inner: value.inner }),
                "GuiUserArea" => SAPComponent::GuiUserArea(GuiUserArea { inner: value.inner }),
                "GuiVComponent" => {
                    SAPComponent::GuiVComponent(GuiVComponent { inner: value.inner })
                }
                "GuiVContainer" => {
                    SAPComponent::GuiVContainer(GuiVContainer { inner: value.inner })
                }
                "GuiVHViewSwitch" => {
                    SAPComponent::GuiVHViewSwitch(GuiVHViewSwitch { inner: value.inner })
                }
                _ => SAPComponent::GuiComponent(value),
            }
        } else {
            SAPComponent::GuiComponent(value)
        }
    }
}

impl From<VARIANT> for SAPComponent {
    fn from(value: VARIANT) -> Self {
        let idisp: &IDispatch = value.variant_into().unwrap();
        Self::from(idisp.clone())
    }
}

com_shim! {
    struct GuiApplication: GuiContainer + GuiComponent {
        // TODO ActiveSession: Object,
        mut AllowSystemMessages: bool,
        mut ButtonbarVisible: bool,
        Children: GuiComponentCollection,
        ConnectionErrorText: String,
        Connections: GuiComponentCollection,
        mut HistoryEnabled: bool,
        MajorVersion: i32,
        MinorVersion: i32,
        NewVisualDesign: bool,
        Patchlevel: i32,
        Revision: i32,
        mut StatusbarVisible: bool,
        mut TitlebarVisible: bool,
        mut ToolbarVisible: bool,
        Utils: GuiUtils,

        fn AddHistoryEntry(String, String) -> bool,
        fn CreateGuiCollection() -> GuiCollection,
        fn DropHistory() -> bool,
        fn Ignore(i16),
        fn OpenConnection(String) -> GuiComponent,
        fn OpenConnectionByConnectionString(String) -> GuiComponent,
    }
}

com_shim! {
    struct GuiBarChart: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        ChartCount: i32,

        fn BarCount(i32) -> i32,
        fn GetBarContent(i32, i32, i32) -> String,
        fn GetGridLineContent(i32, i32, i32) -> String,
        fn GridCount(i32) -> i32,
        fn LinkCount(i32) -> i32,
        fn SendData(String),
    }
}

com_shim! {
    struct GuiBox: GuiVComponent + GuiComponent {
        CharHeight: i32,
        CharLeft: i32,
        CharTop: i32,
        CharWidth: i32,
    }
}

com_shim! {
    struct GuiButton: GuiVComponent + GuiComponent {
        Emphasized: bool,
        LeftLabel: GuiComponent,
        RightLabel: GuiComponent,

        fn Press(),
    }
}

com_shim! {
    struct GuiCalendar: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        endSelection: String,
        mut FirstVisibleDate: String,
        mut FocusDate: String,
        FocusedElement: i32,
        horizontal: bool,
        mut LastVisibleDate: String,
        mut SelectionInterval: String,
        startSelection: String,
        Today: String,

        fn ContextMenu(i32, i32, i32, String, String),
        fn CreateDate(i32, i32, i32),
        fn GetColor(String) -> i32,
        fn GetColorInfo(i32) -> String,
        fn GetDateTooltip(String) -> String,
        fn GetDay(String) -> i32,
        fn GetMonth(String) -> i32,
        fn GetWeekday(String) -> String,
        fn GetWeekNumber(String) -> i32,
        fn GetYear(String) -> i32,
        fn IsWeekend(String) -> bool,
        fn SelectMonth(i32, i32),
        fn SelectRange(String, String),
        fn SelectWeek(i32, i32),
    }
}

com_shim! {
    struct GuiChart: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        fn ValueChange(i32, i32, String, String, bool, String, String, i32),
    }
}

com_shim! {
    struct GuiCheckBox: GuiVComponent + GuiComponent {
        ColorIndex: i32,
        ColorIntensified: i32,
        ColorInverse: bool,
        Flushing: bool,
        IsLeftLabel: bool,
        IsListElement: bool,
        IsRightLabel: bool,
        LeftLabel: GuiComponent,
        RightLabel: GuiComponent,
        RowText: String,
        mut Selected: bool,

        fn GetListProperty(String) -> String,
        fn GetListPropertyNonRec(String) -> String,
    }
}

com_shim! {
    struct GuiCollection {
        Count: i32,
        Length: i32,
        r#Type: String,
        TypeAsNumber: i32,

        // TODO fn Add(Variant),
        fn ElementAt(i32) -> GuiComponent,
    }
}

com_shim! {
    struct GuiColorSelector: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        fn ChangeSelection(i16),
    }
}

com_shim! {
    struct GuiComboBox: GuiVComponent + GuiComponent {
        CharHeight: i32,
        CharLeft: i32,
        CharTop: i32,
        CharWidth: i32,
        CurListBoxEntry: GuiComponent,
        Entries: GuiCollection,
        Flushing: bool,
        Highlighted: bool,
        IsLeftLabel: bool,
        IsListBoxActive: bool,
        IsRightLabel: bool,
        mut Key: String,
        LeftLabel: GuiComponent,
        Required: bool,
        RightLabel: GuiComponent,
        ShowKey: bool,
        Text: String,
        mut Value: String,

        fn SetKeySpace(),
    }
}

com_shim! {
    struct GuiComboBoxControl: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell {
        CurListBoxEntry: GuiComponent,
        Entries: GuiCollection,
        IsListBoxActive: bool,
        LabelText: String,
        mut Selected: String,
        Text: String,

        fn FireSelected(),
    }
}

com_shim! {
    struct GuiComboBoxEntry {
        Key: String,
        Pos: i32,
        Value: String,
    }
}

com_shim! {
    struct GuiComponent {
        ContainerType: bool,
        Id: String,
        Name: String,
        r#Type: String,
        TypeAsNumber: i32,
    }
}

com_shim! {
    struct GuiComponentCollection {
        Count: i32,
        Length: i32,
        r#Type: String,
        TypeAsNumber: i32,

        fn ElementAt(i32) -> GuiComponent,
    }
}

com_shim! {
    struct GuiConnection: GuiContainer + GuiComponent {
        Children: GuiComponentCollection,
        ConnectionString: String,
        Description: String,
        DisabledByServer: bool,
        Sessions: GuiComponentCollection,

        fn CloseConnection(),
        fn CloseSession(String),
    }
}

com_shim! {
    struct GuiContainer: GuiComponent {
        Children: GuiComponentCollection,

        fn FindById(String) -> GuiComponent,
    }
}

com_shim! {
    struct GuiContainerShell: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        AccDescription: String,
    }
}

com_shim! {
    struct GuiCTextField: GuiTextField + GuiVComponent + GuiComponent { }
}

com_shim! {
    struct GuiCustomControl: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        CharHeight: i32,
        CharLeft: i32,
        CharTop: i32,
        CharWidth: i32,
    }
}

com_shim! {
    struct GuiDialogShell: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        Title: String,

        fn Close(),
    }
}

com_shim! {
    struct GuiDockShell: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        AccDescription: String,
        DockerIsVertical: bool,
        mut DockerPixelSize: i32,
    }
}

com_shim! {
    struct GuiEAIViewer2D: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        mut AnnoutationEnabled: i32,
        mut AnnotationMode: i16,
        mut RedliningStream: String,

        fn annotationTextRequest(String),
    }
}

com_shim! {
    struct GuiEAIViewer3D: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell { }
}

com_shim! {
    struct GuiFrameWindow: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        mut ElementVisualizationMode: bool,
        GuiFocus: GuiComponent,
        Handle: i32,
        Iconic: bool,
        SystemFocus: GuiComponent,
        WorkingPaneHeight: i32,
        WorkingPaneWidth: i32,

        fn Close(),
        fn CompBitmap(String, String) -> i32,
        fn HardCopy(String, i16) -> String,
        fn Iconify(),
        fn IsVKeyAllowed(i16) -> bool,
        fn JumpBackward(),
        fn JumpForward(),
        fn Maximize(),
        fn Restore(),
        fn SendVKey(i16),
        fn ShowMessageBox(String, String, i32, i32) -> i32,
        fn TabBackward(),
        fn TabForward(),
    }
}

com_shim! {
    struct GuiGOSShell: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent { }
}

com_shim! {
    struct GuiGraphAdapt: GuiVComponent + GuiVContainer + GuiContainer + GuiComponent + GuiShell { }
}

com_shim! {
    struct GuiGridView: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell {
        ColumnCount: i32,
        // TODO mut ColumnOrder: Object,
        mut CurrentCellColumn: String,
        mut CurrentCellRow: i32,
        mut FirstVisibleColumn: String,
        mut FirstVisibleRow: i32,
        FrozenColumnCount: i32,
        RowCount: i32,
        // TODO mut SelectedCells: Object,
        // TODO mut SelectedColumns: Object,
        mut SelectedRows: String,
        SelectionMode: String,
        Title: String,
        ToolbarButtonCount: i32,
        VisibleRowCount: i32,

        fn ClearSelection(),
        fn Click(i32, String),
        fn ClickCurrentCell(),
        fn ContextMenu(),
        fn CurrentCellMoved(),
        fn DeleteRows(String),
        fn DeselectColumn(String),
        fn DoubleClick(i32, String),
        fn DoubleClickCurrentCell(),
        fn DuplicateRows(String),
        fn GetCellChangeable(i32, String) -> bool,
        fn GetCellCheckBoxChecked(i32, String) -> bool,
        fn GetCellColor(i32, String) -> i32,
        fn GetCellHeight(i32, String) -> i32,
        fn GetCellHotspotType(i32, String) -> String,
        fn GetCellIcon(i32, String) -> String,
        fn GetCellLeft(i32, String) -> i32,
        fn GetCellListBoxCount(i32, String) -> i32,
        fn GetCellListBoxCurIndex(i32, String) -> String,
        fn GetCellMaxLength(i32, String) -> i32,
        fn GetCellState(i32, String) -> String,
        fn GetCellTooltip(i32, String) -> String,
        fn GetCellTop(i32, String) -> i32,
        fn GetCellType(i32, String) -> String,
        fn GetCellValue(i32, String) -> String,
        fn GetCellWidth(i32, String) -> i32,
        fn GetColorInfo(i32) -> String,
        fn GetColumnDataType(String) -> String,
        fn GetColumnOperationType(String) -> String,
        fn GetColumnPosition(String) -> i32,
        fn GetColumnSortType(String) -> String,
        // TODO fn GetColumnTitles(String) -> Object,
        fn GetColumnTooltip(String) -> String,
        fn GetColumnTotalType(String) -> String,
        fn GetDisplayedColumnTitle(String) -> String,
        fn GetRowTotalLevel(i32) -> i32,
        fn GetSymbolInfo(String) -> String,
        fn GetToolbarButtonChecked(i32) -> bool,
        fn GetToolbarButtonEnabled(i32) -> bool,
        fn GetToolbarButtonIcon(i32) -> String,
        fn GetToolbarButtonId(i32) -> String,
        fn GetToolbarButtonText(i32) -> String,
        fn GetToolbarButtonTooltip(i32) -> String,
        fn GetToolbarButtonType(i32) -> String,
        fn GetToolbarFocusButton() -> i32,
        fn HasCellF4Help(i32, String) -> bool,
        fn HistoryCurEntry(i32, String) -> String,
        fn HistoryCurIndex(i32, String) -> i32,
        fn HistoryIsActive(i32, String) -> bool,
        fn HistoryList(i32, String) -> GuiCollection,
        fn InsertRows(String),
        fn IsCellHotspot(i32, String) -> bool,
        fn IsCellSymbol(i32, String) -> bool,
        fn IsCellTotalExpander(i32, String) -> bool,
        fn IsColumnFiltered(String) -> bool,
        fn IsColumnKey(String) -> bool,
        fn IsTotalRowExpanded(i32) -> bool,
        fn ModifyCell(i32, String, String),
        fn ModifyCheckBox(i32, String, bool),
        fn MoveRows(i32, i32, i32),
        fn PressButton(i32, String),
        fn PressButtonCurrentCell(),
        fn PressColumnHeader(String),
        fn PressEnter(),
        fn PressF1(),
        fn PressF4(),
        fn PressToolbarButton(String),
        fn PressToolbarContextButton(String),
        fn PressTotalRow(i32, String),
        fn PressTotalRowCurrentCell(),
        fn SelectAll(),
        fn SelectColumn(String),
        fn SelectionChanged(),
        fn SelectToolbarMenuItem(String),
        fn SetColumnWidth(String, i32),
        fn SetCurrentCell(i32, String),
        fn TriggerModified(),
    }
}

com_shim! {
    struct GuiHTMLViewer: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell {
        // TODO BrowserHandle: Object,
        DocumentComplete: i32,

        fn ContextMenu(),
        fn GetBrowerControlType() -> i32,
        fn SapEvent(String, String, String),
    }
}

com_shim! {
    struct GuiInputFieldControl: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell {
        ButtonTooltip: String,
        FindButtonActivated: bool,
        HistoryCurEntry: String,
        HistoryCurIndex: i32,
        HistoryIsActive: bool,
        HistoryList: GuiCollection,
        LabelText: String,
        PromptText: String,

        fn Submit(),
    }
}

com_shim! {
    struct GuiLabel: GuiVComponent + GuiComponent {
        mut CaretPosition: i32,
        CharHeight: i32,
        CharLeft: i32,
        CharTop: i32,
        CharWidth: i32,
        ColorIndex: i32,
        ColorIntensified: bool,
        ColorInverse: bool,
        DisplayedText: String,
        Highlighted: String,
        IsHotspot: String,
        IsLeftLabel: bool,
        IsListElement: bool,
        IsRightLabel: bool,
        MaxLength: i32,
        Numerical: bool,
        RowText: String,

        fn GetListProperty(String) -> String,
        fn GetListPropertyNonRec(String) -> String,
    }
}

com_shim! {
    struct GuiMainWindow: GuiFrameWindow + GuiVComponent + GuiVContainer + GuiContainer + GuiComponent {
        mut ButtonbarVisible: bool,
        mut StatusbarVisible: bool,
        mut TitlebarVisible: bool,
        mut ToolbarVisible: bool,

        fn ResizeWorkingPane(i32, i32, bool),
        fn ResizeWorkingPaneEx(i32, i32, bool),
    }
}

com_shim! {
    struct GuiMap: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell { }
}

com_shim! {
    struct GuiMenu: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        fn Select(),
    }
}

com_shim! {
    struct GuiMenubar: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent { }
}

com_shim! {
    struct GuiMessageWindow: GuiVComponent + GuiComponent {
        FocusedButton: i32,
        HelpButtonHelpText: String,
        HelpButtonText: String,
        MessageText: String,
        MessageType: i32,
        OKButtonHelpText: String,
        OKButtonText: String,
        ScreenLeft: i32,
        ScreenTop: i32,
        Visible: bool,
    }
}

com_shim! {
    struct GuiModalWindow: GuiFrameWindow + GuiVComponent + GuiVContainer + GuiComponent + GuiContainer {
        fn IsPopupDialog() -> bool,
        fn PopupDialogText() -> String,
    }
}

com_shim! {
    struct GuiNetChart: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell {
        LinkCount: i32,
        NodeCount: i32,

        fn GetLinkContent(i32, i32) -> String,
        fn GetNodeContent(i32, i32) -> String,
        fn SendData(String),
    }
}

com_shim! {
    struct GuiOfficeIntegration: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        // TODO Document: Object,
        HostedApplication: i32,

        fn AppendRow(String, String),
        fn CloseDocument(i32, bool, bool),
        // TODO fn CustomEvent(i32, String, i32, ...),
        fn RemoveContent(String),
        fn SaveDocument(i32, bool),
        fn SetDocument(i32, String),
    }
}

com_shim! {
    struct GuiOkCodeField: GuiVComponent + GuiComponent {
        Opened: bool,

        fn PressF1(),
    }
}

com_shim! {
    struct GuiPasswordField: GuiTextField + GuiVComponent + GuiComponent { }
}

com_shim! {
    struct GuiPicture: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell {
        AltText: String,
        DisplayMode: String,
        Icon: String,
        Url: String,

        fn Click(),
        fn ClickControlArea(i32, i32),
        fn ClickPictureArea(i32, i32),
        fn ContextMenu(i32, i32),
        fn DoubleClick(),
        fn DoubleClickControlArea(i32, i32),
        fn DoubleClickPictureArea(i32, i32),
    }
}

com_shim! {
    struct GuiRadioButton: GuiVComponent + GuiComponent {
        CharHeight: i32,
        CharLeft: i32,
        CharTop: i32,
        CharWidth: i32,
        Flushing: bool,
        GroupCount: i32,
        GroupMembers: GuiComponentCollection,
        GroupPos: i32,
        IsLeftLabel: bool,
        IsRightLabel: bool,
        LeftLabel: GuiComponent,
        RightLabel: GuiComponent,
        Selected: bool,

        fn Select(),
    }
}

com_shim! {
    struct GuiSapChart: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell { }
}

com_shim! {
    struct GuiScrollbar {
        Maximum: i32,
        Minimum: i32,
        PageSize: i32,
        mut Position: i32,
        Range: i32,
    }
}

com_shim! {
    struct GuiScrollContainer: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        HorizontalScrollbar: GuiComponent,
        VerticalScrollbar: GuiComponent,
    }
}

com_shim! {
    struct GuiSession: GuiContainer + GuiComponent {
        mut AccEnhancedTabChain: bool,
        mut AccSymbolReplacement: bool,
        ActiveWindow: GuiComponent,
        mut Busy: bool,
        // TODO mut ErrorList: GuiCollection,
        Info: GuiSessionInfo,
        IsActive: bool,
        IsListBoxActive: bool,
        ListBoxCurrEntry: i32,
        ListBoxCurrEntryHeight: i32,
        ListBoxCurrEntryLeft: i32,
        ListBoxCurrEntryTop: i32,
        ListBoxCurrEntryWidth: i32,
        ListBoxHeight: i32,
        ListBoxLeft: i32,
        ListBoxTop: i32,
        ListBoxWidth: i32,
        mut PassportPreSystemId: String,
        mut PassportSystemId: String,
        mut PassportTransactionId: String,
        ProgressPercent: i32,
        ProgressText: String,
        mut Record: bool,
        mut RecordFile: String,
        mut SaveAsUnicode: bool,
        mut ShowDropdownKeys: bool,
        mut SuppressBackendPopups: bool,
        mut TestToolMode: i32,

        fn AsStdNumberFormat(String) -> String,
        fn ClearErrorList(),
        fn CreateSession(),
        fn EnableJawsEvents(),
        fn EndTransaction(),
        fn FindByPosition(i32, i32) -> GuiComponent,
        fn GetIconResourceName(String) -> String,
        fn GetObjectTree(String) -> String,
        fn GetVKeyDescription(i32) -> String,
        fn LockSessionUI(),
        fn SendCommand(String),
        fn SendCommandAsync(String),
        fn StartTransaction(String),
        fn UnlockSessionUI(),
    }
}

com_shim! {
    struct GuiSessionInfo {
        ApplicationServer: String,
        Client: String,
        Codepage: i32,
        Flushes: i32,
        Group: String,
        GuiCodepage: i32,
        I18NMode: bool,
        InterpretationTime: i32,
        IsLowSpeedConnection: bool,
        Language: String,
        MessageServer: String,
        Program: String,
        ResponseTime: i32,
        RoundTrips: i32,
        ScreenNumber: i32,
        ScriptingModeReadOnly: bool,
        ScriptingModeRecordingDisabled: bool,
        SessionNumber: i32,
        SystemName: String,
        SystemNumber: i32,
        SystemSessionId: String,
        Transaction: String,
        UI_GUIDELINE: String,
        User: String,
    }
}

com_shim! {
    struct GuiShell: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        AccDescription: String,
        DragDropSupported: bool,
        Handle: i32,
        OcxEvents: GuiCollection,
        SubType: String,

        fn SelectContextMenuItem(String),
        fn SelectContextMenuItemByPosition(String),
        fn SelectContextMenuItemByText(String),
    }
}

com_shim! {
    struct GuiSimpleContainer: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        IsListElement: bool,
        IsStepLoop: bool,
        IsStepLoopInTableStructure: bool,
        LoopColCount: i32,
        LoopCurrentCol: i32,
        LoopCurrentColCount: i32,
        LoopCurrentRow: i32,
        LoopRowCount: i32,

        fn GetListProperty(String) -> String,
        fn GetListPropertyNonRec(String) -> String,
    }
}

com_shim! {
    struct GuiSplit: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        IsVertical: i32,

        fn GetColSize(i32) -> i32,
        fn GetRowSize(i32) -> i32,
        fn SetColSize(i32, i32),
        fn SetRowSize(i32, i32),
    }
}

com_shim! {
    struct GuiSplitterContainer: GuiVContainer + GuiVComponent + GuiComponent + GuiContainer + GuiShell {
        IsVertical: bool,
        mut SashPosition: i32,
    }
}

com_shim! {
    struct GuiStage: GuiVComponent + GuiVContainer + GuiContainer + GuiShell + GuiComponent {
        fn ContextMenu(String),
        fn DoubleClick(String),
        fn SelectItems(String),
    }
}

com_shim! {
    struct GuiStatusbar: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer {
        Handle: i32,
        MessageAsPopup: bool,
        MessageHasLongText: i32,
        MessageId: String,
        MessageNumber: String,
        MessageParameter: String,
        MessageType: String,

        fn CreateSupportMessageClick(),
        fn DoubleClick(),
        fn ServiceRequestClick(),
    }
}

com_shim! {
    struct GuiStatusBarLink: GuiVComponent + GuiComponent {
        fn Press(),
    }
}

com_shim! {
    struct GuiStatusPane: GuiVComponent + GuiComponent {
        Children: GuiComponentCollection,
    }
}

com_shim! {
    struct GuiTab: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        fn ScrollToLeft(),
        fn Select(),
    }
}

com_shim! {
    struct GuiTableColumn: GuiComponentCollection {
        DefaultTooltip: String,
        Fixed: bool,
        IconName: String,
        Selected: bool,
        Title: String,
        Tooltip: String,
    }
}

com_shim! {
    struct GuiTableControl: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        CharHeight: i32,
        CharLeft: i32,
        CharTop: i32,
        CharWidth: i32,
        // TODO ColSelectMode: GuiTableSelectionType,
        Columns: GuiCollection,
        CurrentCol: i32,
        CurrentRow: i32,
        HorizontalScrollbar: GuiComponent,
        RowCount: i32,
        Rows: GuiCollection,
        // TODO RowSelectMode: GuiTableSelectionType,
        TableFieldName: String,
        VerticalScrollbar: GuiComponent,
        VisibleRowCount: i32,

        fn ConfigureLayout(),
        fn DeselectAllColumns(),
        fn GetAbsoluteRow(i32) -> GuiTableRow,
        fn GetCell(i32, i32) -> GuiComponent,
        fn ReorderTable(String),
        fn SelectAllColumns(),
    }
}

com_shim! {
    struct GuiTableRow: GuiComponentCollection {
        Selectable: bool,
        mut Selected: bool,
    }
}

com_shim! {
    struct GuiTabStrip: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        CharHeight: i32,
        CharLeft: i32,
        CharTop: i32,
        CharWidth: i32,
        LeftTab: GuiComponent,
        SelectedTab: GuiComponent,
    }
}

com_shim! {
    struct GuiTextedit: GuiShell + GuiVComponent + GuiVContainer + GuiContainer + GuiComponent {
        CurrentColumn: i32,
        CurrentLine: i32,
        mut FirstVisibleLine: i32,
        LastVisibleLine: i32,
        LineCount: i32,
        NumberOfUnprotectedTextParts: i32,
        SelectedText: String,
        SelectionEndColumn: i32,
        SelectionEndLine: i32,
        SelectionIndexEnd: i32,
        SelectionIndexStart: i32,
        SelectionStartColumn: i32,
        SelectionStartLine: i32,

        fn ContextMenu(),
        fn DoubleClick(),
        fn GetLineText(i32) -> String,
        fn GetUnprotectedTextPart(i32) -> String,
        fn IsBreakpointLine(i32) -> bool,
        fn IsCommentLine(i32) -> bool,
        fn IsHighlightedLine(i32) -> bool,
        fn IsProtectedLine(i32) -> bool,
        fn IsSelectedLine(i32) -> bool,
        fn ModifiedStatusChanged(bool),
        fn MultipleFilesDropped(),
        fn PressF1(),
        fn PressF4(),
        fn SetSelectionIndexes(i32, i32),
        fn SetUnprotectedTextPart(i32, String) -> bool,
        fn SingleFileDropped(String),
    }
}

com_shim! {
    struct GuiTextField: GuiVComponent + GuiComponent {
        mut CaretPosition: i32,
        DisplayedText: String,
        Highlighted: bool,
        HistoryCurEntry: String,
        HistoryCurIndex: i32,
        HistoryIsActive: bool,
        HistoryList: GuiCollection,
        IsHotspot: bool,
        IsLeftLabel: bool,
        IsListElement: bool,
        IsOField: bool,
        IsRightLabel: bool,
        LeftLabel: GuiComponent,
        MaxLength: i32,
        Numerical: bool,
        Required: bool,
        RightLabel: GuiComponent,

        fn GetListProperty(String) -> String,
        fn GetListPropertyNonRec(String) -> String,
    }
}

com_shim! {
    struct GuiTitlebar: GuiVComponent + GuiVContainer + GuiContainer + GuiComponent { }
}

com_shim! {
    struct GuiToolbar: GuiVComponent + GuiVContainer + GuiContainer + GuiComponent { }
}

com_shim! {
    struct GuiToolbarControl: GuiShell + GuiVComponent + GuiVContainer + GuiComponent + GuiContainer {
        ButtonCount: i32,
        FocusedButton: i32,

        fn GetButtonChecked(i32) -> bool,
        fn GetButtonEnabled(i32) -> bool,
        fn GetButtonIcon(i32) -> String,
        fn GetButtonId(i32) -> String,
        fn GetButtonText(i32) -> String,
        fn GetButtonTooltip(i32) -> String,
        fn GetButtonType(i32) -> String,
        fn GetMenuItemIdFromPosition(i32) -> String,
        fn PressButton(String),
        fn PressContextButton(String),
        fn SelectMenuItem(String),
        fn SelectMenuItemByText(String),
    }
}

com_shim! {
    struct GuiTree: GuiShell + GuiVContainer + GuiVComponent + GuiComponent + GuiContainer {
        // TODO ColumnOrder: Object,
        HierarchyHeaderWidth: i32,
        SelectedNode: String,
        TopNode: String,

        fn ChangeCheckbox(String, String, bool),
        fn ClickLink(String, String),
        fn CollapseNode(String),
        fn DefaultContextMenu(),
        fn DoubleClickItem(String, String),
        fn DoubleClickNode(String),
        fn EnsureVisibleHorizontalItem(String, String),
        fn ExpandNode(String),
        fn FindNodeKeyByPath(String) -> String,
        fn GetAbapImage(String, String) -> String,
        // TODO fn GetAllNodeKeys() -> Object,
        fn GetCheckBoxState(String, String) -> bool,
        // TODO fn GetColumnCol(String) -> Object,
        // TODO fn GetColumnHeaders() -> Object,
        fn GetColumnIndexFromName(String) -> i32,
        // TODO fn GetColumnNames() -> Object,
        fn GetColumnTitleFromName(String) -> String,
        // TODO fn GetColumnTitles() -> Object,
        fn GetFocusedNodeKey() -> String,
        fn GetHierarchyLevel(String) -> i32,
        fn GetHierarchyTitle() -> String,
        fn GetIsDisabled(String, String) -> bool,
        fn GetIsEditable(String, String) -> bool,
        fn GetIsHighLighted(String, String) -> bool,
        fn GetItemHeight(String, String) -> i32,
        fn GetItemLeft(String, String) -> i32,
        fn GetItemStyle(String, String) -> i32,
        fn GetItemText(String, String) -> String,
        fn GetItemTextColor(String, String) -> u64,
        fn GetItemToolTip(String, String) -> String,
        fn GetItemTop(String, String) -> i32,
        fn GetItemType(String, String) -> i32,
        fn GetItemWidth(String, String) -> i32,
        fn GetListTreeNodeItemCount(String) -> i32,
        fn GetNextNodeKey(String) -> String,
        fn GetNodeAbapImage(String) -> String,
        fn GetNodeChildrenCount(String) -> i32,
        fn GetNodeChildrenCountByPath(String) -> i32,
        fn GetNodeHeight(String) -> i32,
        fn GetNodeIndex(String) -> i32,
        // TODO fn GetNodeItemHeaders(String) -> Object,
        fn GetNodeKeyByPath(String) -> String,
        fn GetNodeLeft(String) -> i32,
        fn GetNodePathByKey(String) -> String,
        // TODO fn GetNodesCol() -> Object,
        fn GetNodeStyle(String) -> i32,
        fn GetNodeTextByKey(String) -> String,
        fn GetNodeTextByPath(String) -> String,
        fn GetNodeTextColor(String) -> u64,
        fn GetNodeToolTip(String) -> String,
        fn GetNodeTop(String) -> i32,
        fn GetNodeWidth(String) -> i32,
        fn GetParent(String) -> String,
        fn GetPreviousNodeKey(String) -> String,
        // TODO fn GetSelectedNodes() -> Object,
        fn GetSelectionMode() -> i16,
        fn GetStyleDescription(i32) -> String,
        // TODO fn GetSubNodesCol(String) -> Object,
        fn GetTreeType() -> i32,
        fn HeaderContextMenu(String),
        fn IsFolder(String) -> bool,
        fn IsFolderExpandable(String) -> bool,
        fn IsFolderExpanded(String) -> bool,
        fn ItemContextMenu(String, String),
        fn NodeContextMenu(String),
        fn PressButton(String, String),
        fn PressHeader(String),
        fn PressKey(String),
        fn SelectColumn(String),
        fn SelectedItemColumn() -> String,
        fn SelectedItemNode() -> String,
        fn SelectItem(String, String),
        fn SelectNode(String),
        fn SetCheckBoxState(String, String, i32),
        fn SetColumnWidth(String, i32),
        fn UnselectAll(),
        fn UnselectColumn(String),
        fn UnselectNode(String),
    }
}

com_shim! {
    struct GuiUserArea: GuiVContainer + GuiVComponent + GuiComponent + GuiContainer {
        HorizontalScrollbar: GuiComponent,
        IsOTFPreview: bool,
        VerticalScrollbar: GuiComponent,

        fn FindByLabel(String, String) -> GuiComponent,
        fn ListNavigate(String),
    }
}

com_shim! {
    struct GuiUtils {
        MESSAGE_OPTION_OK: i32,
        MESSAGE_OPTION_OKCANCEL: i32,
        MESSAGE_OPTION_YESNO: i32,
        MESSAGE_RESULT_CANCEL: i32,
        MESSAGE_RESULT_NO: i32,
        MESSAGE_RESULT_OK: i32,
        MESSAGE_RESULT_YES: i32,
        MESSAGE_TYPE_ERROR: i32,
        MESSAGE_TYPE_INFORMATION: i32,
        MESSAGE_TYPE_PLAIN: i32,
        MESSAGE_TYPE_QUESTION: i32,
        MESSAGE_TYPE_WARNING: i32,

        fn CloseFile(i32),
        fn OpenFile(String) -> i32,
        fn ShowMessageBox(String, String, i32, i32) -> i32,
        fn Write(i32, String),
        fn WriteLine(i32, String),
    }
}

com_shim! {
    struct GuiVComponent: GuiComponent {
        AccLabelCollection: GuiComponentCollection,
        AccText: String,
        AccTextOnRequest: String,
        AccTooltip: String,
        Changeable: bool,
        DefaultTooltip: String,
        Height: i32,
        IconName: String,
        IsSymbolFont: bool,
        Left: i32,
        Modified: bool,
        ParentFrame: GuiComponent,
        ScreenLeft: i32,
        ScreenTop: i32,
        mut Text: String,
        Tooltip: String,
        Top: i32,
        Width: i32,

        fn DumpState(String) -> GuiCollection,
        fn SetFocus(),
        fn Visualize(bool) -> bool,
    }
}

com_shim! {
    struct GuiVContainer: GuiVComponent + GuiComponent + GuiContainer {
        fn FindAllByName(String, String) -> GuiComponentCollection,
        fn FindAllByNameEx(String, i32) -> GuiComponentCollection,
        fn FindByName(String, String) -> GuiComponent,
        fn FindByNameEx(String, String) -> GuiComponent,
    }
}

com_shim! {
    struct GuiVHViewSwitch: GuiVComponent + GuiComponent {}
}
