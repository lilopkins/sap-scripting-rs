use com_shim::{com_shim, IDispatchExt, VariantExt};
use windows::{core::*, Win32::System::Com::*};

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
        if let Ok(kind) = value._type() {
            log::debug!("Converting component {kind} to SAPComponent.");
            match kind.as_str() {
                "GuiApplication" => {
                    SAPComponent::GuiApplication(GuiApplication { inner: value.inner })
                }
                "GuiBarChart" => SAPComponent::GuiBarChart(GuiBarChart { inner: value.inner }),
                "GuiBox" => SAPComponent::GuiBox(GuiBox { inner: value.inner }),
                "GuiButton" => SAPComponent::GuiButton(GuiButton { inner: value.inner }),
                "GuiCalendar" => SAPComponent::GuiCalendar(GuiCalendar { inner: value.inner }),
                "GuiChart" => SAPComponent::GuiChart(GuiChart { inner: value.inner }),
                "GuiCheckBox" => SAPComponent::GuiCheckBox(GuiCheckBox { inner: value.inner }),
                "GuiColorSelector" => {
                    SAPComponent::GuiColorSelector(GuiColorSelector { inner: value.inner })
                }
                "GuiComboBox" => SAPComponent::GuiComboBox(GuiComboBox { inner: value.inner }),
                "GuiComboBoxControl" => {
                    SAPComponent::GuiComboBoxControl(GuiComboBoxControl { inner: value.inner })
                }
                "GuiComboBoxEntry" => {
                    SAPComponent::GuiComboBoxEntry(GuiComboBoxEntry { inner: value.inner })
                }
                "GuiComponent" => SAPComponent::GuiComponent(value.into()),
                "GuiConnection" => {
                    SAPComponent::GuiConnection(GuiConnection { inner: value.inner })
                }
                "GuiContainer" => SAPComponent::GuiContainer(GuiContainer { inner: value.inner }),
                "GuiContainerShell" => {
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
                "GuiEAIViewer2D" => {
                    SAPComponent::GuiEAIViewer2D(GuiEAIViewer2D { inner: value.inner })
                }
                "GuiEAIViewer3D" => {
                    SAPComponent::GuiEAIViewer3D(GuiEAIViewer3D { inner: value.inner })
                }
                "GuiFrameWindow" => {
                    SAPComponent::GuiFrameWindow(GuiFrameWindow { inner: value.inner })
                }
                "GuiGOSShell" => SAPComponent::GuiGOSShell(GuiGOSShell { inner: value.inner }),
                "GuiGraphAdapt" => {
                    SAPComponent::GuiGraphAdapt(GuiGraphAdapt { inner: value.inner })
                }
                "GuiGridView" => SAPComponent::GuiGridView(GuiGridView { inner: value.inner }),
                "GuiHTMLViewer" => {
                    SAPComponent::GuiHTMLViewer(GuiHTMLViewer { inner: value.inner })
                }
                "GuiInputFieldControl" => {
                    SAPComponent::GuiInputFieldControl(GuiInputFieldControl { inner: value.inner })
                }
                "GuiLabel" => SAPComponent::GuiLabel(GuiLabel { inner: value.inner }),
                "GuiMainWindow" => {
                    SAPComponent::GuiMainWindow(GuiMainWindow { inner: value.inner })
                }
                "GuiMap" => SAPComponent::GuiMap(GuiMap { inner: value.inner }),
                "GuiMenu" => SAPComponent::GuiMenu(GuiMenu { inner: value.inner }),
                "GuiMenubar" => SAPComponent::GuiMenubar(GuiMenubar { inner: value.inner }),
                "GuiModalWindow" => {
                    SAPComponent::GuiModalWindow(GuiModalWindow { inner: value.inner })
                }
                "GuiNetChart" => SAPComponent::GuiNetChart(GuiNetChart { inner: value.inner }),
                "GuiOfficeIntegration" => {
                    SAPComponent::GuiOfficeIntegration(GuiOfficeIntegration { inner: value.inner })
                }
                "GuiOkCodeField" => {
                    SAPComponent::GuiOkCodeField(GuiOkCodeField { inner: value.inner })
                }
                "GuiPasswordField" => {
                    SAPComponent::GuiPasswordField(GuiPasswordField { inner: value.inner })
                }
                "GuiPicture" => SAPComponent::GuiPicture(GuiPicture { inner: value.inner }),
                "GuiRadioButton" => {
                    SAPComponent::GuiRadioButton(GuiRadioButton { inner: value.inner })
                }
                "GuiSapChart" => SAPComponent::GuiSapChart(GuiSapChart { inner: value.inner }),
                "GuiScrollbar" => SAPComponent::GuiScrollbar(GuiScrollbar { inner: value.inner }),
                "GuiScrollContainer" => {
                    SAPComponent::GuiScrollContainer(GuiScrollContainer { inner: value.inner })
                }
                "GuiSession" => SAPComponent::GuiSession(GuiSession { inner: value.inner }),
                "GuiShell" => SAPComponent::GuiShell(GuiShell { inner: value.inner }),
                "GuiSimpleContainer" => {
                    SAPComponent::GuiSimpleContainer(GuiSimpleContainer { inner: value.inner })
                }
                "GuiSplit" => SAPComponent::GuiSplit(GuiSplit { inner: value.inner }),
                "GuiSplitterContainer" => {
                    SAPComponent::GuiSplitterContainer(GuiSplitterContainer { inner: value.inner })
                }
                "GuiStage" => SAPComponent::GuiStage(GuiStage { inner: value.inner }),
                "GuiStatusbar" => SAPComponent::GuiStatusbar(GuiStatusbar { inner: value.inner }),
                "GuiStatusPane" => {
                    SAPComponent::GuiStatusPane(GuiStatusPane { inner: value.inner })
                }
                "GuiTab" => SAPComponent::GuiTab(GuiTab { inner: value.inner }),
                "GuiTableControl" => {
                    SAPComponent::GuiTableControl(GuiTableControl { inner: value.inner })
                }
                "GuiTabStrip" => SAPComponent::GuiTabStrip(GuiTabStrip { inner: value.inner }),
                "GuiTextedit" => SAPComponent::GuiTextedit(GuiTextedit { inner: value.inner }),
                "GuiTextField" => SAPComponent::GuiTextField(GuiTextField { inner: value.inner }),
                "GuiTitlebar" => SAPComponent::GuiTitlebar(GuiTitlebar { inner: value.inner }),
                "GuiToolbar" => SAPComponent::GuiToolbar(GuiToolbar { inner: value.inner }),
                "GuiTree" => SAPComponent::GuiTree(GuiTree { inner: value.inner }),
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
        let idisp = value.to_idispatch().unwrap();
        Self::from(idisp.clone())
    }
}

com_shim! {
    class GuiApplication: GuiContainer + GuiComponent {
        // TODO ActiveSession: Object,
        mut AllowSystemMessages: bool,
        mut ButtonbarVisible: bool,
        Children: GuiComponentCollection,
        ConnectionErrorText: String,
        Connections: GuiComponentCollection,
        mut HistoryEnabled: bool,
        MajorVersion: i64,
        MinorVersion: i64,
        NewVisualDesign: bool,
        Patchlevel: i64,
        Revision: i64,
        mut StatusbarVisible: bool,
        mut TitlebarVisible: bool,
        mut ToolbarVisible: bool,
        Utils: GuiUtils,

        fn AddHistoryEntry(String, String) -> bool,
        fn CreateGuiCollection() -> GuiCollection,
        fn DropHistory() -> bool,
        fn Ignore(i32),
        fn OpenConnection(String) -> SAPComponent,
        fn OpenConnectionByConnectionString(String) -> SAPComponent,
    }
}

com_shim! {
    class GuiBarChart: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        ChartCount: i64,

        fn BarCount(i64) -> i64,
        fn GetBarContent(i64, i64, i64) -> String,
        fn GetGridLineContent(i64, i64, i64) -> String,
        fn GridCount(i64) -> i64,
        fn LinkCount(i64) -> i64,
        fn SendData(String),
    }
}

com_shim! {
    class GuiBox: GuiVComponent + GuiComponent {
        CharHeight: i64,
        CharLeft: i64,
        CharTop: i64,
        CharWidth: i64,
    }
}

com_shim! {
    class GuiButton: GuiVComponent + GuiComponent {
        Emphasized: bool,
        LeftLabel: SAPComponent,
        RightLabel: SAPComponent,

        fn Press(),
    }
}

com_shim! {
    class GuiCalendar: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        endSelection: String,
        mut FirstVisibleDate: String,
        mut FocusDate: String,
        FocusedElement: i64,
        horizontal: bool,
        mut LastVisibleDate: String,
        mut SelectionInterval: String,
        startSelection: String,
        Today: String,

        fn ContextMenu(i64, i64, i64, String, String),
        fn CreateDate(i64, i64, i64),
        fn GetColor(String) -> i64,
        fn GetColorInfo(i64) -> String,
        fn GetDateTooltip(String) -> String,
        fn GetDay(String) -> i64,
        fn GetMonth(String) -> i64,
        fn GetWeekday(String) -> String,
        fn GetWeekNumber(String) -> i64,
        fn GetYear(String) -> i64,
        fn IsWeekend(String) -> bool,
        fn SelectMonth(i64, i64),
        fn SelectRange(String, String),
        fn SelectWeek(i64, i64),
    }
}

com_shim! {
    class GuiChart: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        fn ValueChange(i64, i64, String, String, bool, String, String, i64),
    }
}

com_shim! {
    class GuiCheckBox: GuiVComponent + GuiComponent {
        ColorIndex: i64,
        ColorIntensified: i64,
        ColorInverse: bool,
        Flushing: bool,
        IsLeftLabel: bool,
        IsListElement: bool,
        IsRightLabel: bool,
        LeftLabel: SAPComponent,
        RightLabel: SAPComponent,
        RowText: String,
        mut Selected: bool,

        fn GetListProperty(String) -> String,
        fn GetListPropertyNonRec(String) -> String,
    }
}

com_shim! {
    class GuiCollection {
        Count: i64,
        Length: i64,
        Type: String,
        TypeAsNumber: i64,

        // TODO fn Add(Variant),
        fn ElementAt(i64) -> SAPComponent,
    }
}

com_shim! {
    class GuiColorSelector: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        fn ChangeSelection(i32),
    }
}

com_shim! {
    class GuiComboBox: GuiVComponent + GuiComponent {
        CharHeight: i64,
        CharLeft: i64,
        CharTop: i64,
        CharWidth: i64,
        CurListBoxEntry: SAPComponent,
        Entries: GuiCollection,
        Flushing: bool,
        Highlighted: bool,
        IsLeftLabel: bool,
        IsListBoxActive: bool,
        IsRightLabel: bool,
        mut Key: String,
        LeftLabel: SAPComponent,
        Required: bool,
        RightLabel: SAPComponent,
        ShowKey: bool,
        Text: String,
        mut Value: String,

        fn SetKeySpace(),
    }
}

com_shim! {
    class GuiComboBoxControl: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell {
        CurListBoxEntry: SAPComponent,
        Entries: GuiCollection,
        IsListBoxActive: bool,
        LabelText: String,
        mut Selected: String,
        Text: String,

        fn FireSelected(),
    }
}

com_shim! {
    class GuiComboBoxEntry {
        Key: String,
        Pos: i64,
        Value: String,
    }
}

com_shim! {
    class GuiComponent {
        ContainerType: bool,
        Id: String,
        Name: String,
        Type: String,
        TypeAsNumber: i64,
    }
}

com_shim! {
    class GuiComponentCollection {
        Count: i64,
        Length: i64,
        Type: String,
        TypeAsNumber: i64,

        fn ElementAt(i64) -> SAPComponent,
    }
}

com_shim! {
    class GuiConnection: GuiContainer + GuiComponent {
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
    class GuiContainer: GuiComponent {
        Children: GuiComponentCollection,

        fn FindById(String) -> SAPComponent,
    }
}

com_shim! {
    class GuiContainerShell: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        AccDescription: String,
    }
}

com_shim! {
    class GuiCTextField: GuiTextField + GuiVComponent + GuiComponent { }
}

com_shim! {
    class GuiCustomControl: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        CharHeight: i64,
        CharLeft: i64,
        CharTop: i64,
        CharWidth: i64,
    }
}

com_shim! {
    class GuiDialogShell: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        Title: String,

        fn Close(),
    }
}

com_shim! {
    class GuiDockShell: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        AccDescription: String,
        DockerIsVertical: bool,
        mut DockerPixelSize: i64,
    }
}

com_shim! {
    class GuiEAIViewer2D: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        mut AnnoutationEnabled: i64,
        mut AnnotationMode: i32,
        mut RedliningStream: String,

        fn annotationTextRequest(String),
    }
}

com_shim! {
    class GuiEAIViewer3D: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell { }
}

com_shim! {
    class GuiFrameWindow: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        mut ElementVisualizationMode: bool,
        GuiFocus: SAPComponent,
        Handle: i64,
        Iconic: bool,
        SystemFocus: SAPComponent,
        WorkingPaneHeight: i64,
        WorkingPaneWidth: i64,

        fn Close(),
        fn CompBitmap(String, String) -> i64,
        fn Iconify(),
        fn IsVKeyAllowed(i32) -> bool,
        fn JumpBackward(),
        fn JumpForward(),
        fn Maximize(),
        fn Restore(),
        fn SendVKey(i32),
        fn ShowMessageBox(String, String, i64, i64) -> i64,
        fn TabBackward(),
        fn TabForward(),
    }
}

com_shim! {
    class GuiGOSShell: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent { }
}

com_shim! {
    class GuiGraphAdapt: GuiVComponent + GuiVContainer + GuiContainer + GuiComponent + GuiShell { }
}

com_shim! {
    class GuiGridView: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell {
        ColumnCount: i64,
        // TODO mut ColumnOrder: Object,
        mut CurrentCellColumn: String,
        mut CurrentCellRow: i64,
        mut FirstVisibleColumn: String,
        mut FirstVisibleRow: i64,
        FrozenColumnCount: i64,
        RowCount: i64,
        // TODO mut SelectedCells: Object,
        // TODO mut SelectedColumns: Object,
        mut SelectedRows: String,
        SelectionMode: String,
        Title: String,
        ToolbarButtonCount: i64,
        VisibleRowCount: i64,

        fn ClearSelection(),
        fn Click(i64, String),
        fn ClickCurrentCell(),
        fn ContextMenu(),
        fn CurrentCellMoved(),
        fn DeleteRows(String),
        fn DeselectColumn(String),
        fn DoubleClick(i64, String),
        fn DoubleClickCurrentCell(),
        fn DuplicateRows(String),
        fn GetCellChangeable(i64, String) -> bool,
        fn GetCellCheckBoxChecked(i64, String) -> bool,
        fn GetCellColor(i64, String) -> i64,
        fn GetCellHeight(i64, String) -> i64,
        fn GetCellHotspotType(i64, String) -> String,
        fn GetCellIcon(i64, String) -> String,
        fn GetCellLeft(i64, String) -> i64,
        fn GetCellListBoxCount(i64, String) -> i64,
        fn GetCellListBoxCurIndex(i64, String) -> String,
        fn GetCellMaxLength(i64, String) -> i64,
        fn GetCellState(i64, String) -> String,
        fn GetCellTooltip(i64, String) -> String,
        fn GetCellTop(i64, String) -> i64,
        fn GetCellType(i64, String) -> String,
        fn GetCellValue(i64, String) -> String,
        fn GetCellWidth(i64, String) -> i64,
        fn GetColorInfo(i64) -> String,
        fn GetColumnDataType(String) -> String,
        fn GetColumnOperationType(String) -> String,
        fn GetColumnPosition(String) -> i64,
        fn GetColumnSortType(String) -> String,
        // TODO fn GetColumnTitles(String) -> Object,
        fn GetColumnTooltip(String) -> String,
        fn GetColumnTotalType(String) -> String,
        fn GetDisplayedColumnTitle(String) -> String,
        fn GetRowTotalLevel(i64) -> i64,
        fn GetSymbolInfo(String) -> String,
        fn GetToolbarButtonChecked(i64) -> bool,
        fn GetToolbarButtonEnabled(i64) -> bool,
        fn GetToolbarButtonIcon(i64) -> String,
        fn GetToolbarButtonId(i64) -> String,
        fn GetToolbarButtonText(i64) -> String,
        fn GetToolbarButtonTooltip(i64) -> String,
        fn GetToolbarButtonType(i64) -> String,
        fn GetToolbarFocusButton() -> i64,
        fn HasCellF4Help(i64, String) -> bool,
        fn HistoryCurEntry(i64, String) -> String,
        fn HistoryCurIndex(i64, String) -> i64,
        fn HistoryIsActive(i64, String) -> bool,
        fn HistoryList(i64, String) -> GuiCollection,
        fn InsertRows(String),
        fn IsCellHotspot(i64, String) -> bool,
        fn IsCellSymbol(i64, String) -> bool,
        fn IsCellTotalExpander(i64, String) -> bool,
        fn IsColumnFiltered(String) -> bool,
        fn IsColumnKey(String) -> bool,
        fn IsTotalRowExpanded(i64) -> bool,
        fn ModifyCell(i64, String, String),
        fn ModifyCheckBox(i64, String, bool),
        fn MoveRows(i64, i64, i64),
        fn PressButton(i64, String),
        fn PressButtonCurrentCell(),
        fn PressColumnHeader(String),
        fn PressEnter(),
        fn PressF1(),
        fn PressF4(),
        fn PressToolbarButton(String),
        fn PressToolbarContextButton(String),
        fn PressTotalRow(i64, String),
        fn PressTotalRowCurrentCell(),
        fn SelectAll(),
        fn SelectColumn(String),
        fn SelectionChanged(),
        fn SelectToolbarMenuItem(String),
        fn SetColumnWidth(String, i64),
        fn SetCurrentCell(i64, String),
        fn TriggerModified(),
    }
}

com_shim! {
    class GuiHTMLViewer: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell {
        // TODO BrowserHandle: Object,
        DocumentComplete: i64,

        fn ContextMenu(),
        fn GetBrowerControlType() -> i64,
        fn SapEvent(String, String, String),
    }
}

com_shim! {
    class GuiInputFieldControl: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell {
        ButtonTooltip: String,
        FindButtonActivated: bool,
        HistoryCurEntry: String,
        HistoryCurIndex: i64,
        HistoryIsActive: bool,
        HistoryList: GuiCollection,
        LabelText: String,
        PromptText: String,

        fn Submit(),
    }
}

com_shim! {
    class GuiLabel: GuiVComponent + GuiComponent {
        mut CaretPosition: i64,
        CharHeight: i64,
        CharLeft: i64,
        CharTop: i64,
        CharWidth: i64,
        ColorIndex: i64,
        ColorIntensified: bool,
        ColorInverse: bool,
        DisplayedText: String,
        Highlighted: String,
        IsHotspot: String,
        IsLeftLabel: bool,
        IsListElement: bool,
        IsRightLabel: bool,
        MaxLength: i64,
        Numerical: bool,
        RowText: String,

        fn GetListProperty(String) -> String,
        fn GetListPropertyNonRec(String) -> String,
    }
}

com_shim! {
    class GuiMainWindow: GuiFrameWindow + GuiVComponent + GuiVContainer + GuiContainer + GuiComponent {
        mut ButtonbarVisible: bool,
        mut StatusbarVisible: bool,
        mut TitlebarVisible: bool,
        mut ToolbarVisible: bool,

        fn ResizeWorkingPane(i64, i64, bool),
        fn ResizeWorkingPaneEx(i64, i64, bool),
    }
}

com_shim! {
    class GuiMap: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell { }
}

com_shim! {
    class GuiMenu: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        fn Select(),
    }
}

com_shim! {
    class GuiMenubar: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent { }
}

com_shim! {
    class GuiMessageWindow: GuiVComponent + GuiComponent {
        FocusedButton: i64,
        HelpButtonHelpText: String,
        HelpButtonText: String,
        MessageText: String,
        MessageType: i64,
        OKButtonHelpText: String,
        OKButtonText: String,
        ScreenLeft: i64,
        ScreenTop: i64,
        Visible: bool,
    }
}

com_shim! {
    class GuiModalWindow: GuiFrameWindow + GuiVComponent + GuiVContainer + GuiComponent + GuiContainer {
        fn IsPopupDialog() -> bool,
        fn PopupDialogText() -> String,
    }
}

com_shim! {
    class GuiNetChart: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell {
        LinkCount: i64,
        NodeCount: i64,

        fn GetLinkContent(i64, i64) -> String,
        fn GetNodeContent(i64, i64) -> String,
        fn SendData(String),
    }
}

com_shim! {
    class GuiOfficeIntegration: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        // TODO Document: Object,
        HostedApplication: i64,

        fn AppendRow(String, String),
        fn CloseDocument(i64, bool, bool),
        // TODO fn CustomEvent(i64, String, i64, ...),
        fn RemoveContent(String),
        fn SaveDocument(i64, bool),
        fn SetDocument(i64, String),
    }
}

com_shim! {
    class GuiOkCodeField: GuiVComponent + GuiComponent {
        Opened: bool,

        fn PressF1(),
    }
}

com_shim! {
    class GuiPasswordField: GuiTextField + GuiVComponent + GuiComponent { }
}

com_shim! {
    class GuiPicture: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell {
        AltText: String,
        DisplayMode: String,
        Icon: String,
        Url: String,

        fn Click(),
        fn ClickControlArea(i64, i64),
        fn ClickPictureArea(i64, i64),
        fn ContextMenu(i64, i64),
        fn DoubleClick(),
        fn DoubleClickControlArea(i64, i64),
        fn DoubleClickPictureArea(i64, i64),
    }
}

com_shim! {
    class GuiRadioButton: GuiVComponent + GuiComponent {
        CharHeight: i64,
        CharLeft: i64,
        CharTop: i64,
        CharWidth: i64,
        Flushing: bool,
        GroupCount: i64,
        GroupMembers: GuiComponentCollection,
        GroupPos: i64,
        IsLeftLabel: bool,
        IsRightLabel: bool,
        LeftLabel: SAPComponent,
        RightLabel: SAPComponent,
        Selected: bool,

        fn Select(),
    }
}

com_shim! {
    class GuiSapChart: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell { }
}

com_shim! {
    class GuiScrollbar {
        Maximum: i64,
        Minimum: i64,
        PageSize: i64,
        mut Position: i64,
        Range: i64,
    }
}

com_shim! {
    class GuiScrollContainer: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        HorizontalScrollbar: SAPComponent,
        VerticalScrollbar: SAPComponent,
    }
}

com_shim! {
    class GuiSession: GuiContainer + GuiComponent {
        mut AccEnhancedTabChain: bool,
        mut AccSymbolReplacement: bool,
        ActiveWindow: SAPComponent,
        mut Busy: bool,
        // TODO mut ErrorList: GuiCollection,
        Info: GuiSessionInfo,
        IsActive: bool,
        IsListBoxActive: bool,
        ListBoxCurrEntry: i64,
        ListBoxCurrEntryHeight: i64,
        ListBoxCurrEntryLeft: i64,
        ListBoxCurrEntryTop: i64,
        ListBoxCurrEntryWidth: i64,
        ListBoxHeight: i64,
        ListBoxLeft: i64,
        ListBoxTop: i64,
        ListBoxWidth: i64,
        mut PassportPreSystemId: String,
        mut PassportSystemId: String,
        mut PassportTransactionId: String,
        ProgressPercent: i64,
        ProgressText: String,
        mut Record: bool,
        mut RecordFile: String,
        mut SaveAsUnicode: bool,
        mut ShowDropdownKeys: bool,
        mut SuppressBackendPopups: bool,
        mut TestToolMode: i64,

        fn AsStdNumberFormat(String) -> String,
        fn ClearErrorList(),
        fn CreateSession(),
        fn EnableJawsEvents(),
        fn EndTransaction(),
        fn FindByPosition(i64, i64) -> SAPComponent,
        fn GetIconResourceName(String) -> String,
        fn GetObjectTree(String) -> String,
        fn GetVKeyDescription(i64) -> String,
        fn LockSessionUI(),
        fn SendCommand(String),
        fn SendCommandAsync(String),
        fn StartTransaction(String),
        fn UnlockSessionUI(),
    }
}

com_shim! {
    class GuiSessionInfo {
        ApplicationServer: String,
        Client: String,
        Codepage: i64,
        Flushes: i64,
        Group: String,
        GuiCodepage: i64,
        I18NMode: bool,
        InterpretationTime: i64,
        IsLowSpeedConnection: bool,
        Language: String,
        MessageServer: String,
        Program: String,
        ResponseTime: i64,
        RoundTrips: i64,
        ScreenNumber: i64,
        ScriptingModeReadOnly: bool,
        ScriptingModeRecordingDisabled: bool,
        SessionNumber: i64,
        SystemName: String,
        SystemNumber: i64,
        SystemSessionId: String,
        Transaction: String,
        UI_GUIDELINE: String,
        User: String,
    }
}

com_shim! {
    class GuiShell: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        AccDescription: String,
        DragDropSupported: bool,
        Handle: i64,
        OcxEvents: GuiCollection,
        SubType: String,

        fn SelectContextMenuItem(String),
        fn SelectContextMenuItemByPosition(String),
        fn SelectContextMenuItemByText(String),
    }
}

com_shim! {
    class GuiSimpleContainer: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        IsListElement: bool,
        IsStepLoop: bool,
        IsStepLoopInTableStructure: bool,
        LoopColCount: i64,
        LoopCurrentCol: i64,
        LoopCurrentColCount: i64,
        LoopCurrentRow: i64,
        LoopRowCount: i64,

        fn GetListProperty(String) -> String,
        fn GetListPropertyNonRec(String) -> String,
    }
}

com_shim! {
    class GuiSplit: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        IsVertical: i64,

        fn GetColSize(i64) -> i64,
        fn GetRowSize(i64) -> i64,
        fn SetColSize(i64, i64),
        fn SetRowSize(i64, i64),
    }
}

com_shim! {
    class GuiSplitterContainer: GuiVContainer + GuiVComponent + GuiComponent + GuiContainer + GuiShell {
        IsVertical: bool,
        mut SashPosition: i64,
    }
}

com_shim! {
    class GuiStage: GuiVComponent + GuiVContainer + GuiContainer + GuiShell + GuiComponent {
        fn ContextMenu(String),
        fn DoubleClick(String),
        fn SelectItems(String),
    }
}

com_shim! {
    class GuiStatusbar: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer {
        Handle: i64,
        MessageAsPopup: bool,
        MessageHasLongText: i64,
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
    class GuiStatusBarLink: GuiVComponent + GuiComponent {
        fn Press(),
    }
}

com_shim! {
    class GuiStatusPane: GuiVComponent + GuiComponent {
        Children: GuiComponentCollection,
    }
}

com_shim! {
    class GuiTab: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        fn ScrollToLeft(),
        fn Select(),
    }
}

com_shim! {
    class GuiTableColumn: GuiComponentCollection {
        DefaultTooltip: String,
        Fixed: bool,
        IconName: String,
        Selected: bool,
        Title: String,
        Tooltip: String,
    }
}

com_shim! {
    class GuiTableControl: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        CharHeight: i64,
        CharLeft: i64,
        CharTop: i64,
        CharWidth: i64,
        // TODO ColSelectMode: GuiTableSelectionType,
        Columns: GuiCollection,
        CurrentCol: i64,
        CurrentRow: i64,
        HorizontalScrollbar: SAPComponent,
        RowCount: i64,
        Rows: GuiCollection,
        // TODO RowSelectMode: GuiTableSelectionType,
        TableFieldName: String,
        VerticalScrollbar: SAPComponent,
        VisibleRowCount: i64,

        fn ConfigureLayout(),
        fn DeselectAllColumns(),
        fn GetAbsoluteRow(i64) -> SAPComponent,
        fn GetCell(i64, i64) -> SAPComponent,
        fn ReorderTable(String),
        fn SelectAllColumns(),
    }
}

com_shim! {
    class GuiTableRow: GuiComponentCollection {
        Selectable: bool,
        mut Selected: bool,
    }
}

com_shim! {
    class GuiTabStrip: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        CharHeight: i64,
        CharLeft: i64,
        CharTop: i64,
        CharWidth: i64,
        LeftTab: SAPComponent,
        SelectedTab: SAPComponent,
    }
}

com_shim! {
    class GuiTextedit: GuiShell + GuiVComponent + GuiVContainer + GuiContainer + GuiComponent {
        CurrentColumn: i64,
        CurrentLine: i64,
        mut FirstVisibleLine: i64,
        LastVisibleLine: i64,
        LineCount: i64,
        NumberOfUnprotectedTextParts: i64,
        SelectedText: String,
        SelectionEndColumn: i64,
        SelectionEndLine: i64,
        SelectionIndexEnd: i64,
        SelectionIndexStart: i64,
        SelectionStartColumn: i64,
        SelectionStartLine: i64,

        fn ContextMenu(),
        fn DoubleClick(),
        fn GetLineText(i64) -> String,
        fn GetUnprotectedTextPart(i64) -> String,
        fn IsBreakpointLine(i64) -> bool,
        fn IsCommentLine(i64) -> bool,
        fn IsHighlightedLine(i64) -> bool,
        fn IsProtectedLine(i64) -> bool,
        fn IsSelectedLine(i64) -> bool,
        fn ModifiedStatusChanged(bool),
        fn MultipleFilesDropped(),
        fn PressF1(),
        fn PressF4(),
        fn SetSelectionIndexes(i64, i64),
        fn SetUnprotectedTextPart(i64, String) -> bool,
        fn SingleFileDropped(String),
    }
}

com_shim! {
    class GuiTextField: GuiVComponent + GuiComponent {
        mut CaretPosition: i64,
        DisplayedText: String,
        Highlighted: bool,
        HistoryCurEntry: String,
        HistoryCurIndex: i64,
        HistoryIsActive: bool,
        HistoryList: GuiCollection,
        IsHotspot: bool,
        IsLeftLabel: bool,
        IsListElement: bool,
        IsOField: bool,
        IsRightLabel: bool,
        LeftLabel: SAPComponent,
        MaxLength: i64,
        Numerical: bool,
        Required: bool,
        RightLabel: SAPComponent,

        fn GetListProperty(String) -> String,
        fn GetListPropertyNonRec(String) -> String,
    }
}

com_shim! {
    class GuiTitlebar: GuiVComponent + GuiVContainer + GuiContainer + GuiComponent { }
}

com_shim! {
    class GuiToolbar: GuiVComponent + GuiVContainer + GuiContainer + GuiComponent { }
}

com_shim! {
    class GuiToolbarControl: GuiShell + GuiVComponent + GuiVContainer + GuiComponent + GuiContainer {
        ButtonCount: i64,
        FocusedButton: i64,

        fn GetButtonChecked(i64) -> bool,
        fn GetButtonEnabled(i64) -> bool,
        fn GetButtonIcon(i64) -> String,
        fn GetButtonId(i64) -> String,
        fn GetButtonText(i64) -> String,
        fn GetButtonTooltip(i64) -> String,
        fn GetButtonType(i64) -> String,
        fn GetMenuItemIdFromPosition(i64) -> String,
        fn PressButton(String),
        fn PressContextButton(String),
        fn SelectMenuItem(String),
        fn SelectMenuItemByText(String),
    }
}

com_shim! {
    class GuiTree: GuiShell + GuiVContainer + GuiVComponent + GuiComponent + GuiContainer {
        // TODO ColumnOrder: Object,
        HierarchyHeaderWidth: i64,
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
        fn GetColumnIndexFromName(String) -> i64,
        // TODO fn GetColumnNames() -> Object,
        fn GetColumnTitleFromName(String) -> String,
        // TODO fn GetColumnTitles() -> Object,
        fn GetFocusedNodeKey() -> String,
        fn GetHierarchyLevel(String) -> i64,
        fn GetHierarchyTitle() -> String,
        fn GetIsDisabled(String, String) -> bool,
        fn GetIsEditable(String, String) -> bool,
        fn GetIsHighLighted(String, String) -> bool,
        fn GetItemHeight(String, String) -> i64,
        fn GetItemLeft(String, String) -> i64,
        fn GetItemStyle(String, String) -> i64,
        fn GetItemText(String, String) -> String,
        fn GetItemTextColor(String, String) -> u64,
        fn GetItemToolTip(String, String) -> String,
        fn GetItemTop(String, String) -> i64,
        fn GetItemType(String, String) -> i64,
        fn GetItemWidth(String, String) -> i64,
        fn GetListTreeNodeItemCount(String) -> i64,
        fn GetNextNodeKey(String) -> String,
        fn GetNodeAbapImage(String) -> String,
        fn GetNodeChildrenCount(String) -> i64,
        fn GetNodeChildrenCountByPath(String) -> i64,
        fn GetNodeHeight(String) -> i64,
        fn GetNodeIndex(String) -> i64,
        // TODO fn GetNodeItemHeaders(String) -> Object,
        fn GetNodeKeyByPath(String) -> String,
        fn GetNodeLeft(String) -> i64,
        fn GetNodePathByKey(String) -> String,
        // TODO fn GetNodesCol() -> Object,
        fn GetNodeStyle(String) -> i64,
        fn GetNodeTextByKey(String) -> String,
        fn GetNodeTextByPath(String) -> String,
        fn GetNodeTextColor(String) -> u64,
        fn GetNodeToolTip(String) -> String,
        fn GetNodeTop(String) -> i64,
        fn GetNodeWidth(String) -> i64,
        fn GetParent(String) -> String,
        fn GetPreviousNodeKey(String) -> String,
        // TODO fn GetSelectedNodes() -> Object,
        fn GetSelectionMode() -> i32,
        fn GetStyleDescription(i64) -> String,
        // TODO fn GetSubNodesCol(String) -> Object,
        fn GetTreeType() -> i64,
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
        fn SetCheckBoxState(String, String, i64),
        fn SetColumnWidth(String, i64),
        fn UnselectAll(),
        fn UnselectColumn(String),
        fn UnselectNode(String),
    }
}

com_shim! {
    class GuiUserArea: GuiVContainer + GuiVComponent + GuiComponent + GuiContainer {
        HorizontalScrollbar: SAPComponent,
        IsOTFPreview: bool,
        VerticalScrollbar: SAPComponent,

        fn FindByLabel(String, String) -> SAPComponent,
        fn ListNavigate(String),
    }
}

com_shim! {
    class GuiUtils {
        MESSAGE_OPTION_OK: i64,
        MESSAGE_OPTION_OKCANCEL: i64,
        MESSAGE_OPTION_YESNO: i64,
        MESSAGE_RESULT_CANCEL: i64,
        MESSAGE_RESULT_NO: i64,
        MESSAGE_RESULT_OK: i64,
        MESSAGE_RESULT_YES: i64,
        MESSAGE_TYPE_ERROR: i64,
        MESSAGE_TYPE_INFORMATION: i64,
        MESSAGE_TYPE_PLAIN: i64,
        MESSAGE_TYPE_QUESTION: i64,
        MESSAGE_TYPE_WARNING: i64,

        fn CloseFile(i64),
        fn OpenFile(String) -> i64,
        fn ShowMessageBox(String, String, i64, i64) -> i64,
        fn Write(i64, String),
        fn WriteLine(i64, String),
    }
}

com_shim! {
    class GuiVComponent: GuiComponent {
        AccLabelCollection: GuiComponentCollection,
        AccText: String,
        AccTextOnRequest: String,
        AccTooltip: String,
        Changeable: bool,
        DefaultTooltip: String,
        Height: i64,
        IconName: String,
        IsSymbolFont: bool,
        Left: i64,
        Modified: bool,
        ParentFrame: SAPComponent,
        ScreenLeft: i64,
        ScreenTop: i64,
        mut Text: String,
        Tooltip: String,
        Top: i64,
        Width: i64,

        fn DumpState(String) -> GuiCollection,
        fn SetFocus(),
        fn Visualize(bool) -> bool,
    }
}

com_shim! {
    class GuiVContainer: GuiVComponent + GuiComponent + GuiContainer {
        fn FindAllByName(String, String) -> GuiComponentCollection,
        fn FindAllByNameEx(String, i64) -> GuiComponentCollection,
        fn FindByName(String, String) -> SAPComponent,
        fn FindByNameEx(String, String) -> SAPComponent,
    }
}

com_shim! {
    class GuiVHViewSwitch: GuiVComponent + GuiComponent {}
}
