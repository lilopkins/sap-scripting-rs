use com_shim::{com_shim, IDispatchExt, VariantTypeExt};
use windows::{Win32::System::Com::*, Win32::System::Variant::*, core::*};

/// A wrapper over the SAP scripting engine, equivalent to CSapROTWrapper.
pub struct SAPWrapper {
    inner: IDispatch,
}

impl SAPWrapper {
    pub(crate) fn new() -> crate::Result<Self> {
        unsafe {
            let clsid: GUID = CLSIDFromProgID(w!("SapROTWr.SapROTWrapper"))?;
            let p_clsid: *const GUID = &clsid;
            tracing::debug!("CSapROTWrapper CLSID: {:?}", clsid);

            let dispatch: IDispatch =
                CoCreateInstance(p_clsid, None, CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER)?;
            Ok(SAPWrapper { inner: dispatch })
        }
    }

    /// Get the Scripting Engine object from this wrapper.
    pub fn scripting_engine(&self) -> crate::Result<GuiApplication> {
        tracing::debug!("Getting UI ROT entry...");
        let result = self.inner.call(
            "GetROTEntry",
            vec![VARIANT::variant_from("SAPGUI".to_string())],
        )?;

        let sap_gui: &IDispatch = result.variant_into()?;

        tracing::debug!("Getting scripting engine.");
        let scripting_engine = sap_gui.call("GetScriptingEngine", vec![])?;

        Ok(GuiApplication {
            inner: <com_shim::VARIANT as VariantTypeExt<'_, &IDispatch>>::variant_into(
                &scripting_engine,
            )?
            .clone(),
        })
    }
}

pub trait HasSAPType {
    fn sap_type() -> &'static str;
    fn sap_subtype() -> Option<&'static str>;
}

macro_rules! sap_type {
    ($tgt: ty, $type: expr) => {
        impl HasSAPType for $tgt {
            fn sap_type() -> &'static str {
                $type
            }
            fn sap_subtype() -> Option<&'static str> {
                None
            }
        }
    };
    ($tgt: ty, $type: expr, $subtype: expr) => {
        impl HasSAPType for $tgt {
            fn sap_type() -> &'static str {
                $type
            }
            fn sap_subtype() -> Option<&'static str> {
                Some($subtype)
            }
        }
    };
}

impl GuiComponent {
    pub fn downcast<Tgt>(&self) -> Option<Tgt>
    where
        Tgt: HasSAPType + From<IDispatch>,
    {
        if let Ok(mut kind) = self.r_type() {
            tracing::debug!("GuiComponent is {kind}.");
            if kind.as_str() == "GuiShell" {
                if let Ok(sub_kind) = (GuiShell {
                    inner: self.inner.clone(),
                })
                .sub_type()
                {
                    // use subkind if a GuiShell
                    tracing::debug!("Subkind is {sub_kind}");
                    kind = sub_kind;
                }
            }
            let target_kind = Tgt::sap_subtype().unwrap_or_else(|| Tgt::sap_type());
            if kind == target_kind {
                Some(Tgt::from(self.inner.clone()))
            } else {
                None
            }
        } else {
            None
        }
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
sap_type!(GuiApplication, "GuiApplication");

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
sap_type!(GuiBarChart, "GuiShell", "BarChart");

com_shim! {
    struct GuiBox: GuiVComponent + GuiComponent {
        CharHeight: i32,
        CharLeft: i32,
        CharTop: i32,
        CharWidth: i32,
    }
}
sap_type!(GuiBox, "GuiBox");

com_shim! {
    struct GuiButton: GuiVComponent + GuiComponent {
        Emphasized: bool,
        LeftLabel: GuiComponent,
        RightLabel: GuiComponent,

        fn Press(),
    }
}
sap_type!(GuiButton, "GuiButton");

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
sap_type!(GuiCalendar, "GuiShell", "Calendar");

com_shim! {
    struct GuiChart: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        fn ValueChange(i32, i32, String, String, bool, String, String, i32),
    }
}
sap_type!(GuiChart, "GuiShell", "Chart");

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
sap_type!(GuiCheckBox, "GuiCheckBox");

com_shim! {
    struct GuiCollection {
        Count: i32,
        Length: i32,
        r#Type: String,
        TypeAsNumber: i32,

        // TODO fn Add(&IDispatch),
        fn ElementAt(i32) -> GuiComponent,
    }
}

com_shim! {
    struct GuiColorSelector: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        fn ChangeSelection(i16),
    }
}
sap_type!(GuiColorSelector, "GuiShell", "ColorSelector");

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
sap_type!(GuiComboBox, "GuiComboBox");

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
sap_type!(GuiComboBoxControl, "GuiShell", "ComboBoxControl");

com_shim! {
    struct GuiComboBoxEntry {
        Key: String,
        Pos: i32,
        Value: String,
    }
}
sap_type!(GuiComboBoxEntry, "GuiComboBoxEntry");

com_shim! {
    struct GuiComponent {
        ContainerType: bool,
        Id: String,
        Name: String,
        r#Type: String,
        TypeAsNumber: i32,
    }
}
sap_type!(GuiComponent, "GuiComponent");

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
sap_type!(GuiConnection, "GuiConnection");

com_shim! {
    struct GuiContainer: GuiComponent {
        Children: GuiComponentCollection,

        fn FindById(String) -> GuiComponent,
    }
}
sap_type!(GuiContainer, "GuiContainer");

com_shim! {
    struct GuiContainerShell: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        AccDescription: String,
    }
}
sap_type!(GuiContainerShell, "GuiShell", "ContainerShell");

com_shim! {
    struct GuiCTextField: GuiTextField + GuiVComponent + GuiComponent { }
}
sap_type!(GuiCTextField, "GuiCTextField");

com_shim! {
    struct GuiCustomControl: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        CharHeight: i32,
        CharLeft: i32,
        CharTop: i32,
        CharWidth: i32,
    }
}
sap_type!(GuiCustomControl, "GuiCustomControl");

com_shim! {
    struct GuiDialogShell: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        Title: String,

        fn Close(),
    }
}
sap_type!(GuiDialogShell, "GuiDialogShell");

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
sap_type!(GuiEAIViewer2D, "GuiShell", "EAIViewer2D");

com_shim! {
    struct GuiEAIViewer3D: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell { }
}
sap_type!(GuiEAIViewer3D, "GuiShell", "EAIViewer3D");

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
sap_type!(GuiFrameWindow, "GuiFrameWindow");

com_shim! {
    struct GuiGOSShell: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent { }
}
sap_type!(GuiGOSShell, "GuiGOSShell");

com_shim! {
    struct GuiGraphAdapt: GuiVComponent + GuiVContainer + GuiContainer + GuiComponent + GuiShell { }
}
sap_type!(GuiGraphAdapt, "GuiShell", "GraphAdapt");

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
sap_type!(GuiGridView, "GuiShell", "GridView");

com_shim! {
    struct GuiHTMLViewer: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell {
        // TODO BrowserHandle: Object,
        DocumentComplete: i32,

        fn ContextMenu(),
        fn GetBrowerControlType() -> i32,
        fn SapEvent(String, String, String),
    }
}
sap_type!(GuiHTMLViewer, "GuiShell", "HTMLViewer");

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
sap_type!(GuiInputFieldControl, "GuiShell", "InputFieldControl");

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
sap_type!(GuiLabel, "GuiLabel");

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
sap_type!(GuiMainWindow, "GuiMainWindow");

com_shim! {
    struct GuiMap: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell { }
}
sap_type!(GuiMap, "GuiShell", "Map");

com_shim! {
    struct GuiMenu: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        fn Select(),
    }
}
sap_type!(GuiMenu, "GuiMenu");

com_shim! {
    struct GuiMenubar: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent { }
}
sap_type!(GuiMenubar, "GuiMenubar");

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
sap_type!(GuiModalWindow, "GuiModalWindow");

com_shim! {
    struct GuiNetChart: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell {
        LinkCount: i32,
        NodeCount: i32,

        fn GetLinkContent(i32, i32) -> String,
        fn GetNodeContent(i32, i32) -> String,
        fn SendData(String),
    }
}
sap_type!(GuiNetChart, "GuiShell", "NetChart");

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
sap_type!(GuiOfficeIntegration, "GuiShell", "OfficeIntegration");

com_shim! {
    struct GuiOkCodeField: GuiVComponent + GuiComponent {
        Opened: bool,

        fn PressF1(),
    }
}
sap_type!(GuiOkCodeField, "GuiOkCodeField");

com_shim! {
    struct GuiPasswordField: GuiTextField + GuiVComponent + GuiComponent { }
}
sap_type!(GuiPasswordField, "GuiPasswordField");

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
sap_type!(GuiPicture, "GuiShell", "Picture");

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
sap_type!(GuiRadioButton, "GuiRadioButton");

com_shim! {
    struct GuiSapChart: GuiVComponent + GuiVContainer + GuiComponent + GuiContainer + GuiShell { }
}
sap_type!(GuiSapChart, "GuiShell", "SapChart");

com_shim! {
    struct GuiScrollbar {
        Maximum: i32,
        Minimum: i32,
        PageSize: i32,
        mut Position: i32,
        Range: i32,
    }
}
sap_type!(GuiScrollbar, "GuiScrollbar");

com_shim! {
    struct GuiScrollContainer: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        HorizontalScrollbar: GuiComponent,
        VerticalScrollbar: GuiComponent,
    }
}
sap_type!(GuiScrollContainer, "GuiScrollContainer");

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
sap_type!(GuiSession, "GuiSession");

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
sap_type!(GuiShell, "GuiShell");

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
sap_type!(GuiSimpleContainer, "GuiSimpleContainer");

com_shim! {
    struct GuiSplit: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent + GuiShell {
        IsVertical: i32,

        fn GetColSize(i32) -> i32,
        fn GetRowSize(i32) -> i32,
        fn SetColSize(i32, i32),
        fn SetRowSize(i32, i32),
    }
}
sap_type!(GuiSplit, "GuiShell", "Split");

com_shim! {
    struct GuiSplitterContainer: GuiVContainer + GuiVComponent + GuiComponent + GuiContainer + GuiShell {
        IsVertical: bool,
        mut SashPosition: i32,
    }
}
sap_type!(GuiSplitterContainer, "GuiShell", "SplitterContainer");

com_shim! {
    struct GuiStage: GuiVComponent + GuiVContainer + GuiContainer + GuiShell + GuiComponent {
        fn ContextMenu(String),
        fn DoubleClick(String),
        fn SelectItems(String),
    }
}
sap_type!(GuiStage, "GuiShell", "Stage");

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
sap_type!(GuiStatusbar, "GuiStatusbar");

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
sap_type!(GuiStatusPane, "GuiStatusPane");

com_shim! {
    struct GuiTab: GuiVContainer + GuiVComponent + GuiContainer + GuiComponent {
        fn ScrollToLeft(),
        fn Select(),
    }
}
sap_type!(GuiTab, "GuiTab");

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
sap_type!(GuiTableControl, "GuiTableControl");

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
sap_type!(GuiTabStrip, "GuiTabStrip");

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
sap_type!(GuiTextedit, "GuiShell", "Textedit");

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
sap_type!(GuiTextField, "GuiTextField");

com_shim! {
    struct GuiTitlebar: GuiVComponent + GuiVContainer + GuiContainer + GuiComponent { }
}
sap_type!(GuiTitlebar, "GuiTitlebar");

com_shim! {
    struct GuiToolbar: GuiVComponent + GuiVContainer + GuiContainer + GuiComponent { }
}
sap_type!(GuiToolbar, "GuiToolbar");

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
sap_type!(GuiTree, "GuiShell", "Tree");

com_shim! {
    struct GuiUserArea: GuiVContainer + GuiVComponent + GuiComponent + GuiContainer {
        HorizontalScrollbar: GuiComponent,
        IsOTFPreview: bool,
        VerticalScrollbar: GuiComponent,

        fn FindByLabel(String, String) -> GuiComponent,
        fn ListNavigate(String),
    }
}
sap_type!(GuiUserArea, "GuiUserArea");

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
sap_type!(GuiVComponent, "GuiVComponent");

com_shim! {
    struct GuiVContainer: GuiVComponent + GuiComponent + GuiContainer {
        fn FindAllByName(String, String) -> GuiComponentCollection,
        fn FindAllByNameEx(String, i32) -> GuiComponentCollection,
        fn FindByName(String, String) -> GuiComponent,
        fn FindByNameEx(String, String) -> GuiComponent,
    }
}
sap_type!(GuiVContainer, "GuiVContainer");

com_shim! {
    struct GuiVHViewSwitch: GuiVComponent + GuiComponent {}
}
sap_type!(GuiVHViewSwitch, "GuiVHViewSwitch");
