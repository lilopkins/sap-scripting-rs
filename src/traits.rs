use windows::Win32::System::Com::{IDispatch, VARIANT};

use crate::idispatch_ext::IDispatchExt;
use crate::types::*;
use crate::variant_ext::VariantExt;

/// Forward a function with no args no return type to the IDispatch
macro_rules! forward_func {
    ($(#[$attr:meta])* => $snake_name: ident, $name: expr) => {
        $(#[$attr])*
        fn $snake_name(&self) -> crate::Result<()> {
            let _ = self.get_idispatch().call($name, vec![])?;
            Ok(())
        }
    };
}

/// Forward a function with 1 arg and no return type to the IDispatch
macro_rules! forward_func_1_arg {
    ($(#[$attr:meta])* => $snake_name: ident, $name: expr, $arg_name: ident, $arg_ty: ty, $arg_transformer: ident) => {
        $(#[$attr])*
        fn $snake_name(&self, $arg_name: $arg_ty) -> crate::Result<()> {
            let _ = self.get_idispatch().call($name, vec![VARIANT::$arg_transformer($arg_name)])?;
            Ok(())
        }
    };
}

/// Get a property from the IDispatch
macro_rules! get_property {
    ($(#[$attr:meta])* => $snake_name: ident, $name: expr, $kind: ty, $transformer: ident) => {
        $(#[$attr])*
        fn $snake_name(&self) -> crate::Result<$kind> {
            Ok(self.get_idispatch().get($name)?.$transformer()?)
        }
    };
}

/// A component that has an IDispatch value. Every component needs this, and this trait guarantees that.
pub trait HasDispatch<T = Self> {
    /// Get the IDispatch object for low-level access to this component.
    fn get_idispatch(&self) -> &IDispatch;
}

/// A GuiBox is a simple frame with a name (also called a "Group Box"). The items inside the frame are
/// not children of the box. The type prefix is "box".
pub trait GuiBoxMethods<T: GuiVComponentMethods = Self>: GuiVComponentMethods<T> {
    /// Height of the GuiBox in character metric.
    fn char_height(&self) -> crate::Result<i64> {
        Ok(self.get_idispatch().get("CharHeight")?.to_i64()?)
    }

    /// Left coordinate of the GuiBox in character metric.
    fn char_left(&self) -> crate::Result<i64> {
        Ok(self.get_idispatch().get("CharLeft")?.to_i64()?)
    }
    /// Top coordinate of the GuiBox in character metric.
    fn char_top(&self) -> crate::Result<i64> {
        Ok(self.get_idispatch().get("CharTop")?.to_i64()?)
    }
    /// Width of the GuiBox in character metric.
    fn char_width(&self) -> crate::Result<i64> {
        Ok(self.get_idispatch().get("CharWidth")?.to_i64()?)
    }
}

/// GuiComponent is the base class for most classes in the Scripting API. It was designed to allow generic
/// programming, meaning you can work with objects without knowing their exact type.
///
/// Note: Type is named `kind` due to Rust restrictions.
///
/// Note: Parent is not currently implemented.
pub trait GuiComponentMethods<T: HasDispatch = Self>: HasDispatch<T> {
    /// This property is TRUE, if the object is a container and therefore has the Children property.
    fn container_type(&self) -> crate::Result<bool> {
        Ok(self.get_idispatch().get("ContainerType")?.to_bool()?)
    }

    /// An object id is a unique textual identifier for the object. It is built in a URLlike formatting,
    /// starting at the GuiApplication object and drilling down to the respective object.
    fn id(&self) -> crate::Result<String> {
        Ok(self.get_idispatch().get("Id")?.to_string()?)
    }

    /// The name property is especially useful when working with simple scripts that only access dynpro
    /// fields. In that case a field can be found using its name and type information, which is easier to
    /// read than a possibly very long id. However, there is no guarantee that there are no two objects
    /// with the same name and type in a given dynpro.
    fn name(&self) -> crate::Result<String> {
        Ok(self.get_idispatch().get("Name")?.to_string()?)
    }

    /// The type information of GuiComponent can be used to determine which properties and methods an object
    /// supports. The value of the type string is the name of the type taken from this documentation.
    fn kind(&self) -> crate::Result<String> {
        Ok(self.get_idispatch().get("Type")?.to_string()?)
    }

    /// While the Type property is a string value, the TypeAsNumber property is a long value that can
    /// alternatively be used to identify an object's type . It was added for better performance in methods
    /// such as FindByIdEx. Possible values for this property are taken from the GuiComponentTypeenumeration.
    fn kind_as_number(&self) -> crate::Result<i64> {
        Ok(self.get_idispatch().get("TypeAsNumber")?.to_i64()?)
    }
}

/// This interface resembles GuiVContainer. The only difference is that it is not intended for visual objects
/// but rather administrative objects such as connections or sessions. Objects exposing this interface will
/// therefore support GuiComponent but not GuiVComponent. GuiContainer extends the GuiComponent Object.
pub trait GuiContainerMethods<T: GuiComponentMethods = Self>: GuiComponentMethods<T> {
    /// Search through the object's descendants for a given id. If the parameter is a fully qualified id, the
    /// function will first check if the container object's id is a prefix of the id parameter. If that is the
    /// case, this prefix is truncated. If no descendant with the given id can be found the function raises an
    /// exception.
    fn find_by_id<S>(&self, id: S) -> crate::Result<SAPComponent>
    where
        S: AsRef<str>,
    {
        let comp = GuiComponent {
            inner: self
                .get_idispatch()
                .call("FindById", vec![VARIANT::from_str(id.as_ref())])?
                .to_idispatch()?
                .clone(),
        };
        Ok(SAPComponent::from(comp))
    }

    /// This collection contains all direct children of the object.
    fn children(&self, n: u32) -> crate::Result<GuiSession> {
        Ok(GuiSession {
            inner: self
                .get_idispatch()
                .get_named("Children", VARIANT::from_i32(n as i32))?
                .to_idispatch()?
                .clone(),
        })
    }
}

/// A GuiFrameWindow is a high level visual object in the runtime hierarchy. It can be either the main
/// window or a modal popup window. See the GuiMainWindow and GuiModalWindow sections for examples.
/// GuiFrameWindow itself is an abstract interface. GuiFrameWindow extends the GuiVContainer Object.
/// The type prefix is wnd, the name is wnd plus the window number in square brackets.
pub trait GuiFrameWindowMethods<T: GuiVContainerMethods + GuiContainerMethods = Self>: GuiVContainerMethods<T> {
    forward_func! {
        /// The function attempts to close the window. Trying to close the last main window of a session
        /// will not succeed immediately; the dialog ‘Do you really want to log off?’ will be displayed
        /// first.
        => close, "Close"
    }
    forward_func! {
        /// This will set a window to the iconified state. It is not possible to iconify a specific window
        /// of a session; both the main window and all existing modals will be iconfied.
        => iconify, "Iconify"
    }
    /// This function returns True if the virtual key VKey is currently available. The VKeys are defined
    /// in the menu painter.
    fn is_vkey_allowed(&self, vkey: i32) -> crate::Result<bool> {
        Ok(self.get_idispatch().call("IsVKeyAllowed", vec![VARIANT::from_i32(vkey)])?.to_bool()?)
    }
    forward_func! {
        /// Execute the Ctrl+Shift+Tab key on the window to jump backward one block.
        => jump_backward, "JumpBackward"
    }
    forward_func! {
        /// Execute the Ctrl+Tab key on the window to jump forward one block.
        => jump_forward, "JumpForward"
    }
    forward_func! {
        /// This will maximize a window. It is not possible to maximize a modal window; it is always the
        /// main window which will be maximized.
        => maximize, "Maximize"
    }
    forward_func! {
        /// This will restore a window from its iconified state. It is not possible to restore a specific
        /// window of a session; both the main window and all existing modals will be restored.
        => restore, "Restore"
    }
    forward_func_1_arg! {
        /// The virtual key VKey is executed on the window. The VKeys are defined in the menu painter.
        => send_vkey, "SendVKey", vkey, i32, from_i32
    }
    // TODO ShowMessageBox
    forward_func! {
        /// Execute the Shift+Tab key on the window to jump backward one element.
        => tab_backward, "TabBackward"
    }
    forward_func! {
        /// Execute the Tab key on the window to jump forward one element.
        => tab_forward, "TabForward"
    }
    get_property! {
        /// This is the height of the working pane in character metric.
        => working_pane_height, "WorkingPaneHeight", i64, to_i64
    }
    get_property! {
        /// This is the width of the working pane in character metric. The working pane is the area
        /// between the toolbars in the upper area of the window and the status bar at the bottom of
        /// the window.
        => working_pane_width, "WorkingPaneWidth", i64, to_i64
    }
}

pub trait GuiTextFieldMethods<T: GuiVComponentMethods = Self>: GuiVComponentMethods<T> {
    // TODO
}

/// The GuiVComponent interface is exposed by all visual objects, such as windows, buttons or text fields.
/// Like GuiComponent, it is an abstract interface. Any object supporting the GuiVComponent interface also
/// exposes the GuiComponent interface. GuiVComponent extends the GuiComponent Object.
pub trait GuiVComponentMethods<T: GuiComponentMethods = Self>: GuiComponentMethods<T> {
    // /// This function dumps the state of the object. The parameter innerObject may be used to specify
    // /// for which internal object the data should be dumped. Only the most complex components, such as
    // /// the GuiCtrlGridView, support this parameter. All other components always dump their full state.
    // /// All components that support this parameter have in common that they return general information
    // /// about the control’s state if the parameter “innerObject” contains an empty string. The
    // /// available values for the innerObject parameter are specified as part of the class description
    // /// for those components that support it.
    // fn dump_state<S>(&self, inner_object: S) -> crate::Result<GuiCollection>
    // where
    //     S: AsRef<str>,
    // {
    //     Ok(GuiCollection {
    //         inner: self
    //             .get_idispatch()
    //             .call("DumpState", vec![VARIANT::from_str(inner_object.as_ref())])?
    //             .to_idispatch()?,
    //     })
    // }

    /// This function can be used to set the focus onto an object. If a user interacts with SAP GUI,
    /// it moves the focus whenever the interaction is with a new object. Interacting with an object
    /// through the scripting component does not change the focus. There are some cases in which the
    /// SAP application explicitly checks for the focus and behaves differently depending on the
    /// focused object.
    fn set_focus(&self) -> crate::Result<()> {
        let _ = self.get_idispatch().call("SetFocus", vec![])?;
        Ok(())
    }

    /// Calling this method of a component will display a red frame around the specified component
    /// if the parameter on is true. The frame will be removed if on is false. Some components such
    /// as GuiCtrlGridView support displaying the frame around inner objects, such as cells. The
    /// format of the innerObject string is the same as for the dumpState method.
    fn visualize(&self, on: bool) -> crate::Result<()> {
        let _ = self
            .get_idispatch()
            .call("Visualize", vec![VARIANT::from_bool(on)])?;
        Ok(())
    }

    // TODO
    // fn acc_label_collection(&self) -> crate::Result<GuiComponentCollection> {
    //     Ok(self.get_idispatch().get("AccLabelCollection")?)
    // }

    /// An additional text for accessibility support.
    fn acc_text(&self) -> crate::Result<String> {
        Ok(self.get_idispatch().get("AccText")?.to_string()?)
    }

    /// An additional text for accessibility support.
    fn acc_text_on_request(&self) -> crate::Result<String> {
        Ok(self.get_idispatch().get("AccTextOnRequest")?.to_string()?)
    }

    /// An additional tooltip text for accessibility support.
    fn acc_tooltip(&self) -> crate::Result<String> {
        Ok(self.get_idispatch().get("AccTooltip")?.to_string()?)
    }

    /// An object is changeable if it is neither disabled nor read-only.
    fn changeable(&self) -> crate::Result<bool> {
        Ok(self.get_idispatch().get("Changeable")?.to_bool()?)
    }

    /// Tooltip text generated from the short text defined in the data dictionary for the given screen
    /// element type.
    fn default_tooltip(&self) -> crate::Result<String> {
        Ok(self.get_idispatch().get("DefaultTooltip")?.to_string()?)
    }

    /// Height of the component in pixels.
    fn height(&self) -> crate::Result<i64> {
        Ok(self.get_idispatch().get("Height")?.to_i64()?)
    }

    /// If the object has been assigned an icon, then this property is the name of the icon, otherwise
    /// it is an empty string.
    fn icon_name(&self) -> crate::Result<String> {
        Ok(self.get_idispatch().get("IconName")?.to_string()?)
    }

    /// The property is TRUE if the component's text is visualized in the SAP symbol font.
    fn is_symbol_font(&self) -> crate::Result<bool> {
        Ok(self.get_idispatch().get("IsSymbolFont")?.to_bool()?)
    }

    /// Left position of the element in screen coordinates
    fn left(&self) -> crate::Result<i64> {
        Ok(self.get_idispatch().get("Left")?.to_i64()?)
    }

    /// An object is modified if its state has been changed by the user and this change has not yet been
    /// sent to the SAP system.
    fn modified(&self) -> crate::Result<bool> {
        Ok(self.get_idispatch().get("Modified")?.to_bool()?)
    }

    /// If the control is hosted by the Frame object, the value of the property is this frame. Overwise
    /// NULL.
    fn parent_frame(&self) -> crate::Result<GuiComponent> {
        Ok(GuiComponent {
            inner: self
                .get_idispatch()
                .get("ParentFrame")?
                .to_idispatch()?
                .clone(),
        })
    }

    /// The y position of the component in screen coordinates.
    fn screen_left(&self) -> crate::Result<i64> {
        Ok(self.get_idispatch().get("ScreenLeft")?.to_i64()?)
    }

    /// The x position of the component in screen coordinates.
    fn screen_top(&self) -> crate::Result<i64> {
        Ok(self.get_idispatch().get("ScreenTop")?.to_i64()?)
    }

    /// The value of this property very much depends on the type of the object on which it is called.
    /// This is obvious for text fields or menu items. On the other hand this property is empty for toolbar
    /// buttons and is the class id for shells. You can read the text property of a label, but you can’t
    /// change it, whereas you can only set the text property of a password field, but not read it.
    fn text(&self) -> crate::Result<String> {
        Ok(self.get_idispatch().get("Text")?.to_string()?)
    }

    /// The value of this property very much depends on the type of the object on which it is called.
    /// This is obvious for text fields or menu items. On the other hand this property is empty for toolbar
    /// buttons and is the class id for shells. You can read the text property of a label, but you can’t
    /// change it, whereas you can only set the text property of a password field, but not read it.
    fn set_text<S>(&self, value: S) -> crate::Result<()>
    where
        S: AsRef<str>,
    {
        self.get_idispatch()
            .set("Text", VARIANT::from_str(value.as_ref()))?;
        Ok(())
    }

    /// The tooltip contains a text which is designed to help a user understand the meaning of a given text
    /// field or button.
    fn tooltip(&self) -> crate::Result<String> {
        Ok(self.get_idispatch().get("Tooltip")?.to_string()?)
    }

    /// Top coordinate of the element in screen coordinates.
    fn top(&self) -> crate::Result<i64> {
        Ok(self.get_idispatch().get("Top")?.to_i64()?)
    }

    /// Width of the component in pixels.
    fn width(&self) -> crate::Result<i64> {
        Ok(self.get_idispatch().get("Width")?.to_i64()?)
    }
}

/// An object exposes the GuiVContainer interface if it is both visible and can have children. It will then
/// also expose GuiComponent and GuiVComponent. Examples of this interface are windows and subscreens,
/// toolbars or controls having children, such as the splitter control. GuiVContainer extends the GuiContainer
/// Object and the GuiVComponent Object.
pub trait GuiVContainerMethods<T: GuiContainerMethods + GuiVComponentMethods = Self>:
    GuiVComponentMethods<T>
{
    // TODO fn find_all_by_name(&self);
    // TODO fn find_all_by_name_ex(&self);

    /// Unlike FindById, this function does not guarantee a unique result. It will simply return the first
    /// descendant matching both the name and type parameters. This is a more natural description of the
    /// object than the complex id, but it only makes sense on dynpro objects as most other objects do not
    /// have a meaningful name. If no descendant with matching name and type can be found, the function
    /// raises an exception.
    fn find_by_name<S>(&self, name: S, kind: S) -> crate::Result<SAPComponent>
    where
        S: AsRef<str>,
    {
        Ok(SAPComponent::from(GuiComponent {
            inner: self
                .get_idispatch()
                .call(
                    "FindByName",
                    vec![
                        VARIANT::from_str(name.as_ref()),
                        VARIANT::from_str(kind.as_ref()),
                    ],
                )?
                .to_idispatch()?
                .clone(),
        }))
    }

    /// This method works exactly like FindByName, but takes the type parameter with data type long coming
    /// from the GuiComponentType enumeration. See also GuiComponentType.
    fn find_by_name_ex<S>(&self, name: S, kind: i64) -> crate::Result<SAPComponent>
    where
        S: AsRef<str>,
    {
        Ok(SAPComponent::from(GuiComponent {
            inner: self
                .get_idispatch()
                .call(
                    "FindByNameEx",
                    vec![VARIANT::from_str(name.as_ref()), VARIANT::from_i64(kind)],
                )?
                .to_idispatch()?
                .clone(),
        }))
    }
}
