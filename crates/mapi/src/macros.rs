//! Public macros exported from the [outlook-mapi](https://crates.io/crates/outlook-mapi) crate.

#[allow(unused_imports)]
use crate::sys;

/// Declare a variable length struct with the same layout as [`sys::ENTRYID`] and implement casting
/// functions:
///
/// - `fn as_ptr(&self) -> *const sys::ENTRYID`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::ENTRYID`.
///
/// ### Sample
/// ```
/// use outlook_mapi::{sys, SizedENTRYID};
///
/// SizedENTRYID! { EntryId[12] }
///
/// let entry_id = EntryId {
///     abFlags: [0x0, 0x1, 0x2, 0x3],
///     ab: [0x4, 0x5, 0x6, 0x7, 0x8, 0x9, 0xa, 0xb, 0xc, 0xd, 0xe, 0xf],
/// };
///
/// let entry_id: *const sys::ENTRYID = entry_id.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedENTRYID {
    ($name:ident [ $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub abFlags: [u8; 4],
            pub ab: [u8; $count],
        }

        outlook_mapi_macros::impl_sized_struct_casts!($name, $crate::sys::ENTRYID);
    };
}

/// Declare a variable length struct with the same layout as [`sys::SPropTagArray`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::SPropTagArray`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::SPropTagArray`.
///
/// ### Sample
/// ```
/// use outlook_mapi::{sys, SizedSPropTagArray};
///
/// SizedSPropTagArray! { PropTagArray[2] }
///
/// let prop_tag_array = PropTagArray {
///     aulPropTag: [
///         sys::PR_ENTRYID,
///         sys::PR_DISPLAY_NAME_W,
///     ],
///     ..Default::default()
/// };
///
/// let prop_tag_array: *const sys::SPropTagArray = prop_tag_array.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSPropTagArray {
    ($name:ident [ $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub cValues: u32,
            pub aulPropTag: [u32; $count],
        }

        outlook_mapi_macros::impl_sized_struct_casts!($name, $crate::sys::SPropTagArray);

        outlook_mapi_macros::impl_sized_struct_default!($name {
            cValues: $count as u32,
            aulPropTag: [$crate::sys::PR_NULL; $count],
        });
    };
}

/// Declare a variable length struct with the same layout as [`sys::SPropProblemArray`] and
/// implement casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::SPropProblemArray`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::SPropProblemArray`.
///
/// ### Sample
/// ```
/// use outlook_mapi::{sys, SizedSPropProblemArray};
///
/// SizedSPropProblemArray! { PropProblemArray[2] }
///
/// let prop_problem_array = PropProblemArray {
///     aProblem: [
///         sys::SPropProblem {
///             ulIndex: 0,
///             ulPropTag: sys::PR_ENTRYID,
///             scode: sys::MAPI_E_NOT_FOUND.0
///         },
///         sys::SPropProblem {
///             ulIndex: 1,
///             ulPropTag: sys::PR_DISPLAY_NAME_W,
///             scode: sys::MAPI_E_NOT_FOUND.0
///         },
///     ],
///     ..Default::default()
/// };
///
/// let prop_problem_array: *const sys::SPropProblemArray = prop_problem_array.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSPropProblemArray {
    ($name:ident [ $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub cValues: u32,
            pub aProblem: [$crate::sys::SPropProblem; $count],
        }

        outlook_mapi_macros::impl_sized_struct_casts!($name, $crate::sys::SPropProblemArray);

        {
            const DEFAULT_VALUE: $crate::sys::SPropProblem = $crate::sys::SPropProblem {
                ulIndex: 0,
                ulPropTag: $crate::sys::PR_NULL,
                scode: 0,
            };

            outlook_mapi_macros::impl_sized_struct_default!($name {
                cValues: $count as u32,
                aProblem: [DEFAULT_VALUE; $count],
            });
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::ADRLIST`] and implement casting
/// functions:
///
/// - `fn as_ptr(&self) -> *const sys::ADRLIST`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::ADRLIST`.
///
/// ### Sample
/// ```
/// use outlook_mapi::{sys, SizedADRLIST};
///
/// SizedADRLIST! { AdrList[2] }
///
/// let adr_list = AdrList {
///     aEntries: [
///         sys::ADRENTRY {
///             ulReserved1: 0,
///             cValues: 0,
///             rgPropVals: std::ptr::null_mut(),
///         },
///         sys::ADRENTRY {
///             ulReserved1: 0,
///             cValues: 0,
///             rgPropVals: std::ptr::null_mut(),
///         },
///     ],
///     ..Default::default()
/// };
///
/// let adr_list: *const sys::ADRLIST = adr_list.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedADRLIST {
    ($name:ident [ $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub cEntries: u32,
            pub aEntries: [$crate::sys::ADRENTRY; $count],
        }

        outlook_mapi_macros::impl_sized_struct_casts!($name, $crate::sys::ADRLIST);

        {
            const DEFAULT_VALUE: $crate::sys::ADRENTRY = $crate::sys::ADRENTRY {
                ulReserved1: 0,
                cValues: 0,
                rgPropVals: std::ptr::null_mut(),
            };

            outlook_mapi_macros::impl_sized_struct_default!($name {
                cEntries: $count as u32,
                aEntries: [DEFAULT_VALUE; $count],
            });
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::SRowSet`] and implement casting
/// functions:
///
/// - `fn as_ptr(&self) -> *const sys::SRowSet`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::SRowSet`.
///
/// ### Sample
/// ```
/// use std::ptr;
/// use outlook_mapi::{sys, SizedSRowSet};
///
/// SizedSRowSet! { RowSet[2] }
///
/// let row_set = RowSet {
///     aRow: [
///         sys::SRow {
///             ulAdrEntryPad: 0,
///             cValues: 0,
///             lpProps: ptr::null_mut(),
///         },
///         sys::SRow {
///             ulAdrEntryPad: 0,
///             cValues: 0,
///             lpProps: ptr::null_mut(),
///         },
///     ],
///     ..Default::default()
/// };
///
/// let row_set: *const sys::SRowSet = row_set.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSRowSet {
    ($name:ident [ $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub cRows: u32,
            pub aRow: [$crate::sys::SRow; $count],
        }

        outlook_mapi_macros::impl_sized_struct_casts!($name, $crate::sys::SRowSet);

        {
            const DEFAULT_VALUE: $crate::sys::SRow = $crate::sys::SRow {
                ulAdrEntryPad: 0,
                cValues: 0,
                lpProps: std::ptr::null_mut(),
            };

            outlook_mapi_macros::impl_sized_struct_default!($name {
                cRows: $count as u32,
                aRow: [DEFAULT_VALUE; $count],
            });
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::SSortOrderSet`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::SSortOrderSet`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::SSortOrderSet`.
///
/// ### Sample
/// ```
/// use std::ptr;
/// use outlook_mapi::{sys, SizedSSortOrderSet};
///
/// SizedSSortOrderSet! { SortOrderSet[3] }
///
/// let sort_order_set = SortOrderSet {
///     cCategories: 1,
///     cExpanded: 1,
///     aSort: [
///         sys::SSortOrder {
///             ulPropTag: sys::PR_CONVERSATION_TOPIC_W,
///             ulOrder: sys::TABLE_SORT_DESCEND,
///         },
///         sys::SSortOrder {
///             ulPropTag: sys::PR_MESSAGE_DELIVERY_TIME,
///             ulOrder: sys::TABLE_SORT_CATEG_MAX,
///         },
///         sys::SSortOrder {
///             ulPropTag: sys::PR_CONVERSATION_INDEX,
///             ulOrder: sys::TABLE_SORT_ASCEND,
///         },
///     ],
///     ..Default::default()
/// };
///
/// let sort_order_set: *const sys::SSortOrderSet = sort_order_set.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSSortOrderSet {
    ($name:ident [ $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub cSorts: u32,
            pub cCategories: u32,
            pub cExpanded: u32,
            pub aSort: [$crate::sys::SSortOrder; $count],
        }

        outlook_mapi_macros::impl_sized_struct_casts!($name, $crate::sys::SSortOrderSet);

        {
            const DEFAULT_VALUE: $crate::sys::SSortOrder = $crate::sys::SSortOrder {
                ulPropTag: $crate::sys::PR_NULL,
                ulOrder: $crate::sys::TABLE_SORT_ASCEND,
            };

            outlook_mapi_macros::impl_sized_struct_default!($name {
                cSorts: $count as u32,
                cCategories: 0,
                cExpanded: 0,
                aSort: [DEFAULT_VALUE; $count],
            });
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLLABEL`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLLABEL`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLLABEL`
///
/// It also initializes the [`sys::DTBLLABEL::ulFlags`] member and implements either of these
/// accessor to fill in the string buffer, depending on whether it is declared with [`u8`] or
/// [`u16`]:
///
/// - [`u8`]: `fn label_name(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn label_name(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// use outlook_mapi::{sys, SizedDtblLabel};
/// use windows_core::{PCSTR, PCWSTR};
///
/// const LABEL: &str = "Display Table Label";
///
/// SizedDtblLabel! { DisplayTableLabelA[u8; LABEL.len()] }
///
/// let mut display_table_label = DisplayTableLabelA::default();
/// assert_eq!(display_table_label.ulFlags, 0);
///
/// let label: Vec<_> = LABEL.bytes().collect();
/// assert_eq!(LABEL.len(), label.len());
/// display_table_label.label_name().copy_from_slice(label.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_label.lpszLabelName.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
/// }
///
/// let display_table_label: *const sys::DTBLLABEL = display_table_label.as_ptr();
///
/// SizedDtblLabel! { DisplayTableLabelW[u16; LABEL.len()] }
///
/// let mut display_table_label = DisplayTableLabelW::default();
/// assert_eq!(display_table_label.ulFlags, sys::MAPI_UNICODE);
///
/// let label: Vec<_> = LABEL.encode_utf16().collect();
/// assert_eq!(LABEL.len(), label.len());
/// display_table_label.label_name().copy_from_slice(label.as_slice());
/// unsafe {
///     assert_eq!(
///         PCWSTR::from_raw(display_table_label.lpszLabelName.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
/// }
///
/// let display_table_label: *const sys::DTBLLABEL = display_table_label.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblLabel {
    ($name:ident [ $char:ident; $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub ulbLpszLabelName: u32,
            pub ulFlags: u32,
            pub lpszLabelName: [$char; $count + 1],
        }

        outlook_mapi_macros::impl_sized_struct_casts!($name, $crate::sys::DTBLLABEL);

        outlook_mapi_macros::impl_sized_struct_default!($name {
            ulbLpszLabelName: std::mem::size_of::<$crate::sys::DTBLLABEL>() as u32,
            ulFlags: outlook_mapi_macros::display_table_default_flags!(
                $char,
                $crate::sys::MAPI_UNICODE
            ),
            lpszLabelName: [0; $count + 1],
        });

        impl $name {
            pub fn label_name(&mut self) -> &mut [$char] {
                &mut self.lpszLabelName[..$count]
            }
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLEDIT`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLEDIT`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLEDIT`
///
/// It also initializes the [`sys::DTBLEDIT::ulFlags`] member and implements either of these
/// accessor to fill in the string buffer, depending on whether it is declared with [`u8`] or
/// [`u16`]:
///
/// - [`u8`]: `fn chars_allowed(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn chars_allowed(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// use outlook_mapi::{sys, SizedDtblEdit};
/// use windows_core::{PCSTR, PCWSTR};
///
/// const ALLOWED: &str = "Allowed Characters";
///
/// SizedDtblEdit! { DisplayTableEditA[u8; ALLOWED.len()] }
///
/// let mut display_table_edit = DisplayTableEditA::default();
/// assert_eq!(display_table_edit.ulFlags, 0);
///
/// let allowed: Vec<_> = ALLOWED.bytes().collect();
/// assert_eq!(ALLOWED.len(), allowed.len());
/// display_table_edit.chars_allowed().copy_from_slice(allowed.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_edit.lpszCharsAllowed.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         ALLOWED);
/// }
///
/// let display_table_edit: *const sys::DTBLEDIT = display_table_edit.as_ptr();
///
/// SizedDtblEdit! { DisplayTableEditW[u16; ALLOWED.len()] }
///
/// let mut display_table_edit = DisplayTableEditW::default();
/// assert_eq!(display_table_edit.ulFlags, sys::MAPI_UNICODE);
///
/// let allowed: Vec<_> = ALLOWED.encode_utf16().collect();
/// assert_eq!(ALLOWED.len(), allowed.len());
/// display_table_edit.chars_allowed().copy_from_slice(allowed.as_slice());
/// unsafe {
///     assert_eq!(
///         PCWSTR::from_raw(display_table_edit.lpszCharsAllowed.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         ALLOWED);
/// }
///
/// let display_table_edit: *const sys::DTBLEDIT = display_table_edit.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblEdit {
    ($name:ident [ $char:ident; $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub ulbLpszCharsAllowed: u32,
            pub ulFlags: u32,
            pub ulNumCharsAllowed: u32,
            pub ulPropTag: u32,
            pub lpszCharsAllowed: [$char; $count + 1],
        }

        outlook_mapi_macros::impl_sized_struct_casts!($name, $crate::sys::DTBLEDIT);

        outlook_mapi_macros::impl_sized_struct_default!($name {
            ulbLpszCharsAllowed: std::mem::size_of::<$crate::sys::DTBLEDIT>() as u32,
            ulFlags: outlook_mapi_macros::display_table_default_flags!(
                $char,
                $crate::sys::MAPI_UNICODE
            ),
            ulNumCharsAllowed: 0,
            ulPropTag: $crate::sys::PR_NULL,
            lpszCharsAllowed: [0; $count + 1],
        });

        impl $name {
            pub fn chars_allowed(&mut self) -> &mut [$char] {
                &mut self.lpszCharsAllowed[..$count]
            }
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLCOMBOBOX`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLCOMBOBOX`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLCOMBOBOX`
///
/// It also initializes the [`sys::DTBLCOMBOBOX::ulFlags`] member and implements either of these
/// accessor to fill in the string buffer, depending on whether it is declared with [`u8`] or
/// [`u16`]:
///
/// - [`u8`]: `fn chars_allowed(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn chars_allowed(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// use outlook_mapi::{sys, SizedDtblComboBox};
/// use windows_core::{PCSTR, PCWSTR};
///
/// const ALLOWED: &str = "Allowed Characters";
///
/// SizedDtblComboBox! { DisplayTableComboBoxA[u8; ALLOWED.len()] }
///
/// let mut display_table_combo_box = DisplayTableComboBoxA::default();
/// assert_eq!(display_table_combo_box.ulFlags, 0);
///
/// let allowed: Vec<_> = ALLOWED.bytes().collect();
/// assert_eq!(ALLOWED.len(), allowed.len());
/// display_table_combo_box.chars_allowed().copy_from_slice(allowed.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_combo_box.lpszCharsAllowed.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         ALLOWED);
/// }
///
/// let display_table_combo_box: *const sys::DTBLCOMBOBOX = display_table_combo_box.as_ptr();
///
/// SizedDtblComboBox! { DisplayTableComboBoxW[u16; ALLOWED.len()] }
///
/// let mut display_table_combo_box = DisplayTableComboBoxW::default();
/// assert_eq!(display_table_combo_box.ulFlags, sys::MAPI_UNICODE);
///
/// let allowed: Vec<_> = ALLOWED.encode_utf16().collect();
/// assert_eq!(ALLOWED.len(), allowed.len());
/// display_table_combo_box.chars_allowed().copy_from_slice(allowed.as_slice());
/// unsafe {
///     assert_eq!(
///         PCWSTR::from_raw(display_table_combo_box.lpszCharsAllowed.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         ALLOWED);
/// }
///
/// let display_table_combo_box: *const sys::DTBLCOMBOBOX = display_table_combo_box.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblComboBox {
    ($name:ident [ $char:ident; $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub ulbLpszCharsAllowed: u32,
            pub ulFlags: u32,
            pub ulNumCharsAllowed: u32,
            pub ulPRPropertyName: u32,
            pub ulPRTableName: u32,
            pub lpszCharsAllowed: [$char; $count + 1],
        }

        outlook_mapi_macros::impl_sized_struct_casts!($name, $crate::sys::DTBLCOMBOBOX);

        outlook_mapi_macros::impl_sized_struct_default!($name {
            ulbLpszCharsAllowed: std::mem::size_of::<$crate::sys::DTBLCOMBOBOX>() as u32,
            ulFlags: outlook_mapi_macros::display_table_default_flags!(
                $char,
                $crate::sys::MAPI_UNICODE
            ),
            ulNumCharsAllowed: 0,
            ulPRPropertyName: $crate::sys::PR_NULL,
            ulPRTableName: $crate::sys::PR_NULL,
            lpszCharsAllowed: [0; $count + 1],
        });

        impl $name {
            pub fn chars_allowed(&mut self) -> &mut [$char] {
                &mut self.lpszCharsAllowed[..$count]
            }
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLCHECKBOX`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLCHECKBOX`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLCHECKBOX`
///
/// It also initializes the [`sys::DTBLCHECKBOX::ulFlags`] member and implements either of these
/// accessor to fill in the string buffer, depending on whether it is declared with [`u8`] or
/// [`u16`]:
///
/// - [`u8`]: `fn label(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn label(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// use outlook_mapi::{sys, SizedDtblCheckBox};
/// use windows_core::{PCSTR, PCWSTR};
///
/// const LABEL: &str = "Checkbox Label";
///
/// SizedDtblCheckBox! { DisplayTableCheckBoxA[u8; LABEL.len()] }
///
/// let mut display_table_check_box = DisplayTableCheckBoxA::default();
/// assert_eq!(display_table_check_box.ulFlags, 0);
///
/// let label: Vec<_> = LABEL.bytes().collect();
/// assert_eq!(LABEL.len(), label.len());
/// display_table_check_box.label().copy_from_slice(label.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_check_box.lpszLabel.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
/// }
///
/// let display_table_check_box: *const sys::DTBLCHECKBOX = display_table_check_box.as_ptr();
///
/// SizedDtblCheckBox! { DisplayTableCheckBoxW[u16; LABEL.len()] }
///
/// let mut display_table_check_box = DisplayTableCheckBoxW::default();
/// assert_eq!(display_table_check_box.ulFlags, sys::MAPI_UNICODE);
///
/// let label: Vec<_> = LABEL.encode_utf16().collect();
/// assert_eq!(LABEL.len(), label.len());
/// display_table_check_box.label().copy_from_slice(label.as_slice());
/// unsafe {
///     assert_eq!(
///         PCWSTR::from_raw(display_table_check_box.lpszLabel.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
/// }
///
/// let display_table_check_box: *const sys::DTBLCHECKBOX = display_table_check_box.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblCheckBox {
    ($name:ident [ $char:ident; $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub ulbLpszLabel: u32,
            pub ulFlags: u32,
            pub ulPRPropertyName: u32,
            pub lpszLabel: [$char; $count + 1],
        }

        outlook_mapi_macros::impl_sized_struct_casts!($name, $crate::sys::DTBLCHECKBOX);

        outlook_mapi_macros::impl_sized_struct_default!($name {
            ulbLpszLabel: std::mem::size_of::<$crate::sys::DTBLCHECKBOX>() as u32,
            ulFlags: outlook_mapi_macros::display_table_default_flags!(
                $char,
                $crate::sys::MAPI_UNICODE
            ),
            ulPRPropertyName: $crate::sys::PR_NULL,
            lpszLabel: [0; $count + 1],
        });

        impl $name {
            pub fn label(&mut self) -> &mut [$char] {
                &mut self.lpszLabel[..$count]
            }
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLGROUPBOX`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLGROUPBOX`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLGROUPBOX`
///
/// It also initializes the [`sys::DTBLGROUPBOX::ulFlags`] member and implements either of these
/// accessor to fill in the string buffer, depending on whether it is declared with [`u8`] or
/// [`u16`]:
///
/// - [`u8`]: `fn label(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn label(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// use outlook_mapi::{sys, SizedDtblGroupBox};
/// use windows_core::{PCSTR, PCWSTR};
///
/// const LABEL: &str = "Groupbox Label";
///
/// SizedDtblGroupBox! { DisplayTableGroupBoxA[u8; LABEL.len()] }
///
/// let mut display_table_group_box = DisplayTableGroupBoxA::default();
/// assert_eq!(display_table_group_box.ulFlags, 0);
///
/// let label: Vec<_> = LABEL.bytes().collect();
/// assert_eq!(LABEL.len(), label.len());
/// display_table_group_box.label().copy_from_slice(label.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_group_box.lpszLabel.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
/// }
///
/// let display_table_group_box: *const sys::DTBLGROUPBOX = display_table_group_box.as_ptr();
///
/// SizedDtblGroupBox! { DisplayTableGroupBoxW[u16; LABEL.len()] }
///
/// let mut display_table_group_box = DisplayTableGroupBoxW::default();
/// assert_eq!(display_table_group_box.ulFlags, sys::MAPI_UNICODE);
///
/// let label: Vec<_> = LABEL.encode_utf16().collect();
/// assert_eq!(LABEL.len(), label.len());
/// display_table_group_box.label().copy_from_slice(label.as_slice());
/// unsafe {
///     assert_eq!(
///         PCWSTR::from_raw(display_table_group_box.lpszLabel.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
/// }
///
/// let display_table_group_box: *const sys::DTBLGROUPBOX = display_table_group_box.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblGroupBox {
    ($name:ident [ $char:ident; $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub ulbLpszLabel: u32,
            pub ulFlags: u32,
            pub lpszLabel: [$char; $count + 1],
        }

        outlook_mapi_macros::impl_sized_struct_casts!($name, $crate::sys::DTBLGROUPBOX);

        outlook_mapi_macros::impl_sized_struct_default!($name {
            ulbLpszLabel: std::mem::size_of::<$crate::sys::DTBLGROUPBOX>() as u32,
            ulFlags: outlook_mapi_macros::display_table_default_flags!(
                $char,
                $crate::sys::MAPI_UNICODE
            ),
            lpszLabel: [0; $count + 1],
        });

        impl $name {
            pub fn label(&mut self) -> &mut [$char] {
                &mut self.lpszLabel[..$count]
            }
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLBUTTON`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLBUTTON`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLBUTTON`
///
/// It also initializes the [`sys::DTBLBUTTON::ulFlags`] member and implements either of these
/// accessor to fill in the string buffer, depending on whether it is declared with [`u8`] or
/// [`u16`]:
///
/// - [`u8`]: `fn label(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn label(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// use outlook_mapi::{sys, SizedDtblButton};
/// use windows_core::{PCSTR, PCWSTR};
///
/// const LABEL: &str = "Button Label";
///
/// SizedDtblButton! { DisplayTableButtonA[u8; LABEL.len()] }
///
/// let mut display_table_button = DisplayTableButtonA::default();
/// assert_eq!(display_table_button.ulFlags, 0);
///
/// let label: Vec<_> = LABEL.bytes().collect();
/// assert_eq!(LABEL.len(), label.len());
/// display_table_button.label().copy_from_slice(label.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_button.lpszLabel.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
/// }
///
/// let display_table_button: *const sys::DTBLBUTTON = display_table_button.as_ptr();
///
/// SizedDtblButton! { DisplayTableButtonW[u16; LABEL.len()] }
///
/// let mut display_table_button = DisplayTableButtonW::default();
/// assert_eq!(display_table_button.ulFlags, sys::MAPI_UNICODE);
///
/// let label: Vec<_> = LABEL.encode_utf16().collect();
/// assert_eq!(LABEL.len(), label.len());
/// display_table_button.label().copy_from_slice(label.as_slice());
/// unsafe {
///     assert_eq!(
///         PCWSTR::from_raw(display_table_button.lpszLabel.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
/// }
///
/// let display_table_button: *const sys::DTBLBUTTON = display_table_button.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblButton {
    ($name:ident [ $char:ident; $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub ulbLpszLabel: u32,
            pub ulFlags: u32,
            pub ulPRControl: u32,
            pub lpszLabel: [$char; $count + 1],
        }

        outlook_mapi_macros::impl_sized_struct_casts!($name, $crate::sys::DTBLBUTTON);

        outlook_mapi_macros::impl_sized_struct_default!($name {
            ulbLpszLabel: std::mem::size_of::<$crate::sys::DTBLBUTTON>() as u32,
            ulFlags: outlook_mapi_macros::display_table_default_flags!(
                $char,
                $crate::sys::MAPI_UNICODE
            ),
            ulPRControl: $crate::sys::PR_NULL,
            lpszLabel: [0; $count + 1],
        });

        impl $name {
            pub fn label(&mut self) -> &mut [$char] {
                &mut self.lpszLabel[..$count]
            }
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLPAGE`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLPAGE`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLPAGE`
///
/// It also initializes the [`sys::DTBLPAGE::ulFlags`] member and implements either of these
/// accessor to fill in the string buffer, depending on whether it is declared with [`u8`] or
/// [`u16`]:
///
/// - [`u8`]: `fn label(&mut self) -> &mut [u8]`, and `fn context(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn label(&mut self) -> &mut [u16]`, and `fn context(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// use outlook_mapi::{sys, SizedDtblPage};
/// use windows_core::{PCSTR, PCWSTR};
///
/// const LABEL: &str = "Page Label";
/// const COMPONENT: &str = "Page Component";
///
/// SizedDtblPage! { DisplayTablePageA[u8; LABEL.len(); COMPONENT.len()] }
///
/// let mut display_table_page = DisplayTablePageA::default();
/// assert_eq!(display_table_page.ulFlags, 0);
///
/// let label: Vec<_> = LABEL.bytes().collect();
/// assert_eq!(LABEL.len(), label.len());
/// display_table_page.label().copy_from_slice(label.as_slice());
/// let component: Vec<_> = COMPONENT.bytes().collect();
/// assert_eq!(COMPONENT.len(), component.len());
/// display_table_page.component().copy_from_slice(component.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_page.lpszLabel.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
///     assert_eq!(
///         PCSTR::from_raw(display_table_page.lpszComponent.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         COMPONENT);
/// }
///
/// let display_table_page: *const sys::DTBLPAGE = display_table_page.as_ptr();
///
/// SizedDtblPage! { DisplayTablePageW[u16; LABEL.len(); COMPONENT.len()] }
///
/// let mut display_table_page = DisplayTablePageW::default();
/// assert_eq!(display_table_page.ulFlags, sys::MAPI_UNICODE);
///
/// let label: Vec<_> = LABEL.encode_utf16().collect();
/// assert_eq!(LABEL.len(), label.len());
/// display_table_page.label().copy_from_slice(label.as_slice());
/// let component: Vec<_> = COMPONENT.encode_utf16().collect();
/// assert_eq!(COMPONENT.len(), component.len());
/// display_table_page.component().copy_from_slice(component.as_slice());
/// unsafe {
///     assert_eq!(
///         PCWSTR::from_raw(display_table_page.lpszLabel.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
///     assert_eq!(
///         PCWSTR::from_raw(display_table_page.lpszComponent.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         COMPONENT);
/// }
///
/// let display_table_page: *const sys::DTBLPAGE = display_table_page.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblPage {
    ($name:ident [ $char:ident; $count1:expr; $count2:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub ulbLpszLabel: u32,
            pub ulFlags: u32,
            pub ulbLpszComponent: u32,
            pub ulContext: u32,
            pub lpszLabel: [$char; $count1 + 1],
            pub lpszComponent: [$char; $count2 + 1],
        }

        outlook_mapi_macros::impl_sized_struct_casts!($name, $crate::sys::DTBLPAGE);

        outlook_mapi_macros::impl_sized_struct_default!($name {
            ulbLpszLabel: std::mem::size_of::<$crate::sys::DTBLPAGE>() as u32,
            ulFlags: outlook_mapi_macros::display_table_default_flags!(
                $char,
                $crate::sys::MAPI_UNICODE
            ),
            ulbLpszComponent: (std::mem::size_of::<$crate::sys::DTBLPAGE>()
                + std::mem::size_of::<[$char; $count1 + 1]>()) as u32,
            ulContext: 0,
            lpszLabel: [0; $count1 + 1],
            lpszComponent: [0; $count2 + 1],
        });

        impl $name {
            pub fn label(&mut self) -> &mut [$char] {
                &mut self.lpszLabel[..$count1]
            }

            pub fn component(&mut self) -> &mut [$char] {
                &mut self.lpszComponent[..$count2]
            }
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLRADIOBUTTON`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLRADIOBUTTON`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLRADIOBUTTON`
///
/// It also initializes the [`sys::DTBLRADIOBUTTON::ulFlags`] member and implements either of these
/// accessor to fill in the string buffer, depending on whether it is declared with [`u8`] or
/// [`u16`]:
///
/// - [`u8`]: `fn label(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn label(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// use outlook_mapi::{sys, SizedDtblRadioButton};
/// use windows_core::{PCSTR, PCWSTR};
///
/// const LABEL: &str = "Radiobutton Label";
///
/// SizedDtblRadioButton! { DisplayTableRadioButtonA[u8; LABEL.len()] }
///
/// let mut display_table_radio_button = DisplayTableRadioButtonA::default();
/// assert_eq!(display_table_radio_button.ulFlags, 0);
///
/// let label: Vec<_> = LABEL.bytes().collect();
/// assert_eq!(LABEL.len(), label.len());
/// display_table_radio_button.label().copy_from_slice(label.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_radio_button.lpszLabel.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
/// }
///
/// let display_table_radio_button: *const sys::DTBLRADIOBUTTON = display_table_radio_button.as_ptr();
///
/// SizedDtblRadioButton! { DisplayTableRadioButtonW[u16; LABEL.len()] }
///
/// let mut display_table_radio_button = DisplayTableRadioButtonW::default();
/// assert_eq!(display_table_radio_button.ulFlags, sys::MAPI_UNICODE);
///
/// let label: Vec<_> = LABEL.encode_utf16().collect();
/// assert_eq!(LABEL.len(), label.len());
/// display_table_radio_button.label().copy_from_slice(label.as_slice());
/// unsafe {
///     assert_eq!(
///         PCWSTR::from_raw(display_table_radio_button.lpszLabel.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
/// }
///
/// let display_table_radio_button: *const sys::DTBLRADIOBUTTON = display_table_radio_button.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblRadioButton {
    ($name:ident [ $char:ident; $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub ulbLpszLabel: u32,
            pub ulFlags: u32,
            pub ulcButtons: u32,
            pub ulPropTag: u32,
            pub lReturnValue: i32,
            pub lpszLabel: [$char; $count + 1],
        }

        outlook_mapi_macros::impl_sized_struct_casts!($name, $crate::sys::DTBLRADIOBUTTON);

        outlook_mapi_macros::impl_sized_struct_default!($name {
            ulbLpszLabel: std::mem::size_of::<$crate::sys::DTBLRADIOBUTTON>() as u32,
            ulFlags: outlook_mapi_macros::display_table_default_flags!(
                $char,
                $crate::sys::MAPI_UNICODE
            ),
            ulcButtons: 0,
            ulPropTag: $crate::sys::PR_NULL,
            lReturnValue: 0,
            lpszLabel: [0; $count + 1],
        });

        impl $name {
            pub fn label(&mut self) -> &mut [$char] {
                &mut self.lpszLabel[..$count]
            }
        }
    };
}

#[cfg(test)]
mod tests {
    use crate::*;
    use std::mem;

    #[test]
    fn sized_entry_id_1() {
        SizedENTRYID!(EntryId[1]);
        assert_eq!(mem::size_of::<sys::ENTRYID>(), mem::size_of::<EntryId>(),);
    }

    #[test]
    fn sized_prop_tag_array_1() {
        SizedSPropTagArray!(PropTagArray[1]);
        assert_eq!(
            mem::size_of::<sys::SPropTagArray>(),
            mem::size_of::<PropTagArray>(),
        );
    }

    #[test]
    fn sized_prop_problem_array_1() {
        SizedSPropProblemArray!(PropProblemArray[1]);
        assert_eq!(
            mem::size_of::<sys::SPropProblemArray>(),
            mem::size_of::<PropProblemArray>(),
        );
    }

    #[test]
    fn sized_adr_list_1() {
        SizedADRLIST!(AdrList[1]);
        assert_eq!(mem::size_of::<sys::ADRLIST>(), mem::size_of::<AdrList>(),);
    }

    #[test]
    fn sized_row_set_1() {
        SizedSRowSet!(RowSet[1]);
        assert_eq!(mem::size_of::<sys::SRowSet>(), mem::size_of::<RowSet>(),);
    }

    #[test]
    fn sized_sort_order_set_1() {
        SizedSSortOrderSet!(SortOrderSet[1]);
        assert_eq!(
            mem::size_of::<sys::SSortOrderSet>(),
            mem::size_of::<SortOrderSet>(),
        );
    }
}
