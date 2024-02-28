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
