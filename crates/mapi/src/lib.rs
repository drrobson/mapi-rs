pub use outlook_mapi_sys::Microsoft::Office::Outlook::MAPI::Win32 as sys;

pub mod mapi_initialize;
pub mod mapi_logon;
pub mod row;
pub mod row_set;

pub use mapi_initialize::*;
pub use mapi_logon::*;
pub use row::*;
pub use row_set::*;

#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedENTRYID {
    ({ flags: $flags:expr, bytes: $bytes:expr }) => {
        std::mem::transmute::<_, &mut $crate::sys::ENTRYID>(&mut {
            #[repr(C)]
            struct EntryId {
                flags: [u8; 4],
                bytes: [u8; $bytes.len()],
            }
            EntryId {
                flags: $flags,
                bytes: $bytes,
            }
        })
    };
}

#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSPropTagArray {
    ($tags:expr) => {
        std::mem::transmute::<_, &mut $crate::sys::SPropTagArray>(&mut {
            const COUNT: usize = $tags.len();
            #[repr(C)]
            struct PropTagArray {
                count: u32,
                tags: [u32; COUNT],
            }
            PropTagArray {
                count: COUNT as u32,
                tags: $tags,
            }
        })
    };
}

#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSPropProblemArray {
    ($problems:expr) => {
        std::mem::transmute::<_, &mut $crate::sys::SPropProblemArray>(&mut {
            const COUNT: usize = $problems.len();
            #[repr(C)]
            struct ProblemArray {
                count: u32,
                problems: [$crate::sys::SPropProblem; COUNT],
            }
            ProblemArray {
                count: COUNT as u32,
                problems: $problems,
            }
        })
    };
}

#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedADRLIST {
    ($entries:expr) => {
        std::mem::transmute::<_, &mut $crate::sys::ADRLIST>(&mut {
            const COUNT: usize = $entries.len();
            #[repr(C)]
            struct AdrList {
                count: u32,
                entries: [$crate::sys::ADRENTRY; COUNT],
            }
            AdrList {
                count: COUNT as u32,
                entries: $entries,
            }
        })
    };
}

#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSRowSet {
    ($rows:expr) => {
        std::mem::transmute::<_, &mut $crate::sys::SRowSet>(&mut {
            const COUNT: usize = $rows.len();
            #[repr(C)]
            struct RowSet {
                count: u32,
                rows: [$crate::sys::SRow; COUNT],
            }
            RowSet {
                count: COUNT as u32,
                rows: $rows,
            }
        })
    };
}

#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSSortOrderSet {
    ({ categories: $categories:expr, expanded: $expanded:expr, sorts: $sorts:expr }) => {
        std::mem::transmute::<_, &mut $crate::sys::SSortOrderSet>(&mut {
            const COUNT: usize = $sorts.len();

            let count = COUNT;
            let categories = $categories;
            let expanded = $expanded;

            assert!(categories <= count);
            assert!(expanded <= categories);

            #[repr(C)]
            struct SortOrderSet {
                count: u32,
                categories: u32,
                expanded: u32,
                sorts: [$crate::sys::SSortOrder; COUNT],
            }
            SortOrderSet {
                count: count as u32,
                categories: categories as u32,
                expanded: expanded as u32,
                sorts: $sorts,
            }
        })
    };
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::{mem, ptr};

    #[test]
    fn sized_macros() {
        assert_eq!(
            mem::size_of::<sys::ENTRYID>(),
            mem::size_of_val(unsafe { SizedENTRYID!({ flags: [0; 4], bytes: [0; 1] }) })
        );
        assert_eq!(
            mem::size_of::<sys::SPropTagArray>(),
            mem::size_of_val(unsafe { SizedSPropTagArray!([sys::PR_NULL]) })
        );
        assert_eq!(
            mem::size_of::<sys::SPropProblemArray>(),
            mem::size_of_val(unsafe {
                SizedSPropProblemArray!([sys::SPropProblem {
                    ulIndex: 0,
                    ulPropTag: sys::PR_NULL,
                    scode: sys::MAPI_E_NOT_FOUND.0
                }])
            })
        );
        assert_eq!(
            mem::size_of::<sys::ADRLIST>(),
            mem::size_of_val(unsafe {
                SizedADRLIST!([sys::ADRENTRY {
                    ulReserved1: 0,
                    cValues: 0,
                    rgPropVals: ptr::null_mut(),
                }])
            })
        );
        assert_eq!(
            mem::size_of::<sys::SRowSet>(),
            mem::size_of_val(unsafe {
                SizedSRowSet!([sys::SRow {
                    ulAdrEntryPad: 0,
                    cValues: 0,
                    lpProps: ptr::null_mut(),
                }])
            })
        );
        assert_eq!(
            mem::size_of::<sys::SSortOrderSet>(),
            mem::size_of_val(unsafe {
                SizedSSortOrderSet!({
                    categories: 0,
                    expanded: 0,
                    sorts: [sys::SSortOrder {
                        ulPropTag: sys::PR_NULL,
                        ulOrder: sys::TABLE_SORT_ASCEND,
                    }]
                })
            })
        );
    }
}
