#[macro_use]
extern crate outlook_mapi;

use core::{ptr, slice};
use outlook_mapi::{
    mapi_initialize, mapi_logon, row_set, Microsoft::Office::Outlook::MAPI::Win32::*,
};
use std::sync::Arc;
use windows_core::*;

fn main() -> Result<()> {
    println!("Initializing MAPI...");
    let initialized =
        mapi_initialize::Initialize::new(Default::default()).expect("failed to initialize MAPI");
    println!("Trying to logon to the default profile...");
    let logon = mapi_logon::Logon::new(
        Arc::new(initialized),
        Default::default(),
        None,
        None,
        mapi_logon::Flags {
            extended: true,
            unicode: true,
            logon_ui: true,
            use_default: true,
            ..Default::default()
        },
    )
    .expect("should be able to logon to the default MAPI profile");
    println!("Success!");

    // Now try to list the stores in the default MAPI profile.
    let mut rows: row_set::RowSet = Default::default();
    unsafe {
        let stores_table = logon.session.GetMsgStoresTable(0)?;
        HrQueryAllRows(
            &stores_table,
            SizedSPropTagArray!(2, PR_ENTRYID, PR_DISPLAY_NAME_W),
            ptr::null_mut(),
            SizedSSortOrderSet!(
                1,
                0,
                0,
                SSortOrder {
                    ulPropTag: PR_DISPLAY_NAME_W,
                    ulOrder: TABLE_SORT_ASCEND,
                }
            ),
            50,
            rows.as_mut_ptr(),
        )?;
    }

    println!("Found {rows} stores", rows = rows.len());
    for (idx, row) in rows.into_iter().enumerate() {
        assert_eq!(2, row.len());
        let mut values = row.iter();
        let entry_id = values.next().expect("missing entry ID");
        assert_eq!(entry_id.ulPropTag, PR_ENTRYID);
        let display_name = values.next().expect("missing display name");
        assert_eq!(display_name.ulPropTag, PR_DISPLAY_NAME_W);
        unsafe {
            let entry_id =
                slice::from_raw_parts(entry_id.Value.bin.lpb, entry_id.Value.bin.cb as usize);
            let display_name = String::from_utf16(display_name.Value.lpszW.as_wide())?;
            println!(
                "Store {idx}: {display_name} ({entry_id} byte ID)",
                idx = idx + 1,
                entry_id = entry_id.len()
            );
        }
    }

    Ok(())
}
