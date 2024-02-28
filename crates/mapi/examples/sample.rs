use core::ptr;
use outlook_mapi::{sys::*, *};
use windows_core::*;

fn main() -> Result<()> {
    println!("Initializing MAPI...");
    let initialized = Initialize::new(Default::default()).expect("failed to initialize MAPI");
    println!("Trying to logon to the default profile...");
    let logon = Logon::new(
        initialized,
        Default::default(),
        None,
        None,
        LogonFlags {
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
    let mut rows: RowSet = Default::default();
    unsafe {
        let stores_table = logon.session.GetMsgStoresTable(0)?;
        HrQueryAllRows(
            &stores_table,
            SizedSPropTagArray!([PR_ENTRYID, PR_DISPLAY_NAME_W]),
            ptr::null_mut(),
            SizedSSortOrderSet!({
                categories: 0,
                expanded: 0,
                sorts: [SSortOrder {
                    ulPropTag: PR_DISPLAY_NAME_W,
                    ulOrder: TABLE_SORT_ASCEND,
                }]
            }),
            50,
            rows.as_mut_ptr(),
        )?;
    }

    println!("Found {rows} stores", rows = rows.len());
    for (idx, row) in rows.into_iter().enumerate() {
        // Use 1-based indices for messages.
        let idx = idx + 1;

        assert_eq!(2, row.len());
        let mut values = row.iter();

        let Some(PropValue {
            tag: PR_ENTRYID,
            value: PropValueData::Binary(entry_id),
        }) = values.next()
        else {
            eprintln!("Store {idx}: missing entry ID");
            continue;
        };

        let Some(PropValue {
            tag: PR_DISPLAY_NAME_W,
            value: PropValueData::Unicode(display_name),
        }) = values.next()
        else {
            eprintln!("Store {idx}: missing display name");
            continue;
        };
        let display_name = unsafe { display_name.to_string() }
            .unwrap_or_else(|err| format!("bad display name: {err}"));

        println!(
            "Store {idx}: {display_name} ({entry_id} byte ID)",
            entry_id = entry_id.len()
        );
    }

    Ok(())
}
