use crate::{sys::*, *};
use core::{ptr, slice};

pub struct RowSet {
    rows: *mut SRowSet,
}

impl RowSet {
    pub fn as_mut_ptr(&mut self) -> *mut *mut SRowSet {
        &mut self.rows
    }

    pub fn is_empty(&self) -> bool {
        unsafe {
            self.rows
                .as_ref()
                .map(|rows| rows.cRows == 0)
                .unwrap_or(true)
        }
    }

    pub fn len(&self) -> usize {
        unsafe {
            self.rows
                .as_ref()
                .map(|rows| rows.cRows as usize)
                .unwrap_or_default()
        }
    }
}

impl Default for RowSet {
    fn default() -> Self {
        Self {
            rows: ptr::null_mut(),
        }
    }
}

impl IntoIterator for RowSet {
    type Item = Row;
    type IntoIter = <Vec<Self::Item> as IntoIterator>::IntoIter;

    fn into_iter(self) -> Self::IntoIter {
        unsafe {
            if let Some(rows) = self.rows.as_mut() {
                let count = rows.cRows as usize;
                let data: &mut [SRow] = slice::from_raw_parts_mut(rows.aRow.as_mut_ptr(), count);
                let data = data.iter_mut().map(Row::new).collect();
                data
            } else {
                vec![]
            }
        }
        .into_iter()
    }
}

impl Drop for RowSet {
    fn drop(&mut self) {
        if !self.rows.is_null() {
            unsafe {
                FreeProws(self.rows);
            }
        }
    }
}
