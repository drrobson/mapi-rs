use crate::sys::*;
use core::{mem, slice};
use std::ptr;

pub struct Row {
    count: usize,
    props: *mut SPropValue,
}

impl Row {
    pub fn new(row: &mut SRow) -> Self {
        Self {
            count: mem::replace(&mut row.cValues, 0) as usize,
            props: mem::replace(&mut row.lpProps, ptr::null_mut()),
        }
    }

    pub fn is_empty(&self) -> bool {
        self.count == 0 || self.props.is_null()
    }

    pub fn len(&self) -> usize {
        self.count
    }

    pub fn iter(&self) -> impl Iterator<Item = &SPropValue> {
        if self.props.is_null() {
            vec![]
        } else {
            unsafe {
                let data: &[SPropValue] = slice::from_raw_parts(self.props, self.count);
                let data = data.iter().collect();
                data
            }
        }
        .into_iter()
    }
}

impl Drop for Row {
    fn drop(&mut self) {
        if !self.props.is_null() {
            unsafe {
                MAPIFreeBuffer(mem::transmute(self.props));
            }
        }
    }
}
