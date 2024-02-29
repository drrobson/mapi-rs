//! Define [`PropValue`] and [`PropValueData`].

use crate::{sys, PropTag};
use core::{ffi, mem, slice};
use windows::Win32::{
    Foundation::{E_INVALIDARG, E_POINTER, FILETIME},
    System::Com::CY,
};
use windows_core::*;

/// Wrapper for a [`sys::SPropValue`] structure which allows pattern matching on [`PropValueData`].
pub struct PropValue<'a> {
    pub tag: PropTag,
    pub value: PropValueData<'a>,
}

/// Enum with values from the original [`sys::SPropValue::Value`] union.
pub enum PropValueData<'a> {
    /// [`sys::PT_I2`] or [`sys::PT_SHORT`]
    Short(i16),

    /// [`sys::PT_I4`] or [`sys::PT_LONG`]
    Long(i32),

    /// [`sys::PT_PTR`] or [`sys::PT_FILE_HANDLE`]
    Pointer(*mut ffi::c_void),

    /// [`sys::PT_R4`] or [`sys::PT_FLOAT`]
    Float(f32),

    /// [`sys::PT_R8`] or [`sys::PT_DOUBLE`]
    Double(f64),

    /// [`sys::PT_BOOLEAN`]
    Boolean(u16),

    /// [`sys::PT_CURRENCY`]
    Currency(i64),

    /// [`sys::PT_APPTIME`]
    AppTime(f64),

    /// [`sys::PT_SYSTIME`]
    FileTime(FILETIME),

    /// [`sys::PT_STRING8`]
    AnsiString(PCSTR),

    /// [`sys::PT_BINARY`]
    Binary(&'a [u8]),

    /// [`sys::PT_UNICODE`]
    Unicode(PCWSTR),

    /// [`sys::PT_CLSID`]
    Guid(&'a GUID),

    /// [`sys::PT_I8`] or [`sys::PT_LONGLONG`]
    LargeInteger(i64),

    /// [`sys::PT_MV_SHORT`]
    ShortArray(&'a [i16]),

    /// [`sys::PT_MV_LONG`]
    LongArray(&'a [i32]),

    /// [`sys::PT_MV_FLOAT`]
    FloatArray(&'a [f32]),

    /// [`sys::PT_MV_DOUBLE`]
    DoubleArray(&'a [f64]),

    /// [`sys::PT_MV_CURRENCY`]
    CurrencyArray(&'a [CY]),

    /// [`sys::PT_MV_APPTIME`]
    AppTimeArray(&'a [f64]),

    /// [`sys::PT_MV_SYSTIME`]
    FileTimeArray(&'a [FILETIME]),

    /// [`sys::PT_MV_BINARY`]
    BinaryArray(&'a [sys::SBinary]),

    /// [`sys::PT_MV_STRING8`]
    AnsiStringArray(&'a [PCSTR]),

    /// [`sys::PT_MV_UNICODE`]
    UnicodeArray(&'a [PCWSTR]),

    /// [`sys::PT_MV_CLSID`]
    GuidArray(&'a [GUID]),

    /// [`sys::PT_MV_LONGLONG`]
    LargeIntegerArray(&'a [i64]),

    /// [`sys::PT_ERROR`]
    Error(HRESULT),

    /// [`sys::PT_NULL`] or [`sys::PT_OBJECT`]
    Object(i32),
}

impl<'a> From<&'a sys::SPropValue> for PropValue<'a> {
    /// Convert a [`sys::SPropValue`] reference into a friendlier [`PropValue`] type, which often
    /// supports safe access to the [`sys::SPropValue::Value`] union.
    fn from(value: &sys::SPropValue) -> Self {
        let tag = PropTag::from(value.ulPropTag);
        let prop_type = tag.prop_type() as u32 & !sys::MV_INSTANCE;
        let data = unsafe {
            match prop_type {
                sys::PT_SHORT => PropValueData::Short(value.Value.i),
                sys::PT_LONG => PropValueData::Long(value.Value.l),
                sys::PT_PTR => PropValueData::Pointer(value.Value.lpv),
                sys::PT_FLOAT => PropValueData::Float(value.Value.flt),
                sys::PT_DOUBLE => PropValueData::Double(value.Value.dbl),
                sys::PT_BOOLEAN => PropValueData::Boolean(value.Value.b),
                sys::PT_CURRENCY => PropValueData::Currency(value.Value.cur.int64),
                sys::PT_APPTIME => PropValueData::AppTime(value.Value.at),
                sys::PT_SYSTIME => PropValueData::FileTime(value.Value.ft),
                sys::PT_STRING8 => {
                    if value.Value.lpszA.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::AnsiString(PCSTR::from_raw(value.Value.lpszA.as_ptr()))
                    }
                }
                sys::PT_BINARY => {
                    if value.Value.bin.lpb.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::Binary(slice::from_raw_parts(
                            value.Value.bin.lpb,
                            value.Value.bin.cb as usize,
                        ))
                    }
                }
                sys::PT_UNICODE => {
                    if value.Value.lpszW.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::Unicode(PCWSTR::from_raw(value.Value.lpszW.as_ptr()))
                    }
                }
                sys::PT_CLSID => value
                    .Value
                    .lpguid
                    .as_ref()
                    .map(PropValueData::Guid)
                    .unwrap_or(PropValueData::Error(E_POINTER)),
                sys::PT_LONGLONG => PropValueData::LargeInteger(value.Value.li),
                sys::PT_MV_SHORT => {
                    if value.Value.MVi.lpi.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::ShortArray(slice::from_raw_parts(
                            value.Value.MVi.lpi,
                            value.Value.MVi.cValues as usize,
                        ))
                    }
                }
                sys::PT_MV_LONG => {
                    if value.Value.MVl.lpl.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::LongArray(slice::from_raw_parts(
                            value.Value.MVl.lpl,
                            value.Value.MVl.cValues as usize,
                        ))
                    }
                }
                sys::PT_MV_FLOAT => {
                    if value.Value.MVflt.lpflt.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::FloatArray(slice::from_raw_parts(
                            value.Value.MVflt.lpflt,
                            value.Value.MVflt.cValues as usize,
                        ))
                    }
                }
                sys::PT_MV_DOUBLE => {
                    if value.Value.MVdbl.lpdbl.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::DoubleArray(slice::from_raw_parts(
                            value.Value.MVdbl.lpdbl,
                            value.Value.MVdbl.cValues as usize,
                        ))
                    }
                }
                sys::PT_MV_CURRENCY => {
                    if value.Value.MVcur.lpcur.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::CurrencyArray(slice::from_raw_parts(
                            value.Value.MVcur.lpcur,
                            value.Value.MVcur.cValues as usize,
                        ))
                    }
                }
                sys::PT_MV_APPTIME => {
                    if value.Value.MVat.lpat.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::AppTimeArray(slice::from_raw_parts(
                            value.Value.MVat.lpat,
                            value.Value.MVat.cValues as usize,
                        ))
                    }
                }
                sys::PT_MV_SYSTIME => {
                    if value.Value.MVft.lpft.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::FileTimeArray(slice::from_raw_parts(
                            value.Value.MVft.lpft,
                            value.Value.MVft.cValues as usize,
                        ))
                    }
                }
                sys::PT_MV_BINARY => {
                    if value.Value.MVbin.lpbin.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::BinaryArray(slice::from_raw_parts(
                            value.Value.MVbin.lpbin,
                            value.Value.MVbin.cValues as usize,
                        ))
                    }
                }
                sys::PT_MV_STRING8 => {
                    if value.Value.MVszA.lppszA.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::AnsiStringArray(slice::from_raw_parts(
                            mem::transmute(value.Value.MVszA.lppszA),
                            value.Value.MVszA.cValues as usize,
                        ))
                    }
                }
                sys::PT_MV_UNICODE => {
                    if value.Value.MVszW.lppszW.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::UnicodeArray(slice::from_raw_parts(
                            mem::transmute(value.Value.MVszW.lppszW),
                            value.Value.MVszW.cValues as usize,
                        ))
                    }
                }
                sys::PT_MV_CLSID => {
                    if value.Value.MVguid.lpguid.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::GuidArray(slice::from_raw_parts(
                            value.Value.MVguid.lpguid,
                            value.Value.MVguid.cValues as usize,
                        ))
                    }
                }
                sys::PT_MV_LONGLONG => {
                    if value.Value.MVli.lpli.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::LargeIntegerArray(slice::from_raw_parts(
                            value.Value.MVli.lpli,
                            value.Value.MVli.cValues as usize,
                        ))
                    }
                }
                sys::PT_ERROR => PropValueData::Error(HRESULT(value.Value.err)),
                sys::PT_NULL | sys::PT_OBJECT => PropValueData::Object(value.Value.x),
                _ => PropValueData::Error(E_INVALIDARG),
            }
        };
        PropValue { tag, value: data }
    }
}
