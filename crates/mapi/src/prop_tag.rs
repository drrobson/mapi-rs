//! Utilities for accessing the `PROP_TYPE` and `PROP_ID` portions of a `u32` `PROP_TAG`.

/// Simple wrapper for a MAPI `PROP_TAG`.
#[repr(transparent)]
pub struct PropTag(pub u32);

impl PropTag {
    /// Combine the `PROP_ID` and `PROP_TYPE` to form a [`PropTag`].
    pub const fn new(prop_id: u16, prop_type: u16) -> Self {
        Self(((prop_id as u32) << 16) | (prop_type as u32))
    }

    /// Extract the `PROP_ID` portion of the [`PropTag`].
    pub const fn prop_id(&self) -> u16 {
        ((self.0 & 0xFFFF_0000) >> 16) as u16
    }

    /// Extract the `PROP_TYPE` portion of the [`PropTag`].
    pub const fn prop_type(&self) -> u16 {
        (self.0 & 0xFFFF) as u16
    }
}

impl From<u32> for PropTag {
    /// Wrap a constant `PROP_TAG` value in a [`PropTag`].
    fn from(value: u32) -> Self {
        Self(value)
    }
}

impl From<PropTag> for u32 {
    /// Get a constant `PROP_TAG` value from a [`PropTag`].
    fn from(value: PropTag) -> Self {
        value.0
    }
}
