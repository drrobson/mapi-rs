//! Define [`MAPIBuffer`] and [`MAPIOutParam`].
//!
//! Smart pointer types for memory allocated with [`sys::MAPIAllocateBuffer`], which must be freed
//! with [`sys::MAPIFreeBuffer`], or [`sys::MAPIAllocateMore`], which is chained to another
//! allocation and must not outlive that allocation or be separately freed.

use crate::sys;
use core::{
    ffi,
    marker::PhantomData,
    mem::{self, MaybeUninit},
    ptr, slice,
};
use windows::Win32::Foundation::E_OUTOFMEMORY;
use windows_core::{Error, HRESULT};

/// Errors which can be returned from this module.
#[derive(Debug)]
pub enum MAPIAllocError {
    /// The underlying [`sys::MAPIAllocateBuffer`] and [`sys::MAPIAllocateMore`] take a `u32`
    /// parameter specifying the size of the buffer. If you exceed the capacity of a `u32`, it will
    /// be impossible to make the necessary allocation.
    SizeOverflow(usize),

    /// MAPI APIs like to work with input and output buffers using `*const u8` and `*mut u8` rather
    /// than strongly typed pointers. In C++, this requires a lot of `reinterpret_cast` operations.
    /// For the accessors on this type, we'll allow transmuting the buffer to the desired type, as
    /// long as it fits in the allocation. If you don't allocate enough space for the number of
    /// elements you are accessing, it will return this error.
    OutOfBoundsAccess,

    /// There are no [documented](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/mapiallocatebuffer)
    /// conditions where [`sys::MAPIAllocateBuffer`] or [`sys::MAPIAllocateMore`] will return an
    /// error, but if they do fail, this will propagate the [`Error`] result. If the allocation
    /// function returns `null` with no other error, it will treat that as [`E_OUTOFMEMORY`].
    AllocationFailed(Error),

    /// Once [`MAPIBuffer::assume_init`] or [`MAPIBuffer::assume_init_slice`] has been called once,
    /// we assume that the buffer has been fully initialized. If you call either of those functions
    /// more than once, it will return this error.
    AlreadyInitialized,

    /// You must call [`MAPIBuffer::assume_init`] or [`MAPIBuffer::assume_init_slice`] before any
    /// calls to [`MAPIBuffer::as_mut`] or [`MAPIBuffer::as_mut_slice`]. If you don't, those calls
    /// will return this error.
    NotYetInitialized,
}

enum Buffer<T>
where
    T: Sized,
{
    Uninit(*mut MaybeUninit<T>),
    Ready(*mut T),
}

enum MAPIAlloc<'a, T>
where
    T: Sized,
{
    Root {
        buffer: Buffer<T>,
        byte_count: usize,
    },
    More {
        buffer: Buffer<T>,
        byte_count: usize,
        root: *mut ffi::c_void,
        phantom: PhantomData<&'a T>,
    },
}

impl<'a, T> MAPIAlloc<'a, T>
where
    T: Sized,
{
    fn new(count: usize) -> Result<Self, MAPIAllocError> {
        let byte_count = count * mem::size_of::<T>();
        Ok(Self::Root {
            buffer: unsafe {
                let mut alloc = ptr::null_mut();
                HRESULT::from_win32(sys::MAPIAllocateBuffer(
                    u32::try_from(byte_count)
                        .map_err(|_| MAPIAllocError::SizeOverflow(byte_count))?,
                    &mut alloc,
                ) as u32)
                .ok()
                .map_err(MAPIAllocError::AllocationFailed)?;
                if alloc.is_null() {
                    return Err(MAPIAllocError::AllocationFailed(Error::from_hresult(
                        E_OUTOFMEMORY,
                    )));
                }
                Buffer::Uninit(alloc as *mut _)
            },
            byte_count,
        })
    }

    fn chain<P>(&'a self, count: usize) -> Result<MAPIAlloc<'a, P>, MAPIAllocError>
    where
        P: Sized,
    {
        let root = match self {
            Self::Root { buffer, .. } => match buffer {
                Buffer::Uninit(alloc) => *alloc as *mut _,
                Buffer::Ready(alloc) => *alloc as *mut _,
            },
            Self::More { root, .. } => *root,
        };
        let byte_count = count * mem::size_of::<T>();
        Ok(MAPIAlloc::More {
            buffer: unsafe {
                let mut alloc = ptr::null_mut();
                HRESULT::from_win32(sys::MAPIAllocateMore(
                    u32::try_from(byte_count)
                        .map_err(|_| MAPIAllocError::SizeOverflow(byte_count))?,
                    root,
                    &mut alloc,
                ) as u32)
                .ok()
                .map_err(MAPIAllocError::AllocationFailed)?;
                if alloc.is_null() {
                    return Err(MAPIAllocError::AllocationFailed(Error::from_hresult(
                        E_OUTOFMEMORY,
                    )));
                }
                Buffer::Uninit(alloc as *mut _)
            },
            byte_count,
            root,
            phantom: PhantomData,
        })
    }

    fn into<P>(self) -> Result<MAPIAlloc<'a, P>, MAPIAllocError> {
        let result = match self {
            Self::Root {
                buffer: Buffer::Uninit(alloc),
                byte_count,
            } => Ok(MAPIAlloc::Root {
                buffer: Buffer::Uninit(alloc as *mut _),
                byte_count,
            }),
            Self::More {
                buffer: Buffer::Uninit(alloc),
                byte_count,
                root,
                ..
            } => Ok(MAPIAlloc::More {
                buffer: Buffer::Uninit(alloc as *mut _),
                byte_count,
                root,
                phantom: PhantomData,
            }),
            _ => Err(MAPIAllocError::AlreadyInitialized),
        };
        if result.is_ok() {
            mem::forget(self);
        }
        result
    }

    fn uninit(&mut self) -> Result<&mut MaybeUninit<T>, MAPIAllocError> {
        let (alloc, byte_count) = match self {
            Self::Root {
                buffer: Buffer::Uninit(alloc),
                byte_count,
            } => (alloc, byte_count),
            Self::More {
                buffer: Buffer::Uninit(alloc),
                byte_count,
                ..
            } => (alloc, byte_count),
            _ => return Err(MAPIAllocError::AlreadyInitialized),
        };
        if mem::size_of::<T>() > *byte_count {
            return Err(MAPIAllocError::OutOfBoundsAccess);
        }
        Ok(unsafe { &mut *(*alloc) })
    }

    fn uninit_slice(&mut self, count: usize) -> Result<&mut [MaybeUninit<T>], MAPIAllocError> {
        let (alloc, byte_count) = match self {
            Self::Root {
                buffer: Buffer::Uninit(alloc),
                byte_count,
            } => (alloc, byte_count),
            Self::More {
                buffer: Buffer::Uninit(alloc),
                byte_count,
                ..
            } => (alloc, byte_count),
            _ => return Err(MAPIAllocError::AlreadyInitialized),
        };
        if mem::size_of::<T>() * count > *byte_count {
            return Err(MAPIAllocError::OutOfBoundsAccess);
        }
        Ok(unsafe { slice::from_raw_parts_mut(*alloc, count) })
    }

    unsafe fn assume_init(&mut self) -> Result<&mut T, MAPIAllocError> {
        let (buffer, byte_count) = match self {
            Self::Root { buffer, byte_count } => (buffer, byte_count),
            Self::More {
                buffer, byte_count, ..
            } => (buffer, byte_count),
        };
        if mem::size_of::<T>() > *byte_count {
            return Err(MAPIAllocError::OutOfBoundsAccess);
        }
        let result;
        *buffer = match buffer {
            Buffer::Uninit(alloc) => {
                result = *alloc as *mut T;
                Buffer::Ready(result)
            }
            _ => return Err(MAPIAllocError::AlreadyInitialized),
        };
        Ok(&mut *result)
    }

    unsafe fn assume_init_slice(&mut self, count: usize) -> Result<&mut [T], MAPIAllocError> {
        let (buffer, byte_count) = match self {
            Self::Root { buffer, byte_count } => (buffer, byte_count),
            Self::More {
                buffer, byte_count, ..
            } => (buffer, byte_count),
        };
        if mem::size_of::<T>() * count > *byte_count {
            return Err(MAPIAllocError::OutOfBoundsAccess);
        }
        let result;
        *buffer = match buffer {
            Buffer::Uninit(alloc) => {
                result = *alloc as *mut T;
                Buffer::Ready(result)
            }
            Buffer::Ready(_) => return Err(MAPIAllocError::AlreadyInitialized),
        };
        Ok(slice::from_raw_parts_mut(result, count))
    }

    fn as_mut(&mut self) -> Result<&mut T, MAPIAllocError> {
        let (alloc, byte_count) = match self {
            Self::Root {
                buffer: Buffer::Ready(alloc),
                byte_count,
            } => (alloc, byte_count),
            Self::More {
                buffer: Buffer::Ready(alloc),
                byte_count,
                ..
            } => (alloc, byte_count),
            _ => return Err(MAPIAllocError::NotYetInitialized),
        };
        if mem::size_of::<T>() > *byte_count {
            Err(MAPIAllocError::OutOfBoundsAccess)
        } else {
            Ok(unsafe { &mut **alloc })
        }
    }

    fn as_mut_slice(&mut self, count: usize) -> Result<&mut [T], MAPIAllocError> {
        let (alloc, byte_count) = match self {
            Self::Root {
                buffer: Buffer::Ready(alloc),
                byte_count,
            } => (alloc, byte_count),
            Self::More {
                buffer: Buffer::Ready(alloc),
                byte_count,
                ..
            } => (alloc, byte_count),
            _ => return Err(MAPIAllocError::NotYetInitialized),
        };
        if mem::size_of::<T>() * count > *byte_count {
            Err(MAPIAllocError::OutOfBoundsAccess)
        } else {
            Ok(unsafe { slice::from_raw_parts_mut(*alloc, count) })
        }
    }
}

impl<T> Drop for MAPIAlloc<'_, T> {
    fn drop(&mut self) {
        if let Self::Root { buffer, .. } = self {
            let alloc = match mem::replace(buffer, Buffer::Uninit(ptr::null_mut())) {
                Buffer::Uninit(alloc) => alloc as *mut T,
                Buffer::Ready(alloc) => alloc,
            };
            if !alloc.is_null() {
                #[cfg(test)]
                unreachable!();
                #[cfg(not(test))]
                unsafe {
                    sys::MAPIFreeBuffer(alloc as *mut _);
                }
            }
        }
    }
}

/// Wrapper type for an allocation with either [`sys::MAPIAllocateBuffer`] or
/// [`sys::MAPIAllocateMore`].
pub struct MAPIBuffer<'a, T>(MAPIAlloc<'a, T>)
where
    T: Sized;

impl<'a, T> MAPIBuffer<'a, T> {
    /// Create a new allocation with enough room for `count` elements of type `T` with a call to
    /// [`sys::MAPIAllocateBuffer`]. The buffer is freed as soon as the [`MAPIBuffer`] is dropped.
    ///
    /// If you call [`MAPIBuffer::chain`] to create any more allocations with
    /// [`sys::MAPIAllocateMore`], their lifetimes are constrained to the lifetime of this
    /// allocation and they will all be freed together in a single call to [`sys::MAPIFreeBuffer`].
    pub fn new(count: usize) -> Result<Self, MAPIAllocError> {
        Ok(Self(MAPIAlloc::new(count)?))
    }

    /// Create a new allocation with enough room for `count` elements of type `P` with a call to
    /// [`sys::MAPIAllocateMore`]. The result is a separate allocation that is not freed until
    /// `self` is dropped at the beginning of the chain.
    ///
    /// You may call [`MAPIBuffer::chain`] on the result as well, they will both share a root
    /// allocation created with [`MAPIBuffer::new`].
    pub fn chain<P>(&'a self, count: usize) -> Result<MAPIBuffer<'a, P>, MAPIAllocError> {
        Ok(MAPIBuffer::<'a, P>(self.0.chain::<P>(count)?))
    }

    /// Convert an uninitialized allocation to another type. You can use this, for example, to
    /// perform an allocation with extra space in a `&mut [u8]` buffer, and then cast that to a
    /// specific type. This is useful with the `CbNewXXX` functions in [`crate::sized_types`].
    pub fn into<P>(self) -> Result<MAPIBuffer<'a, P>, MAPIAllocError> {
        Ok(MAPIBuffer::<'a, P>(self.0.into::<P>()?))
    }

    /// Get an uninitialized out-parameter with enough room for a single element of type `T`.
    pub fn uninit(&mut self) -> Result<&mut MaybeUninit<T>, MAPIAllocError> {
        self.0.uninit()
    }

    /// Get an uninitialized out-parameter with enough room for `count` elements of type `T`.
    pub fn uninit_slice(&mut self, count: usize) -> Result<&mut [MaybeUninit<T>], MAPIAllocError> {
        self.0.uninit_slice(count)
    }

    /// Once the buffer is known to be completely filled in, get a reference to a single element of
    /// type `T`.
    ///
    /// # Safety
    ///
    /// Like [`MaybeUninit`], the caller must ensure that the buffer is completely initialized
    /// before calling [`MAPIBuffer::assume_init`]. It is undefined behavior to leave it
    /// uninitialized once we start accessing it.
    pub unsafe fn assume_init(&mut self) -> Result<&mut T, MAPIAllocError> {
        self.0.assume_init()
    }

    /// Once the buffer is known to be completely filled in, get a reference to a slice with
    /// `count` elements of type `T`.
    ///
    /// # Safety
    ///
    /// Like [`MaybeUninit`], the caller must ensure that the buffer is completely initialized
    /// before calling [`MAPIBuffer::assume_init_slice`]. It is undefined behavior to leave it
    /// uninitialized once we start accessing it.
    pub unsafe fn assume_init_slice(&mut self, count: usize) -> Result<&mut [T], MAPIAllocError> {
        self.0.assume_init_slice(count)
    }

    /// Access a single element of type `T` once it has been initialized with `assume_init`.
    pub fn as_mut(&mut self) -> Result<&mut T, MAPIAllocError> {
        self.0.as_mut()
    }

    /// Access a slice with `count` elements of type `T` once it has been initialized with
    /// `assume_init_slice`.
    pub fn as_mut_slice(&mut self, count: usize) -> Result<&mut [T], MAPIAllocError> {
        self.0.as_mut_slice(count)
    }
}

/// Hold an out-pointer for MAPI APIs which perform their own buffer allocations. This version does
/// not perform any validation of the buffer size, so the typed accessors are inherently unsafe.
pub struct MAPIOutParam<T>(*mut T)
where
    T: Sized;

impl<T> MAPIOutParam<T>
where
    T: Sized,
{
    /// Get a `*mut *mut T` suitable for use with a MAPI API that fills in an out-pointer
    /// with a newly allocated buffer.
    pub fn as_mut_ptr(&mut self) -> *mut *mut T {
        &mut self.0
    }

    /// Access a single element of type `T`.
    ///
    /// # Safety
    ///
    /// This version does not perform any validation of the buffer size, so the typed accessors are
    /// inherently unsafe. The only thing it handles is a `null` check.
    pub unsafe fn as_mut(&mut self) -> Option<&mut T> {
        self.0.as_mut()
    }

    /// Access a slice with `count` elements of type `T`.
    ///
    /// # Safety
    ///
    /// This version does not perform any validation of the buffer size, so the typed accessors are
    /// inherently unsafe. The only thing it handles is a `null` check.
    pub unsafe fn as_mut_slice(&mut self, count: usize) -> Option<&mut [T]> {
        if self.0.is_null() {
            None
        } else {
            Some(slice::from_raw_parts_mut(self.0, count))
        }
    }
}

impl<T> Default for MAPIOutParam<T>
where
    T: Sized,
{
    fn default() -> Self {
        Self(ptr::null_mut())
    }
}

impl<T> Drop for MAPIOutParam<T>
where
    T: Sized,
{
    fn drop(&mut self) {
        if !self.0.is_null() {
            #[cfg(test)]
            unreachable!();
            #[cfg(not(test))]
            unsafe {
                sys::MAPIFreeBuffer(self.0 as *mut _);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::*;

    SizedSPropTagArray! { TestTags[2] }

    const TEST_TAGS: TestTags = TestTags {
        cValues: 2,
        aulPropTag: [sys::PR_INSTANCE_KEY, sys::PR_SUBJECT_W],
    };

    #[test]
    fn buffer_uninit() {
        let mut buffer: MaybeUninit<TestTags> = MaybeUninit::uninit();
        let mut mapi_buffer = MAPIBuffer(MAPIAlloc::Root {
            buffer: Buffer::Uninit(&mut buffer),
            byte_count: mem::size_of::<TestTags>(),
        });
        assert!(mapi_buffer.uninit().is_ok());
        mem::forget(mapi_buffer);
    }

    #[test]
    fn buffer_into() {
        let mut buffer: [MaybeUninit<u8>; mem::size_of::<TestTags>()] =
            [MaybeUninit::uninit(); CbNewSPropTagArray(2)];
        let mut mapi_buffer = MAPIBuffer(MAPIAlloc::Root {
            buffer: Buffer::Uninit(buffer.as_mut_ptr()),
            byte_count: buffer.len(),
        });
        assert!(mapi_buffer.uninit().is_ok());
        let mut mapi_buffer = mapi_buffer.into::<TestTags>().expect("into failed");
        assert!(mapi_buffer.uninit().is_ok());
        mem::forget(mapi_buffer);
    }

    #[test]
    fn buffer_assume_init() {
        let mut buffer = MaybeUninit::uninit();
        let mut mapi_buffer = MAPIBuffer(MAPIAlloc::Root {
            buffer: Buffer::Uninit(&mut buffer),
            byte_count: mem::size_of_val(&buffer),
        });
        let buffer: &mut TestTags =
            unsafe { mapi_buffer.assume_init() }.expect("assume_init failed");
        *buffer = TEST_TAGS;
        let test_tags = mapi_buffer.as_mut().expect("as_mut failed");
        assert_eq!(TEST_TAGS.cValues, test_tags.cValues);
        assert_eq!(TEST_TAGS.aulPropTag, test_tags.aulPropTag);
        mem::forget(mapi_buffer);
    }
}
