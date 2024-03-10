//! Define [`MAPIBuffer`] and [`MAPIOutParam`].
//!
//! Smart pointer types for memory allocated with [`sys::MAPIAllocateBuffer`], which must be freed
//! with [`sys::MAPIFreeBuffer`], or [`sys::MAPIAllocateMore`], which is chained to another
//! allocation and must not outlive that allocation or be separately freed.

use crate::sys;
use core::{
    ffi,
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
    Uninit(*mut T),
    Ready(*mut T),
}

trait TypedAlloc {
    fn as_mut_ptr(&self) -> *mut ffi::c_void;
}

enum MAPIAlloc<'a, T>
where
    T: Sized,
{
    Root(Buffer<T>, usize),
    More(Buffer<T>, usize, &'a dyn TypedAlloc),
}

impl<T> TypedAlloc for MAPIAlloc<'_, T> {
    fn as_mut_ptr(&self) -> *mut ffi::c_void {
        match self {
            Self::Root(buffer, _) => match buffer {
                Buffer::Uninit(alloc) => *alloc as *mut _,
                Buffer::Ready(alloc) => *alloc as *mut _,
            },
            Self::More(_, _, alloc) => alloc.as_mut_ptr(),
        }
    }
}

impl<'a, T> MAPIAlloc<'a, T>
where
    T: Sized,
{
    fn new(count: usize) -> Result<Self, MAPIAllocError> {
        let byte_count = count * mem::size_of::<T>();
        Ok(Self::Root(
            unsafe {
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
        ))
    }

    fn chain<P>(&'a self, count: usize) -> Result<MAPIAlloc<'a, P>, MAPIAllocError>
    where
        P: Sized,
    {
        let root = self.as_mut_ptr();
        let byte_count = count * mem::size_of::<T>();
        Ok(MAPIAlloc::<'a, P>::More(
            unsafe {
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
            self,
        ))
    }

    fn uninit(&mut self) -> Result<&mut MaybeUninit<T>, MAPIAllocError> {
        let (alloc, byte_count) = match self {
            Self::Root(Buffer::Uninit(alloc), byte_count) => (alloc, byte_count),
            Self::More(Buffer::Uninit(alloc), byte_count, _) => (alloc, byte_count),
            _ => return Err(MAPIAllocError::AlreadyInitialized),
        };
        if mem::size_of::<T>() > *byte_count {
            return Err(MAPIAllocError::OutOfBoundsAccess);
        }
        Ok(unsafe { &mut *(*alloc as *mut _) })
    }

    fn uninit_slice(&mut self, count: usize) -> Result<&mut [MaybeUninit<T>], MAPIAllocError> {
        let (alloc, byte_count) = match self {
            Self::Root(Buffer::Uninit(alloc), byte_count) => (alloc, byte_count),
            Self::More(Buffer::Uninit(alloc), byte_count, _) => (alloc, byte_count),
            _ => return Err(MAPIAllocError::AlreadyInitialized),
        };
        if mem::size_of::<T>() * count > *byte_count {
            return Err(MAPIAllocError::OutOfBoundsAccess);
        }
        Ok(unsafe { slice::from_raw_parts_mut(*alloc as *mut _, count) })
    }

    unsafe fn assume_init(&mut self) -> Result<&mut T, MAPIAllocError> {
        let (buffer, byte_count) = match self {
            Self::Root(buffer, byte_count) => (buffer, byte_count),
            Self::More(buffer, byte_count, _) => (buffer, byte_count),
        };
        if mem::size_of::<T>() != *byte_count {
            return Err(MAPIAllocError::OutOfBoundsAccess);
        }
        let mut result = MaybeUninit::uninit();
        *buffer = match buffer {
            Buffer::Uninit(alloc) => {
                result.write(*alloc);
                Buffer::Ready(*alloc)
            }
            Buffer::Ready(_) => return Err(MAPIAllocError::AlreadyInitialized),
        };
        Ok(&mut *(result.assume_init()))
    }

    unsafe fn assume_init_slice(&mut self, count: usize) -> Result<&mut [T], MAPIAllocError> {
        let (buffer, byte_count) = match self {
            Self::Root(buffer, byte_count) => (buffer, byte_count),
            Self::More(buffer, byte_count, _) => (buffer, byte_count),
        };
        if mem::size_of::<T>() * count != *byte_count {
            return Err(MAPIAllocError::OutOfBoundsAccess);
        }
        let mut result = MaybeUninit::uninit();
        *buffer = match buffer {
            Buffer::Uninit(alloc) => {
                result.write(*alloc);
                Buffer::Ready(*alloc)
            }
            Buffer::Ready(_) => return Err(MAPIAllocError::AlreadyInitialized),
        };
        Ok(slice::from_raw_parts_mut(result.assume_init(), count))
    }

    fn as_mut(&mut self) -> Result<&mut T, MAPIAllocError> {
        let (alloc, byte_count) = match self {
            Self::Root(Buffer::Ready(alloc), byte_count) => (alloc, byte_count),
            Self::More(Buffer::Ready(alloc), byte_count, _) => (alloc, byte_count),
            _ => return Err(MAPIAllocError::NotYetInitialized),
        };
        if mem::size_of::<T>() != *byte_count {
            Err(MAPIAllocError::OutOfBoundsAccess)
        } else {
            Ok(unsafe { &mut *(*alloc) })
        }
    }

    fn as_mut_slice(&mut self, count: usize) -> Result<&mut [T], MAPIAllocError> {
        let (alloc, byte_count) = match self {
            Self::Root(Buffer::Ready(alloc), byte_count) => (alloc, byte_count),
            Self::More(Buffer::Ready(alloc), byte_count, _) => (alloc, byte_count),
            _ => return Err(MAPIAllocError::NotYetInitialized),
        };
        if mem::size_of::<T>() * count != *byte_count {
            Err(MAPIAllocError::OutOfBoundsAccess)
        } else {
            Ok(unsafe { slice::from_raw_parts_mut(*alloc, count) })
        }
    }
}

impl<T> Drop for MAPIAlloc<'_, T> {
    fn drop(&mut self) {
        if let Self::Root(buffer, _) = self {
            let alloc = match mem::replace(buffer, Buffer::Uninit(ptr::null_mut())) {
                Buffer::Uninit(alloc) => alloc,
                Buffer::Ready(alloc) => alloc,
            };
            if !alloc.is_null() {
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
            unsafe {
                sys::MAPIFreeBuffer(self.0 as *mut _);
            }
        }
    }
}
