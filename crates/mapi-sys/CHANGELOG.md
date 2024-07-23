# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.5.6](https://github.com/wravery/mapi-rs/compare/outlook-mapi-sys-v0.5.5...outlook-mapi-sys-v0.5.6) - 2024-07-23

### Added
- Pick up DEFINE_OLEGUID macros in winmd

### Other
- *(deps)* Cleanup implied windows-rs features
- release
- Merge branch 'main' of https://github.com/wravery/mapi-rs
- *(deps)* Update windows-rs to 0.58
- release
- *(deps)* Update windows-rs to 0.57
- release
- Bump MSRV according to new clippy warnings
- release
- Pick up latest HRESULT definitions
- release
- Update windows-rs to 0.56
- release
- release
- update windows-rs dependencies
- Onboard relase-plz update
- Bump versions for latest changes
- Doc cleanup pass
- Make SizedXXX macros declare a custom type (like MAPI does)
- Simplify scoping of cfg attr
- Bump versions for next update
- Move load_mapi module to outlook-mapi-sys
- Get delay loading of olmapi32.dll working
- Rename mapi crate to outlook-mapi
- Rename mapi-sys crate to outlook-mapi-sys
- Clarify error message in build.rs
- Replace remaining referenecs/links to webview2 with MAPI
- Link with fake import libs from mapi-scrubbed project
- Control direct linking to olmapi32 with feature flag
- Initial project setup based on webview2-rs

## [0.5.5](https://github.com/wravery/mapi-rs/compare/outlook-mapi-sys-v0.5.4...outlook-mapi-sys-v0.5.5) - 2024-07-23

### Other
- Merge branch 'main' of https://github.com/wravery/mapi-rs
- *(deps)* Update windows-rs to 0.58

## [0.5.4](https://github.com/wravery/mapi-rs/compare/outlook-mapi-sys-v0.5.3...outlook-mapi-sys-v0.5.4) - 2024-07-15

### Added
- Expose a new function to attempt to load the Outlook MAPI subsystem DLL

## [0.5.3](https://github.com/wravery/mapi-rs/compare/outlook-mapi-sys-v0.5.2...outlook-mapi-sys-v0.5.3) - 2024-06-12

### Other
- *(deps)* Update windows-rs to 0.57

## [0.5.2](https://github.com/wravery/mapi-rs/compare/outlook-mapi-sys-v0.5.1...outlook-mapi-sys-v0.5.2) - 2024-05-17

### Other
- Bump MSRV according to new clippy warnings

## [0.5.1](https://github.com/wravery/mapi-rs/compare/outlook-mapi-sys-v0.5.0...outlook-mapi-sys-v0.5.1) - 2024-05-08

### Other
- Pick up latest HRESULT definitions

## [0.5.0](https://github.com/wravery/mapi-rs/compare/outlook-mapi-sys-v0.4.3...outlook-mapi-sys-v0.5.0) - 2024-04-12

### Other
- Update windows-rs to 0.56

## [0.4.3](https://github.com/wravery/mapi-rs/compare/outlook-mapi-sys-v0.4.2...outlook-mapi-sys-v0.4.3) - 2024-03-10

### Added
- Pick up DEFINE_OLEGUID macros in winmd

## [0.4.2](https://github.com/wravery/mapi-rs/compare/outlook-mapi-sys-v0.4.1...outlook-mapi-sys-v0.4.2) - 2024-03-01

### Other
- update windows-rs dependencies

## [0.4.1](https://github.com/wravery/mapi-rs/compare/outlook-mapi-sys-v0.4.0...outlook-mapi-sys-v0.4.1) - 2024-02-29
- Added CHANGELOG.md
