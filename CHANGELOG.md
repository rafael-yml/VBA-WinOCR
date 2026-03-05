# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## 1.1.0 - 2626-03-05
Forked from [DanysysTeam/VBA-UWPOCR](https://github.com/DanysysTeam/VBA-UWPOCR).

- Replaced `Kernel32.dll` (`RtlMoveMemory`) with `msvcrt.dll` (`memcpy`) 
  to avoid triggering Windows Defender on certain corporate environments
- Added timeout/retry counter to `WaitForAsyncInterface`
- Fixed `RECT` type fields from `Single` to `Long`
- Fixed `ResultArray` indexing bug in word array mode

## 1.0.0 - 2024-04-04
### First Release :tada:
