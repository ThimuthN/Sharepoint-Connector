# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0] - 2026-05-14

### Added
- **New `--folder-url` parameter** - Simplified usage by accepting full folder URLs
  - Users can now copy-paste SharePoint folder URLs directly
  - Example: `--folder-url "https://company.sharepoint.com/Documents/Inbox"`
- **Dual URL format support** - Both simple and flexible URL approaches work
  - Simple: `--folder-url` with just the filename
  - Flexible: `--sharepoint-url` + `--remote-path` for granular control
- **Large file support** - Upload files up to 100 MB
  - Previous limit: 4 MB
  - New limit: 100 MB with 5-minute timeout
  - Automatic retry logic for network failures
- **Wrapper function fixes** - Fixed parameter passing in CLI wrappers
  - `_save_profile()` now accepts `token_store_cls` parameter
  - `_ensure_profile_token()` now accepts `auth_cls` parameter

### Improved
- **Timeout settings** - Increased from 30s to 300s (5 minutes)
  - Better support for large file uploads
  - More reliable on slower networks
- **Upload session handling** - Now uses simple uploads for all files up to 100 MB
  - More reliable than session-based chunked uploads
  - Faster and simpler implementation
- **Progress logging** - Better feedback during upload operations
  - Upload progress shown as percentage
  - File size and speed information logged
- **Documentation** - Comprehensive README with examples
  - Quick start guide (2-minute setup)
  - Azure app registration instructions
  - Usage examples for all operations
  - Troubleshooting guide
  - Real-world RPA examples (UiPath, Python, Bash)

### Fixed
- **URL encoding** - Proper handling of special characters in filenames
- **Error messages** - More actionable error descriptions

## [1.0.0] - 2026-05-01

### Added
- **Initial release** with core functionality:
  - Upload files to SharePoint/OneDrive
  - Download files
  - List folder contents
  - Delete files
  - Move files between folders
  - Create nested folders
  - Check if files exist
- **One-time OAuth setup** - Browser-based authentication with PKCE flow
- **Encrypted credential storage** - Local token encryption
- **Automatic token refresh** - Silent background token refresh
- **CLI interface** - Command-line tool for all operations
- **Python SDK** - High-level API for Python bots
- **Profile support** - Multiple named profiles for different accounts/locations
- **Conflict handling** - Options to overwrite, rename, or fail on conflicts
- **JSON output** - Machine-readable output for bot parsing
- **Retry logic** - Automatic retry for transient failures

---

## Migration Guide

### From v1.0.0 to v1.1.0

No breaking changes! Your existing setup will continue to work.

**New:** You can now use the simpler `--folder-url` approach:

**Old way (still works):**
```bash
python -m rpa_sharepoint_connector run --profile default --op upload \
  --sharepoint-url "https://company.sharepoint.com" \
  --remote-path "Documents/Inbox/file.pdf" \
  --local-path file.pdf
```

**New way (simpler):**
```bash
python -m rpa_sharepoint_connector run --profile default --op upload \
  --folder-url "https://company.sharepoint.com/Documents/Inbox" \
  --local-path file.pdf \
  --remote-path file.pdf
```

Both approaches work and produce identical results!

---

## Future Roadmap

- [ ] Chunked uploads for files > 100 MB
- [ ] Batch operations (upload multiple files at once)
- [ ] OneDrive for Business support improvements
- [ ] Share files with specific users
- [ ] Copy files to another location
- [ ] File version history
- [ ] Search files by name/content
- [ ] Performance optimizations

---

## Known Issues

None currently reported. Please [file an issue](https://github.com/yourusername/rpa-sharepoint-connector/issues) if you encounter any problems.

---

## Support

For questions or issues, please:
1. Check the [README](README.md) troubleshooting section
2. Search [existing issues](https://github.com/yourusername/rpa-sharepoint-connector/issues)
3. Open a new issue with details about your problem
