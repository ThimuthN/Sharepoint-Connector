# What's New

## 1.1.0 (May 14, 2026)

**New**
- `--folder-url` parameter - just copy-paste the SharePoint folder URL from your browser
- Both URL formats work now - use simple mode or the detailed mode, whatever you prefer
- Files up to 100 MB now (was 4 MB before)
- 5-minute timeout for uploads (was 30 seconds) - way better for larger files
- Better progress logging so you know what's happening

**Fixed**
- URL encoding issues with special characters in filenames
- CLI wrapper functions weren't accepting all the right parameters

**Docs**
- New comprehensive README with 2-minute quick start
- Step-by-step Azure setup instructions
- Examples for UiPath, Python, and Bash
- Troubleshooting section with common errors

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
