# Changelog

Each release's section is published as its GitHub release notes and update notes, and the app shows this history in About unArchive. Before tagging a release, rename **Unreleased** to the new version and date.

## [Unreleased]

### Added
- A File menu with Check for Updates and About unArchive. About shows the installed version and the release history.
- An Update available indicator beside the version number. It opens the new version's release notes with the option to install it.
- Create all channels as new, an import option that builds every channel from its CPS template using only the settings in the workbook.
- An import summary listing channels kept, renamed, moved, changed in type, created, and removed.
- Channel Pool drop-downs for Assigned Zone and Channel Type, and Contact Name, Privacy Enabled, Privacy Key Alias, and RAS Key Alias for Digital pool channels.

### Changed
- A channel row without its hidden record ID is always a new channel, so a reused channel name no longer brings back an old channel's settings.
- Import refuses copied rows, record IDs from the other channel tab, and rows whose hidden IDs were left out of a sort, and names each row.
- Scan List shows and sets each channel's actual CPS scan list.

### Fixed
- Receive Only, Privacy Enabled, and Auto Roam accept only True or False, keep the stored value when left blank, and are refused on channel types that do not have them.
- Key aliases left on a channel changed to Analog are reported instead of silently dropped.
- Receive-only channels no longer need a contact.

## [0.2.2] - 2026-10-12

### Changed
- The window opens large enough to show the whole interface without scrolling on a typical laptop screen.

## [0.2.1] - 2026-10-12

### Added
- Start a new layout can replace the pasted contacts with numbered Capacity Plus, Linked Capacity Plus, or Digital contacts, including All Call contacts, and can clear the Channel Pool.
- Each contact block can start at a chosen Call ID.

### Changed
- Capacity Plus Call IDs can be any number from 1 through 16776415, except 255, which is reserved for All Call.

## [0.2.0] - 2026-10-12

### Added
- The installed version is shown at the top of the app, with Check for updates and Install update.
- Updates are downloaded from GitHub releases and installed only when signed with the project's key.

## [0.1.0] - 2026-10-12

### Added
- First release: turns CPS clipboard XML into an editable Excel workbook and rebuilds CPS XML from it, with a new zone layout builder.
