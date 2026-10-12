; Give shortcuts a versioned icon-file path so Windows cannot reuse its old
; cached icon for the executable path after an update.
!macro NSIS_HOOK_POSTINSTALL
  CreateShortCut "$SMPROGRAMS\unArchive.lnk" "$INSTDIR\unarchive.exe" "" "$INSTDIR\unArchive-shortcut-v0.1.4.ico" 0
  CreateShortCut "$DESKTOP\unArchive.lnk" "$INSTDIR\unarchive.exe" "" "$INSTDIR\unArchive-shortcut-v0.1.4.ico" 0
!macroend
