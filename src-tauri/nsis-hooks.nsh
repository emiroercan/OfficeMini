; Give each registered file type its own icon (Tauri uses the app icon for all of them).
!macro NSIS_HOOK_POSTINSTALL
  WriteRegStr SHELL_CONTEXT "Software\Classes\Markdown Document\DefaultIcon" "" "$INSTDIR\icons\md.ico"
  WriteRegStr SHELL_CONTEXT "Software\Classes\Word Document\DefaultIcon" "" "$INSTDIR\icons\docx.ico"
  WriteRegStr SHELL_CONTEXT "Software\Classes\Excel Workbook\DefaultIcon" "" "$INSTDIR\icons\xlsx.ico"
  WriteRegStr SHELL_CONTEXT "Software\Classes\CSV Table\DefaultIcon" "" "$INSTDIR\icons\csv.ico"
  ; Tell Explorer that icons changed.
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend
