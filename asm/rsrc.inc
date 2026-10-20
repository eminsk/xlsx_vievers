; ============================================================================
; Excel Viewer Pro — PE Resource Section
; ============================================================================

include 'C:\asm\hdd\INCLUDE\MACRO\RESOURCE.INC'

section '.rsrc' resource data readable
  directory 3, icons, 14, group_icons
  resource icons, 1, 0409h, icon_data
  resource group_icons, 1, 0409h, main_icon
  icon main_icon, icon_data, 'app_icon.ico'


