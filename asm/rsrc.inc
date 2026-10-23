; ============================================================================
; Excel Viewer Pro — PE Resource Section (Multi-Resolution 7-Icon Suite)
; ============================================================================

include 'C:\asm\hdd\INCLUDE\MACRO\RESOURCE.INC'

section '.rsrc' resource data readable
  directory 3, icons, 14, group_icons

  resource icons,\
    1, 0409h, icon_data1,\
    2, 0409h, icon_data2,\
    3, 0409h, icon_data3,\
    4, 0409h, icon_data4,\
    5, 0409h, icon_data5,\
    6, 0409h, icon_data6,\
    7, 0409h, icon_data7

  resource group_icons,\
    1, 0409h, main_icon

  align 4
  icon_data1 dd RVA ico_raw1, 817, 0, 0
  ico_raw1 file 'app_icon.ico':118, 817
  align 4

  icon_data2 dd RVA ico_raw2, 1402, 0, 0
  ico_raw2 file 'app_icon.ico':935, 1402
  align 4

  icon_data3 dd RVA ico_raw3, 2089, 0, 0
  ico_raw3 file 'app_icon.ico':2337, 2089
  align 4

  icon_data4 dd RVA ico_raw4, 3131, 0, 0
  ico_raw4 file 'app_icon.ico':4426, 3131
  align 4

  icon_data5 dd RVA ico_raw5, 4061, 0, 0
  ico_raw5 file 'app_icon.ico':7557, 4061
  align 4

  icon_data6 dd RVA ico_raw6, 8126, 0, 0
  ico_raw6 file 'app_icon.ico':11618, 8126
  align 4

  icon_data7 dd RVA ico_raw7, 18555, 0, 0
  ico_raw7 file 'app_icon.ico':19744, 18555
  align 4

  main_icon dd RVA grp_header, 6 + 7 * 14, 0, 0
  grp_header:
    dw 0, 1, 7
    file 'app_icon.ico':6, 12
    dw 1
    file 'app_icon.ico':22, 12
    dw 2
    file 'app_icon.ico':38, 12
    dw 3
    file 'app_icon.ico':54, 12
    dw 4
    file 'app_icon.ico':70, 12
    dw 5
    file 'app_icon.ico':86, 12
    dw 6
    file 'app_icon.ico':102, 12
    dw 7
  align 4
