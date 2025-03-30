' ----------------------------------------
' Constants
' ----------------------------------------
Const DEBUG_MODE As Boolean = False
Const WORKSHEET_NAME As String = "worksheet"
Const GRAPH_NAME As String = "graph"

' Not organized yet
Const HIDE_LEGEND As Long = 0
Const AXIS_X As Long = 1
Const AXIS_Y As Long = 2
Const TICK_THICKNESS_00 As Variant = &H00000000
Const TITLE_HIDE As Long = 0
Const SSA_COLOR_ALPHA As Long = &H000008a7&
Const HORIZONTAL As Long = 0
Const VERTICAL As Long = 1
Const LABEL_PTS_00 As Variant = "0"
Const LABEL_PTS_07 As String = "97"
Const LABEL_PTS_08 As String = "111"
Const TICK_THICKNESS_08 As Variant = &H00000008
Const TICK_WIDTH_20 As Variant = &H00000020

' Axis scale types
Const SAA_TYPE_LINEAR = 1
Const SAA_TYPE_COMMON = 2
Const SAA_TYPE_LOG = 3
Const SAA_TYPE_PROBABILITY = 4
Const SAA_TYPE_PROBIT = 5
Const SAA_TYPE_LOGIT = 6
Const SAA_TYPE_CATEGORY = 7
Const SAA_TYPE_DATETIME = 8

' For Plot Wizard
' ----------------------------------------
' Columns
Const _PLOT_TYPE_EXPLANATION_COL As Long = 0
Const PLOT_TYPE_COL As Long = 1
' Rows
Const PLOT_TYPE_ROW As Long = 0
Const PLOT_STYLE_ROW As Long = 1
Const PLOT_DATA_TYPE_ROW As Long = 2
Const _PLOT_COLUMNS_PER_PLOT_ROW As Long = 3 ' Spacer
Const _PLOT_PLOT_COLUMNS_COUNT_ARRAY_ROW As Long = 4 ' Spacer
Const PLOT_DATA_SOURCE_ROW As Long = 5
Const PLOT_POLARUNITS_ROW As Long = 6
Const PLOT_ANGLEUNITS_ROW As Long = 7
Const PLOT_MIN_ANGLE_ROW As Long = 8
Const PLOT_MAX_ANGLE_ROW As Long = 8
Const PLOT_UNKONWN1_ROW As Long = 9
Const PLOT_GROUP_STYLE_ROW As Long = 10
Const PLOT_USE_AUTOMATIC_LEGENDS_ROW As Long = 11

' Graph Parameters
' ----------------------------------------
' Columns
Const _GRAPH_PARAMS_EXPLANATION_COL As Long = 2
Const GRAPH_PARAMS_COL As Long = 3
' Rows
Const X_LABEL_ROW As Long = 0
Const X_MM_ROW As Long = 1
Const X_SCALE_TYPE_ROW As Long = 2
Const X_MIN_ROW As Long = 3
Const X_MAX_ROW As Long = 4
Const _X_TICKS_ROW As Long = 5
Const Y_LABEL_ROW As Long = 6
Const Y_MM_ROW As Long = 7
Const Y_SCALE_TYPE_ROW As Long = 8
Const Y_MIN_ROW As Long = 9
Const Y_MAX_ROW As Long = 10
Const _Y_TICKS_ROW As Long = 11

' Ticks (Not handled by macros but embedded in JNB file)
' ----------------------------------------
Const _X_TICKS_COL As Long = 7
Const _Y_TICKS_COL As Long = 14

' Data Columns
' ----------------------------------------
' General
Const DATA_START_COL As Long = 16
Const DATA_CHUNK_SIZE As Long = 16
Const DATA_MAX_NUM_CHUNKS As Long = 13
' In-chunk indices
Const DATA_ID_X AS LONG = 0
Const DATA_ID_X_ERR AS LONG = 1
Const DATA_ID_X_UPPER AS LONG = 2
Const DATA_ID_X_LOWER AS LONG = 3
Const DATA_ID_Y AS LONG = 4
Const DATA_ID_Y_ERR AS LONG = 5
Const DATA_ID_Y_UPPER AS LONG = 6
Const DATA_ID_Y_LOWER AS LONG = 7
Const DATA_ID_RGBA AS LONG = 8
