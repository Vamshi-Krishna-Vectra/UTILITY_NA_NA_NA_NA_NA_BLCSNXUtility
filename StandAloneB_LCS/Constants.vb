Module Constants
    Public Const MODULE_VERSION As String = "V1.0"
    Public Const EXCEL_OPEN_PASSWORD As String = "APRRVBIW"
    Public Const SHEETNAMEWRITE As String = "F_data"
    Public Const SHEETNAMEDRAWDETAILS As String = "Drw_data"
    Public Const SHEETFACEVECTORDETAILS As String = "FaceVec"
    Public Const BURNOUT_FACEVEC_SHEET As String = "BO_FaceVec"
    Public Const BURNOUT_FDATA_SHEET As String = "BO_F_data"
    Public Const BURNOUT_BODYNAMES_SHEET As String = "BO_Body_Name"
    Public Const MISCINFOSHEETNAME As String = "Misc_Info"
    Public Const BODYSHEETNAME As String = "Body_Name"
    Public Const HOLEVECTORSHEETNAME As String = "Hole_Vector"
    Public Const CONFIGSHEETNAME As String = "CONFIG"
    Public Const MATING_BODY_SHEET_NAME As String = "Mating_Body_Info"
    'Public Const MATING_BODY_FACE_INFO_TAB_1_SHEET_NAME As String = "Mating_Body_Face_Info_1"
    'Public Const MATING_BODY_FACE_INFO_TAB_2_SHEET_NAME As String = "Mating_Body_Face_Info_2"
    'Public Const MATING_BODY_FACE_INFO_TAB_3_SHEET_NAME As String = "Mating_Body_Face_Info_3"

    'Added - Amitabh - 9/22/16 - New format of reporting
    Public Const MATING_BODY_FACES_PARENT_HEADER As String = "Parent Body Mating Face"
    Public Const MATING_BODY_FACES_CHILD_HEADER As String = "Child Body Mating Face"

    Public Const WRORKBOOK_NAME As String = "AutoDim_DataParse"
    Public Const XLSM As String = ".xlsm"

    Public Const FEATURE_TYPE As String = "Feature Type"
    Public Const FEATURE_NAME As String = "Feature Name"
    Public Const FACE_NAME As String = "Face Name"
    Public Const FACE_TYPE As String = "Face Type"
    Public Const FACE_AREA As String = "Face Area"
    Public Const PRE_FAB As String = "Pre-Fab"
    Public Const FEAT_NAME As String = "Feat Name"
    Public Const FACECENTERX As String = "Center-X"
    Public Const FACECENTERY As String = "Center-Y"
    Public Const FACECENTERZ As String = "Centr-Z"
    Public Const FACERADIUS As String = "Radius"
    Public Const FACEDIRECTION As String = "Direction"
    Public Const EDGE_NAME As String = "Edge Name"
    Public Const EDGETYPE As String = "Edge Type"
    Public Const EDGEDIA As String = "Edge Diameter"
    Public Const EDGECENTERX As String = "Edge Center X"
    Public Const EDGECENTERY As String = "Edge Center Y"
    Public Const EDGECENTERZ As String = "Edge Center Z"
    Public Const START_FACE As String = "Start Face"
    Public Const VERTEX1X As String = "Vertex 1-X"
    Public Const VERTEX1Y As String = "Vertex 1-Y"
    Public Const VERTEX1Z As String = "Vertex 1-Z"

    Public Const VERTEX2X As String = "Vertex 2-X"
    Public Const VERTEX2Y As String = "Vertex 2-Y"
    Public Const VERTEX2Z As String = "Vertex 2-Z"

    Public Const EDGE_LENGTH As String = "Edge Length"

    Public Const HOLEX As String = "Hole-X"
    Public Const HOLEY As String = "Hole-Y"
    Public Const HOLEZ As String = "Hole-Z"
    Public Const DEPTH As String = "Hole Depth"
    Public Const HOLE_SIZE As String = "HOLE_SIZE"

    Public Const FLAME_CUT_FACE As String = "FLAME_CUT_FACE"
    Public Const FLAME_CUT_FACE_ATTR As String = "B_FLAMECUT_FACE"
    Public Const FLAME_CUT_FACE_ATTR_VALUE As String = "Y"

    Public Const HOLE_DIA As String = "Hole Dia"
    Public Const START_HOLE As String = "Start Hole"
    Public Const HOLE_PARENT As String = "Hole Parent"
    Public Const THREAD_PARENT1 As String = "Thread Parent 1"
    Public Const THREAD_PARENT2 As String = "Thread Parent 2"
    Public Const CALLOUT As String = "Thread Callout"
    Public Const EDGE_CURVATURE As String = "Edge Curvature"
    Public Const BLENDRADIUS As String = "Blend Radius"
    Public Const VECTORX As String = "Vector-X"
    Public Const VECTORY As String = "Vector-Y"
    Public Const VECTORZ As String = "Vector-Z"

    Public Const CONCAVE As String = "Concave"
    Public Const CONVEX As String = "Convex"

    'Public Const MACHINESURFACE As String = "Finish Tol"

    Public Const COLUMN_TITLE_ROW_NUM As Integer = 2
    Public Const START_ROW_WRITE As Integer = 3
    Public Const START_COL_WRITE As Integer = 1
    Public Const PARENT_COL_COUNT As Integer = 3

    Public Const DRAWINGSHEETNAME As String = "Object"
    Public Const VIEWNAME As String = "View Name"
    Public Const VIEWSCALE As String = "View Scale"
    Public Const VIEWORIGINX As String = "View Origin-X"
    Public Const VIEWORIGINY As String = "View Origin-Y"
    Public Const VIEWORIGINZ As String = "View Origin-Z"
    Public Const Xx As String = "Xx"
    Public Const Xy As String = "Xy"
    Public Const Xz As String = "Xz"
    Public Const Yx As String = "Yx"
    Public Const Yy As String = "Yy"
    Public Const Yz As String = "Yz"
    Public Const Zx As String = "Zx"
    Public Const Zy As String = "Zy"
    Public Const Zz As String = "Zz"

    Public Const XC As String = "Xc"
    Public Const YC As String = "Yc"
    Public Const ZC As String = "Zc"
    Public Const TYPE As String = "Type"

    Public Const DRAW_COLUMN_TITLE_ROW_NUM As Integer = 1
    Public Const FACE_VECTOR_TITLE_ROW_NUM As Integer = 1
    Public Const MISC_INFO_TITLE_ROW_NUM As Integer = 1
    Public Const HOLE_VECTOR_TITLE_ROW_NUM As Integer = 1
    Public Const BODY_INFO_TITLE_ROW_NUM As Integer = 1

    Public Const DRAW_START_ROW_WRITE As Integer = 2
    Public Const DRAW_START_COL_WRITE As Integer = 1

    Public Const MISC_INFO_START_ROW_WRITE As Integer = 3
    Public Const MISC_INFO_START_COL_WRITE As Integer = 1

    Public Const HOLE_VECT_INFO_START_ROW_WRITE As Integer = 2
    Public Const HOLE_VECT_INFO_START_COL_WRITE As Integer = 1

    Public Const BODY_INFO_START_ROW_WRITE As Integer = 2
    Public Const BODY_INFO_START_COL_WRITE As Integer = 1

    'Public Const INPUT_FILE_PATH As String = "C:\Vectra\Input Files"
    'Public Const INPUT_FOLDER_PATH_TEMP As String = "C:\Vectra\Input Files\Temp"
    'Public Const INPUT_FOLDER_PATH_ERROR As String = "C:\Vectra\Sweep_Data\Error"
    'Public Const OUTPUT_FILE_PATH As String = "C:\Vectra\Output Files"

    Public Const ERROR_REPORT_FOLDER_PATH As String = "C:\Vectra\Output Files"
    Public Const DB_PART_NAME_IN_TEMPLATE As String = "DB_PART_NAME"
    'Public Const DB_PART_NAME As String = "DETAIL_NAME"
    Public Const NAME As String = "Name"
    Public Const VALUE As String = "Value"
    Public Const INCH_TO_MM As Double = 25.4

    Public Const HOLE_NAME As String = "Hole Name"

    Public Const BODY_NAME As String = "Body Name"
    Public Const SHAPE_NAME_INTEMPLATE As String = "SHAPE"
    'Public Const SHAPE As String = "DETAIL_SHAPE"
    Public Const SUB_DET_NUM_IN_TEMPLATE As String = "SUB_DET_NUM"
    'Public Const SUB_DET_NUM As String = "SUB_DETAIL_NUMBER"
    'Public Const TOOL_CLASS As String = "TOOL_CLASS"
    'Public Const TOOL_ID As String = "TOOL_ID"
    'Public Const P_MASS As String = "MASS"
    'Public Const PURCH_OPTION As String = "PURCHASE_OPTION"
    Public Const PURCHASE As String = "P"
    Public Const MAKE_DETAIL As String = "M"
    Public Const COMM As String = "COMM"
    Public Const STD As String = "STD"
    Public Const DESIGN_SOURCE As String = "DESIGN_SOURCE"
    Public Const COMPONENT_NAME As String = "Component Name"
    Public Const LAYER As String = "Layer"

    Public Const STITLE As String = "FACE"
    'Public Const SHORIZVALUE As String = "StartFaceHorizontal"
    'Public Const SVERTICALVALUE As String = "StartFaceVertical"
    'Public Const SLATERALVALUE As String = "Lateral"
    'Public Const SSTARTHOLEVALUE As String = "StartHole"
    'Public Const FINISHTOLATTRNAME As String = "FINISH_TOL"
    Public Const FINISHTOLATTRVALUE As String = "MICRONS"
    'Public Const PLANAR_MATING_FACE_TOLERANCE_VALUE As String = "3.2 MICRONS [VV]"
    'code added Dec-05-2017
    'Public Const RELIEF_CUT_FACE_TOLERANCE_VALUE As String = "6.4 MICRONS [V]"
    'Public Const SHORIZVALUEFAB As String = "StartFaceHorizontal_Fab"
    'Public Const SVERTICALVALUEFAB As String = "StartFaceVertical_Fab"
    'Public Const SLATERALVALUEFAB As String = "Lateral_Fab"

    'Public Const H_FAB As String = "X_FAB"
    'Public Const V_FAB As String = "Y_FAB"
    'Public Const L_FAB As String = "Z_FAB"
    Public Const NOTAPPLICABLE As String = "N/A"

    Public Const CREATED_DATE As String = "Created Date"
    Public Const MODIFIED_DATE As String = "Modified Date"
    Public Const HISTORY_SHEET_NAME As String = "Hist"
    Public Const HISTORY_SHEET_TITLE_ROW_NOS As Integer = 1

    'Public Const STOCK_SIZE As String = "STOCK_SIZE"
    Public Const SUB_DET_NUM_START_VALUE As Integer = 1

    'Public Const STOCK_SIZE_METRIC As String = "STOCK_SIZE_METRIC"
    Public Const STOCK_SIZE_METRIC_INFO_ROW_NOS As Integer = 3
    Public Const STOCK_SIZE_METRIC_INFO_COLUMN_NOS As Integer = 14
    'Public Const QTY As String = "QUANTITY"

    'For the View Directon Cosines
    Public Const Xxc As String = "Xxc"
    Public Const Xyc As String = "Xyc"
    Public Const Xzc As String = "Xzc"
    Public Const Yxc As String = "Yxc"
    Public Const Yyc As String = "Yyc"
    Public Const Yzc As String = "Yzc"
    Public Const Zxc As String = "Zxc"
    Public Const Zyc As String = "Zyc"
    Public Const Zzc As String = "Zzc"
    Public Const MODEL_VIEW_NAME As String = "Model View Name"
    Public Const TITLE_ROW_NUM_VIEW_DIR_COS As Integer = 1
    Public Const VIEW_DIR_COS_SHEET_NAME As String = "Model View Cosines"

    'For Batch Process
    'Public Const NXPART_FILES_INPUT_FOLDER_PATH As String = "C:\Vectra\Input Files\NXParts"
    'Public Const EXCEPTION_REPORT_FILE_NAME As String = "Exception Report.txt"

    'For the 3D Exception Report
    Public Const THREE_DIMENSIONAL_MODEL_EXCEPTION_REPORT_FILE_NAME As String = "3D Exception Report.txt"

    'Pre Fab Hole Attributes
    Public Const PRE_FAB_HOLE_ATTR_TITLE As String = "B_PREFAB"
    Public Const PRE_FAB_HOLE_ATTR_VALUE As String = "Y"
    Public Const FEAT_NAME_FACE_ATTR As String = "FEAT_NAME"

    'For computing the bounding box
    Public Const MIN_POINTX As String = "Min_point X"
    Public Const MIN_POINTY As String = "Min_point Y"
    Public Const MIN_POINTZ As String = "Min_point Z"
    Public Const VECTORXX As String = "Vector XX"
    Public Const VECTORXY As String = "Vector XY"
    Public Const VECTORXZ As String = "Vector XZ"
    Public Const Magnitude_X As String = "Magnitude X"
    Public Const VECTORYX As String = "Vector YX"
    Public Const VECTORYY As String = "Vector YY"
    Public Const VECTORYZ As String = "Vector YZ"
    Public Const Magnitude_Y As String = "Magnitude Y"
    Public Const VECTORZX As String = "Vector ZX"
    Public Const VECTORZY As String = "Vector ZY"
    Public Const VECTORZZ As String = "Vector ZZ"
    Public Const Magnitude_Z As String = "Magnitude Z"

    'Attribute to check whether the sweep data has been previously run
    Public Const VECTRA_SDR As String = "VECTRA_SDR"
    Public Const VECTRA_SDR_RUN_YES As Integer = 1

    'For Error log reporting to be stored in the HUB system
    Public Const ERROR_LOG_FILE_PREFIX As String = "Err_Log_"
    'Public Const ERROR_LOG_FILE_PATH As String = "J:\Error Log\"

    'For reading the CONFIG file
    'Public Const FOLDER_PATH As String = "C:\Vectra\Config\Drawing Automation\"
    'Public Const CONFIG_FILE_NAME As String = "Drawing_Config.txt"
    Public Const CONFIG_FOLDER_NAME As String = "Config"
    Public Const OEM_SUPPLIER_TEXT_FILE_NAME As String = "Oem_Supplier.txt"
    Public Const OUTPUT_FOLDER_PATH_TEXT_FILE_NAME As String = "Output_Folder_Path.txt"
    Public Const ATTRIBUTE_XML_CONFIG_FILE_NAME As String = "Oem_Supplier_Attribute_Mapping.xml"
    'Public Const PART_SWEEP_DATA_TEXT_FILE As String = "Part_Sweep_Data_List.txt"
    Public Const DESIGN_SOURCE_CONFIG_FILE_NAME As String = "Oem_Supplier_DesignSource.xml"
    Public Const BODY_TO_BODY_MATING_TOLERANCE As Double = 0.5
    Public Const BODY_TO_BODY_MATING_TOLERANCE_RELAXED As Double = 1.0
    Public Const PLANAR_FACE As String = "PLANAR"
    Public Const CYLINDRICAL_FACE As String = "CYLINDRICAL"
    Public Const OTHER_TYPES_FACE As String = "OTHER_TYPES"
    Public Const PLANAR_MATING_FACE_TOLERANCE As Double = 0.5
    Public Const CYLINDRICAL_MATING_FACE_TOLERANCE As Double = 0.1
    Public Const FACE_TO_FACE_MATING_TOLERANCE_RELAXED As Double = 1.0

    'Attribute Values reading from ADA
    Public Const B_WELDMENT_TYPE As String = "B_WELDMENT_TYPE"

    Public Const B_SHORIZVALUEMC_FRAME1 As String = "B_X_MC_DATUM_START"
    Public Const B_SVERTICALVALUEMC_FRAME1 As String = "B_Y_MC_DATUM_START"
    Public Const B_SLATERALVALUEMC_FRAME1 As String = "B_Z_MC_DATUM_START"

    'CODE ADDED - 4/18/16  - AMITABH - For FRAME-2
    Public Const B_SHORIZVALUEMC_FRAME2 As String = "B_U_MC_DATUM_START"
    Public Const B_SVERTICALVALUEMC_FRAME2 As String = "B_V_MC_DATUM_START"
    Public Const B_SLATERALVALUEMC_FRAME2 As String = "B_W_MC_DATUM_START"

    Public Const B_1ST_MC_FACE_X As String = "B_1ST_MC_FACE_X"
    Public Const B_1ST_MC_FACE_Y As String = "B_1ST_MC_FACE_Y"

    'CODE ADDED - 4/18/16  - AMITABH - For FRAME-2
    Public Const B_1ST_MC_FACE_U As String = "B_1ST_MC_FACE_U"
    Public Const B_1ST_MC_FACE_V As String = "B_1ST_MC_FACE_V"

    Public Const B_X_FAB As String = "B_X_FAB_ORIGIN"
    Public Const B_Y_FAB As String = "B_Y_FAB_ORIGIN"
    Public Const B_Z_FAB As String = "B_Z_FAB_ORIGIN"
    Public Const B_U_FAB As String = "B_U_FAB_ORIGIN"
    Public Const B_V_FAB As String = "B_V_FAB_ORIGIN"
    Public Const B_W_FAB As String = "B_W_FAB_ORIGIN"

    Public Const B_CUT_X_FACE As String = "B_X_CUT_ORIGIN"
    Public Const B_CUT_Y_FACE As String = "B_Y_CUT_ORIGIN"
    Public Const B_CUT_Z_FACE As String = "B_Z_CUT_ORIGIN"

    Public Const B_DATUM_HOLE_FRAME1 As String = "B_Datum Hole_FRAME1"
    Public Const B_DATUM_FACE_FRAME1 As String = "B_Datum Face_FRAME1"

    'CODE ADDED - 4/18/16  - AMITABH - For FRAME-2
    Public Const B_DATUM_HOLE_FRAME2 As String = "B_Datum Hole_FRAME2"
    Public Const B_DATUM_FACE_FRAME2 As String = "B_Datum Face_FRAME2"

    Public Const VECTRA_TITLE As String = "Vectra"
    Public Const ASSIGN_ATTR As String = "TRUE"

    Public Const BODY_NAME_START_COL_WRITE As Integer = 1
    Public Const CUT_X_FACE As String = "X_CUT_ORIGIN"
    Public Const CUT_Y_FACE As String = "Y_CUT_ORIGIN"
    Public Const CUT_Z_FACE As String = "Z_CUT_ORIGIN"
    Public Const PRIMARY_VIEW_NAME As String = "B_PRIMARY1"
    Public Const SECOND_PRIMARY_VIEW_NAME As String = "B_PRIMARY2"

    'NC PART CONTACT FACE DETERMINATION
    Public Const NC_PART_CONTACT_FACE As String = "NC Part Contact Face"
    Public Const NC_PART_CONTACT_FACE_STRAT_COL_WRITE As Integer = 1
    Public Const NC_Contact_FACE_ATTRIBUTE As String = "B_NCPartContactFace"
    'To Maintain consistency, value is changed from Y to TRUE across the OEM
    'Public Const NC_PCF_ATTR_VALUE As String = "TRUE"
    Public Const NC_PART_CONTACT_FACE_START_ROW_WRITE As Integer = 2
    Public Const FLOOR_MOUNT_FACE_START_ROW_WRITE As Integer = 2
    Public Const P_MAT_IN_TEMPLATE As String = "P_MAT"
    'Public Const P_MAT As String = "Lieferant/Werkstoff_W060"
    'Public Const P_MAT As String = "Werkstoff"

    'TEMPLATE FOLDER NAME
    Public Const TEMPLATE_FOLDER As String = "Templates"
    'Public Const SWEEP_DATA_TEMPLATE_NAME As String = "SweepData"
    'VECTRA SUPPORTING DIRECTORIES
    Public Const TEMP_UNIT_SWEEP_DATA As String = "Temp_Unit_Sweep_Data"
    Public Const DIMENSION_DATABASE_BACKUP As String = "Dimension_Database_Backup"
    Public Const PART_SWEEP_DATA_BACKUP As String = "Part_Sweep_Data_Backup"
    Public Const UNIT_SWEEP_DATA_BACKUP As String = "Unit_Sweep_Data_Backup"
    Public Const UNIT_DATABASES As String = "Unit_Databases"
    Public Const UNPROCESSED_PARTS As String = "UnProcessed_Parts"
    Public Const UNIT_SWEEP_DATA As String = "Unit_Sweep_Data"
    Public Const DRAWINGS As String = "Drawings"
    Public Const DIMENSION_DATABASES As String = "Dimension_Databases"
    Public Const PART_SWEEP_DATA_WITH_ADA As String = "Part_Sweep_Data_With_ADA"

    Public Const DESIGNER_LOG_FILE As String = "Designer_Log.txt"
    'Code added - Amitabh - 11/14/16
    Public Const LOG_FOLDER As String = "Log"

    Public Const FRONT_VIEW As String = "FRONT"
    Public Const LEFT_VIEW As String = "LEFT"
    Public Const TOP_VIEW As String = "TOP"
    Public Const BOTTOM_VIEW As String = "BOTTOM"
    Public Const RIGHT_VIEW As String = "RIGHT"
    Public Const REAR_VIEW As String = "BACK"

    Public Const sTempSheet As String = "SHT50"

    Public Const VISIBLEEDGE_SHEETNAME As String = "Visible_Edge_Names"
    Public Const EDGENAMES_HEADER As String = "Edge Names"
    Public Const FRONT_HEADER As String = "Front"
    Public Const LEFT_HEADER As String = "Left"
    Public Const RIGHT_HEADER As String = "Right"
    Public Const TOP_HEADER As String = "Top"
    Public Const BOTTOM_HEADER As String = "Bottom"
    Public Const REAR_HEADER As String = "Back"

    Public Const EDGE_VISIBLE As String = "Y"
    Public Const EDGE_NOT_VISIBLE As String = "N"

    Public Const PRIMARY1_FRONTVIEW As String = "PRIMARY1_FRONTVIEW"
    Public Const PRIMARY1_LEFTVIEW As String = "PRIMARY1_LEFTVIEW"
    Public Const PRIMARY1_TOPVIEW As String = "PRIMARY1_TOPVIEW"
    Public Const PRIMARY1_BOTTOMVIEW As String = "PRIMARY1_BOTTOMVIEW"
    Public Const PRIMARY1_RIGHTVIEW As String = "PRIMARY1_RIGHTVIEW"
    Public Const PRIMARY1_REARVIEW As String = "PRIMARY1_REARVIEW"

    Public Const PRIMARY2_FRONTVIEW As String = "PRIMARY2_FRONTVIEW"
    Public Const PRIMARY2_LEFTVIEW As String = "PRIMARY2_LEFTVIEW"
    Public Const PRIMARY2_TOPVIEW As String = "PRIMARY2_TOPVIEW"
    Public Const PRIMARY2_BOTTOMVIEW As String = "PRIMARY2_BOTTOMVIEW"
    Public Const PRIMARY2_RIGHTVIEW As String = "PRIMARY2_RIGHTVIEW"
    Public Const PRIMARY2_REARVIEW As String = "PRIMARY2_REARVIEW"

    Public Const B_PRIMARY1 As String = "B_PRIMARY1"
    Public Const B_PRIMARY2 As String = "B_PRIMARY2"

    Public Const PRIMARY1FRONT_HEADER As String = "Primary1_Front"
    Public Const PRIMARY1LEFT_HEADER As String = "Primary1_Left"
    Public Const PRIMARY1RIGHT_HEADER As String = "Primary1_Right"
    Public Const PRIMARY1TOP_HEADER As String = "Primary1_Top"
    Public Const PRIMARY1BOTTOM_HEADER As String = "Primary1_Bottom"
    Public Const PRIMARY1REAR_HEADER As String = "Primary1_Back"

    Public Const PRIMARY2FRONT_HEADER As String = "Primary2_Front"
    Public Const PRIMARY2LEFT_HEADER As String = "Primary2_Left"
    Public Const PRIMARY2RIGHT_HEADER As String = "Primary2_Right"
    Public Const PRIMARY2TOP_HEADER As String = "Primary2_Top"
    Public Const PRIMARY2BOTTOM_HEADER As String = "Primary2_Bottom"
    Public Const PRIMARY2REAR_HEADER As String = "Primary2_Back"

    'CODE ADDED - 6/13/16 - Ignore Wire mesh in sweep data
    Public Const WIRE_MESH_SHAPE As String = "WIRE MESH"

    Public Const FLAT As String = "FLAT"
    Public Const ROUND_SHAPE As String = "ROUND"
    Public Const ROUND_TUBING_SHAPE As String = "RND TUBG"
    Public Const DOWEL_HOLE As String = "DOWEL_HOLE"
    Public Const MINIMUM_NUM_OF_FLOOR_MOUNT_BODY As Integer = 4
    'Public Const B_FLOOR_MOUNT_VIEW As String = "B_FLOOR_MOUNT"
    Public Const TEMP_SCALE As Double = 0.5

    '10 degree angular tolerance
    Public Const ANGULAR_TOLERANCE_FOR_CONCENTRIC_CHECK As Double = 0.015192247
    Public Const ONE_DEG As Double = 0.000152


    Public Const BODY_LCS_SHEET_NAME As String = "Body_LCS"
    Public Const PART_LCS_VIEW_NAME As String = "B_LCS"

    'Constant values used for checking concentricity of given faces.
    Public Const CONCENTRIC_DISTANCE_TOLERANCE_BET_FACE_CENTERS As Double = 5
    Public Const BENT_BRACKET_CONCENTRIC_FACES_DISTANCE_TOLERANCE As Double = 12.5

    Public Const AUTO2D_RUN_DATE_TIME_ATTR As String = "B_AUTO2D_RUN_DATE"
    'code added Jun-29-2017
    Public Const PART_DETAIL_ATTRIBUTE As String = "B_DETAIL"
    Public Const PART_DETAIL_YES As String = "Y"
    Public Const PART_DETAIL_NO As String = "N"

    'Public Const CLIENT_STOCK_SIZE_ATTR As String = "Fertigmaße/Bestellbezeichnung_W060"
    'Public Const CLIENT_STOCK_SIZE_ATTR As String = "Rohmass"

    'Public Const LIST_OF_PARTS_IN_WELDMENT_FILE_NAME As String = "List_Of_Parts_in_Weldment.txt"

    Public Const PART_SWEEP_DATA_FILES_PROCESSED_TEXT_FILE_NAME As String = "Part_Sweep_Status_File.txt"

    Public Const TOOL_NAME As String = "TOOL_NAME"

    Public Const CAR_DIVISION As String = "CAR"
    Public Const TRUCK_DIVISION As String = "TRUCK"
    Public Const DIVISION_NAME As String = "DIVISION"

    Public Const CAR_DIVISION_FEATURE_GROUP As String = "ROUGH_PART"
    Public Const TRUCK_DIVISION_FEATURE_GROUP As String = "PART_DESIGN"
    Public Const DEGREE As Double = 0.017453292519943282

    'OEM NAME
    Public Const GM_OEM_NAME As String = "GM"
    Public Const CHRYSLER_OEM_NAME As String = "CHRYSLER"
    Public Const DAIMLER_OEM_NAME As String = "DAIMLER"
    Public Const FIAT_OEM_NAME As String = "FIAT"
    Public Const GESTAMP_OEM_NAME As String = "GESTAMP"

    'SUPPLIER NAME
    Public Const COMAU_NAME As String = "COMAU"
    Public Const VALIANT_NAME As String = "VALIANT"

    Public Const LOG_FILE As String = "Log.txt"
    Public Const PROCESS_ID_FILE As String = "Process_ID.txt"

    Public Const PLATE As String = "PLATE"
    Public Const SQUARE As String = "SQUARE"
    Public Const SQUARE_TUBG As String = "SQUARE TUBG"
    Public Const RECT_TUBG As String = "RECT TUBG"

    'Public Const FLOOR_MOUNT_ATTRIBUTE As String = "B_FLOOR_MOUNTING_FACE"
    Public Const FLOOR_MOUNT_ATTRIBUTE As String = "B_FLOOR_MOUNT_FACE"
    Public Const FLOOR_MOUNT_DIR As String = "B_FLOOR_MOUNT_DIR"

    Public Const B_PART_TYPE As String = "B_PART_TYPE"
    Public Const WELDED_ASS As String = "WELDED_ASS"
    Public Const WELDED_CHILD As String = "WELDED_CHILD"
    Public Const NON_WELDED_ASS As String = "SINGLE_DETAIL"

    Public Const BURNOUT_FEATURE_GROUP As String = "CORPO"
    Public Const BURNOUT_STRING As String = "CM_Flame_Cut"
    Public Const BURNOUT_YES As String = "Y"

    'Code added Mar-292-2019
    Public Const B_ADA_TYPE As String = "B_ADA_TYPE"
    Public Const AUXILIARY_FRAME As String = "AUXILIARY FRAME"

    Public Const STD_PROJ_FRONTVIEW As String = "STD_PROJ_FRONTVIEW"
    Public Const STD_PROJ_LEFTVIEW As String = "STD_PROJ_LEFTVIEW"
    Public Const STD_PROJ_TOPVIEW As String = "STD_PROJ_TOPVIEW"
    Public Const STD_PROJ_BOTTOMVIEW As String = "STD_PROJ1_BOTTOMVIEW"
    Public Const STD_PROJ_RIGHTVIEW As String = "STD_PROJ_RIGHTVIEW"
    Public Const STD_PROJ_REARVIEW As String = "STD_PROJ_REARVIEW"


    Public Const ONE_DEG_TOLERANCE As Double = 0.000152

    'Code added Jun-28-2019
    'Common names that are read from the xml file
    Public Const XML_PARTNAME As String = "PartName"
    Public Const XML_SHAPE As String = "Shape"
    Public Const XML_SUBDETAILNUMBER As String = "SubDetailNumber"
    Public Const XML_MASS As String = "Mass"
    Public Const XML_PURCHASEOPTION As String = "PurchaseOption"
    Public Const XML_STOCKSIZE As String = "StockSize"
    Public Const XML_STOCKSIZEMETRIC As String = "StockSizeMetric"
    Public Const XML_FINISHTOLERANCE As String = "FinishTolerance"
    Public Const XML_QUANTITY As String = "Quantity"
    Public Const XML_MATERIAL As String = "Material"
    Public Const XML_FINISHTOLVALUE1 As String = "FinishTolValue1"
    Public Const XML_FINISHTOLVALUE2 As String = "FinishTolValue2"
    Public Const XML_FINISHTOLVALUE3 As String = "FinishTolValue3"
    Public Const XML_BOM As String = "Bom"
    Public Const XML_CLIENTPARTNAME As String = "ClientPartName"
    Public Const XML_DETAILSHEET1 As String = "DetailSheet1"
    Public Const XML_LAYOUTSHEET1 As String = "LayoutSheet1"
    Public Const XML_LAYOUTSHEET2 As String = "LayoutSheet2"
    Public Const XML_LAYOUTSHEET3 As String = "LayoutSheet3"
    Public Const XML_SHEET As String = "Sheet"
    Public Const XML_MMTOINCHCONVERSATION As String = "MMToInchConversation"
    Public Const XML_CLIENTSTOCKSIZE As String = "ClientStockSize"
    Public Const XML_TOOLCLASS As String = "ToolClass"
    Public Const XML_TOOLID As String = "ToolID"
    Public Const XML_ALTPURCH As String = "AltPurch"

    Public Const ErrorFileName As String = "ErrorNamesToProcessPSDAgain.txt"
End Module
