CLASS zcl_excel_style_fill DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

*"* public components of class ZCL_EXCEL_STYLE_FILL
*"* do not include other source files here!!!
    CONSTANTS c_fill_none type zif_excel_data_decl=>zexcel_fill_type VALUE 'none'. "#EC NOTEXT
    CONSTANTS c_fill_solid type zif_excel_data_decl=>zexcel_fill_type VALUE 'solid'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_linear type zif_excel_data_decl=>zexcel_fill_type VALUE 'linear'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_path type zif_excel_data_decl=>zexcel_fill_type VALUE 'path'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkdown type zif_excel_data_decl=>zexcel_fill_type VALUE 'darkDown'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkgray type zif_excel_data_decl=>zexcel_fill_type VALUE 'darkGray'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkgrid type zif_excel_data_decl=>zexcel_fill_type VALUE 'darkGrid'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkhorizontal type zif_excel_data_decl=>zexcel_fill_type VALUE 'darkHorizontal'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darktrellis type zif_excel_data_decl=>zexcel_fill_type VALUE 'darkTrellis'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkup type zif_excel_data_decl=>zexcel_fill_type VALUE 'darkUp'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkvertical type zif_excel_data_decl=>zexcel_fill_type VALUE 'darkVertical'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_gray0625 type zif_excel_data_decl=>zexcel_fill_type VALUE 'gray0625'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_gray125 type zif_excel_data_decl=>zexcel_fill_type VALUE 'gray125'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lightdown type zif_excel_data_decl=>zexcel_fill_type VALUE 'lightDown'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lightgray type zif_excel_data_decl=>zexcel_fill_type VALUE 'lightGray'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lightgrid type zif_excel_data_decl=>zexcel_fill_type VALUE 'lightGrid'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lighthorizontal type zif_excel_data_decl=>zexcel_fill_type VALUE 'lightHorizontal'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lighttrellis type zif_excel_data_decl=>zexcel_fill_type VALUE 'lightTrellis'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lightup type zif_excel_data_decl=>zexcel_fill_type VALUE 'lightUp'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lightvertical type zif_excel_data_decl=>zexcel_fill_type VALUE 'lightVertical'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_mediumgray type zif_excel_data_decl=>zexcel_fill_type VALUE 'mediumGray'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_horizontal90 type zif_excel_data_decl=>zexcel_fill_type VALUE 'horizontal90'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_horizontal270 type zif_excel_data_decl=>zexcel_fill_type VALUE 'horizontal270'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_horizontalb type zif_excel_data_decl=>zexcel_fill_type VALUE 'horizontalb'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_vertical type zif_excel_data_decl=>zexcel_fill_type VALUE 'vertical'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_fromcenter type zif_excel_data_decl=>zexcel_fill_type VALUE 'fromCenter'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_diagonal45 type zif_excel_data_decl=>zexcel_fill_type VALUE 'diagonal45'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_diagonal45b type zif_excel_data_decl=>zexcel_fill_type VALUE 'diagonal45b'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_diagonal135 type zif_excel_data_decl=>zexcel_fill_type VALUE 'diagonal135'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_diagonal135b type zif_excel_data_decl=>zexcel_fill_type VALUE 'diagonal135b'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_cornerlt type zif_excel_data_decl=>zexcel_fill_type VALUE 'cornerLT'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_cornerlb type zif_excel_data_decl=>zexcel_fill_type VALUE 'cornerLB'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_cornerrt type zif_excel_data_decl=>zexcel_fill_type VALUE 'cornerRT'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_cornerrb type zif_excel_data_decl=>zexcel_fill_type VALUE 'cornerRB'. "#EC NOTEXT
    DATA gradtype type zif_excel_data_decl=>zexcel_s_gradient_type .
    DATA filltype type zif_excel_data_decl=>zexcel_fill_type .
    DATA rotation type zif_excel_data_decl=>zexcel_rotation .
    DATA fgcolor type zif_excel_data_decl=>zexcel_s_style_color .
    DATA bgcolor type zif_excel_data_decl=>zexcel_s_style_color .

    METHODS constructor .
    METHODS get_structure
      RETURNING
        VALUE(es_fill) type zif_excel_data_decl=>zexcel_s_style_fill .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_STYLE_FILL
*"* do not include other source files here!!!
  PRIVATE SECTION.

    METHODS build_gradient .
    METHODS check_filltype_is_gradient
      RETURNING
        VALUE(rv_is_gradient) TYPE abap_bool .
ENDCLASS.



CLASS zcl_excel_style_fill IMPLEMENTATION.


  METHOD build_gradient.
    CHECK check_filltype_is_gradient( ) EQ abap_true.
    CLEAR gradtype.
    CASE filltype.
      WHEN c_fill_gradient_horizontal90.
        gradtype-degree = '90'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_horizontal270.
        gradtype-degree = '270'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_horizontalb.
        gradtype-degree = '90'.
        gradtype-position1 = '0'.
        gradtype-position2 = '0.5'.
        gradtype-position3 = '1'.
      WHEN c_fill_gradient_vertical.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_fromcenter.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-bottom = '0.5'.
        gradtype-top = '0.5'.
        gradtype-left = '0.5'.
        gradtype-right = '0.5'.
      WHEN c_fill_gradient_diagonal45.
        gradtype-degree = '45'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_diagonal45b.
        gradtype-degree = '45'.
        gradtype-position1 = '0'.
        gradtype-position2 = '0.5'.
        gradtype-position3 = '1'.
      WHEN c_fill_gradient_diagonal135.
        gradtype-degree = '135'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_diagonal135b.
        gradtype-degree = '135'.
        gradtype-position1 = '0'.
        gradtype-position2 = '0.5'.
        gradtype-position3 = '1'.
      WHEN c_fill_gradient_cornerlt.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_cornerlb.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-bottom = '1'.
        gradtype-top = '1'.
      WHEN c_fill_gradient_cornerrt.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-left = '1'.
        gradtype-right = '1'.
      WHEN c_fill_gradient_cornerrb.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-bottom = '1'.
        gradtype-top = '1'.
        gradtype-left = '1'.
        gradtype-right = '1'.
    ENDCASE.

  ENDMETHOD.                    "build_gradient


  METHOD check_filltype_is_gradient.
    CASE filltype.
      WHEN c_fill_gradient_horizontal90 OR
           c_fill_gradient_horizontal270 OR
           c_fill_gradient_horizontalb OR
           c_fill_gradient_vertical OR
           c_fill_gradient_fromcenter OR
           c_fill_gradient_diagonal45 OR
           c_fill_gradient_diagonal45b OR
           c_fill_gradient_diagonal135 OR
           c_fill_gradient_diagonal135b OR
           c_fill_gradient_cornerlt OR
           c_fill_gradient_cornerlb OR
           c_fill_gradient_cornerrt OR
           c_fill_gradient_cornerrb.
        rv_is_gradient = abap_true.
    ENDCASE.
  ENDMETHOD.                    "check_filltype_is_gradient


  METHOD constructor.
    filltype = zcl_excel_style_fill=>c_fill_none.
    fgcolor-theme     = zcl_excel_style_color=>c_theme_not_set.
    fgcolor-indexed   = zcl_excel_style_color=>c_indexed_not_set.
    bgcolor-theme     = zcl_excel_style_color=>c_theme_not_set.
    bgcolor-indexed   = zcl_excel_style_color=>c_indexed_sys_foreground.
    rotation = 0.

  ENDMETHOD.                    "CONSTRUCTOR


  METHOD get_structure.
    es_fill-rotation  = me->rotation.
    es_fill-filltype  = me->filltype.
    es_fill-fgcolor   = me->fgcolor.
    es_fill-bgcolor   = me->bgcolor.
    me->build_gradient( ).
    es_fill-gradtype = me->gradtype.
  ENDMETHOD.                    "GET_STRUCTURE
ENDCLASS.
