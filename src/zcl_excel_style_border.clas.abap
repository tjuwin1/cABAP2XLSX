CLASS zcl_excel_style_border DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

*"* public components of class ZCL_EXCEL_STYLE_BORDER
*"* do not include other source files here!!!
    DATA border_style type zif_excel_data_decl=>zexcel_border .
    DATA border_color type zif_excel_data_decl=>zexcel_s_style_color .
    CONSTANTS c_border_none type zif_excel_data_decl=>zexcel_border VALUE 'none'. "#EC NOTEXT
    CONSTANTS c_border_dashdot type zif_excel_data_decl=>zexcel_border VALUE 'dashDot'. "#EC NOTEXT
    CONSTANTS c_border_dashdotdot type zif_excel_data_decl=>zexcel_border VALUE 'dashDotDot'. "#EC NOTEXT
    CONSTANTS c_border_dashed type zif_excel_data_decl=>zexcel_border VALUE 'dashed'. "#EC NOTEXT
    CONSTANTS c_border_dotted type zif_excel_data_decl=>zexcel_border VALUE 'dotted'. "#EC NOTEXT
    CONSTANTS c_border_double type zif_excel_data_decl=>zexcel_border VALUE 'double'. "#EC NOTEXT
    CONSTANTS c_border_hair type zif_excel_data_decl=>zexcel_border VALUE 'hair'. "#EC NOTEXT
    CONSTANTS c_border_medium type zif_excel_data_decl=>zexcel_border VALUE 'medium'. "#EC NOTEXT
    CONSTANTS c_border_mediumdashdot type zif_excel_data_decl=>zexcel_border VALUE 'mediumDashDot'. "#EC NOTEXT
    CONSTANTS c_border_mediumdashdotdot type zif_excel_data_decl=>zexcel_border VALUE 'mediumDashDotDot'. "#EC NOTEXT
    CONSTANTS c_border_mediumdashed type zif_excel_data_decl=>zexcel_border VALUE 'mediumDashed'. "#EC NOTEXT
    CONSTANTS c_border_slantdashdot type zif_excel_data_decl=>zexcel_border VALUE 'slantDashDot'. "#EC NOTEXT
    CONSTANTS c_border_thick type zif_excel_data_decl=>zexcel_border VALUE 'thick'. "#EC NOTEXT
    CONSTANTS c_border_thin type zif_excel_data_decl=>zexcel_border VALUE 'thin'. "#EC NOTEXT

    METHODS constructor .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_STYLE_BORDER
*"* do not include other source files here!!!
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_excel_style_border IMPLEMENTATION.


  METHOD constructor.
    border_style = zcl_excel_style_border=>c_border_none.
    border_color-theme     = zcl_excel_style_color=>c_theme_not_set.
    border_color-indexed   = zcl_excel_style_color=>c_indexed_not_set.
  ENDMETHOD.
ENDCLASS.
