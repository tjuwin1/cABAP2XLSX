CLASS zcl_excel_data_validation DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
*"* public components of class ZCL_EXCEL_DATA_VALIDATION
*"* do not include other source files here!!!

    DATA errorstyle type zif_excel_data_decl=>zexcel_data_val_error_style .
    DATA operator type zif_excel_data_decl=>zexcel_data_val_operator .
    DATA allowblank TYPE abap_boolean VALUE 'X'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA cell_column type zif_excel_data_decl=>zexcel_cell_column_alpha .
    DATA cell_column_to type zif_excel_data_decl=>zexcel_cell_column_alpha .
    DATA cell_row type zif_excel_data_decl=>zexcel_cell_row .
    DATA cell_row_to type zif_excel_data_decl=>zexcel_cell_row .
    CONSTANTS c_type_custom type zif_excel_data_decl=>zexcel_data_val_type VALUE 'custom'. "#EC NOTEXT
    CONSTANTS c_type_list type zif_excel_data_decl=>zexcel_data_val_type VALUE 'list'. "#EC NOTEXT
    DATA showerrormessage TYPE abap_bool VALUE 'X'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA showinputmessage TYPE abap_bool VALUE 'X'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA type type zif_excel_data_decl=>zexcel_data_val_type .
    DATA formula1 type zif_excel_data_decl=>zexcel_validation_formula1 .
    DATA formula2 type zif_excel_data_decl=>zexcel_validation_formula1 .
    CONSTANTS c_type_none type zif_excel_data_decl=>zexcel_data_val_type VALUE 'none'. "#EC NOTEXT
    CONSTANTS c_type_date type zif_excel_data_decl=>zexcel_data_val_type VALUE 'date'. "#EC NOTEXT
    CONSTANTS c_type_decimal type zif_excel_data_decl=>zexcel_data_val_type VALUE 'decimal'. "#EC NOTEXT
    CONSTANTS c_type_textlength type zif_excel_data_decl=>zexcel_data_val_type VALUE 'textLength'. "#EC NOTEXT
    CONSTANTS c_type_time type zif_excel_data_decl=>zexcel_data_val_type VALUE 'time'. "#EC NOTEXT
    CONSTANTS c_type_whole type zif_excel_data_decl=>zexcel_data_val_type VALUE 'whole'. "#EC NOTEXT
    CONSTANTS c_style_stop type zif_excel_data_decl=>zexcel_data_val_error_style VALUE 'stop'. "#EC NOTEXT
    CONSTANTS c_style_warning type zif_excel_data_decl=>zexcel_data_val_error_style VALUE 'warning'. "#EC NOTEXT
    CONSTANTS c_style_information type zif_excel_data_decl=>zexcel_data_val_error_style VALUE 'information'. "#EC NOTEXT
    CONSTANTS c_operator_between type zif_excel_data_decl=>zexcel_data_val_operator VALUE 'between'. "#EC NOTEXT
    CONSTANTS c_operator_equal type zif_excel_data_decl=>zexcel_data_val_operator VALUE 'equal'. "#EC NOTEXT
    CONSTANTS c_operator_greaterthan type zif_excel_data_decl=>zexcel_data_val_operator VALUE 'greaterThan'. "#EC NOTEXT
    CONSTANTS c_operator_greaterthanorequal type zif_excel_data_decl=>zexcel_data_val_operator VALUE 'greaterThanOrEqual'. "#EC NOTEXT
    CONSTANTS c_operator_lessthan type zif_excel_data_decl=>zexcel_data_val_operator VALUE 'lessThan'. "#EC NOTEXT
    CONSTANTS c_operator_lessthanorequal type zif_excel_data_decl=>zexcel_data_val_operator VALUE 'lessThanOrEqual'. "#EC NOTEXT
    CONSTANTS c_operator_notbetween type zif_excel_data_decl=>zexcel_data_val_operator VALUE 'notBetween'. "#EC NOTEXT
    CONSTANTS c_operator_notequal type zif_excel_data_decl=>zexcel_data_val_operator VALUE 'notEqual'. "#EC NOTEXT
    DATA showdropdown TYPE abap_bool .
    DATA errortitle TYPE string .
    DATA error TYPE string .
    DATA prompttitle TYPE string .
    DATA prompt TYPE string .

    METHODS constructor .
*"* protected components of class ZCL_EXCEL_DATA_VALIDATION
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_DATA_VALIDATION
*"* do not include other source files here!!!
  PROTECTED SECTION.
  PRIVATE SECTION.
*"* private components of class ZCL_EXCEL_DATA_VALIDATION
*"* do not include other source files here!!!
ENDCLASS.



CLASS zcl_excel_data_validation IMPLEMENTATION.


  METHOD constructor.
    " Initialise instance variables
    formula1          = ''.
    formula2          = ''.
    type              = me->c_type_none.
    errorstyle        = me->c_style_stop.
    operator          = ''.
    allowblank        = abap_false.
    showdropdown      = abap_false.
    showinputmessage  = abap_true.
    showerrormessage  = abap_true.
    errortitle        = ''.
    error             = ''.
    prompttitle       = ''.
    prompt            = ''.
* inizialize dimension range
    cell_row     = 1.
    cell_column  = 'A'.
  ENDMETHOD.
ENDCLASS.
