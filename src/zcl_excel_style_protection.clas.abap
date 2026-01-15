CLASS zcl_excel_style_protection DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_STYLE_PROTECTION
*"* do not include other source files here!!!
  PUBLIC SECTION.

    CONSTANTS c_protection_hidden type zif_excel_data_decl=>zexcel_cell_protection VALUE '1'. "#EC NOTEXT
    CONSTANTS c_protection_locked type zif_excel_data_decl=>zexcel_cell_protection VALUE '1'. "#EC NOTEXT
    CONSTANTS c_protection_unhidden type zif_excel_data_decl=>zexcel_cell_protection VALUE '0'. "#EC NOTEXT
    CONSTANTS c_protection_unlocked type zif_excel_data_decl=>zexcel_cell_protection VALUE '0'. "#EC NOTEXT
    DATA hidden type zif_excel_data_decl=>zexcel_cell_protection .
    DATA locked type zif_excel_data_decl=>zexcel_cell_protection .

    METHODS constructor .
    METHODS get_structure
      RETURNING
        VALUE(ep_protection) type zif_excel_data_decl=>zexcel_s_style_protection .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_STYLE_PROTECTION
*"* do not include other source files here!!!
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_excel_style_protection IMPLEMENTATION.


  METHOD constructor.
    locked = me->c_protection_locked.
    hidden = me->c_protection_unhidden.
  ENDMETHOD.


  METHOD get_structure.
    ep_protection-locked = me->locked.
    ep_protection-hidden = me->hidden.
  ENDMETHOD.
ENDCLASS.
