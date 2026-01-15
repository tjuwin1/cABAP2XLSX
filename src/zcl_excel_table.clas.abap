class zcl_excel_table definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_TABLE
*"* do not include other source files here!!!
  public section.

    constants builtinstyle_dark1 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleDark1'. "#EC NOTEXT
    constants builtinstyle_dark2 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleDark2'. "#EC NOTEXT
    constants builtinstyle_dark3 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleDark3'. "#EC NOTEXT
    constants builtinstyle_dark4 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleDark4'. "#EC NOTEXT
    constants builtinstyle_dark5 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleDark5'. "#EC NOTEXT
    constants builtinstyle_dark6 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleDark6'. "#EC NOTEXT
    constants builtinstyle_dark7 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleDark7'. "#EC NOTEXT
    constants builtinstyle_dark8 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleDark8'. "#EC NOTEXT
    constants builtinstyle_dark9 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleDark9'. "#EC NOTEXT
    constants builtinstyle_dark10 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleDark10'. "#EC NOTEXT
    constants builtinstyle_dark11 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleDark11'. "#EC NOTEXT
    constants builtinstyle_light1 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight1'. "#EC NOTEXT
    constants builtinstyle_light2 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight2'. "#EC NOTEXT
    constants builtinstyle_light3 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight3'. "#EC NOTEXT
    constants builtinstyle_light4 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight4'. "#EC NOTEXT
    constants builtinstyle_light5 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight5'. "#EC NOTEXT
    constants builtinstyle_pivot_light16 type zif_excel_data_decl=>zexcel_table_style value 'PivotStyleLight16'. "#EC NOTEXT
    constants builtinstyle_light6 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight6'. "#EC NOTEXT
    constants totals_function_average type zif_excel_data_decl=>zexcel_table_totals_function value 'average'. "#EC NOTEXT
    constants builtinstyle_light7 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight7'. "#EC NOTEXT
    constants totals_function_count type zif_excel_data_decl=>zexcel_table_totals_function value 'count'. "#EC NOTEXT
    constants builtinstyle_light8 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight8'. "#EC NOTEXT
    constants totals_function_custom type zif_excel_data_decl=>zexcel_table_totals_function value 'custom'. "#EC NOTEXT
    constants builtinstyle_light9 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight9'. "#EC NOTEXT
    constants totals_function_max type zif_excel_data_decl=>zexcel_table_totals_function value 'max'. "#EC NOTEXT
    constants builtinstyle_light10 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight10'. "#EC NOTEXT
    constants totals_function_min type zif_excel_data_decl=>zexcel_table_totals_function value 'min'. "#EC NOTEXT
    constants builtinstyle_light11 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight11'. "#EC NOTEXT
    constants totals_function_sum type zif_excel_data_decl=>zexcel_table_totals_function value 'sum'. "#EC NOTEXT
    constants builtinstyle_light12 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight12'. "#EC NOTEXT
    data fieldcat type zif_excel_data_decl=>zexcel_t_fieldcatalog .
    constants builtinstyle_light13 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight13'. "#EC NOTEXT
    constants builtinstyle_light14 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight14'. "#EC NOTEXT
    data settings type zif_excel_data_decl=>zexcel_s_table_settings .
    constants builtinstyle_light15 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight15'. "#EC NOTEXT
    constants builtinstyle_light16 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight16'. "#EC NOTEXT
    constants builtinstyle_light17 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight17'. "#EC NOTEXT
    constants builtinstyle_light18 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight18'. "#EC NOTEXT
    constants builtinstyle_light19 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight19'. "#EC NOTEXT
    constants builtinstyle_light20 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight20'. "#EC NOTEXT
    constants builtinstyle_light21 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleLight21'. "#EC NOTEXT
    constants builtinstyle_medium1 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium1'. "#EC NOTEXT
    constants builtinstyle_medium2 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium2'. "#EC NOTEXT
    constants builtinstyle_medium3 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium3'. "#EC NOTEXT
    constants builtinstyle_medium4 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium4'. "#EC NOTEXT
    constants builtinstyle_medium5 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium5'. "#EC NOTEXT
    constants builtinstyle_medium6 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium6'. "#EC NOTEXT
    constants builtinstyle_medium7 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium7'. "#EC NOTEXT
    constants builtinstyle_medium8 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium8'. "#EC NOTEXT
    constants builtinstyle_medium9 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium9'. "#EC NOTEXT
    constants builtinstyle_medium10 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium10'. "#EC NOTEXT
    constants builtinstyle_medium11 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium11'. "#EC NOTEXT
    constants builtinstyle_medium12 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium12'. "#EC NOTEXT
    constants builtinstyle_medium13 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium13'. "#EC NOTEXT
    constants builtinstyle_medium14 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium14'. "#EC NOTEXT
    constants builtinstyle_medium15 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium15'. "#EC NOTEXT
    constants builtinstyle_medium16 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium16'. "#EC NOTEXT
    constants builtinstyle_medium17 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium17'. "#EC NOTEXT
    constants builtinstyle_medium18 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium18'. "#EC NOTEXT
    constants builtinstyle_medium19 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium19'. "#EC NOTEXT
    constants builtinstyle_medium20 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium20'. "#EC NOTEXT
    constants builtinstyle_medium21 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium21'. "#EC NOTEXT
    constants builtinstyle_medium22 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium22'. "#EC NOTEXT
    constants builtinstyle_medium23 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium23'. "#EC NOTEXT
    constants builtinstyle_medium24 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium24'. "#EC NOTEXT
    constants builtinstyle_medium25 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium26'. "#EC NOTEXT
    constants builtinstyle_medium27 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium27'. "#EC NOTEXT

    methods get_totals_formula
      importing
        !ip_column        type clike
        !ip_function      type zif_excel_data_decl=>zexcel_table_totals_function
      returning
        value(ep_formula) type string
      raising
        zcx_excel .
    methods has_totals
      returning
        value(ep_result) type abap_bool .
    methods set_data
      importing
        !ir_data type standard table .
    methods get_id
      returning
        value(ov_id) type i .
    methods set_id
      importing
        !iv_id type i .
    methods get_name
      returning
        value(ov_name) type string .
    methods get_reference
      importing
        !ip_include_totals_row type abap_bool default abap_true
      returning
        value(ov_reference)    type string
      raising
        zcx_excel .
    methods get_bottom_row_integer
      returning
        value(ev_row) type i .
    methods get_right_column_integer
      returning
        value(ev_column) type i
      raising
        zcx_excel .
*"* protected components of class ZCL_EXCEL_TABLE
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_TABLE
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_TABLE
*"* do not include other source files here!!!
  protected section.
  private section.

    data id type i .
    data name type string .
    data table_data type ref to data .
    data builtinstyle_medium28 type zif_excel_data_decl=>zexcel_table_style value 'TableStyleMedium28'. "#EC NOTEXT .  .  . " .
endclass.



class zcl_excel_table implementation.


  method get_bottom_row_integer.
    data: lv_table_lines type i.
    field-symbols: <fs_table> type standard table.

    if settings-bottom_right_row is not initial.
*    ev_row =  zcl_excel_common=>convert_column2int( settings-bottom_right_row ). " del issue #246
      ev_row =  settings-bottom_right_row .                                         " ins issue #246
      return.
    endif.

    assign table_data->* to <fs_table>.
    lv_table_lines = lines( <fs_table> ).
    if lv_table_lines = 0.
      lv_table_lines = 1. "table needs at least 1 data row
    endif.

    ev_row = settings-top_left_row + lv_table_lines.

    if me->has_totals( ) = abap_true."  ????  AND ip_include_totals_row = abap_true.
      ev_row += 1.
    endif.
  endmethod.


  method get_id.
    ov_id = id.
  endmethod.


  method get_name.

    if me->name is initial.
      me->name = zcl_excel_common=>number_to_excel_string( ip_value = me->id ).
      concatenate 'table' me->name into me->name.
    endif.

    ov_name = me->name.
  endmethod.


  method get_reference.
    data: lv_left_column_int       type zif_excel_data_decl=>zexcel_cell_column,
          lv_right_column_int      type zif_excel_data_decl=>zexcel_cell_column,
          lv_table_lines           type i,
          lv_left_column           type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_right_column          type zif_excel_data_decl=>zexcel_cell_column_alpha,
          ls_field_catalog         type zif_excel_data_decl=>zexcel_s_fieldcatalog,
          lv_bottom_row            type zif_excel_data_decl=>zexcel_cell_row,
          lv_top_row_string(10)    type c,
          lv_bottom_row_string(10) type c.

    field-symbols: <fs_table> type standard table.

*column
    lv_left_column_int = zcl_excel_common=>convert_column2int( settings-top_left_column ).
    lv_right_column_int = lv_left_column_int - 1.
    lv_right_column_int += lines( fieldcat ).
    lv_left_column  = zcl_excel_common=>convert_column2alpha( lv_left_column_int ).
    lv_right_column = zcl_excel_common=>convert_column2alpha( lv_right_column_int ).

*row
    assign table_data->* to <fs_table>.
    lv_table_lines = lines( <fs_table> ).
    if lv_table_lines = 0.
      lv_table_lines = 1. "table needs at least 1 data row
    endif.
    lv_bottom_row = settings-top_left_row + lv_table_lines .

    if me->has_totals( ) = abap_true and ip_include_totals_row = abap_true.
      lv_bottom_row += 1.
    endif.

    lv_top_row_string = |{ settings-top_left_row }|.
    lv_bottom_row_string = |{ lv_bottom_row }|.

    concatenate lv_left_column lv_top_row_string
                ':'
                lv_right_column lv_bottom_row_string into ov_reference.

  endmethod.


  method get_right_column_integer.
    data: ls_field_catalog  type zif_excel_data_decl=>zexcel_s_fieldcatalog.

    if settings-bottom_right_column is not initial.
      ev_column =  zcl_excel_common=>convert_column2int( settings-bottom_right_column ).
      return.
    endif.

    ev_column =  zcl_excel_common=>convert_column2int( settings-top_left_column ) + lines(  fieldcat ).

  endmethod.


  method get_totals_formula.
    constants: lc_function_id_sum     type string value '109',
               lc_function_id_min     type string value '105',
               lc_function_id_max     type string value '104',
               lc_function_id_count   type string value '103',
               lc_function_id_average type string value '101'.

    data: lv_function_id type string.

    case ip_function.
      when zcl_excel_table=>totals_function_sum.
        lv_function_id = lc_function_id_sum.

      when zcl_excel_table=>totals_function_min.
        lv_function_id = lc_function_id_min.

      when zcl_excel_table=>totals_function_max.
        lv_function_id = lc_function_id_max.

      when zcl_excel_table=>totals_function_count.
        lv_function_id = lc_function_id_count.

      when zcl_excel_table=>totals_function_average.
        lv_function_id = lc_function_id_average.

      when zcl_excel_table=>totals_function_custom. " issue #292
        return.

      when others.
        zcx_excel=>raise_text( 'Invalid totals formula. See ZCL_ for possible values' ).
    endcase.

    concatenate 'SUBTOTAL(' lv_function_id ',[' ip_column '])' into ep_formula.
  endmethod.


  method has_totals.
    data: ls_field_catalog    type zif_excel_data_decl=>zexcel_s_fieldcatalog.

    ep_result = abap_false.

    loop at fieldcat into ls_field_catalog.
      if ls_field_catalog-totals_function is not initial.
        ep_result = abap_true.
        exit.
      endif.
    endloop.

  endmethod.


  method set_data.

    data lr_temp type ref to data.

    field-symbols: <lt_table_temp> type any table,
                   <lt_table>      type any table.

    lr_temp = ref #( ir_data ).
    assign lr_temp->* to <lt_table_temp>.
    create data table_data like <lt_table_temp>.
    assign me->table_data->* to <lt_table>.
    <lt_table> = <lt_table_temp>.

  endmethod.


  method set_id.
    id = iv_id.
  endmethod.
endclass.
