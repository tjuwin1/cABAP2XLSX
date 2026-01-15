class zcl_excel_common definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_COMMON
*"* do not include other source files here!!!
  public section.
    types: begin of ty_uom,
             uom type i_unitofmeasure-UnitOfMeasure,
             dec type i_unitofmeasure-UnitOfMeasureNumberOfDecimals,
             uome type i_unitofmeasure-UnitOfMeasure_E,
           end of ty_uom,

           begin of ty_curr,
             curr type waers,
             dec  type i_currency-decimals,
           end of ty_curr.

    constants c_excel_baseline_date type d value '19000101'. "#EC NOTEXT
    class-data c_excel_numfmt_offset type int1 value 164. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    constants c_excel_sheet_max_col type int4 value 16384.  "#EC NOTEXT
    constants c_excel_sheet_min_col type int4 value 1.      "#EC NOTEXT
    constants c_excel_sheet_max_row type int4 value 1048576. "#EC NOTEXT
    constants c_excel_sheet_min_row type int4 value 1.      "#EC NOTEXT
    class-data c_spras_en type spras value 'E'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    "CLASS-DATA o_conv TYPE REF TO cl_abap_conv_out_ce .
    constants c_excel_1900_leap_year type d value '19000228'. "#EC NOTEXT
    class-data c_xlsx_file_filter type string value 'Excel Workbook (*.xlsx)|*.xlsx|'. "#EC NOTEXT .  .  .  .  .  .  . " .
    class-data lt_uoms type standard table of ty_uom.
    class-data lt_currs type standard table of ty_curr.
    class-methods class_constructor .
    class-methods describe_structure
      importing
        !io_struct type ref to cl_abap_structdescr.
*      RETURNING
*        VALUE(rt_dfies) TYPE ddfields .
    class-methods convert_column2alpha
      importing
        !ip_column       type simple
      returning
        value(ep_column) type zif_excel_data_decl=>zexcel_cell_column_alpha
      raising
        zcx_excel .
    class-methods convert_column2int
      importing
        !ip_column       type simple
      returning
        value(ep_column) type zif_excel_data_decl=>zexcel_cell_column
      raising
        zcx_excel .
    class-methods convert_column_a_row2columnrow
      importing
        !i_column          type simple
        !i_row             type zif_excel_data_decl=>zexcel_cell_row
      returning
        value(e_columnrow) type string
      raising
        zcx_excel.
    class-methods convert_columnrow2column_a_row
      importing
        !i_columnrow  type clike
      exporting
        !e_column     type zif_excel_data_decl=>zexcel_cell_column_alpha
        !e_column_int type zif_excel_data_decl=>zexcel_cell_column
        !e_row        type zif_excel_data_decl=>zexcel_cell_row
      raising
        zcx_excel.
    class-methods convert_range2column_a_row
      importing
        !i_range            type clike
        !i_allow_1dim_range type abap_bool default abap_false
      exporting
        !e_column_start     type zif_excel_data_decl=>zexcel_cell_column_alpha
        !e_column_start_int type zif_excel_data_decl=>zexcel_cell_column
        !e_column_end       type zif_excel_data_decl=>zexcel_cell_column_alpha
        !e_column_end_int   type zif_excel_data_decl=>zexcel_cell_column
        !e_row_start        type zif_excel_data_decl=>zexcel_cell_row
        !e_row_end          type zif_excel_data_decl=>zexcel_cell_row
        !e_sheet            type clike
      raising
        zcx_excel .
    class-methods convert_columnrow2column_o_row
      importing
        !i_columnrow type clike
      exporting
        !e_column    type zif_excel_data_decl=>zexcel_cell_column_alpha
        !e_row       type zif_excel_data_decl=>zexcel_cell_row .
    class-methods clone_ixml_with_namespaces
      importing
        element       type ref to if_ixml_element
      returning
        value(result) type ref to if_ixml_element.
    class-methods date_to_excel_string
      importing
        !ip_value       type d
      returning
        value(ep_value) type zif_excel_data_decl=>zexcel_cell_value .
    class-methods encrypt_password
      importing
        !i_pwd                 type zif_excel_data_decl=>zexcel_aes_password
      returning
        value(r_encrypted_pwd) type zif_excel_data_decl=>zexcel_aes_password .
    class-methods escape_string
      importing
        !ip_value               type clike
      returning
        value(ep_escaped_value) type string .
    class-methods unescape_string
      importing
        !iv_escaped                type clike
      returning
        value(ev_unescaped_string) type string
      raising
        zcx_excel .
    "! <p class="shorttext synchronized" lang="en">Convert date from Excel format to SAP</p>
    "! @parameter ip_value | String being an Excel number representing a date (e.g. 45141 means 2023/08/03,
    "!                       45141.58832 means 2023/08/03 14:07:11). Important: if the input is date +
    "!                       time, use the additional parameter IP_EXACT = 'X'.
    "! @parameter ip_exact | If the input value also contains the time i.e. a fractional part exists
    "!                       (e.g. 45141.58832 means 2023/08/03 14:07:11), ip_exact = 'X' will
    "!                       return the exact date (e.g. 2023/08/03), while ip_exact = ' ' (default) will
    "!                       return the rounded-up date (e.g. 2023/08/04). NB: this rounding-up doesn't
    "!                       happen if the time is before 12:00:00.
    "! @parameter ep_value | Date corresponding to the input Excel number. It returns a null date if
    "!                       the input value contains non-numeric characters.
    "! @raising zcx_excel | The numeric input corresponds to a date before 1900/1/1 or after 9999/12/31.
    class-methods excel_string_to_date
      importing
        !ip_value       type zif_excel_data_decl=>zexcel_cell_value
        !ip_exact       type abap_bool default abap_false
      returning
        value(ep_value) type d
      raising
        zcx_excel .
    class-methods excel_string_to_time
      importing
        !ip_value       type zif_excel_data_decl=>zexcel_cell_value
      returning
        value(ep_value) type t
      raising
        zcx_excel .
    class-methods excel_string_to_number
      importing
        !ip_value       type zif_excel_data_decl=>zexcel_cell_value
      returning
        value(ep_value) type f
      raising
        zcx_excel .
    class-methods get_fieldcatalog
      importing
        !ip_table              type standard table
      returning
        value(ep_fieldcatalog) type zif_excel_data_decl=>zexcel_t_fieldcatalog .
    class-methods number_to_excel_string
      importing
        value(ip_value)  type numeric
        ip_currency      type waers_curc optional
        ip_unitofmeasure type meins optional
      returning
        value(ep_value)  type zif_excel_data_decl=>zexcel_cell_value .
    class-methods recursive_class_to_struct
      importing
        !i_source  type any
      changing
        !e_target  type data
        !e_targetx type data .
    class-methods recursive_struct_to_class
      importing
        !i_source  type data
        !i_sourcex type data
      changing
        !e_target  type any .
    class-methods time_to_excel_string
      importing
        !ip_value       type t
      returning
        value(ep_value) type zif_excel_data_decl=>zexcel_cell_value .
    class-methods utclong_to_excel_string
      importing
        !ip_utclong     type any
      returning
        value(ep_value) type zif_excel_data_decl=>zexcel_cell_value .
    types: t_char10 type c length 10.
    types: t_char255 type c length 255.
    class-methods split_file
      importing
        !ip_file         type t_char255
      exporting
        !ep_file         type t_char255
        !ep_extension    type t_char10
        !ep_dotextension type t_char10 .
    class-methods calculate_cell_distance
      importing
        !iv_reference_cell type clike
        !iv_current_cell   type clike
      exporting
        !ev_row_difference type i
        !ev_col_difference type i
      raising
        zcx_excel .
    class-methods determine_resulting_formula
      importing
        !iv_reference_cell          type clike
        !iv_reference_formula       type clike
        !iv_current_cell            type clike
      returning
        value(ev_resulting_formula) type string
      raising
        zcx_excel .
    class-methods shift_formula
      importing
        !iv_reference_formula       type clike
        value(iv_shift_cols)        type i
        value(iv_shift_rows)        type i
      returning
        value(ev_resulting_formula) type string
      raising
        zcx_excel .
    class-methods is_cell_in_range
      importing
        !ip_column         type simple
        !ip_row            type zif_excel_data_decl=>zexcel_cell_row
        !ip_range          type clike
      returning
        value(rp_in_range) type abap_bool
      raising
        zcx_excel .
*"* protected components of class ZCL_EXCEL_COMMON
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_COMMON
*"* do not include other source files here!!!
  protected section.
  private section.

    class-data c_excel_col_module type int2 value 64. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    class-data sv_prev_in1  type zif_excel_data_decl=>zexcel_cell_column.
    class-data sv_prev_out1 type zif_excel_data_decl=>zexcel_cell_column_alpha.
    class-data sv_prev_in2  type c length 10.
    class-data sv_prev_out2 type zif_excel_data_decl=>zexcel_cell_column.
    class-methods structure_case
      importing
        !is_component  type abap_componentdescr
      changing
        !xt_components type abap_component_tab .
    class-methods structure_recursive
      importing
        !is_component        type abap_componentdescr
      returning
        value(rt_components) type abap_component_tab .
    types ty_char1 type c length 1.
    class-methods char2hex
      importing
        !i_char      type ty_char1
      returning
        value(r_hex) type zif_excel_data_decl=>zexcel_pwd_hash .
    class-methods shl01
      importing
        !i_pwd_hash       type zif_excel_data_decl=>zexcel_pwd_hash
      returning
        value(r_pwd_hash) type zif_excel_data_decl=>zexcel_pwd_hash .
    class-methods shr14
      importing
        !i_pwd_hash       type zif_excel_data_decl=>zexcel_pwd_hash
      returning
        value(r_pwd_hash) type zif_excel_data_decl=>zexcel_pwd_hash .
endclass.



class zcl_excel_common implementation.


  method calculate_cell_distance.

    data: lv_reference_row       type i,
          lv_reference_col_alpha type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_reference_col       type i,
          lv_current_row         type i,
          lv_current_col_alpha   type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_current_col         type i.

*--------------------------------------------------------------------*
* Split reference  cell into numerical row/column representation
*--------------------------------------------------------------------*
    convert_columnrow2column_a_row( exporting
                                      i_columnrow = iv_reference_cell
                                    importing
                                      e_column    = lv_reference_col_alpha
                                      e_row       = lv_reference_row ).
    lv_reference_col = convert_column2int( lv_reference_col_alpha ).

*--------------------------------------------------------------------*
* Split current  cell into numerical row/column representation
*--------------------------------------------------------------------*
    convert_columnrow2column_a_row( exporting
                                      i_columnrow = iv_current_cell
                                    importing
                                      e_column    = lv_current_col_alpha
                                      e_row       = lv_current_row ).
    lv_current_col = convert_column2int( lv_current_col_alpha ).

*--------------------------------------------------------------------*
* Calculate row and column difference
* Positive:   Current cell below    reference cell
*         or  Current cell right of reference cell
* Negative:   Current cell above    reference cell
*         or  Current cell left  of reference cell
*--------------------------------------------------------------------*
    ev_row_difference = lv_current_row - lv_reference_row.
    ev_col_difference = lv_current_col - lv_reference_col.

  endmethod.


  method char2hex.
    "@TODO: Juwin to Fix
*    IF o_conv IS NOT BOUND.
*      o_conv = cl_abap_conv_out_ce=>create( endian   = 'L'
*                                            ignore_cerr = abap_true
*                                            replacement = '#' ).
*    ENDIF.
*
*    CALL METHOD o_conv->reset( ).
*    CALL METHOD o_conv->write( data = i_char ).
*    r_hex+1 = o_conv->get_buffer( ). " x'65' must be x'0065'

  endmethod.


  method class_constructor.
    c_xlsx_file_filter = 'Excel Workbook (*.xlsx)|*.xlsx|'(005).
  endmethod.


  method convert_column2alpha.

    data: lv_uccpi  type i,
          lv_text   type c length 2,
          lv_module type int4,
          lv_column type zif_excel_data_decl=>zexcel_cell_column.

* Propagate zcx_excel if error occurs           " issue #155 - less restrictive typing for ip_column
    lv_column = convert_column2int( ip_column ).  " issue #155 - less restrictive typing for ip_column

*--------------------------------------------------------------------*
* Check whether column is in allowed range for EXCEL to handle ( 1-16384 )
*--------------------------------------------------------------------*
    if   lv_column > 16384
      or lv_column < 1.
      zcx_excel=>raise_text( 'Index out of bounds' ).
    endif.

*--------------------------------------------------------------------*
* Look up for previous succesfull cached result
*--------------------------------------------------------------------*
    if lv_column = sv_prev_in1 and sv_prev_out1 is not initial.
      ep_column = sv_prev_out1.
      return.
    else.
      clear sv_prev_out1.
      sv_prev_in1 = lv_column.
    endif.

*--------------------------------------------------------------------*
* Build alpha representation of column
*--------------------------------------------------------------------*
    ep_column = xco_cp_xlsx=>coordinate->for_numeric_value( lv_column )->get_alphabetic_value( ).
*    WHILE lv_column GT 0.
*      lv_module = ( lv_column - 1 ) MOD 26.
*      lv_uccpi  = 65 + lv_module.
*
*      lv_column = ( lv_column - lv_module ) / 26.
*
*      lv_text   = cl_abap_conv_in_ce=>uccpi( lv_uccpi ).
*      CONCATENATE lv_text ep_column INTO ep_column.
*
*    ENDWHILE.

*--------------------------------------------------------------------*
* Save succesfull output into cache
*--------------------------------------------------------------------*
    sv_prev_out1 = ep_column.

  endmethod.


  method convert_column2int.
    data: lv_column       type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_column_c     type c length 10,
          lv_errormessage type string.                          " Can't pass '...'(abc) to exception-class

    lv_column_c = ip_column.
    translate lv_column_c to upper case.                      " Fix #246
    condense lv_column_c no-gaps.
    if lv_column_c eq ''.
      message e800(zabap2xlsx) into lv_errormessage.
      zcx_excel=>raise_symsg( ).
    endif.

    try.
        if lv_column_c co '1234567890 '.                      " Fix #164
          ep_column = lv_column_c.                            " Fix #164
*--------------------------------------------------------------------*
* Maximum column for EXCEL:  XFD = 16384    " if anyone has a reference for this information - please add here instead of this comment
*--------------------------------------------------------------------*
          if ep_column > 16384 or ep_column < 1.
            lv_errormessage = 'Index out of bounds'(004).
            zcx_excel=>raise_text( lv_errormessage ).
          endif.
          return.
        else.
          ep_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( lv_column_c )->get_numeric_value( ).
        endif.
      catch cx_sy_conversion_no_number.                 "#EC NO_HANDLER
        " Try the character-approach if approach via number has failed
    endtry.
  endmethod.


  method convert_column_a_row2columnrow.
    data: lv_row_alpha    type string,
          lv_column_alpha type zif_excel_data_decl=>zexcel_cell_column_alpha.

    lv_row_alpha = i_row.
    lv_column_alpha = zcl_excel_common=>convert_column2alpha( i_column ).
    shift lv_row_alpha right deleting trailing space.
    shift lv_row_alpha left deleting leading space.
    concatenate lv_column_alpha lv_row_alpha into e_columnrow.

  endmethod.


  method convert_columnrow2column_a_row.
*--------------------------------------------------------------------*
    "issue #256 - replacing char processing with regex
*--------------------------------------------------------------------*
* Stefan Schmoecker, 2013-08-11
*    Allow input to be CLIKE instead of STRING
*--------------------------------------------------------------------*

    data: pane_cell_row_a type string,
          lv_columnrow    type string.

    lv_columnrow = i_columnrow.    " Get rid of trailing blanks

    find regex '^(\D+)(\d+)$' in lv_columnrow submatches e_column
                                                         pane_cell_row_a.
    if e_column_int is supplied.
      e_column_int = convert_column2int( ip_column = e_column ).
    endif.
    e_row = pane_cell_row_a.

  endmethod.


  method convert_range2column_a_row.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-12-07
*              - ...
* changes: renaming variables to naming conventions
*          aligning code
*          added exceptionclass
*          added errorhandling for invalid range
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
* issue#241 - error when sheetname contains "!"
*           - sheetname should be returned unescaped
*              - Stefan Schmoecker,                          2012-12-07
* changes: changed coding to support sheetnames with "!"
*          unescaping sheetname
*--------------------------------------------------------------------*
* issue#155 - lessening restrictions of input parameters
*              - Stefan Schmoecker,                          2012-12-07
* changes: i_range changed to clike
*          e_sheet changed to clike
*--------------------------------------------------------------------*

    data: lv_sheet           type string,
          lv_range           type string,
          lv_columnrow_start type string,
          lv_columnrow_end   type string,
          lv_position        type i,
          lv_errormessage    type string.                          " Can't pass '...'(abc) to exception-class


*--------------------------------------------------------------------*
* Split input range into sheetname and Area
* 4 cases - a) input empty --> nothing to do
*         - b) sheetname existing - starts with '            example 'Sheet 1'!$B$6:$D$13
*         - c) sheetname existing - does not start with '    example Sheet1!$B$6:$D$13
*         - d) no sheetname - just area                      example $B$6:$D$13
*--------------------------------------------------------------------*
* Initialize output parameters
    clear: e_column_start,
           e_column_end,
           e_row_start,
           e_row_end,
           e_sheet.

    if i_range is initial.                                " a) input empty --> nothing to do
      return.

    elseif i_range(1) = `'`.                              " b) sheetname existing - starts with '
      find regex '\![^\!]*$' in i_range match offset lv_position.  " Find last !
      if sy-subrc = 0.
        lv_sheet = i_range(lv_position).
        add 1 to lv_position.
        lv_range = i_range.
        shift lv_range left by lv_position places.
      else.
        lv_errormessage = 'Invalid range'(001).
        zcx_excel=>raise_text( lv_errormessage ).
      endif.

    elseif i_range cs '!'.                                " c) sheetname existing - does not start with '
      split i_range at '!' into lv_sheet lv_range.
      " begin Dennis Schaaf
      if lv_range cp '*#REF*'.
        lv_errormessage = 'Invalid range'(001).
        zcx_excel=>raise_text( lv_errormessage ).
      endif.
      " end Dennis Schaaf
    else.                                                 " d) no sheetname - just area
      lv_range = i_range.
    endif.

    replace all occurrences of '$' in lv_range with ''.
    split lv_range at ':' into lv_columnrow_start lv_columnrow_end.

    if i_allow_1dim_range = abap_true.
      convert_columnrow2column_o_row( exporting i_columnrow = lv_columnrow_start
                                      importing e_column    = e_column_start
                                                e_row       = e_row_start ).
      convert_columnrow2column_o_row( exporting i_columnrow = lv_columnrow_end
                                      importing e_column    = e_column_end
                                                e_row       = e_row_end ).
    else.
      convert_columnrow2column_a_row( exporting i_columnrow = lv_columnrow_start
                                      importing e_column    = e_column_start
                                                e_row       = e_row_start ).
      convert_columnrow2column_a_row( exporting i_columnrow = lv_columnrow_end
                                      importing e_column    = e_column_end
                                                e_row       = e_row_end ).
    endif.

    if e_column_start_int is supplied and e_column_start is not initial.
      e_column_start_int = convert_column2int( e_column_start ).
    endif.
    if e_column_end_int is supplied and e_column_end is not initial.
      e_column_end_int = convert_column2int( e_column_end ).
    endif.

    e_sheet = unescape_string( lv_sheet ).                  " Return in unescaped form
  endmethod.


  method convert_columnrow2column_o_row.

    data: row       type string.
    data: columnrow type string.

    clear e_column.

    columnrow = i_columnrow.

    find regex '^(\D*)(\d*)$' in columnrow submatches e_column
                                                      row.

    e_row = row.

  endmethod.


  method clone_ixml_with_namespaces.

    types: begin of ty_name_value,
             name  type string,
             value type string,
           end of ty_name_value.

    data: iterator    type ref to if_ixml_node_iterator,
          node        type ref to if_ixml_node,
          xmlns       type ty_name_value,
          xmlns_table type table of ty_name_value.
    field-symbols <xmlns> type ty_name_value.

    iterator = element->create_iterator( ).
    result ?= element->clone( ).
    node = iterator->get_next( ).
    while node is bound.
      xmlns-name = node->get_namespace_prefix( ).
      xmlns-value = node->get_namespace_uri( ).
      collect xmlns into xmlns_table.
      node = iterator->get_next( ).
    endwhile.

    loop at xmlns_table assigning <xmlns>.
      result->set_attribute_ns( prefix = 'xmlns' name = <xmlns>-name value = <xmlns>-value ).
    endloop.

  endmethod.


  method date_to_excel_string.
    data: lv_date_diff         type i.

    check ip_value is not initial
      and ip_value <> space.
    " Needed hack caused by the problem that:
    " Excel 2000 incorrectly assumes that the year 1900 is a leap year
    " http://support.microsoft.com/kb/214326/en-us
    if ip_value > c_excel_1900_leap_year.
      lv_date_diff = ip_value - c_excel_baseline_date + 2.
    else.
      lv_date_diff = ip_value - c_excel_baseline_date + 1.
    endif.
    ep_value = zcl_excel_common=>number_to_excel_string( ip_value = lv_date_diff ).
  endmethod.


  method describe_structure.
    "@TODO: Juwin to throw error
*    DATA: lt_components TYPE abap_component_tab,
*          lt_comps      TYPE abap_component_tab,
*          ls_component  TYPE abap_componentdescr,
*          lo_elemdescr  TYPE REF TO cl_abap_elemdescr,
*          ls_dfies      TYPE dfies,
*          l_position    LIKE ls_dfies-position.
*
*    "for DDIC structure get the info directly
*    IF io_struct->is_ddic_type( ) = abap_true.
*      rt_dfies = io_struct->get_ddic_field_list( ).
*    ELSE.
*      lt_components = io_struct->get_components( ).
*
*      LOOP AT lt_components INTO ls_component.
*        structure_case( EXPORTING is_component  = ls_component
*                        CHANGING  xt_components = lt_comps   ) .
*      ENDLOOP.
*      LOOP AT lt_comps INTO ls_component.
*        CLEAR ls_dfies.
*        IF ls_component-type->kind = cl_abap_typedescr=>kind_elem. "E Elementary Type
*          ADD 1 TO l_position.
*          lo_elemdescr ?= ls_component-type.
*          IF lo_elemdescr->is_ddic_type( ) = abap_true.
*            ls_dfies           = lo_elemdescr->get_ddic_field( ).
*            ls_dfies-fieldname = ls_component-name.
*            ls_dfies-position  = l_position.
*          ELSE.
*            ls_dfies-fieldname = ls_component-name.
*            ls_dfies-position  = l_position.
*            ls_dfies-inttype   = lo_elemdescr->type_kind.
*            ls_dfies-leng      = lo_elemdescr->length.
*            ls_dfies-outputlen = lo_elemdescr->length.
*            ls_dfies-decimals  = lo_elemdescr->decimals.
*            ls_dfies-fieldtext = ls_component-name.
*            ls_dfies-reptext   = ls_component-name.
*            ls_dfies-scrtext_s = ls_component-name.
*            ls_dfies-scrtext_m = ls_component-name.
*            ls_dfies-scrtext_l = ls_component-name.
*            ls_dfies-dynpfld   = abap_true.
*          ENDIF.
*          INSERT ls_dfies INTO TABLE rt_dfies.
*        ENDIF.
*      ENDLOOP.
*    ENDIF.
  endmethod.


  method determine_resulting_formula.

    data: lv_row_difference type i,
          lv_col_difference type i.

*--------------------------------------------------------------------*
* Calculate distance of reference and current cell
*--------------------------------------------------------------------*
    calculate_cell_distance( exporting
                               iv_reference_cell = iv_reference_cell
                               iv_current_cell   = iv_current_cell
                             importing
                               ev_row_difference = lv_row_difference
                               ev_col_difference = lv_col_difference ).

*--------------------------------------------------------------------*
* and shift formula by using the row- and columndistance
*--------------------------------------------------------------------*
    ev_resulting_formula = shift_formula( iv_reference_formula = iv_reference_formula
                                          iv_shift_rows        = lv_row_difference
                                          iv_shift_cols        = lv_col_difference ).

  endmethod.                    "determine_resulting_formula


  method encrypt_password.

    data lv_curr_offset            type i.
    data lv_curr_char              type c length 1.
    data lv_curr_hex               type zif_excel_data_decl=>zexcel_pwd_hash.
    data lv_pwd_len                type zif_excel_data_decl=>zexcel_pwd_hash.
    data lv_pwd_hash               type zif_excel_data_decl=>zexcel_pwd_hash.

    constants:
      lv_0x7fff type zif_excel_data_decl=>zexcel_pwd_hash value '7FFF',
      lv_0x0001 type zif_excel_data_decl=>zexcel_pwd_hash value '0001',
      lv_0xce4b type zif_excel_data_decl=>zexcel_pwd_hash value 'CE4B'.

    data lv_pwd            type zif_excel_data_decl=>zexcel_aes_password.

    lv_pwd = i_pwd.

    lv_pwd_len = strlen( lv_pwd ).
    lv_curr_offset = lv_pwd_len - 1.

    while lv_curr_offset ge 0.

      lv_curr_char = lv_pwd+lv_curr_offset(1).
      lv_curr_hex = char2hex( lv_curr_char ).

      lv_pwd_hash = (  shr14( lv_pwd_hash ) bit-and lv_0x0001 ) bit-or ( shl01( lv_pwd_hash ) bit-and lv_0x7fff ).

      lv_pwd_hash = lv_pwd_hash bit-xor lv_curr_hex.
      subtract 1 from lv_curr_offset.
    endwhile.

    lv_pwd_hash = (  shr14( lv_pwd_hash ) bit-and lv_0x0001 ) bit-or ( shl01( lv_pwd_hash ) bit-and lv_0x7fff ).
    lv_pwd_hash = lv_pwd_hash bit-xor lv_0xce4b.
    lv_pwd_hash = lv_pwd_hash bit-xor lv_pwd_len.

    r_encrypted_pwd = lv_pwd_hash.

  endmethod.


  method escape_string.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-12-08
*              - ...
* changes: aligning code
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
* issue#242 - Support escaping for white-spaces
*           - Escaping also necessary when ' encountered in input
*              - Stefan Schmoecker,                          2012-12-08
* changes: switched check if escaping is necessary to regular expression
*          and moved the "REPLACE"
*--------------------------------------------------------------------*
* issue#155 - lessening restrictions of input parameters
*              - Stefan Schmoecker,                          2012-12-08
* changes: ip_value changed to clike
*--------------------------------------------------------------------*
    data:       lv_value                        type string.

*--------------------------------------------------------------------*
* There exist various situations when a space will be used to separate
* different parts of a string. When we have a string consisting spaces
* that will cause errors unless we "escape" the string by putting ' at
* the beginning and at the end of the string.
*--------------------------------------------------------------------*


*--------------------------------------------------------------------*
* When allowing clike-input parameters we might encounter trailing
* "real" blanks .  These are automatically eliminated when moving
* the input parameter to a string.
* Now any remaining spaces ( white-spaces or normal spaces ) should
* trigger the escaping as well as any '
*--------------------------------------------------------------------*
    lv_value = ip_value.


    find regex `\s|'|-` in lv_value.  " \s finds regular and white spaces
    if sy-subrc = 0.
      replace all occurrences of `'` in lv_value with `''`.
      concatenate `'` lv_value `'` into lv_value .
    endif.

    ep_escaped_value = lv_value.

  endmethod.


  method excel_string_to_date.
    data: lv_date_int type i.
    data lv_error_text type string.

    check ip_value is not initial and ip_value cn ' 0'.

    try.
        if ip_exact = abap_false.
          lv_date_int = ip_value.
        else.
          lv_date_int = trunc( ip_value ).
        endif.
        if lv_date_int not between 1 and 2958465.
          zcx_excel=>raise_text( 'Unable to interpret date' ).
        endif.
        ep_value = lv_date_int + c_excel_baseline_date - 2.
        " Needed hack caused by the problem that:
        " Excel 2000 incorrectly assumes that the year 1900 is a leap year
        " http://support.microsoft.com/kb/214326/en-us
        if ep_value < c_excel_1900_leap_year.
          ep_value = ep_value + 1.
        endif.
      catch cx_sy_conversion_error.
        lv_error_text = |String "{ ip_value }" is not a valid Excel date|.
        zcx_excel=>raise_text( lv_error_text ).
    endtry.
  endmethod.


  method excel_string_to_number.

* If we encounter anything more complicated in EXCEL we might have to extend this
* But currently this works fine - even for numbers in scientific notation

    ep_value = ip_value.

  endmethod.


  method excel_string_to_time.
    data: lv_seconds_in_day type i,
          lv_day_fraction   type f,
          lc_seconds_in_day type i value 86400.

    try.

        lv_day_fraction = frac( ip_value ).
        lv_seconds_in_day = lv_day_fraction * lc_seconds_in_day.

        ep_value = lv_seconds_in_day.

      catch cx_sy_conversion_error.
        zcx_excel=>raise_text( 'Unable to interpret time' ).
    endtry.
  endmethod.

  method get_fieldcatalog.
    data lo_tabledescr type ref to cl_abap_tabledescr.
    data lo_strucdescr type ref to cl_abap_structdescr.

    lo_tabledescr ?= cl_abap_tabledescr=>describe_by_data( ip_table ).
    lo_strucdescr ?= lo_tabledescr->get_table_line_type( ).
    loop at lo_strucdescr->components into data(ls_comp).
      append value #( fieldname = ls_comp-name
                      scrtext_l = ls_comp-name
                      position  = lines( ep_fieldcatalog ) + 1
                      abap_type = ls_comp-type_kind )
             to ep_fieldcatalog.
    endloop.
  endmethod.


  method is_cell_in_range.
    data lv_column_start    type zif_excel_data_decl=>zexcel_cell_column_alpha.
    data lv_column_end      type zif_excel_data_decl=>zexcel_cell_column_alpha.
    data lv_row_start       type zif_excel_data_decl=>zexcel_cell_row.
    data lv_row_end         type zif_excel_data_decl=>zexcel_cell_row.
    data lv_column_start_i  type zif_excel_data_decl=>zexcel_cell_column.
    data lv_column_end_i    type zif_excel_data_decl=>zexcel_cell_column.
    data lv_column_i        type zif_excel_data_decl=>zexcel_cell_column.


* Split range and convert columns
    convert_range2column_a_row(
      exporting
        i_range        = ip_range
      importing
        e_column_start = lv_column_start
        e_column_end   = lv_column_end
        e_row_start    = lv_row_start
        e_row_end      = lv_row_end ).

    lv_column_start_i = convert_column2int( ip_column = lv_column_start ).
    lv_column_end_i   = convert_column2int( ip_column = lv_column_end ).

    lv_column_i = convert_column2int( ip_column = ip_column ).

* Check if cell is in range
    if lv_column_i >= lv_column_start_i and
       lv_column_i <= lv_column_end_i   and
       ip_row      >= lv_row_start      and
       ip_row      <= lv_row_end.
      rp_in_range = abap_true.
    endif.
  endmethod.


  method number_to_excel_string.
    data: lv_value_c type c length 100.

    if ip_currency is not initial.
      if lt_currs is initial.
        select currency, decimals from i_currency order by currency into table @lt_currs.
      endif.
      lv_value_c = |{ ip_value currency = ip_currency }|.
    elseif ip_unitofmeasure is not initial.
      if lt_uoms is initial.
        select unitofmeasure, unitofmeasurenumberofdecimals, UnitOfMeasure_E from i_unitofmeasure order by unitofmeasure into table @lt_uoms.
      endif.
      read table lt_uoms into data(ls_uom) with key uom = ip_unitofmeasure binary search.
      lv_value_c = |{ ip_value decimals = ls_uom-dec }|.
    else.
      lv_value_c = ip_value.
    endif.
    replace all occurrences of ',' in lv_value_c with '.'.

    ep_value = lv_value_c.
    condense ep_value.

    if ip_value eq 0.
      ep_value = '0'.
    endif.
  endmethod.


  method recursive_class_to_struct.
    " # issue 139
* is working for me - but after looking through this coding I guess
* I'll rewrite this to a version w/o recursion
* This is private an no one using it so far except me, so no need to hurry
    data: descr          type ref to cl_abap_structdescr,
          wa_component   like line of descr->components,
          attribute_name like wa_component-name,
          flag_class     type abap_bool.

    field-symbols: <field>     type any,
                   <fieldx>    type any,
                   <attribute> type any.


    descr ?= cl_abap_structdescr=>describe_by_data( e_target ).

    loop at descr->components into wa_component.

* Assign structure and X-structure
      assign component wa_component-name of structure e_target  to <field>.
      assign component wa_component-name of structure e_targetx to <fieldx>.
* At least one field in the structure should be marked - otherwise continue with next field
      clear flag_class.
* maybe source is just a structure - try assign component...
      assign component wa_component-name of structure i_source  to <attribute>.
      if sy-subrc <> 0.
* not - then it is an attribute of the class - use different assign then
        concatenate 'i_source->' wa_component-name into attribute_name.
        assign (attribute_name) to <attribute>.
        if sy-subrc <> 0.
          exit.
        endif.  " Should not happen if structure is built properly - otherwise just exit to create no dumps
        flag_class = abap_true.
      endif.

      case wa_component-type_kind.
        when cl_abap_structdescr=>typekind_struct1 or cl_abap_structdescr=>typekind_struct2.  " Structure --> use recursio
          zcl_excel_common=>recursive_class_to_struct( exporting i_source  = <attribute>
                                                       changing  e_target  = <field>
                                                                 e_targetx = <fieldx> ).
        when others.
          <field> = <attribute>.
          <fieldx> = abap_true.

      endcase.
    endloop.

  endmethod.


  method recursive_struct_to_class.
    " # issue 139
* is working for me - but after looking through this coding I guess
* I'll rewrite this to a version w/o recursion
* This is private an no one using it so far except me, so no need to hurry
    data: descr          type ref to cl_abap_structdescr,
          wa_component   like line of descr->components,
          attribute_name like wa_component-name,
          flag_class     type abap_bool,
          o_border       type ref to zcl_excel_style_border.

    field-symbols: <field>     type any,
                   <fieldx>    type any,
                   <attribute> type any.


    descr ?= cl_abap_structdescr=>describe_by_data( i_source ).

    loop at descr->components into wa_component.

* Assign structure and X-structure
      assign component wa_component-name of structure i_source  to <field>.
      assign component wa_component-name of structure i_sourcex to <fieldx>.
* At least one field in the structure should be marked - otherwise continue with next field
      check <fieldx> ca abap_true.
      clear flag_class.
* maybe target is just a structure - try assign component...
      assign component wa_component-name of structure e_target  to <attribute>.
      if sy-subrc <> 0.
* not - then it is an attribute of the class - use different assign then
        concatenate 'E_TARGET->' wa_component-name into attribute_name.
        assign (attribute_name) to <attribute>.
        if sy-subrc <> 0.exit.endif.  " Should not happen if structure is built properly - otherwise just exit to create no dumps
        flag_class = abap_true.
      endif.

      case wa_component-type_kind.
        when cl_abap_structdescr=>typekind_struct1 or cl_abap_structdescr=>typekind_struct2.  " Structure --> use recursion
          " To avoid dump with attribute GRADTYPE of class ZCL_EXCEL_STYLE_FILL
          " quick and really dirty fix -> check the attribute name
          " Border has to be initialized somewhere else
          if wa_component-name eq 'GRADTYPE'.
            flag_class = abap_false.
          endif.

          if flag_class = abap_true and <attribute> is initial.
* Only borders will be passed as unbound references.  But since we want to set a value we have to create an instance
            create object o_border.
            <attribute> = o_border.
          endif.
          zcl_excel_common=>recursive_struct_to_class( exporting i_source  = <field>
                                                                 i_sourcex = <fieldx>
                                                       changing  e_target  = <attribute> ).
        when others.
          check <fieldx> = abap_true.  " Marked for change
          <attribute> = <field>.

      endcase.
    endloop.

  endmethod.


  method shift_formula.

    constants: lcv_operators            type string value '+-/*^%=<>&, !',
               lcv_letters              type string value 'ABCDEFGHIJKLMNOPQRSTUVWXYZ$',
               lcv_digits               type string value '0123456789',
               lcv_cell_reference_error type string value '#REF!'.

    data: lv_tcnt          type i,         " Counter variable
          lv_tlen          type i,         " Temp variable length
          lv_cnt           type i,         " Counter variable
          lv_cnt2          type i,         " Counter variable
          lv_offset1       type i,         " Character offset
          lv_numchars      type i,         " Number of characters counter
          lv_tchar(1)      type c,         " Temp character
          lv_tchar2(1)     type c,         " Temp character
          lv_cur_form      type string,    " Formula for current cell
          lv_ref_cell_addr type string,    " Reference cell address
          lv_tcol1         type string,    " Temp column letter
          lv_tcol2         type string,    " Temp column letter
          lv_tcoln         type i,         " Temp column number
          lv_trow1         type string,    " Temp row number
          lv_trow2         type string,    " Temp row number
          lv_flen          type i,         " Length of reference formula
          lv_tlen2         type i,         " Temp variable length
          lv_substr1       type string,    " Substring variable
          lv_abscol        type string,    " Absolute column symbol
          lv_absrow        type string,    " Absolute row symbol
          lv_ref_formula   type string,
          lv_compare_1     type string,
          lv_compare_2     type string,
          lv_level         type i,         " Level of groups [..[..]..] or {..}

          lv_errormessage  type string.

*--------------------------------------------------------------------*
* When copying a cell in EXCEL to another cell any inherent formulas
* are copied as well.  Cell-references in the formula are being adjusted
* by the distance of the new cell to the original one
*--------------------------------------------------------------------*
* §1 Parse reference formula character by character
* §2 Identify Cell-references
* §3 Shift cell-reference
* §4 Build resulting formula
*--------------------------------------------------------------------*

    lv_ref_formula = iv_reference_formula.
*--------------------------------------------------------------------*
* No distance --> Reference = resulting cell/formula
*--------------------------------------------------------------------*
    if    iv_shift_cols = 0
      and iv_shift_rows = 0.
      ev_resulting_formula = lv_ref_formula.
      return. " done
    endif.


    lv_flen     = strlen( lv_ref_formula ).
    lv_numchars = 1.

*--------------------------------------------------------------------*
* §1 Parse reference formula character by character
*--------------------------------------------------------------------*
    do lv_flen times.

      clear: lv_tchar,
             lv_substr1,
             lv_ref_cell_addr.
      lv_cnt2 = lv_cnt + 1.
      if lv_cnt2 > lv_flen.
        exit. " Done
      endif.

*--------------------------------------------------------------------*
* Here we have the current character in the formula
*--------------------------------------------------------------------*
      lv_tchar = lv_ref_formula+lv_cnt(1).

*--------------------------------------------------------------------*
* Operators or opening parenthesis will separate possible cellreferences
*--------------------------------------------------------------------*
      if    (    lv_tchar ca lcv_operators
              or lv_tchar ca '(' )
        and lv_cnt2 = 1.
        lv_substr1  = lv_ref_formula+lv_offset1(1).
        concatenate lv_cur_form lv_substr1 into lv_cur_form.
        lv_cnt      = lv_cnt + 1.
        lv_offset1  = lv_cnt.
        lv_numchars = 1.
        continue.       " --> next character in formula can be analyzed
      endif.

*--------------------------------------------------------------------*
* Quoted literal text holds no cell reference --> advance to end of text
*--------------------------------------------------------------------*
      if lv_tchar eq '"'.
        lv_cnt      = lv_cnt + 1.
        lv_numchars = lv_numchars + 1.
        lv_tchar     = lv_ref_formula+lv_cnt(1).
        while lv_tchar ne '"'.

          lv_cnt      = lv_cnt + 1.
          lv_numchars = lv_numchars + 1.
          lv_tchar    = lv_ref_formula+lv_cnt(1).

        endwhile.
        lv_cnt2    = lv_cnt + 1.
        lv_substr1 = lv_ref_formula+lv_offset1(lv_numchars).
        concatenate lv_cur_form lv_substr1 into lv_cur_form.
        lv_cnt     = lv_cnt + 1.
        if lv_cnt = lv_flen.
          exit.
        endif.
        lv_offset1  = lv_cnt.
        lv_numchars = 1.
        lv_tchar    = lv_ref_formula+lv_cnt(1).
        lv_cnt2     = lv_cnt + 1.
        continue.       " --> next character in formula can be analyzed
      endif.


*--------------------------------------------------------------------*
* Groups - Ignore values inside blocks [..[..]..] and {..}
*     R1C1-Style Cell Reference: R[1]C[1]
*     Cell References: 'C:\[Source.xlsx]Sheet1'!$A$1
*     Array constants: {1,3.5,TRUE,"Hello"}
*     "Intra table reference": Flights[[#This Row],[Air fare]]
*--------------------------------------------------------------------*
      if lv_tchar ca '[]{}' or lv_level > 0.
        if lv_tchar ca '[{'.
          lv_level = lv_level + 1.
        elseif lv_tchar ca ']}'.
          lv_level = lv_level - 1.
        endif.
        if lv_cnt2 = lv_flen.
          lv_substr1 = iv_reference_formula+lv_offset1(lv_numchars).
          concatenate lv_cur_form lv_substr1 into lv_cur_form.
          exit.
        endif.
        lv_numchars = lv_numchars + 1.
        lv_cnt   = lv_cnt   + 1.
        lv_cnt2  = lv_cnt   + 1.
        continue.
      endif.

*--------------------------------------------------------------------*
* Operators or parenthesis or last character in formula will separate possible cellreferences
*--------------------------------------------------------------------*
      if   lv_tchar ca lcv_operators
        or lv_tchar ca '():'
        or lv_cnt2  =  lv_flen.
        if lv_cnt > 0.
          lv_substr1 = lv_ref_formula+lv_offset1(lv_numchars).
*--------------------------------------------------------------------*
* Check for text concatenation and functions
*--------------------------------------------------------------------*
          if ( lv_tchar ca lcv_operators and lv_tchar eq lv_substr1 ) or lv_tchar eq '('.
            concatenate lv_cur_form lv_substr1 into lv_cur_form.
            lv_cnt = lv_cnt + 1.
            lv_offset1 = lv_cnt.
            lv_cnt2 = lv_cnt + 1.
            lv_numchars = 1.
            continue.       " --> next character in formula can be analyzed
          endif.

          lv_tlen = lv_cnt2 - lv_offset1.
*--------------------------------------------------------------------*
* Exclude mathematical operators and closing parentheses
*--------------------------------------------------------------------*
          if   lv_tchar ca lcv_operators
            or lv_tchar ca ':)'.
            if    lv_cnt2     = lv_flen
              and lv_numchars = 1.
              concatenate lv_cur_form lv_substr1 into lv_cur_form.
              lv_cnt      = lv_cnt + 1.
              lv_offset1  = lv_cnt.
              lv_cnt2     = lv_cnt + 1.
              lv_numchars = 1.
              continue.       " --> next character in formula can be analyzed
            else.
              lv_tlen = lv_tlen - 1.
            endif.
          endif.
*--------------------------------------------------------------------*
* Capture reference cell address
*--------------------------------------------------------------------*
          try.
              lv_ref_cell_addr = lv_ref_formula+lv_offset1(lv_tlen). "Ref cell address
            catch cx_root.
              lv_errormessage = 'Internal error in Class ZCL_EXCEL_COMMON Method SHIFT_FORMULA Spot 1 '.  " Change to messageclass if possible
              zcx_excel=>raise_text( lv_errormessage ).
          endtry.

*--------------------------------------------------------------------*
* Split cell address into characters and numbers
*--------------------------------------------------------------------*
          clear: lv_tlen,
                 lv_tcnt,
                 lv_tcol1,
                 lv_trow1.
          lv_tlen = strlen( lv_ref_cell_addr ).
          if lv_tlen <> 0.
            clear: lv_tcnt.
            do lv_tlen times.
              clear: lv_tchar2.
              lv_tchar2 = lv_ref_cell_addr+lv_tcnt(1).
              if lv_tchar2 ca lcv_letters.
                concatenate lv_tcol1 lv_tchar2 into lv_tcol1.
              elseif lv_tchar2 ca lcv_digits.
                concatenate lv_trow1 lv_tchar2 into lv_trow1.
              endif.
              lv_tcnt = lv_tcnt + 1.
            enddo.
          endif.

          " Is valid column & row ?
          if lv_tcol1 is not initial and lv_trow1 is not initial.
            " COLUMN + ROW
            concatenate lv_tcol1 lv_trow1 into lv_compare_1.
            " Original condensed string
            lv_compare_2 = lv_ref_cell_addr.
            condense lv_compare_2.
            if lv_compare_1 <> lv_compare_2.
              clear: lv_trow1, lv_tchar2.
            endif.
          endif.

*--------------------------------------------------------------------*
* Check for invalid cell address
*--------------------------------------------------------------------*
          if lv_tcol1 is initial or lv_trow1 is initial.
            concatenate lv_cur_form lv_substr1 into lv_cur_form.
            lv_cnt = lv_cnt + 1.
            lv_offset1 = lv_cnt.
            lv_cnt2 = lv_cnt + 1.
            lv_numchars = 1.
            continue.
          endif.
*--------------------------------------------------------------------*
* Check for range names
*--------------------------------------------------------------------*
          clear: lv_tlen.
          lv_tlen = strlen( lv_tcol1 ).
          if lv_tlen gt 3.
            concatenate lv_cur_form lv_substr1 into lv_cur_form.
            lv_cnt = lv_cnt + 1.
            lv_offset1 = lv_cnt.
            lv_cnt2 = lv_cnt + 1.
            lv_numchars = 1.
            continue.
          endif.
*--------------------------------------------------------------------*
* Check for valid row
*--------------------------------------------------------------------*
          if lv_trow1 gt 1048576.
            concatenate lv_cur_form lv_substr1 into lv_cur_form.
            lv_cnt = lv_cnt + 1.
            lv_offset1 = lv_cnt.
            lv_cnt2 = lv_cnt + 1.
            lv_numchars = 1.
            continue.
          endif.
*--------------------------------------------------------------------*
* Check for absolute column or row reference
*--------------------------------------------------------------------*
          clear: lv_tcol2,
                 lv_trow2,
                 lv_abscol,
                 lv_absrow.
          lv_tlen2 = strlen( lv_tcol1 ) - 1.
          if lv_tcol1 is not initial.
            lv_abscol = lv_tcol1(1).
          endif.
          if lv_tlen2 ge 0.
            lv_absrow = lv_tcol1+lv_tlen2(1).
          endif.
          if lv_abscol eq '$' and lv_absrow eq '$'.
            lv_tlen2 = lv_tlen2 - 1.
            if lv_tlen2 > 0.
              lv_tcol1 = lv_tcol1+1(lv_tlen2).
            endif.
            lv_tlen2 = lv_tlen2 + 1.
          elseif lv_abscol eq '$'.
            lv_tcol1 = lv_tcol1+1(lv_tlen2).
          elseif lv_absrow eq '$'.
            lv_tcol1 = lv_tcol1(lv_tlen2).
          endif.
*--------------------------------------------------------------------*
* Check for valid column
*--------------------------------------------------------------------*
          try.
              lv_tcoln = zcl_excel_common=>convert_column2int( lv_tcol1 ) + iv_shift_cols.
            catch zcx_excel.
              concatenate lv_cur_form lv_substr1 into lv_cur_form.
              lv_cnt = lv_cnt + 1.
              lv_offset1 = lv_cnt.
              lv_cnt2 = lv_cnt + 1.
              lv_numchars = 1.
              continue.
          endtry.
*--------------------------------------------------------------------*
* Check whether there is a referencing problem
*--------------------------------------------------------------------*
          lv_trow2 = lv_trow1 + iv_shift_rows.
          " Remove the space used for the sign
          condense lv_trow2.
          if   ( lv_tcoln < 1 and lv_abscol <> '$' )   " Maybe we should add here max-column and max row-tests as well.
            or ( lv_trow2 < 1 and lv_absrow <> '$' ).  " Check how EXCEL behaves in this case
*--------------------------------------------------------------------*
* Referencing problem encountered --> set error
*--------------------------------------------------------------------*
            concatenate lv_cur_form lcv_cell_reference_error into lv_cur_form.
          else.
*--------------------------------------------------------------------*
* No referencing problems --> adjust row and column
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* Adjust column
*--------------------------------------------------------------------*
            if lv_abscol eq '$'.
              concatenate lv_cur_form lv_abscol lv_tcol1 into lv_cur_form.
            elseif iv_shift_cols eq 0.
              concatenate lv_cur_form lv_tcol1 into lv_cur_form.
            else.
              try.
                  lv_tcol2 = zcl_excel_common=>convert_column2alpha( lv_tcoln ).
                  concatenate lv_cur_form lv_tcol2 into lv_cur_form.
                catch zcx_excel.
                  concatenate lv_cur_form lv_substr1 into lv_cur_form.
                  lv_cnt = lv_cnt + 1.
                  lv_offset1 = lv_cnt.
                  lv_cnt2 = lv_cnt + 1.
                  lv_numchars = 1.
                  continue.
              endtry.
            endif.
*--------------------------------------------------------------------*
* Adjust row
*--------------------------------------------------------------------*
            if lv_absrow eq '$'.
              concatenate lv_cur_form lv_absrow lv_trow1 into lv_cur_form.
            elseif iv_shift_rows = 0.
              concatenate lv_cur_form lv_trow1 into lv_cur_form.
            else.
              concatenate lv_cur_form lv_trow2 into lv_cur_form.
            endif.
          endif.

          lv_numchars = 0.
          if   lv_tchar ca lcv_operators
            or lv_tchar ca ':)'.
            concatenate lv_cur_form lv_tchar into lv_cur_form respecting blanks.
          endif.
          lv_offset1 = lv_cnt2.
        endif.
      endif.
      lv_numchars = lv_numchars + 1.
      lv_cnt   = lv_cnt   + 1.
      lv_cnt2  = lv_cnt   + 1.

    enddo.



*--------------------------------------------------------------------*
* Return resulting formula
*--------------------------------------------------------------------*
    if lv_cur_form is not initial.
      ev_resulting_formula = lv_cur_form.
    endif.

  endmethod.


  method shl01.

    data:
      lv_bit      type i,
      lv_curr_pos type i value 2,
      lv_prev_pos type i value 1.

    do 15 times.
      get bit lv_curr_pos of i_pwd_hash into lv_bit.
      set bit lv_prev_pos of r_pwd_hash to lv_bit.
      add 1 to lv_curr_pos.
      add 1 to lv_prev_pos.
    enddo.
    set bit 16 of r_pwd_hash to 0.

  endmethod.


  method shr14.

    data:
      lv_bit      type i,
      lv_curr_pos type i,
      lv_next_pos type i.

    r_pwd_hash = i_pwd_hash.

    do 14 times.
      lv_curr_pos = 15.
      lv_next_pos = 16.

      do 15 times.
        get bit lv_curr_pos of r_pwd_hash into lv_bit.
        set bit lv_next_pos of r_pwd_hash to lv_bit.
        subtract 1 from lv_curr_pos.
        subtract 1 from lv_next_pos.
      enddo.
      set bit 1 of r_pwd_hash to 0.
    enddo.

  endmethod.


  method split_file.

    data: lt_hlp type table of text255,
          ls_hlp type text255.

    data: lf_ext(10)     type c,
          lf_dot_ext(10) type c.
    data: lf_anz type i,
          lf_len type i.
** ---------------------------------------------------------------------

    clear: lt_hlp,
           ep_file,
           ep_extension,
           ep_dotextension.

** Split the whole file at '.'
    split ip_file at '.' into table lt_hlp.

** get the extenstion from the last line of table
    lf_anz = lines( lt_hlp ).
    if lf_anz <= 1.
      ep_file = ip_file.
      return.
    endif.

    read table lt_hlp into ls_hlp index lf_anz.
    ep_extension = ls_hlp.
    lf_ext =  ls_hlp.
    if not lf_ext is initial.
      concatenate '.' lf_ext into lf_dot_ext.
    endif.
    ep_dotextension = lf_dot_ext.

** get only the filename
    lf_len = strlen( ip_file ) - strlen( lf_dot_ext ).
    if lf_len > 0.
      ep_file = ip_file(lf_len).
    endif.

  endmethod.


  method structure_case.
    data: lt_comp_str        type abap_component_tab.

    case is_component-type->kind.
      when cl_abap_typedescr=>kind_elem. "E Elementary Type
        insert is_component into table xt_components.
      when cl_abap_typedescr=>kind_table. "T Table
        insert is_component into table xt_components.
      when cl_abap_typedescr=>kind_struct. "S Structure
        lt_comp_str = structure_recursive( is_component = is_component ).
        insert lines of lt_comp_str into table xt_components.
      when others. "cl_abap_typedescr=>kind_ref or  cl_abap_typedescr=>kind_class or  cl_abap_typedescr=>kind_intf.
* We skip it. for now.
    endcase.
  endmethod.


  method structure_recursive.
    data: lo_struct     type ref to cl_abap_structdescr,
          lt_components type abap_component_tab,
          ls_components type abap_componentdescr.

    lo_struct ?= is_component-type.
    lt_components = lo_struct->get_components( ).

    loop at lt_components into ls_components.
      structure_case( exporting is_component  = ls_components
                      changing  xt_components = rt_components ) .
    endloop.

  endmethod.


  method time_to_excel_string.
    data: lv_seconds_in_day type i,
          lv_day_fraction   type f,
          lc_time_baseline  type t value '000000',
          lc_seconds_in_day type i value 86400.

    lv_seconds_in_day = ip_value - lc_time_baseline.
    lv_day_fraction = lv_seconds_in_day / lc_seconds_in_day.
    ep_value = zcl_excel_common=>number_to_excel_string( ip_value = lv_day_fraction ).
  endmethod.


  method unescape_string.

    constants   lcv_regex                       type string value `^'[^']`    & `|` &  " Beginning single ' OR
                                                                  `[^']'$`    & `|` &  " Trailing single '  OR
                                                                  `[^']'[^']`.         " Single ' somewhere in between


    data:       lv_errormessage                 type string.                          " Can't pass '...'(abc) to exception-class

*--------------------------------------------------------------------*
* This method is used to extract the "real" string from an escaped string.
* An escaped string can be identified by a beginning ' which must be
* accompanied by a trailing '
* All '' in between beginning and trailing ' are treated as single '
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* When allowing clike-input parameters we might encounter trailing
* "real" blanks .  These are automatically eliminated when moving
* the input parameter to a string.
*--------------------------------------------------------------------*
    ev_unescaped_string = iv_escaped.           " Pass through if not escaped

    check ev_unescaped_string is not initial.   " Nothing to do if empty
    check ev_unescaped_string(1) = `'`.         " Nothing to do if not escaped

*--------------------------------------------------------------------*
* Remove leading and trailing '
*--------------------------------------------------------------------*
    replace regex `^'(.*)'$` in ev_unescaped_string with '$1'.
    if sy-subrc <> 0.
      lv_errormessage = 'Input not properly escaped - &'(002).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

*--------------------------------------------------------------------*
* Any remaining single ' should not be here
*--------------------------------------------------------------------*
    find regex lcv_regex in ev_unescaped_string.
    if sy-subrc = 0.
      lv_errormessage = 'Input not properly escaped - &'(002).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

*--------------------------------------------------------------------*
* Replace '' with '
*--------------------------------------------------------------------*
    replace all occurrences of `''` in ev_unescaped_string with `'`.


  endmethod.

  method utclong_to_excel_string.
    data lv_timestamp type timestamp.
    data lv_date type d.
    data lv_time type t.

    " The data type UTCLONG and the method UTCLONG2TSTMP_SHORT are not available before ABAP 7.54
    "   -> Need of a dynamic call to avoid compilation error before ABAP 7.54

    call method cl_abap_tstmp=>('UTCLONG2TSTMP_SHORT')
      exporting
        utclong   = ip_utclong
      receiving
        timestamp = lv_timestamp.
    convert time stamp lv_timestamp time zone 'UTC   ' into date lv_date time lv_time.
    ep_value = |{ date_to_excel_string( lv_date ) + time_to_excel_string( lv_time ) }|.
  endmethod.

endclass.
