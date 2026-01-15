class zcl_excel_srv_demos definition
  public
  create public .

  public section.

    interfaces if_http_service_extension .
  protected section.
  private section.
    methods demo1 returning value(ro_excel) type ref to zcl_excel.
    methods demo2 returning value(ro_excel) type ref to zcl_excel.
    methods demo3 returning value(ro_excel) type ref to zcl_excel.
    methods demo4 returning value(ro_excel) type ref to zcl_excel.
    methods demo5 returning value(ro_excel) type ref to zcl_excel.
    methods demo6 returning value(ro_excel) type ref to zcl_excel.
    methods demo7 returning value(ro_excel) type ref to zcl_excel.
    methods demo8 returning value(ro_excel) type ref to zcl_excel.
    methods demo9 returning value(ro_excel) type ref to zcl_excel.
    methods demo10 returning value(ro_excel) type ref to zcl_excel.
    methods demo11 returning value(ro_excel) type ref to zcl_excel.
    methods demo12 returning value(ro_excel) type ref to zcl_excel.
    methods write importing value(io_excel) type ref to zcl_excel
                  returning value(rv_excel) type xstring.
endclass.



class zcl_excel_srv_demos implementation.
  method if_http_service_extension~handle_request.
    data lv_xlsx type ref to zcl_excel.
    data lv_html type string.

    data(lv_demo) = request->get_form_field( 'DEMO' ).
    case lv_demo.
      when space.
        lv_html = |<html><head><title>Excel Demos</title></head><body><ul>|.
        do 12 times.
          lv_html = lv_html &&
          |<li><a href="/sap/bc/http/sap/ZSRV_EXCEL_DEMOS?DEMO=DEMO{ sy-index }" target="_blank">Demo{ sy-index }</a></li>|.
        enddo.
        lv_html = |{ lv_html }</ul></body></html>|.
        response->set_text( lv_html ).
      when others.
        response->set_content_type( 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ).
        call method me->(lv_demo)
          receiving
            ro_excel = lv_xlsx.
        response->set_binary( write( lv_xlsx ) ).
    endcase.
  endmethod.
  method demo4.

    data: lo_worksheet  type ref to zcl_excel_worksheet,
          lo_style_cond type ref to zcl_excel_style_cond.

    data: ls_iconset    type zif_excel_data_decl=>zexcel_conditional_iconset.

    create object ro_excel.

    " Get active sheet
    lo_worksheet = ro_excel->get_active_worksheet( ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.


    ls_iconset-iconset                  = zcl_excel_style_cond=>c_iconset_3trafficlights2.
    ls_iconset-cfvo1_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset-cfvo1_value              = '0'.
    ls_iconset-cfvo2_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset-cfvo2_value              = '33'.
    ls_iconset-cfvo3_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset-cfvo3_value              = '66'.
    ls_iconset-showvalue                = zcl_excel_style_cond=>c_showvalue_true.

    lo_style_cond->mode_iconset  = ls_iconset.
    lo_style_cond->set_range( ip_start_column  = 'C'
                                     ip_start_row     = 4
                                     ip_stop_column   = 'C'
                                     ip_stop_row      = 8 ).


    lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = 100 ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'C' ip_value = 1000 ).
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'C' ip_value = 150 ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'C' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 8 ip_column = 'C' ip_value = 500 ).


    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset-iconset           = zcl_excel_style_cond=>c_iconset_3trafficlights2.
    ls_iconset-showvalue         = zcl_excel_style_cond=>c_showvalue_false.
    lo_style_cond->mode_iconset  = ls_iconset.
    lo_style_cond->set_range( ip_start_column  = 'E'
                              ip_start_row     = 4
                              ip_stop_column   = 'E'
                              ip_stop_row      = 8 ).


    lo_worksheet->set_cell( ip_row = 4 ip_column = 'E' ip_value = 100 ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'E' ip_value = 1000 ).
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'E' ip_value = 150 ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'E' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 8 ip_column = 'E' ip_value = 500 ).
  endmethod.
  method demo6.

    data: lo_worksheet type ref to zcl_excel_worksheet,
          lv_row       type i,
          lv_formula   type string.


    create object ro_excel.

    " Get active sheet
    lo_worksheet = ro_excel->get_active_worksheet( ).

*--------------------------------------------------------------------*
*  Get some testdata
*--------------------------------------------------------------------*
    lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = 100  ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'C' ip_value = 1000  ).
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'C' ip_value = 150 ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'C' ip_value = -10  ).
    lo_worksheet->set_cell( ip_row = 8 ip_column = 'C' ip_value = 500  ).


*--------------------------------------------------------------------*
*  Demonstrate using formulas
*--------------------------------------------------------------------*
    lo_worksheet->set_cell( ip_row = 9 ip_column = 'C' ip_formula = 'SUM(C4:C8)' ).


*--------------------------------------------------------------------*
* Demonstrate standard EXCEL-behaviour when copying a formula to another cell
* by calculating the resulting formula to put into another cell
*--------------------------------------------------------------------*
    do 10 times.

      lv_formula = zcl_excel_common=>shift_formula( iv_reference_formula = 'SUM(C4:C8)'
                                                  iv_shift_cols        = 0                " Offset in Columns - here we copy in same column --> 0
                                                  iv_shift_rows        = sy-index ).      " Offset in Row     - here we copy downward --> sy-index
      lv_row = 9 + sy-index.                                                                " Absolute row = sy-index rows below reference cell
      lo_worksheet->set_cell( ip_row = lv_row ip_column = 'C' ip_formula = lv_formula ).

    enddo.
  endmethod.
  method demo7.

    data: lo_worksheet  type ref to zcl_excel_worksheet,
          lo_style_cond type ref to zcl_excel_style_cond.

    data: ls_iconset3    type zif_excel_data_decl=>zexcel_conditional_iconset,
          ls_iconset4    type zif_excel_data_decl=>zexcel_conditional_iconset,
          ls_iconset5    type zif_excel_data_decl=>zexcel_conditional_iconset,
          ls_databar     type zif_excel_data_decl=>zexcel_conditional_databar,
          ls_colorscale2 type zif_excel_data_decl=>zexcel_conditional_colorscale,
          ls_colorscale3 type zif_excel_data_decl=>zexcel_conditional_colorscale.

    create object ro_excel.

    ls_iconset3-cfvo1_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset3-cfvo1_value              = '0'.
    ls_iconset3-cfvo2_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset3-cfvo2_value              = '33'.
    ls_iconset3-cfvo3_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset3-cfvo3_value              = '66'.
    ls_iconset3-showvalue                = zcl_excel_style_cond=>c_showvalue_true.

    ls_iconset4-cfvo1_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset4-cfvo1_value              = '0'.
    ls_iconset4-cfvo2_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset4-cfvo2_value              = '25'.
    ls_iconset4-cfvo3_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset4-cfvo3_value              = '50'.
    ls_iconset4-cfvo4_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset4-cfvo4_value              = '75'.
    ls_iconset4-showvalue                = zcl_excel_style_cond=>c_showvalue_true.

    ls_iconset5-cfvo1_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset5-cfvo1_value              = '0'.
    ls_iconset5-cfvo2_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset5-cfvo2_value              = '20'.
    ls_iconset5-cfvo3_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset5-cfvo3_value              = '40'.
    ls_iconset5-cfvo4_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset5-cfvo4_value              = '60'.
    ls_iconset5-cfvo5_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset5-cfvo5_value              = '80'.
    ls_iconset5-showvalue                = zcl_excel_style_cond=>c_showvalue_true.

    ls_databar-cfvo1_type               = zcl_excel_style_cond=>c_cfvo_type_min.
    ls_databar-cfvo1_value              = '0'.
    ls_databar-cfvo2_type               = zcl_excel_style_cond=>c_cfvo_type_max.
    ls_databar-cfvo2_value              = '0'.
    ls_databar-colorrgb                 = 'FF638EC6'.

    ls_colorscale2-cfvo1_type               = zcl_excel_style_cond=>c_cfvo_type_min.
    ls_colorscale2-cfvo1_value              = '0'.
    ls_colorscale2-cfvo2_type               = zcl_excel_style_cond=>c_cfvo_type_percentile.
    ls_colorscale2-cfvo2_value              = '50'.
    ls_colorscale2-colorrgb1                = 'FFF8696B'.
    ls_colorscale2-colorrgb2                = 'FF63BE7B'.

    ls_colorscale3-cfvo1_type               = zcl_excel_style_cond=>c_cfvo_type_min.
    ls_colorscale3-cfvo1_value              = '0'.
    ls_colorscale3-cfvo2_type               = zcl_excel_style_cond=>c_cfvo_type_percentile.
    ls_colorscale3-cfvo2_value              = '50'.
    ls_colorscale3-cfvo3_type               = zcl_excel_style_cond=>c_cfvo_type_max.
    ls_colorscale3-cfvo3_value              = '0'.
    ls_colorscale3-colorrgb1                = 'FFF8696B'.
    ls_colorscale3-colorrgb2                = 'FFFFEB84'.
    ls_colorscale3-colorrgb3                = 'FF63BE7B'.

    " Get active sheet
    lo_worksheet = ro_excel->get_active_worksheet( ).

* ICONSET

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.

    ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3arrows.

    lo_style_cond->mode_iconset  = ls_iconset3.
    lo_style_cond->set_range( ip_start_column  = 'B'
                                     ip_start_row     = 5
                                     ip_stop_column   = 'B'
                                     ip_stop_row      = 9 ).

    lo_worksheet->set_cell( ip_row = 4 ip_column = 'B' ip_value = 'C_ICONSET_3ARROWS' ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'B' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'B' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'B' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 8 ip_column = 'B' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 9 ip_column = 'B' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3arrowsgray.
    lo_style_cond->mode_iconset  = ls_iconset3.
    lo_style_cond->set_range( ip_start_column  = 'C'
                                     ip_start_row     = 5
                                     ip_stop_column   = 'C'
                                     ip_stop_row      = 9 ).

    lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = 'C_ICONSET_3ARROWSGRAY' ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'C' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'C' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'C' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 8 ip_column = 'C' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 9 ip_column = 'C' ip_value = 50 ).
    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3flags.
    lo_style_cond->mode_iconset  = ls_iconset3.
    lo_style_cond->set_range( ip_start_column  = 'D'
                                     ip_start_row     = 5
                                     ip_stop_column   = 'D'
                                     ip_stop_row      = 9 ).

    lo_worksheet->set_cell( ip_row = 4 ip_column = 'D' ip_value = 'C_ICONSET_3FLAGS' ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'D' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'D' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'D' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 8 ip_column = 'D' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 9 ip_column = 'D' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3trafficlights.
    lo_style_cond->mode_iconset  = ls_iconset3.
    lo_style_cond->set_range( ip_start_column  = 'E'
                                     ip_start_row     = 5
                                     ip_stop_column   = 'E'
                                     ip_stop_row      = 9 ).

    lo_worksheet->set_cell( ip_row = 4 ip_column = 'E' ip_value = 'C_ICONSET_3TRAFFICLIGHTS' ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'E' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'E' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'E' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 8 ip_column = 'E' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 9 ip_column = 'E' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3trafficlights2.
    lo_style_cond->mode_iconset  = ls_iconset3.
    lo_style_cond->set_range( ip_start_column  = 'F'
                                     ip_start_row     = 5
                                     ip_stop_column   = 'F'
                                     ip_stop_row      = 9 ).

    lo_worksheet->set_cell( ip_row = 4 ip_column = 'F' ip_value = 'C_ICONSET_3TRAFFICLIGHTS2' ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'F' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'F' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'F' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 8 ip_column = 'F' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 9 ip_column = 'F' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3signs.
    lo_style_cond->mode_iconset  = ls_iconset3.
    lo_style_cond->set_range( ip_start_column  = 'G'
                                     ip_start_row     = 5
                                     ip_stop_column   = 'G'
                                     ip_stop_row      = 9 ).

    lo_worksheet->set_cell( ip_row = 4 ip_column = 'G' ip_value = 'C_ICONSET_3SIGNS' ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'G' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'G' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'G' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 8 ip_column = 'G' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 9 ip_column = 'G' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3symbols.
    lo_style_cond->mode_iconset  = ls_iconset3.
    lo_style_cond->set_range( ip_start_column  = 'H'
                                     ip_start_row     = 5
                                     ip_stop_column   = 'H'
                                     ip_stop_row      = 9 ).

    lo_worksheet->set_cell( ip_row = 4 ip_column = 'H' ip_value = 'C_ICONSET_3SYMBOLS' ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'H' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'H' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'H' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 8 ip_column = 'H' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 9 ip_column = 'H' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset3-iconset                  = zcl_excel_style_cond=>c_iconset_3symbols2.
    lo_style_cond->mode_iconset  = ls_iconset3.
    lo_style_cond->set_range( ip_start_column  = 'I'
                                     ip_start_row     = 5
                                     ip_stop_column   = 'I'
                                     ip_stop_row      = 9 ).

    lo_worksheet->set_cell( ip_row = 4 ip_column = 'I' ip_value = 'C_ICONSET_3SYMBOLS2' ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'I' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'I' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'I' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 8 ip_column = 'I' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 9 ip_column = 'I' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset4-iconset                  = zcl_excel_style_cond=>c_iconset_4arrows.
    lo_style_cond->mode_iconset  = ls_iconset4.
    lo_style_cond->set_range( ip_start_column  = 'B'
                                     ip_start_row     = 12
                                     ip_stop_column   = 'B'
                                     ip_stop_row      = 16 ).

    lo_worksheet->set_cell( ip_row = 11 ip_column = 'B' ip_value = 'C_ICONSET_4ARROWS' ).
    lo_worksheet->set_cell( ip_row = 12 ip_column = 'B' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 13 ip_column = 'B' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 14 ip_column = 'B' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 15 ip_column = 'B' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 16 ip_column = 'B' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset4-iconset                  = zcl_excel_style_cond=>c_iconset_4arrowsgray.
    lo_style_cond->mode_iconset  = ls_iconset4.
    lo_style_cond->set_range( ip_start_column  = 'C'
                                     ip_start_row     = 12
                                     ip_stop_column   = 'C'
                                     ip_stop_row      = 16 ).

    lo_worksheet->set_cell( ip_row = 11 ip_column = 'C' ip_value = 'C_ICONSET_4ARROWSGRAY' ).
    lo_worksheet->set_cell( ip_row = 12 ip_column = 'C' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 13 ip_column = 'C' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 14 ip_column = 'C' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 15 ip_column = 'C' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 16 ip_column = 'C' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset4-iconset                  = zcl_excel_style_cond=>c_iconset_4redtoblack.
    lo_style_cond->mode_iconset  = ls_iconset4.
    lo_style_cond->set_range( ip_start_column  = 'D'
                                     ip_start_row     = 12
                                     ip_stop_column   = 'D'
                                     ip_stop_row      = 16 ).

    lo_worksheet->set_cell( ip_row = 11 ip_column = 'D' ip_value = 'C_ICONSET_4REDTOBLACK' ).
    lo_worksheet->set_cell( ip_row = 12 ip_column = 'D' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 13 ip_column = 'D' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 14 ip_column = 'D' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 15 ip_column = 'D' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 16 ip_column = 'D' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset4-iconset                  = zcl_excel_style_cond=>c_iconset_4rating.
    lo_style_cond->mode_iconset  = ls_iconset4.
    lo_style_cond->set_range( ip_start_column  = 'E'
                                     ip_start_row     = 12
                                     ip_stop_column   = 'E'
                                     ip_stop_row      = 16 ).

    lo_worksheet->set_cell( ip_row = 11 ip_column = 'E' ip_value = 'C_ICONSET_4RATING' ).
    lo_worksheet->set_cell( ip_row = 12 ip_column = 'E' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 13 ip_column = 'E' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 14 ip_column = 'E' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 15 ip_column = 'E' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 16 ip_column = 'E' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset4-iconset                  = zcl_excel_style_cond=>c_iconset_4trafficlights.
    lo_style_cond->mode_iconset  = ls_iconset4.
    lo_style_cond->set_range( ip_start_column  = 'F'
                                     ip_start_row     = 12
                                     ip_stop_column   = 'F'
                                     ip_stop_row      = 16 ).

    lo_worksheet->set_cell( ip_row = 11 ip_column = 'F' ip_value = 'C_ICONSET_4TRAFFICLIGHTS' ).
    lo_worksheet->set_cell( ip_row = 12 ip_column = 'F' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 13 ip_column = 'F' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 14 ip_column = 'F' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 15 ip_column = 'F' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 16 ip_column = 'F' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset5-iconset                  = zcl_excel_style_cond=>c_iconset_5arrows.
    lo_style_cond->mode_iconset  = ls_iconset5.
    lo_style_cond->set_range( ip_start_column  = 'B'
                                     ip_start_row     = 19
                                     ip_stop_column   = 'B'
                                     ip_stop_row      = 23 ).

    lo_worksheet->set_cell( ip_row = 18 ip_column = 'B' ip_value = 'C_ICONSET_5ARROWS' ).
    lo_worksheet->set_cell( ip_row = 19 ip_column = 'B' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 20 ip_column = 'B' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 21 ip_column = 'B' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 22 ip_column = 'B' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 23 ip_column = 'B' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset5-iconset                  = zcl_excel_style_cond=>c_iconset_5arrowsgray.
    lo_style_cond->mode_iconset  = ls_iconset5.
    lo_style_cond->set_range( ip_start_column  = 'C'
                                     ip_start_row     = 19
                                     ip_stop_column   = 'C'
                                     ip_stop_row      = 23 ).

    lo_worksheet->set_cell( ip_row = 18 ip_column = 'C' ip_value = 'C_ICONSET_5ARROWSGRAY' ).
    lo_worksheet->set_cell( ip_row = 19 ip_column = 'C' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 20 ip_column = 'C' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 21 ip_column = 'C' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 22 ip_column = 'C' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 23 ip_column = 'C' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset5-iconset                  = zcl_excel_style_cond=>c_iconset_5rating.
    lo_style_cond->mode_iconset  = ls_iconset5.
    lo_style_cond->set_range( ip_start_column  = 'D'
                                     ip_start_row     = 19
                                     ip_stop_column   = 'D'
                                     ip_stop_row      = 23 ).

    lo_worksheet->set_cell( ip_row = 18 ip_column = 'D' ip_value = 'C_ICONSET_5RATING' ).
    lo_worksheet->set_cell( ip_row = 19 ip_column = 'D' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 20 ip_column = 'D' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 21 ip_column = 'D' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 22 ip_column = 'D' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 23 ip_column = 'D' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
    lo_style_cond->priority      = 1.
    ls_iconset5-iconset                  = zcl_excel_style_cond=>c_iconset_5quarters.
    lo_style_cond->mode_iconset  = ls_iconset5.
    lo_style_cond->set_range( ip_start_column  = 'E'
                                     ip_start_row     = 19
                                     ip_stop_column   = 'E'
                                     ip_stop_row      = 23 ).

* DATABAR

    lo_worksheet->set_cell( ip_row = 25 ip_column = 'B' ip_value = 'DATABAR' ).
    lo_worksheet->set_cell( ip_row = 26 ip_column = 'B' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 27 ip_column = 'B' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 28 ip_column = 'B' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 29 ip_column = 'B' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 30 ip_column = 'B' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule            = zcl_excel_style_cond=>c_rule_databar.
    lo_style_cond->priority        = 1.
    lo_style_cond->mode_databar = ls_databar.
    lo_style_cond->set_range( ip_start_column  = 'B'
                                     ip_start_row     = 26
                                     ip_stop_column   = 'B'
                                     ip_stop_row      = 30 ).

* COLORSCALE

    lo_worksheet->set_cell( ip_row = 25 ip_column = 'C' ip_value = 'COLORSCALE 2 COLORS' ).
    lo_worksheet->set_cell( ip_row = 26 ip_column = 'C' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 27 ip_column = 'C' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 28 ip_column = 'C' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 29 ip_column = 'C' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 30 ip_column = 'C' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule            = zcl_excel_style_cond=>c_rule_colorscale.
    lo_style_cond->priority        = 1.
    lo_style_cond->mode_colorscale = ls_colorscale2.
    lo_style_cond->set_range( ip_start_column  = 'C'
                                     ip_start_row     = 26
                                     ip_stop_column   = 'C'
                                     ip_stop_row      = 30 ).


    lo_worksheet->set_cell( ip_row = 25 ip_column = 'D' ip_value = 'COLORSCALE 3 COLORS' ).
    lo_worksheet->set_cell( ip_row = 26 ip_column = 'D' ip_value = 10 ).
    lo_worksheet->set_cell( ip_row = 27 ip_column = 'D' ip_value = 20 ).
    lo_worksheet->set_cell( ip_row = 28 ip_column = 'D' ip_value = 30 ).
    lo_worksheet->set_cell( ip_row = 29 ip_column = 'D' ip_value = 40 ).
    lo_worksheet->set_cell( ip_row = 30 ip_column = 'D' ip_value = 50 ).

    lo_style_cond = lo_worksheet->add_new_style_cond( ).
    lo_style_cond->rule            = zcl_excel_style_cond=>c_rule_colorscale.
    lo_style_cond->priority        = 1.
    lo_style_cond->mode_colorscale = ls_colorscale3.
    lo_style_cond->set_range( ip_start_column  = 'D'
                                     ip_start_row     = 26
                                     ip_stop_column   = 'D'
                                     ip_stop_row      = 30 ).

  endmethod.
  method demo9.

    data lo_worksheet type ref to zcl_excel_worksheet.
    data lo_column    type ref to zcl_excel_column.
    data lo_row       type ref to zcl_excel_row.

    " Creates active sheet
    create object ro_excel.

    " Get active sheet
    lo_worksheet = ro_excel->get_active_worksheet( ).
    lo_worksheet->set_title( 'Sheet1' ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world in AutoSize column' ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = 'Hello world in a column width size 50' ).
    lo_worksheet->set_cell( ip_column = 'D' ip_row = 4 ip_value = 'Hello world (hidden column)' ).
    lo_worksheet->set_cell( ip_column = 'F' ip_row = 2 ip_value = 'Outline column level 0' ).
    lo_worksheet->set_cell( ip_column = 'G' ip_row = 2 ip_value = 'Outline column level 1' ).
    lo_worksheet->set_cell( ip_column = 'H' ip_row = 2 ip_value = 'Outline column level 2' ).
    lo_worksheet->set_cell( ip_column = 'I' ip_row = 2 ip_value = 'Small' ).


    lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Hello world (hidden row)' ).
    lo_worksheet->set_cell( ip_column = 'E' ip_row = 5 ip_value = 'Hello world in a row height size 20' ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 9 ip_value = 'Simple outline rows 10-16 ( collapsed )' ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 19 ip_value = '3 Outlines - Outlinelevel 1 is collapsed' ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 19 ip_value = 'One of the two inner outlines is expanded, one collapsed' ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 20 ip_value = 'Inner outline level - expanded' ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 24 ip_value = 'Inner outline level - lines 25-28 are collapsed' ).

    lo_worksheet->zif_excel_sheet_properties~summarybelow = zif_excel_sheet_properties=>c_below_off. " By default is on
    lo_worksheet->zif_excel_sheet_properties~summaryright = zif_excel_sheet_properties=>c_right_off. " By default is on

    " Column Settings
    " Auto size
    lo_column = lo_worksheet->get_column( ip_column = 'B' ).
    lo_column->set_auto_size( ip_auto_size = abap_true ).
    lo_column = lo_worksheet->get_column( ip_column = 'I' ).
    lo_column->set_auto_size( ip_auto_size = abap_true ).
    " Manual Width
    lo_column = lo_worksheet->get_column( ip_column = 'C' ).
    lo_column->set_width( ip_width = 50 ).
    lo_column = lo_worksheet->get_column( ip_column = 'D' ).
    lo_column->set_visible( ip_visible = abap_false ).
    " Implementation in the Writer is not working yet ===== TODO =====
    lo_column = lo_worksheet->get_column( ip_column = 'F' ).
    lo_column->set_outline_level( ip_outline_level = 0 ).
    lo_column = lo_worksheet->get_column( ip_column = 'G' ).
    lo_column->set_outline_level( ip_outline_level = 1 ).
    lo_column = lo_worksheet->get_column( ip_column = 'H' ).
    lo_column->set_outline_level( ip_outline_level = 2 ).

    lo_row = lo_worksheet->get_row( ip_row = 1 ).
    lo_row->set_visible( ip_visible = abap_false ).
    lo_row = lo_worksheet->get_row( ip_row = 5 ).
    lo_row->set_row_height( ip_row_height = 20 ).

* Define an outline rows 10-16, collapsed on startup
    lo_worksheet->set_row_outline( iv_row_from = 10
                                 iv_row_to   = 16
                                 iv_collapsed = abap_true ).  " collapsed

* Define an inner outline rows 21-22, expanded when outer outline becomes extended
    lo_worksheet->set_row_outline( iv_row_from = 21
                                 iv_row_to   = 22
                                 iv_collapsed = abap_false ). " expanded

* Define an inner outline rows 25-28, collapsed on startup
    lo_worksheet->set_row_outline( iv_row_from = 25
                                 iv_row_to   = 28
                                 iv_collapsed = abap_true ).  " collapsed

* Define an outer outline rows 20-30, collapsed on startup
    lo_worksheet->set_row_outline( iv_row_from = 20
                                 iv_row_to   = 30
                                 iv_collapsed = abap_true ).  " collapsed

* Hint:  the order you create the outlines can be arbitrary
*        You can start with inner outlines or with outer outlines

*--------------------------------------------------------------------*
* Hide columns right of column M
*--------------------------------------------------------------------*
    lo_worksheet->zif_excel_sheet_properties~hide_columns_from = 'M'.

  endmethod.
  method demo10.

    data: lo_worksheet              type ref to zcl_excel_worksheet,
          lv_style_bold_border_guid type zif_excel_data_decl=>zexcel_cell_style,
          lo_style_bold_border      type ref to zcl_excel_style,
          lo_border_dark            type ref to zcl_excel_style_border.

    create object ro_excel.

    " Get active sheet
    lo_worksheet = ro_excel->get_active_worksheet( ).
    lo_worksheet->set_title( 'sheet1' ).

    create object lo_border_dark.
    lo_border_dark->border_color-rgb = zcl_excel_style_color=>c_black.
    lo_border_dark->border_style = zcl_excel_style_border=>c_border_thin.

    lo_style_bold_border = ro_excel->add_new_style( ).
    lo_style_bold_border->font->bold = abap_true.
    lo_style_bold_border->font->italic = abap_false.
    lo_style_bold_border->font->color-rgb = zcl_excel_style_color=>c_black.
    lo_style_bold_border->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_style_bold_border->borders->allborders = lo_border_dark.
    lv_style_bold_border_guid = lo_style_bold_border->get_guid( ).

    lo_worksheet->set_cell( ip_row = 2 ip_column = 'A' ip_value = 'Test' ).

    lo_worksheet->set_cell( ip_row = 2 ip_column = 'B' ip_value = 'Banana' ip_style = lv_style_bold_border_guid ).
    lo_worksheet->set_cell( ip_row = 2 ip_column = 'C' ip_value = '' ip_style = lv_style_bold_border_guid ).
    lo_worksheet->set_cell( ip_row = 2 ip_column = 'D' ip_value = '' ip_style = lv_style_bold_border_guid ).
    lo_worksheet->set_cell( ip_row = 2 ip_column = 'E' ip_value = '' ip_style = lv_style_bold_border_guid ).
    lo_worksheet->set_cell( ip_row = 2 ip_column = 'F' ip_value = '' ip_style = lv_style_bold_border_guid ).
    lo_worksheet->set_cell( ip_row = 2 ip_column = 'G' ip_value = '' ip_style = lv_style_bold_border_guid ).
    lo_worksheet->set_cell( ip_row = 4 ip_column = 'B' ip_value = 'Apple' ip_style = lv_style_bold_border_guid ).
    lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = '' ip_style = lv_style_bold_border_guid ).
    lo_worksheet->set_cell( ip_row = 4 ip_column = 'D' ip_value = '' ip_style = lv_style_bold_border_guid ).
    lo_worksheet->set_cell( ip_row = 4 ip_column = 'E' ip_value = '' ip_style = lv_style_bold_border_guid ).
    lo_worksheet->set_cell( ip_row = 4 ip_column = 'F' ip_value = '' ip_style = lv_style_bold_border_guid ).
    lo_worksheet->set_cell( ip_row = 4 ip_column = 'G' ip_value = '' ip_style = lv_style_bold_border_guid ).

    lo_worksheet->set_merge( ip_row = 4 ip_column_start = 'B' ip_column_end = 'G' ).

    " Test also if merge works when oher merged chells are empty
    lo_worksheet->set_merge( ip_range = 'B6:G6' ip_value = 'Tomato' ).

    " Test the patch provided by Victor Alekhin to merge cells in one column
    lo_worksheet->set_merge( ip_range = 'B8:G10' ip_value = 'Merge cells also over multiple rows by Victor Alekhin' ).

    " Test the patch provided by Alexander Budeyev with different column merges
    lo_worksheet->set_cell( ip_row = 12 ip_column = 'B' ip_value = 'Merge cells with different merges by Alexander Budeyev' ).
    lo_worksheet->set_cell( ip_row = 13 ip_column = 'B' ip_value = 'Test' ).

    lo_worksheet->set_cell( ip_row = 13 ip_column = 'D' ip_value = 'Banana' ip_style = lv_style_bold_border_guid ).
    lo_worksheet->set_cell( ip_row = 14 ip_column = 'D' ip_value = '' ip_style = lv_style_bold_border_guid ).
    lo_worksheet->set_cell( ip_row = 13 ip_column = 'E' ip_value = 'Apple' ip_style = lv_style_bold_border_guid ).
    lo_worksheet->set_cell( ip_row = 13 ip_column = 'F' ip_value = '' ip_style = lv_style_bold_border_guid ).

    " Test merge (issue)
    lo_worksheet->set_merge( ip_row = 13 ip_column_start = 'B' ip_column_end = 'C' ip_row_to = 15 ).
    lo_worksheet->set_merge( ip_row = 13 ip_column_start = 'D' ip_column_end = 'D' ip_row_to = 14 ).
    lo_worksheet->set_merge( ip_row = 13 ip_column_start = 'E' ip_column_end = 'F' ).

    " Test area with merge
    lo_worksheet->set_area( ip_row = 18 ip_row_to = 19 ip_column_start = 'B' ip_column_end = 'G' ip_style = lv_style_bold_border_guid
                            ip_value = 'Merge cells with new area method by Helmut Bohr ' ip_merge = abap_true ).

    " Test area without merge
    lo_worksheet->set_area( ip_row = 21 ip_row_to = 22 ip_column_start = 'B' ip_column_end = 'G' ip_style = lv_style_bold_border_guid
                            ip_value = 'Test area' ).
  endmethod.
  method demo11.

    data: lo_worksheet             type ref to zcl_excel_worksheet,
          lo_style_center          type ref to zcl_excel_style,
          lo_style_right           type ref to zcl_excel_style,
          lo_style_left            type ref to zcl_excel_style,
          lo_style_general         type ref to zcl_excel_style,
          lo_style_bottom          type ref to zcl_excel_style,
          lo_style_middle          type ref to zcl_excel_style,
          lo_style_top             type ref to zcl_excel_style,
          lo_style_justify         type ref to zcl_excel_style,
          lo_style_mixed           type ref to zcl_excel_style,
          lo_style_mixed_wrap      type ref to zcl_excel_style,
          lo_style_rotated         type ref to zcl_excel_style,
          lo_style_shrink          type ref to zcl_excel_style,
          lo_style_indent          type ref to zcl_excel_style,
          lv_style_center_guid     type zif_excel_data_decl=>zexcel_cell_style,
          lv_style_right_guid      type zif_excel_data_decl=>zexcel_cell_style,
          lv_style_left_guid       type zif_excel_data_decl=>zexcel_cell_style,
          lv_style_general_guid    type zif_excel_data_decl=>zexcel_cell_style,
          lv_style_bottom_guid     type zif_excel_data_decl=>zexcel_cell_style,
          lv_style_middle_guid     type zif_excel_data_decl=>zexcel_cell_style,
          lv_style_top_guid        type zif_excel_data_decl=>zexcel_cell_style,
          lv_style_justify_guid    type zif_excel_data_decl=>zexcel_cell_style,
          lv_style_mixed_guid      type zif_excel_data_decl=>zexcel_cell_style,
          lv_style_mixed_wrap_guid type zif_excel_data_decl=>zexcel_cell_style,
          lv_style_rotated_guid    type zif_excel_data_decl=>zexcel_cell_style,
          lv_style_shrink_guid     type zif_excel_data_decl=>zexcel_cell_style,
          lv_style_indent_guid     type zif_excel_data_decl=>zexcel_cell_style.

    data lo_row type ref to zcl_excel_row.

    create object ro_excel.

    " Get active sheet
    lo_worksheet = ro_excel->get_active_worksheet( ).
    lo_worksheet->set_title( 'sheet1' ).

    "Center
    lo_style_center = ro_excel->add_new_style( ).
    lo_style_center->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lv_style_center_guid = lo_style_center->get_guid( ).
    "Right
    lo_style_right = ro_excel->add_new_style( ).
    lo_style_right->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_right.
    lv_style_right_guid = lo_style_right->get_guid( ).
    "Left
    lo_style_left = ro_excel->add_new_style( ).
    lo_style_left->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_left.
    lv_style_left_guid = lo_style_left->get_guid( ).
    "General
    lo_style_general = ro_excel->add_new_style( ).
    lo_style_general->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_general.
    lv_style_general_guid = lo_style_general->get_guid( ).
    "Bottom
    lo_style_bottom = ro_excel->add_new_style( ).
    lo_style_bottom->alignment->vertical = zcl_excel_style_alignment=>c_vertical_bottom.
    lv_style_bottom_guid = lo_style_bottom->get_guid( ).
    "Middle
    lo_style_middle = ro_excel->add_new_style( ).
    lo_style_middle->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    lv_style_middle_guid = lo_style_middle->get_guid( ).
    "Top
    lo_style_top = ro_excel->add_new_style( ).
    lo_style_top->alignment->vertical = zcl_excel_style_alignment=>c_vertical_top.
    lv_style_top_guid = lo_style_top->get_guid( ).
    "Justify
    lo_style_justify = ro_excel->add_new_style( ).
    lo_style_justify->alignment->vertical = zcl_excel_style_alignment=>c_vertical_justify.
    lv_style_justify_guid = lo_style_justify->get_guid( ).

    "Shrink
    lo_style_shrink = ro_excel->add_new_style( ).
    lo_style_shrink->alignment->shrinktofit = abap_true.
    lv_style_shrink_guid = lo_style_shrink->get_guid( ).

    "Indent
    lo_style_indent = ro_excel->add_new_style( ).
    lo_style_indent->alignment->indent = 5.
    lv_style_indent_guid = lo_style_indent->get_guid( ).

    "Middle / Centered / Wrap
    lo_style_mixed_wrap = ro_excel->add_new_style( ).
    lo_style_mixed_wrap->alignment->horizontal   = zcl_excel_style_alignment=>c_horizontal_center.
    lo_style_mixed_wrap->alignment->vertical     = zcl_excel_style_alignment=>c_vertical_center.
    lo_style_mixed_wrap->alignment->wraptext     = abap_true.
    lv_style_mixed_wrap_guid = lo_style_mixed_wrap->get_guid( ).

    "Middle / Centered / Wrap
    lo_style_mixed = ro_excel->add_new_style( ).
    lo_style_mixed->alignment->horizontal   = zcl_excel_style_alignment=>c_horizontal_center.
    lo_style_mixed->alignment->vertical     = zcl_excel_style_alignment=>c_vertical_center.
    lv_style_mixed_guid = lo_style_mixed->get_guid( ).

    "Center
    lo_style_rotated = ro_excel->add_new_style( ).
    lo_style_rotated->alignment->horizontal   = zcl_excel_style_alignment=>c_horizontal_center.
    lo_style_rotated->alignment->vertical     = zcl_excel_style_alignment=>c_vertical_center.
    lo_style_rotated->alignment->textrotation = 165.                        " -75Ã‚Â° == 90Ã‚Â° + 75Ã‚Â°
    lv_style_rotated_guid = lo_style_rotated->get_guid( ).


    " Set row size for first 7 rows to 40
    do 7 times.
      lo_row = lo_worksheet->get_row( sy-index ).
      lo_row->set_row_height( 40 ).
    enddo.

    "Horizontal alignment
    lo_worksheet->set_cell( ip_row = 4 ip_column = 'B' ip_value = 'Centered Text' ip_style = lv_style_center_guid ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'B' ip_value = 'Right Text'    ip_style = lv_style_right_guid ).
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'B' ip_value = 'Left Text'     ip_style = lv_style_left_guid ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'B' ip_value = 'General Text'  ip_style = lv_style_general_guid ).

    " Shrink & indent
    lo_worksheet->set_cell( ip_row = 4 ip_column = 'F' ip_value = 'Text shrinked' ip_style = lv_style_shrink_guid ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'F' ip_value = 'Text indented' ip_style = lv_style_indent_guid ).

    "Vertical alignment

    lo_worksheet->set_cell( ip_row = 4 ip_column = 'D' ip_value = 'Bottom Text'    ip_style = lv_style_bottom_guid ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'D' ip_value = 'Middle Text'    ip_style = lv_style_middle_guid ).
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'D' ip_value = 'Top Text'       ip_style = lv_style_top_guid ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'D' ip_value = 'Justify Text'   ip_style = lv_style_justify_guid ).

    " Wrapped
    lo_worksheet->set_cell( ip_row = 10 ip_column = 'B'
                          ip_value = 'This is a wrapped text centered in the middle'
                          ip_style = lv_style_mixed_wrap_guid ).

    " Rotated
    lo_worksheet->set_cell( ip_row = 10 ip_column = 'D'
                          ip_value = 'This is a centered text rotated by -75Ã‚Â°'
                          ip_style = lv_style_rotated_guid ).

    " forced line break
    data: lv_value type string.
    concatenate 'This is a wrapped text centered in the middle' cl_abap_char_utilities=>cr_lf
    'and a manuall line break.' into lv_value.
    lo_worksheet->set_cell( ip_row = 11 ip_column = 'B'
                          ip_value = lv_value
                          ip_style = lv_style_mixed_guid ).

  endmethod.

  method demo5.
    data: begin of ls_out,
            material type matnr,
            mattext  type string,
            qnty     type kwmeng,
            meins    type meins,
            amnt     type dmbtr,
            currn    type waers,
          end of ls_out,
          lt_out           like standard table of ls_out,
          lt_field_catalog type zif_excel_data_decl=>zexcel_t_fieldcatalog.

    data lo_worksheet type ref to zcl_excel_worksheet.
    ro_excel = new #( ).

    " Get active sheet
    lo_worksheet = ro_excel->get_active_worksheet( ).
    lt_out = value #(
      ( material = '897234' mattext = 'sddf' qnty = '12.768' meins = 'ST' amnt = '1233.23' currn = 'USD' )
      ( material = '897234' mattext = 'sddf' qnty = '12.768' meins = 'KGM' amnt = '1233.23' currn = 'JPY' )
      ( material = '897234' mattext = 'sddf' qnty = '12.768' meins = 'DGP' amnt = '1233.23' currn = 'BHD' )
      ( material = '897234' mattext = 'sddf' qnty = '12.768' meins = '13' amnt = '1233.23' currn = 'CAD' ) ).
    lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = lt_out ).
    lt_field_catalog[ fieldname = 'AMNT' ]-currency_column = 'CURRN'.
    lt_field_catalog[ fieldname = 'QNTY' ]-unit_column = 'MEINS'.
    lt_field_catalog[ fieldname = 'MATERIAL' ]-text_column = 'MATTEXT'.
    delete lt_field_catalog where fieldname = 'CURRN' or fieldname = 'MEINS' or fieldname = 'MATTEXT'.


    lo_worksheet->calculate_column_widths(  ).
    lo_worksheet->bind_table( ip_table          = lt_out
                              it_field_catalog = lt_field_catalog
                              is_table_settings = value zif_excel_data_decl=>zexcel_s_table_settings(
                                                            table_name       = 'MyDataTable'
                                                            top_left_column  = 'A'
                                                            top_left_row     = 1
                                                            show_row_stripes = abap_true
                                                             ) ).

    do lo_worksheet->get_highest_column( ) times.
      lo_worksheet->get_column( zcl_excel_common=>convert_column2alpha( sy-index ) )->set_auto_size( abap_true ).
    enddo.
    lo_worksheet->calculate_column_widths( ).
  endmethod.

  method demo8.
  endmethod.
  method demo12.

    data: lo_worksheet type ref to zcl_excel_worksheet,
          lo_column    type ref to zcl_excel_column.

    data: lv_value  type string,
          lv_count  type i value 10,
          lv_packed type p length 16 decimals 1 value '1234567890.5'.

    constants: lc_typekind_string type abap_typekind value cl_abap_typedescr=>typekind_string,
               lc_typekind_packed type abap_typekind value cl_abap_typedescr=>typekind_packed,
               lc_typekind_num    type abap_typekind value cl_abap_typedescr=>typekind_num,
               lc_typekind_date   type abap_typekind value cl_abap_typedescr=>typekind_date.

    " Creates active sheet
    create object ro_excel.

    " Get active sheet
    lo_worksheet = ro_excel->get_active_worksheet( ).
    lo_worksheet->set_title( ip_title = 'Cell data types' ).

    lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Number as String'
                          ip_abap_type = lc_typekind_string ).
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = '11'
                          ip_abap_type = lc_typekind_string ).

    lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'String'
                          ip_abap_type = lc_typekind_string ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = ' String with leading spaces'
                          ip_abap_type = lc_typekind_string ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'String without leading spaces'
                          ip_abap_type = lc_typekind_string ).

    lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = 'Packed'
                          ip_abap_type = lc_typekind_string ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 2 ip_value = '50000.01-'
                          ip_abap_type = lc_typekind_packed ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = '5000.02'
                          ip_abap_type = lc_typekind_packed ).

    lo_worksheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = 'Number with Percentage'
                          ip_abap_type = lc_typekind_string ).
    lo_worksheet->set_cell( ip_column = 'D' ip_row = 2 ip_value = '0 %'
                          ip_abap_type = lc_typekind_num ).
    lo_worksheet->set_cell( ip_column = 'D' ip_row = 3 ip_value = '50 %'
                          ip_abap_type = lc_typekind_num ).

    lo_worksheet->set_cell( ip_column = 'E' ip_row = 1 ip_value = 'Date'
                          ip_abap_type = lc_typekind_string ).
    lo_worksheet->set_cell( ip_column = 'E' ip_row = 2 ip_value = '20110831'
                          ip_abap_type = lc_typekind_date ).

    while lv_count <= 15.
      lv_value = lv_count.
      concatenate 'Positive Value with' lv_value 'Digits' into lv_value separated by space.
      lo_worksheet->set_cell( ip_column = 'B' ip_row = lv_count ip_value = lv_value
                            ip_abap_type = lc_typekind_string ).
      lo_worksheet->set_cell( ip_column = 'C' ip_row = lv_count ip_value = lv_packed
                            ip_abap_type = lc_typekind_packed ).
      concatenate 'Positive Value with' lv_value 'Digits formated as string' into lv_value separated by space.
      lo_worksheet->set_cell( ip_column = 'D' ip_row = lv_count ip_value = lv_value
                            ip_abap_type = lc_typekind_string ).
      lo_worksheet->set_cell( ip_column = 'E' ip_row = lv_count ip_value = lv_packed
                            ip_abap_type = lc_typekind_string ).
      lv_packed = lv_packed * 10.
      lv_count  = lv_count + 1.
    endwhile.

    lo_column = lo_worksheet->get_column( ip_column = 'A' ).
    lo_column->set_auto_size( abap_true ).
    lo_column = lo_worksheet->get_column( ip_column = 'B' ).
    lo_column->set_auto_size( abap_true ).
    lo_column = lo_worksheet->get_column( ip_column = 'C' ).
    lo_column->set_auto_size( abap_true ).
    lo_column = lo_worksheet->get_column( ip_column = 'D' ).
    lo_column->set_auto_size( abap_true ).
    lo_column = lo_worksheet->get_column( ip_column = 'E' ).
    lo_column->set_auto_size( abap_true ).


  endmethod.

  method demo3.

    data:
      lo_worksheet type ref to zcl_excel_worksheet,
      lo_hyperlink type ref to zcl_excel_hyperlink,
      lv_tabcolor  type zif_excel_data_decl=>zexcel_s_tabcolor,
      ls_header    type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
      ls_footer    type zif_excel_data_decl=>zexcel_s_worksheet_head_foot.

    " Creates active sheet
    create object ro_excel.

    " Get active sheet
    lo_worksheet = ro_excel->get_active_worksheet( ).
    lo_worksheet->set_title( ip_title = 'Sheet1' ).
    lo_worksheet->zif_excel_sheet_properties~selected = zif_excel_sheet_properties=>c_selected.
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'This is the first sheet' ).
* Set color to tab with sheetname   - Red
    lv_tabcolor-rgb = zcl_excel_style_color=>create_new_argb( ip_red   = 'FF'
                                                            ip_green = '00'
                                                            ip_blu   = '00' ).
    lo_worksheet->set_tabcolor( lv_tabcolor ).

    lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = 'Sheet2!B2' ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'This is link to second sheet' ip_hyperlink = lo_hyperlink ).

    " Page printing settings
    lo_worksheet->sheet_setup->set_page_margins( ip_header = '1' ip_footer = '1' ip_unit = 'cm' ).
    lo_worksheet->sheet_setup->black_and_white   = 'X'.
    lo_worksheet->sheet_setup->fit_to_page       = 'X'.  " you should turn this on to activate fit_to_height and fit_to_width
    lo_worksheet->sheet_setup->fit_to_height     = 0.    " used only if ip_fit_to_page = 'X'
    lo_worksheet->sheet_setup->fit_to_width      = 2.    " used only if ip_fit_to_page = 'X'
    lo_worksheet->sheet_setup->orientation       = zcl_excel_sheet_setup=>c_orientation_landscape.
    lo_worksheet->sheet_setup->page_order        = zcl_excel_sheet_setup=>c_ord_downthenover.
    lo_worksheet->sheet_setup->paper_size        = zcl_excel_sheet_setup=>c_papersize_a4.
    lo_worksheet->sheet_setup->scale             = 80.   " used only if ip_fit_to_page = SPACE

    " Header and Footer
    ls_header-right_value = 'print date &D'.
    ls_header-right_font-size = 8.
    ls_header-right_font-name = zcl_excel_style_font=>c_name_arial.

    ls_footer-left_value = '&Z&F'. "Path / Filename
    ls_footer-left_font = ls_header-right_font.
    ls_footer-right_value = 'page &P of &N'. "page x of y
    ls_footer-right_font = ls_header-right_font.

    lo_worksheet->sheet_setup->set_header_footer( ip_odd_header  = ls_header
                                                ip_odd_footer  = ls_footer ).


    lo_worksheet = ro_excel->add_new_worksheet( ).
    lo_worksheet->set_title( ip_title = 'Sheet2' ).
* Set color to tab with sheetname   - Green
    lv_tabcolor-rgb = zcl_excel_style_color=>create_new_argb( ip_red   = '00'
                                                            ip_green = 'FF'
                                                            ip_blu   = '00' ).
    lo_worksheet->set_tabcolor( lv_tabcolor ).
    lo_worksheet->zif_excel_sheet_properties~selected = zif_excel_sheet_properties=>c_selected.
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'This is the second sheet' ).
    lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = 'Sheet1!B2' ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'This is link to first sheet' ip_hyperlink = lo_hyperlink ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 4 ip_value = 'Sheet3 is hidden' ).

    lo_worksheet->sheet_setup->set_header_footer( ip_odd_header  = ls_header
                                                ip_odd_footer  = ls_footer ).

    lo_worksheet = ro_excel->add_new_worksheet( ).
    lo_worksheet->set_title( ip_title = 'Sheet3' ).
* Set color to tab with sheetname   - Blue
    lv_tabcolor-rgb = zcl_excel_style_color=>create_new_argb( ip_red   = '00'
                                                            ip_green = '00'
                                                            ip_blu   = 'FF' ).
    lo_worksheet->set_tabcolor( lv_tabcolor ).
    lo_worksheet->zif_excel_sheet_properties~hidden = zif_excel_sheet_properties=>c_hidden.

    lo_worksheet->sheet_setup->set_header_footer( ip_odd_header  = ls_header
                                                ip_odd_footer  = ls_footer ).

    lo_worksheet = ro_excel->add_new_worksheet( ).
    lo_worksheet->set_title( ip_title = 'Sheet4' ).
* Set color to tab with sheetname   - other color
    lv_tabcolor-rgb = zcl_excel_style_color=>create_new_argb( ip_red   = '00'
                                                            ip_green = 'FF'
                                                            ip_blu   = 'FF' ).
    lo_worksheet->set_tabcolor( lv_tabcolor ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Cell B3 has value 0' ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 0 ).
    lo_worksheet->zif_excel_sheet_properties~show_zeros = zif_excel_sheet_properties=>c_hidezero.

    lo_worksheet->sheet_setup->set_header_footer( ip_odd_header  = ls_header
                                                ip_odd_footer  = ls_footer ).

    ro_excel->set_active_sheet_index_by_name( 'Sheet1' ).
  endmethod.
  method write.
    data lo_writer type ref to zif_excel_writer.

    lo_writer = new zcl_excel_writer_2007( ).
    return lo_writer->write_file( io_excel ).
  endmethod.
  method demo2.

    data:
      lo_worksheet              type ref to zcl_excel_worksheet,
      lo_style_bold             type ref to zcl_excel_style,
      lo_style_underline        type ref to zcl_excel_style,
      lo_style_filled           type ref to zcl_excel_style,
      lo_style_filled_green     type ref to zcl_excel_style,
      lo_style_filled_turquoise type ref to zcl_excel_style,
      lo_style_border           type ref to zcl_excel_style,
      lo_style_button           type ref to zcl_excel_style,
      lo_border_dark            type ref to zcl_excel_style_border,
      lo_border_light           type ref to zcl_excel_style_border,
      lo_style_gr_cornerlb      type ref to zcl_excel_style,
      lo_style_gr_cornerlt      type ref to zcl_excel_style,
      lo_style_gr_cornerrb      type ref to zcl_excel_style,
      lo_style_gr_cornerrt      type ref to zcl_excel_style,
      lo_style_gr_horizontal90  type ref to zcl_excel_style,
      lo_style_gr_horizontal270 type ref to zcl_excel_style,
      lo_style_gr_horizontalb   type ref to zcl_excel_style,
      lo_style_gr_vertical      type ref to zcl_excel_style,
      lo_style_gr_vertical2     type ref to zcl_excel_style,
      lo_style_gr_fromcenter    type ref to zcl_excel_style,
      lo_style_gr_diagonal45    type ref to zcl_excel_style,
      lo_style_gr_diagonal45b   type ref to zcl_excel_style,
      lo_style_gr_diagonal135   type ref to zcl_excel_style,
      lo_style_gr_diagonal135b  type ref to zcl_excel_style.

    data: lo_row type ref to zcl_excel_row.

    " Creates active sheet
    create object ro_excel.

    " Create border object
    create object lo_border_dark.
    lo_border_dark->border_color-rgb = zcl_excel_style_color=>c_black.
    lo_border_dark->border_style = zcl_excel_style_border=>c_border_thin.
    create object lo_border_light.
    lo_border_light->border_color-rgb = zcl_excel_style_color=>c_gray.
    lo_border_light->border_style = zcl_excel_style_border=>c_border_thin.
    " Create a bold / italic style
    lo_style_bold               = ro_excel->add_new_style( ).
    lo_style_bold->font->bold   = abap_true.
    lo_style_bold->font->italic = abap_true.
    lo_style_bold->font->name   = zcl_excel_style_font=>c_name_arial.
    lo_style_bold->font->scheme = zcl_excel_style_font=>c_scheme_none.
    lo_style_bold->font->color-rgb  = zcl_excel_style_color=>c_red.
    " Create an underline double style
    lo_style_underline                        = ro_excel->add_new_style( ).
    lo_style_underline->font->underline       = abap_true.
    lo_style_underline->font->underline_mode  = zcl_excel_style_font=>c_underline_double.
    lo_style_underline->font->name            = zcl_excel_style_font=>c_name_roman.
    lo_style_underline->font->scheme          = zcl_excel_style_font=>c_scheme_none.
    lo_style_underline->font->family          = zcl_excel_style_font=>c_family_roman.
    " Create filled style yellow
    lo_style_filled                 = ro_excel->add_new_style( ).
    lo_style_filled->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
    lo_style_filled->fill->fgcolor-theme  = zcl_excel_style_color=>c_theme_accent6.
    " Create border with button effects
    lo_style_button                   = ro_excel->add_new_style( ).
    lo_style_button->borders->right   = lo_border_dark.
    lo_style_button->borders->down    = lo_border_dark.
    lo_style_button->borders->left    = lo_border_light.
    lo_style_button->borders->top     = lo_border_light.
    "Create style with border
    lo_style_border                         = ro_excel->add_new_style( ).
    lo_style_border->borders->allborders    = lo_border_dark.
    lo_style_border->borders->diagonal      = lo_border_dark.
    lo_style_border->borders->diagonal_mode = zcl_excel_style_borders=>c_diagonal_both.
    " Create filled style green
    lo_style_filled_green                     = ro_excel->add_new_style( ).
    lo_style_filled_green->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_style_filled_green->fill->fgcolor-rgb  = zcl_excel_style_color=>c_green.
    lo_style_filled_green->font->name         = zcl_excel_style_font=>c_name_cambria.
    lo_style_filled_green->font->scheme       = zcl_excel_style_font=>c_scheme_major.

    " Create filled with gradients
    lo_style_gr_cornerlb                     = ro_excel->add_new_style(  ).
    lo_style_gr_cornerlb->fill->filltype     = zcl_excel_style_fill=>c_fill_gradient_cornerlb.
    lo_style_gr_cornerlb->fill->fgcolor-rgb  = zcl_excel_style_color=>c_blue.
    lo_style_gr_cornerlb->fill->bgcolor-rgb  = zcl_excel_style_color=>c_white.
    lo_style_gr_cornerlb->font->name         = zcl_excel_style_font=>c_name_cambria.
    lo_style_gr_cornerlb->font->scheme       = zcl_excel_style_font=>c_scheme_major.

    lo_style_gr_cornerlt                     = ro_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
    lo_style_gr_cornerlt->fill->filltype     = zcl_excel_style_fill=>c_fill_gradient_cornerlt.

    lo_style_gr_cornerrb                     = ro_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
    lo_style_gr_cornerrb->fill->filltype     = zcl_excel_style_fill=>c_fill_gradient_cornerrb.

    lo_style_gr_cornerrt                     = ro_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
    lo_style_gr_cornerrt->fill->filltype     = zcl_excel_style_fill=>c_fill_gradient_cornerrt.

    lo_style_gr_horizontal90                 = ro_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
    lo_style_gr_horizontal90->fill->filltype = zcl_excel_style_fill=>c_fill_gradient_horizontal90.

    lo_style_gr_horizontal270                = ro_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
    lo_style_gr_horizontal270->fill->filltype = zcl_excel_style_fill=>c_fill_gradient_horizontal270.

    lo_style_gr_horizontalb                  = ro_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
    lo_style_gr_horizontalb->fill->filltype  = zcl_excel_style_fill=>c_fill_gradient_horizontalb.

    lo_style_gr_vertical                     = ro_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
    lo_style_gr_vertical->fill->filltype     = zcl_excel_style_fill=>c_fill_gradient_vertical.

    lo_style_gr_vertical2                    = ro_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
    lo_style_gr_vertical2->fill->filltype    = zcl_excel_style_fill=>c_fill_gradient_vertical.
    lo_style_gr_vertical2->fill->fgcolor-rgb = zcl_excel_style_color=>c_white.
    lo_style_gr_vertical2->fill->bgcolor-rgb = zcl_excel_style_color=>c_blue.

    lo_style_gr_fromcenter                   = ro_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
    lo_style_gr_fromcenter->fill->filltype   = zcl_excel_style_fill=>c_fill_gradient_fromcenter.

    lo_style_gr_diagonal45                   = ro_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
    lo_style_gr_diagonal45->fill->filltype   = zcl_excel_style_fill=>c_fill_gradient_diagonal45.

    lo_style_gr_diagonal45b                  = ro_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
    lo_style_gr_diagonal45b->fill->filltype  = zcl_excel_style_fill=>c_fill_gradient_diagonal45b.

    lo_style_gr_diagonal135                  = ro_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
    lo_style_gr_diagonal135->fill->filltype  = zcl_excel_style_fill=>c_fill_gradient_diagonal135.

    lo_style_gr_diagonal135b                 = ro_excel->add_new_style( io_clone_of = lo_style_gr_cornerlb ).
    lo_style_gr_diagonal135b->fill->filltype = zcl_excel_style_fill=>c_fill_gradient_diagonal135b.



    " Create filled style turquoise using legacy excel ver <= 2003 palette. (https://github.com/abap2xlsx/abap2xlsx/issues/92)
    lo_style_filled_turquoise                 = ro_excel->add_new_style( ).
    ro_excel->legacy_palette->set_color( "replace built-in color from palette with out custom RGB turquoise
      ip_index =     16
      ip_color =     '0040E0D0' ).

    lo_style_filled_turquoise->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
    lo_style_filled_turquoise->fill->fgcolor-indexed  = 16.

    " Get active sheet
    lo_worksheet = ro_excel->get_active_worksheet( ).
    lo_worksheet->set_title( ip_title = 'Styles' ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world' ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = 'Bold text'            ip_style = lo_style_bold ).
    lo_worksheet->set_cell( ip_column = 'D' ip_row = 4 ip_value = 'Underlined text'      ip_style = lo_style_underline ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 5 ip_value = 'Filled text'          ip_style = lo_style_filled ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 6 ip_value = 'Borders'              ip_style = lo_style_border ).
    lo_worksheet->set_cell( ip_column = 'D' ip_row = 7 ip_value = 'I''m not a button :)' ip_style = lo_style_button ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 9 ip_value = 'Modified color for Excel 2003' ip_style = lo_style_filled_turquoise ).
    " Fill the cell and apply one style
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 6 ip_value = 'Filled text'          ip_style = lo_style_filled ).
    " Change the style
    lo_worksheet->set_cell_style( ip_column = 'B' ip_row = 6 ip_style = lo_style_filled_green ).
    " Add Style to an empty cell to test Fix for Issue
    "#44 Exception ZCX_EXCEL thrown when style is set for an empty cell
    " https://github.com/abap2xlsx/abap2xlsx/issues/44
    lo_worksheet->set_cell_style( ip_column = 'E' ip_row = 6 ip_style = lo_style_filled_green ).


    lo_worksheet->set_cell( ip_column = 'B' ip_row = 10  ip_style = lo_style_gr_cornerlb ip_value = zcl_excel_style_fill=>c_fill_gradient_cornerlb ).
    lo_row = lo_worksheet->get_row( ip_row = 10 ).
    lo_row->set_row_height( ip_row_height = 30 ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 11  ip_style = lo_style_gr_cornerlt ip_value = zcl_excel_style_fill=>c_fill_gradient_cornerlt ).
    lo_row = lo_worksheet->get_row( ip_row = 11 ).
    lo_row->set_row_height( ip_row_height = 30 ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 12  ip_style = lo_style_gr_cornerrb ip_value = zcl_excel_style_fill=>c_fill_gradient_cornerrb ).
    lo_row = lo_worksheet->get_row( ip_row = 12 ).
    lo_row->set_row_height( ip_row_height = 30 ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 13  ip_style = lo_style_gr_cornerrt ip_value = zcl_excel_style_fill=>c_fill_gradient_cornerrt ).
    lo_row = lo_worksheet->get_row( ip_row = 13 ).
    lo_row->set_row_height( ip_row_height = 30 ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 14  ip_style = lo_style_gr_horizontal90 ip_value = zcl_excel_style_fill=>c_fill_gradient_horizontal90 ).
    lo_row = lo_worksheet->get_row( ip_row = 14 ).
    lo_row->set_row_height( ip_row_height = 30 ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 15  ip_style = lo_style_gr_horizontal270 ip_value = zcl_excel_style_fill=>c_fill_gradient_horizontal270 ).
    lo_row = lo_worksheet->get_row( ip_row = 15 ).
    lo_row->set_row_height( ip_row_height = 30 ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 16  ip_style = lo_style_gr_horizontalb ip_value = zcl_excel_style_fill=>c_fill_gradient_horizontalb ).
    lo_row = lo_worksheet->get_row( ip_row = 16 ).
    lo_row->set_row_height( ip_row_height = 30 ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 17  ip_style = lo_style_gr_vertical ip_value = zcl_excel_style_fill=>c_fill_gradient_vertical ).
    lo_row = lo_worksheet->get_row( ip_row = 17 ).
    lo_row->set_row_height( ip_row_height = 30 ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 18  ip_style = lo_style_gr_vertical2 ip_value = zcl_excel_style_fill=>c_fill_gradient_vertical ).
    lo_row = lo_worksheet->get_row( ip_row = 18 ).
    lo_row->set_row_height( ip_row_height = 30 ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 19  ip_style = lo_style_gr_fromcenter ip_value = zcl_excel_style_fill=>c_fill_gradient_fromcenter ).
    lo_worksheet->set_cell( ip_column = 'E' ip_row = 19  ip_style = lo_style_gr_fromcenter ip_value = 'little off fromCenter' ).
    lo_worksheet->change_cell_style( ip_column = 'E' ip_row = 19 ip_fill_filltype = zcl_excel_style_fill=>c_fill_none
                                                               ip_fill_gradtype_type = zcl_excel_style_fill=>c_fill_gradient_path
                                                               ip_fill_gradtype_position1 = '0'
                                                               ip_fill_gradtype_position2 = '1'
                                                               ip_fill_gradtype_bottom = '0.4'
                                                               ip_fill_gradtype_top = '0.3'
                                                               ip_fill_gradtype_left = '0.3'
                                                               ip_fill_gradtype_right = '0.4' ).
    lo_row = lo_worksheet->get_row( ip_row = 19 ).
    lo_row->set_row_height( ip_row_height = 30 ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 20  ip_style = lo_style_gr_diagonal45 ip_value = zcl_excel_style_fill=>c_fill_gradient_diagonal45 ).
    lo_row = lo_worksheet->get_row( ip_row = 20 ).
    lo_row->set_row_height( ip_row_height = 30 ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 21  ip_style = lo_style_gr_diagonal45b ip_value = zcl_excel_style_fill=>c_fill_gradient_diagonal45b ).
    lo_row = lo_worksheet->get_row( ip_row = 21 ).
    lo_row->set_row_height( ip_row_height = 30 ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 22  ip_style = lo_style_gr_diagonal135 ip_value = zcl_excel_style_fill=>c_fill_gradient_diagonal135 ).
    lo_row = lo_worksheet->get_row( ip_row = 22 ).
    lo_row->set_row_height( ip_row_height = 30 ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 23  ip_style = lo_style_gr_diagonal135b ip_value = zcl_excel_style_fill=>c_fill_gradient_diagonal135b ).
    lo_row = lo_worksheet->get_row( ip_row = 23 ).
    lo_row->set_row_height( ip_row_height = 30 ).

  endmethod.

  method demo1.
    data lo_worksheet type ref to zcl_excel_worksheet.
    data lo_hyperlink type ref to zcl_excel_hyperlink.
    data lo_column    type ref to zcl_excel_column.
    data lv_date      type d.
    data lv_time      type t.

    ro_excel = new #( ).

    " Get active sheet
    lo_worksheet = ro_excel->get_active_worksheet( ).
    lo_worksheet->set_cell( ip_column = 'B'
                            ip_row    = 2
                            ip_value  = 'Hello world' ).
    lv_date = '20211231'.
    lv_time = '055817'.
    lo_worksheet->set_cell( ip_column = 'B'
                            ip_row    = 3
                            ip_value  = lv_date ).
    lo_worksheet->set_cell( ip_column = 'C'
                            ip_row    = 3
                            ip_value  = lv_time ).
    lo_hyperlink = zcl_excel_hyperlink=>create_external_link( iv_url = 'https://abap2xlsx.github.io/abap2xlsx' ).
    lo_worksheet->set_cell( ip_columnrow = 'B4'
                            ip_value     = 'Click here to visit abap2xlsx homepage'
                            ip_hyperlink = lo_hyperlink ).

    lo_worksheet->set_cell( ip_column = 'B'
                            ip_row    = 6
                            ip_value  = '你好，世界' ).
    lo_worksheet->set_cell( ip_column = 'C'
                            ip_row    = 6
                            ip_value  = '(Chinese)' ).
    lo_worksheet->set_cell( ip_column = 'B'
                            ip_row    = 7
                            ip_value  = 'नमस्ते दुनिया' ).
    lo_worksheet->set_cell( ip_column = 'C'
                            ip_row    = 7
                            ip_value  = '(Hindi)' ).
    lo_worksheet->set_cell( ip_column = 'B'
                            ip_row    = 8
                            ip_value  = 'Hola Mundo' ).
    lo_worksheet->set_cell( ip_column = 'C'
                            ip_row    = 8
                            ip_value  = '(Spanish)' ).
    lo_worksheet->set_cell( ip_column = 'B'
                            ip_row    = 9
                            ip_value  = 'مرحبا بالعالم' ).
    lo_worksheet->set_cell( ip_column = 'C'
                            ip_row    = 9
                            ip_value  = '(Arabic)' ).
    lo_worksheet->set_cell( ip_column = 'B'
                            ip_row    = 10
                            ip_value  = 'ওহে বিশ্ব ' ).
    lo_worksheet->set_cell( ip_column = 'C'
                            ip_row    = 10
                            ip_value  = '(Bengali)' ).
    lo_worksheet->set_cell( ip_column = 'B'
                            ip_row    = 11
                            ip_value  = 'Bonjour le monde' ).
    lo_worksheet->set_cell( ip_column = 'C'
                            ip_row    = 11
                            ip_value  = '(French)' ).
    lo_worksheet->set_cell( ip_column = 'B'
                            ip_row    = 12
                            ip_value  = 'Olá Mundo' ).
    lo_worksheet->set_cell( ip_column = 'C'
                            ip_row    = 12
                            ip_value  = '(Portuguese)' ).
    lo_worksheet->set_cell( ip_column = 'B'
                            ip_row    = 13
                            ip_value  = 'Привет, мир' ).
    lo_worksheet->set_cell( ip_column = 'C'
                            ip_row    = 13
                            ip_value  = '(Russian)' ).
    lo_worksheet->set_cell( ip_column = 'B'
                            ip_row    = 14
                            ip_value  = 'ہیلو دنیا' ).
    lo_worksheet->set_cell( ip_column = 'C'
                            ip_row    = 14
                            ip_value  = '(Urdu)' ).
    lo_worksheet->set_cell( ip_column = 'B'
                            ip_row    = 15
                            ip_value  = '👋🌎, 👋🌍, 👋🌏' ).
    lo_worksheet->set_cell( ip_column = 'C'
                            ip_row    = 15
                            ip_value  = '(Emoji waving hand + 3 parts of the world)' ).

    lo_column = lo_worksheet->get_column( ip_column = 'B' ).
    lo_column->set_width( ip_width = 11 ).

  endmethod.
endclass.
