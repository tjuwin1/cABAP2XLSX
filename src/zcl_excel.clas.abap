class zcl_excel definition
  public
  create public .

  public section.

*"* public components of class ZCL_EXCEL
*"* do not include other source files here!!!
    interfaces zif_excel_book_properties .
    interfaces zif_excel_book_protection .
    interfaces zif_excel_book_vba_project .

    data legacy_palette type ref to zcl_excel_legacy_palette read-only .
    data security type ref to zcl_excel_security .
    data use_template type abap_bool .
    constants version type c length 10 value '7.16.0'.      "#EC NOTEXT

    methods add_new_autofilter
      importing
        !io_sheet            type ref to zcl_excel_worksheet
      returning
        value(ro_autofilter) type ref to zcl_excel_autofilter
      raising
        zcx_excel .
    methods add_new_comment
      returning
        value(eo_comment) type ref to zcl_excel_comment .
    methods add_new_drawing
      importing
        !ip_type          type zif_excel_data_decl=>zexcel_drawing_type default zcl_excel_drawing=>type_image
        !ip_title         type clike optional
      returning
        value(eo_drawing) type ref to zcl_excel_drawing .
    methods add_new_range
      returning
        value(eo_range) type ref to zcl_excel_range .
    methods add_new_style
      importing
        !ip_guid        type zif_excel_data_decl=>zexcel_cell_style optional
        !io_clone_of    type ref to zcl_excel_style optional
          preferred parameter !ip_guid
      returning
        value(eo_style) type ref to zcl_excel_style .
    methods add_new_worksheet
      importing
        !ip_title           type zif_excel_data_decl=>zexcel_sheet_title optional
      returning
        value(eo_worksheet) type ref to zcl_excel_worksheet
      raising
        zcx_excel .
    methods add_static_styles .
    methods constructor .
    methods delete_worksheet
      importing
        !io_worksheet type ref to zcl_excel_worksheet
      raising
        zcx_excel .
    methods delete_worksheet_by_index
      importing
        !iv_index type numeric
      raising
        zcx_excel .
    methods delete_worksheet_by_name
      importing
        !iv_title type clike
      raising
        zcx_excel .
    methods get_active_sheet_index
      returning
        value(r_active_worksheet) type zif_excel_data_decl=>zexcel_active_worksheet .
    methods get_active_worksheet
      returning
        value(eo_worksheet) type ref to zcl_excel_worksheet .
    methods get_autofilters_reference
      returning
        value(ro_autofilters) type ref to zcl_excel_autofilters .
    methods get_default_style
      returning
        value(ep_style) type zif_excel_data_decl=>zexcel_cell_style .
    methods get_drawings_iterator
      importing
        !ip_type           type zif_excel_data_decl=>zexcel_drawing_type
      returning
        value(eo_iterator) type ref to zcl_excel_collection_iterator .
    methods get_next_table_id
      returning
        value(ep_id) type i .
    methods get_ranges_iterator
      returning
        value(eo_iterator) type ref to zcl_excel_collection_iterator .
    methods get_static_cellstyle_guid
      importing
        !ip_cstyle_complete  type zif_excel_data_decl=>zexcel_s_cstyle_complete
        !ip_cstylex_complete type zif_excel_data_decl=>zexcel_s_cstylex_complete
      returning
        value(ep_guid)       type zif_excel_data_decl=>zexcel_cell_style .
    methods get_styles_iterator
      returning
        value(eo_iterator) type ref to zcl_excel_collection_iterator .
    methods get_style_from_guid
      importing
        !ip_guid        type zif_excel_data_decl=>zexcel_cell_style
      returning
        value(eo_style) type ref to zcl_excel_style .
    methods get_style_index_in_styles
      importing
        !ip_guid        type zif_excel_data_decl=>zexcel_cell_style
      returning
        value(ep_index) type i
      raising
        zcx_excel .
    methods get_style_to_guid
      importing
        !ip_guid               type zif_excel_data_decl=>zexcel_cell_style
      returning
        value(ep_stylemapping) type zif_excel_data_decl=>zexcel_s_stylemapping
      raising
        zcx_excel .
    methods get_theme
      exporting
        !eo_theme type ref to zcl_excel_theme .
    methods get_worksheets_iterator
      returning
        value(eo_iterator) type ref to zcl_excel_collection_iterator .
    methods get_worksheets_name
      returning
        value(ep_name) type zif_excel_data_decl=>zexcel_worksheets_name .
    methods get_worksheets_size
      returning
        value(ep_size) type i .
    methods get_worksheet_by_index
      importing
        !iv_index           type numeric
      returning
        value(eo_worksheet) type ref to zcl_excel_worksheet .
    methods get_worksheet_by_name
      importing
        !ip_sheet_name      type zif_excel_data_decl=>zexcel_sheet_title
      returning
        value(eo_worksheet) type ref to zcl_excel_worksheet .
    methods set_active_sheet_index
      importing
        !i_active_worksheet type zif_excel_data_decl=>zexcel_active_worksheet
      raising
        zcx_excel .
    methods set_active_sheet_index_by_name
      importing
        !i_worksheet_name type zif_excel_data_decl=>zexcel_worksheets_name .
    methods set_default_style
      importing
        !ip_style type zif_excel_data_decl=>zexcel_cell_style
      raising
        zcx_excel .
    methods set_theme
      importing
        !io_theme type ref to zcl_excel_theme .
    methods fill_template
      importing
        !iv_data type ref to zcl_excel_template_data
      raising
        zcx_excel .
  protected section.

    data worksheets type ref to zcl_excel_worksheets .
  private section.

    data autofilters type ref to zcl_excel_autofilters .
    data charts type ref to zcl_excel_drawings .
    data default_style type zif_excel_data_decl=>zexcel_cell_style .
*"* private components of class ZCL_EXCEL
*"* do not include other source files here!!!
    data drawings type ref to zcl_excel_drawings .
    data ranges type ref to zcl_excel_ranges .
    data styles type ref to zcl_excel_styles .
    data t_stylemapping1 type zif_excel_data_decl=>zexcel_t_stylemapping1 .
    data t_stylemapping2 type zif_excel_data_decl=>zexcel_t_stylemapping2 .
    data theme type ref to zcl_excel_theme .
    data comments type ref to zcl_excel_comments .

    methods stylemapping_dynamic_style
      importing
        !ip_style        type ref to zcl_excel_style
      returning
        value(eo_style2) type zif_excel_data_decl=>zexcel_s_stylemapping .
endclass.



class zcl_excel implementation.


  method add_new_autofilter.
* Check for autofilter reference: new or overwrite; only one per sheet
    ro_autofilter = autofilters->add( io_sheet ) .
  endmethod.


  method add_new_comment.
    create object eo_comment.

    comments->add( eo_comment ).
  endmethod.


  method add_new_drawing.
* Create default blank worksheet
    create object eo_drawing
      exporting
        ip_type  = ip_type
        ip_title = ip_title.

    case ip_type.
      when 'image'.
        drawings->add( eo_drawing ).
      when 'hd_ft'.
        drawings->add( eo_drawing ).
      when 'chart'.
        charts->add( eo_drawing ).
    endcase.
  endmethod.


  method add_new_range.
* Create default blank range
    create object eo_range.
    ranges->add( eo_range ).
  endmethod.


  method add_new_style.
* Start of deletion # issue 139 - Dateretention of cellstyles
*  CREATE OBJECT eo_style.
*  styles->add( eo_style ).
* End of deletion # issue 139 - Dateretention of cellstyles
* Start of insertion # issue 139 - Dateretention of cellstyles
* Create default style
    create object eo_style
      exporting
        ip_guid     = ip_guid
        io_clone_of = io_clone_of.
    styles->add( eo_style ).

    data: style2 type zif_excel_data_decl=>zexcel_s_stylemapping.
* Copy to new representations
    style2 = stylemapping_dynamic_style( eo_style ).
    insert style2 into table t_stylemapping1.
    insert style2 into table t_stylemapping2.
* End of insertion # issue 139 - Dateretention of cellstyles

  endmethod.


  method add_new_worksheet.

* Create default blank worksheet
    create object eo_worksheet
      exporting
        ip_excel = me
        ip_title = ip_title.

    worksheets->add( eo_worksheet ).
    worksheets->active_worksheet = worksheets->size( ).
  endmethod.


  method add_static_styles.
    " # issue 139
    field-symbols: <style1> like line of t_stylemapping1,
                   <style2> like line of t_stylemapping2.
    data: style type ref to zcl_excel_style.
    LOOP AT me->t_stylemapping1 ASSIGNING <style1> USING KEY added_to_iterator
        WHERE added_to_iterator = abap_false.
      READ TABLE me->t_stylemapping2 ASSIGNING <style2> WITH TABLE KEY guid = <style1>-guid.
      CHECK sy-subrc = 0.  " Should always be true since these tables are being filled parallel

      style = me->add_new_style( <style1>-guid ).

      zcl_excel_common=>recursive_struct_to_class( EXPORTING i_source  = <style1>-complete_style
                                                             i_sourcex = <style1>-complete_stylex
                                                   CHANGING  e_target  = style ).

    ENDLOOP.
  endmethod.


  method constructor.
    data: lo_style      type ref to zcl_excel_style.

* Inizialize instance objects
    create object security.
    create object worksheets.
    create object ranges.
    create object styles.
    create object drawings
      exporting
        ip_type = zcl_excel_drawing=>type_image.
    create object charts
      exporting
        ip_type = zcl_excel_drawing=>type_chart.
    create object comments.
    create object legacy_palette.
    create object autofilters.

    me->zif_excel_book_protection~initialize( ).
    me->zif_excel_book_properties~initialize( ).

    try.
        me->add_new_worksheet( ).

        lo_style = me->add_new_style( ). " Standard style
        me->set_default_style( lo_style->get_guid(  ) ).

        lo_style = me->add_new_style( ). " Standard style with fill gray125
        lo_style->fill->filltype = zcl_excel_style_fill=>c_fill_pattern_gray125.
      catch zcx_excel. " suppress syntax check error
        assert 1 = 2.  " some error processing anyway
    endtry.
  endmethod.


  method delete_worksheet.

    data: lo_worksheet    type ref to zcl_excel_worksheet,
          l_size          type i,
          lv_errormessage type string.

    l_size = get_worksheets_size( ).
    if l_size = 1.  " Only 1 worksheet left --> check whether this is the worksheet to be deleted
      lo_worksheet = me->get_worksheet_by_index( 1 ).
      if lo_worksheet = io_worksheet.
        lv_errormessage = 'Deleting last remaining worksheet is not allowed'(002).
        zcx_excel=>raise_text( lv_errormessage ).
      endif.
    endif.

    me->worksheets->remove( io_worksheet ).

  endmethod.


  method delete_worksheet_by_index.

    data: lo_worksheet    type ref to zcl_excel_worksheet,
          lv_errormessage type string.

    lo_worksheet = me->get_worksheet_by_index( iv_index ).
    if lo_worksheet is not bound.
      lv_errormessage = 'Worksheet not existing'(001).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.
    me->delete_worksheet( lo_worksheet ).

  endmethod.


  method delete_worksheet_by_name.

    data: lo_worksheet    type ref to zcl_excel_worksheet,
          lv_errormessage type string.

    lo_worksheet = me->get_worksheet_by_name( iv_title ).
    if lo_worksheet is not bound.
      lv_errormessage = 'Worksheet not existing'(001).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.
    me->delete_worksheet( lo_worksheet ).

  endmethod.


  method fill_template.

    data: lo_template_filler type ref to zcl_excel_fill_template.

    field-symbols:
      <lv_sheet>     type zif_excel_data_decl=>zexcel_sheet_title,
      <lv_data_line> type zcl_excel_template_data=>ts_template_data_sheet.


    lo_template_filler = zcl_excel_fill_template=>create( me ).

    loop at lo_template_filler->mt_sheet assigning <lv_sheet>.

      read table iv_data->mt_data assigning <lv_data_line> with key sheet = <lv_sheet>.
      check sy-subrc = 0.
      lo_template_filler->fill_sheet( <lv_data_line> ).

    endloop.

  endmethod.


  method get_active_sheet_index.
    r_active_worksheet = me->worksheets->active_worksheet.
  endmethod.


  method get_active_worksheet.

    eo_worksheet = me->worksheets->get( me->worksheets->active_worksheet ).

  endmethod.


  method get_autofilters_reference.

    ro_autofilters = autofilters.

  endmethod.


  method get_default_style.
    ep_style = me->default_style.
  endmethod.


  method get_drawings_iterator.

    case ip_type.
      when zcl_excel_drawing=>type_image.
        eo_iterator = me->drawings->get_iterator( ).
      when zcl_excel_drawing=>type_chart.
        eo_iterator = me->charts->get_iterator( ).
      when others.
    endcase.

  endmethod.


  method get_next_table_id.
    data: lo_worksheet    type ref to zcl_excel_worksheet,
          lo_iterator     type ref to zcl_excel_collection_iterator,
          lv_tables_count type i.

    lo_iterator = me->get_worksheets_iterator( ).
    while lo_iterator->has_next( ) eq abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).

      lv_tables_count = lo_worksheet->get_tables_size( ).
      add lv_tables_count to ep_id.

    endwhile.

    add 1 to ep_id.

  endmethod.


  method get_ranges_iterator.

    eo_iterator = me->ranges->get_iterator( ).

  endmethod.


  method get_static_cellstyle_guid.
    " # issue 139
    data: style like line of me->t_stylemapping1.

    read table me->t_stylemapping1 into style
      with table key dynamic_style_guid = style-guid  " no dynamic style  --> look for initial guid here
                     complete_style     = ip_cstyle_complete
                     complete_stylex    = ip_cstylex_complete.
    if sy-subrc <> 0.
      style-complete_style  = ip_cstyle_complete.
      style-complete_stylex = ip_cstylex_complete.
      style-guid = zcl_excel_obsolete_func_wrap=>guid_create( ). " ins issue #379 - replacement for outdated function call
      insert style into table me->t_stylemapping1.
      insert style into table me->t_stylemapping2.

    endif.

    ep_guid = style-guid.
  endmethod.


  method get_styles_iterator.

    eo_iterator = me->styles->get_iterator( ).

  endmethod.


  method get_style_from_guid.

    data: lo_style    type ref to zcl_excel_style,
          lo_iterator type ref to zcl_excel_collection_iterator.

    lo_iterator = styles->get_iterator( ).
    while lo_iterator->has_next( ) = abap_true.
      lo_style ?= lo_iterator->get_next( ).
      if lo_style->get_guid( ) = ip_guid.
        eo_style = lo_style.
        return.
      endif.
    endwhile.

  endmethod.


  method get_style_index_in_styles.
    data: index type i.
    data: lo_iterator type ref to zcl_excel_collection_iterator,
          lo_style    type ref to zcl_excel_style.

    check ip_guid is not initial.


    lo_iterator = me->get_styles_iterator( ).
    while lo_iterator->has_next( ) = 'X'.
      add 1 to index.
      lo_style ?= lo_iterator->get_next( ).
      if lo_style->get_guid( ) = ip_guid.
        ep_index = index.
        exit.
      endif.
    endwhile.

    if ep_index is initial.
      zcx_excel=>raise_text( 'Index not found' ).
    else.
      subtract 1 from ep_index.  " In excel list starts with "0"
    endif.
  endmethod.


  method get_style_to_guid.
    data: lo_style type ref to zcl_excel_style.
    " # issue 139
    read table me->t_stylemapping2 into ep_stylemapping with table key guid = ip_guid.
    if sy-subrc <> 0.
      zcx_excel=>raise_text( 'GUID not found' ).
    endif.

    if ep_stylemapping-dynamic_style_guid is not initial.
      lo_style = me->get_style_from_guid( ip_guid ).
      zcl_excel_common=>recursive_class_to_struct( exporting i_source = lo_style
                                                   changing  e_target =  ep_stylemapping-complete_style
                                                             e_targetx = ep_stylemapping-complete_stylex ).
    endif.
  endmethod.


  method get_theme.
    eo_theme = theme.
  endmethod.


  method get_worksheets_iterator.

    eo_iterator = me->worksheets->get_iterator( ).

  endmethod.


  method get_worksheets_name.

    ep_name = me->worksheets->name.

  endmethod.


  method get_worksheets_size.

    ep_size = me->worksheets->size( ).

  endmethod.


  method get_worksheet_by_index.


    data: lv_index type zif_excel_data_decl=>zexcel_active_worksheet.

    lv_index = iv_index.
    eo_worksheet = me->worksheets->get( lv_index ).

  endmethod.


  method get_worksheet_by_name.

    data: lv_index type zif_excel_data_decl=>zexcel_active_worksheet,
          l_size   type i.

    l_size = get_worksheets_size( ).

    do l_size times.
      lv_index = sy-index.
      eo_worksheet = me->worksheets->get( lv_index ).
      if eo_worksheet->get_title( ) = ip_sheet_name.
        return.
      endif.
    enddo.

    clear eo_worksheet.

  endmethod.


  method set_active_sheet_index.
    data: lo_worksheet    type ref to zcl_excel_worksheet,
          lv_errormessage type string.

*--------------------------------------------------------------------*
* Check whether worksheet exists
*--------------------------------------------------------------------*
    lo_worksheet = me->get_worksheet_by_index( i_active_worksheet ).
    if lo_worksheet is not bound.
      lv_errormessage = 'Worksheet not existing'(001).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

    me->worksheets->active_worksheet = i_active_worksheet.

  endmethod.


  method set_active_sheet_index_by_name.

    data: ws_it    type ref to zcl_excel_collection_iterator,
          ws       type ref to zcl_excel_worksheet,
          lv_title type zif_excel_data_decl=>zexcel_sheet_title,
          count    type i value 1.

    ws_it = me->worksheets->get_iterator( ).

    while ws_it->has_next( ) = abap_true.
      ws ?= ws_it->get_next( ).
      lv_title = ws->get_title( ).
      if lv_title = i_worksheet_name.
        me->worksheets->active_worksheet = count.
        exit.
      endif.
      count = count + 1.
    endwhile.

  endmethod.


  method set_default_style.
    me->default_style = ip_style.
  endmethod.


  method set_theme.
    theme = io_theme.
  endmethod.


  method stylemapping_dynamic_style.
    " # issue 139
    eo_style2-dynamic_style_guid  = ip_style->get_guid( ).
    eo_style2-guid                = eo_style2-dynamic_style_guid.
    eo_style2-added_to_iterator   = abap_true.

* don't care about attributes here, since this data may change
* dynamically

  endmethod.


  method zif_excel_book_properties~initialize.
    data: lv_timestamp type timestampl.

    me->zif_excel_book_properties~application     = 'Microsoft Excel'.
    me->zif_excel_book_properties~appversion      = '12.0000'.

    get time stamp field lv_timestamp.
    me->zif_excel_book_properties~created         = lv_timestamp.
    me->zif_excel_book_properties~creator         = sy-uname.
    me->zif_excel_book_properties~description     = zcl_excel=>version.
    me->zif_excel_book_properties~modified        = lv_timestamp.
    me->zif_excel_book_properties~lastmodifiedby  = sy-uname.
  endmethod.


  method zif_excel_book_protection~initialize.
    me->zif_excel_book_protection~protected      = zif_excel_book_protection=>c_unprotected.
    me->zif_excel_book_protection~lockrevision   = zif_excel_book_protection=>c_unlocked.
    me->zif_excel_book_protection~lockstructure  = zif_excel_book_protection=>c_unlocked.
    me->zif_excel_book_protection~lockwindows    = zif_excel_book_protection=>c_unlocked.
    clear me->zif_excel_book_protection~workbookpassword.
    clear me->zif_excel_book_protection~revisionspassword.
  endmethod.


  method zif_excel_book_vba_project~set_codename.
    me->zif_excel_book_vba_project~codename = ip_codename.
  endmethod.


  method zif_excel_book_vba_project~set_codename_pr.
    me->zif_excel_book_vba_project~codename_pr = ip_codename_pr.
  endmethod.


  method zif_excel_book_vba_project~set_vbaproject.
    me->zif_excel_book_vba_project~vbaproject = ip_vbaproject.
  endmethod.
endclass.
