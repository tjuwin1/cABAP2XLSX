class zcl_excel_writer_2007 definition
  public
  create public .

  public section.

*"* public components of class ZCL_EXCEL_WRITER_2007
*"* do not include other source files here!!!
    interfaces zif_excel_writer .
    methods constructor.

  protected section.

*"* protected components of class ZCL_EXCEL_WRITER_2007
*"* do not include other source files here!!!
    types: begin of mty_column_formula_used,
             id type zif_excel_data_decl=>zexcel_s_cell_data-column_formula_id,
             si type string,
             "! type: shared, etc.
             t  type string,
           end of mty_column_formula_used,
           mty_column_formulas_used type hashed table of mty_column_formula_used with unique key id.
    constants c_content_types type string value '[Content_Types].xml'. "#EC NOTEXT
    constants c_docprops_app type string value 'docProps/app.xml'. "#EC NOTEXT
    constants c_docprops_core type string value 'docProps/core.xml'. "#EC NOTEXT
    constants c_relationships type string value '_rels/.rels'. "#EC NOTEXT
    constants c_xl_calcchain type string value 'xl/calcChain.xml'. "#EC NOTEXT
    constants c_xl_drawings type string value 'xl/drawings/drawing#.xml'. "#EC NOTEXT
    constants c_xl_drawings_rels type string value 'xl/drawings/_rels/drawing#.xml.rels'. "#EC NOTEXT
    constants c_xl_relationships type string value 'xl/_rels/workbook.xml.rels'. "#EC NOTEXT
    constants c_xl_sharedstrings type string value 'xl/sharedStrings.xml'. "#EC NOTEXT
    constants c_xl_sheet type string value 'xl/worksheets/sheet#.xml'. "#EC NOTEXT
    constants c_xl_sheet_rels type string value 'xl/worksheets/_rels/sheet#.xml.rels'. "#EC NOTEXT
    constants c_xl_styles type string value 'xl/styles.xml'. "#EC NOTEXT
    constants c_xl_theme type string value 'xl/theme/theme1.xml'. "#EC NOTEXT
    constants c_xl_workbook type string value 'xl/workbook.xml'. "#EC NOTEXT
    data excel type ref to zcl_excel .
    data shared_strings type zif_excel_data_decl=>zexcel_t_shared_string .
    data styles_cond_mapping type zif_excel_data_decl=>zexcel_t_styles_cond_mapping .
    data styles_mapping type zif_excel_data_decl=>zexcel_t_styles_mapping .
    constants c_xl_comments type string value 'xl/comments#.xml'. "#EC NOTEXT
    constants cl_xl_drawing_for_comments type string value 'xl/drawings/vmlDrawing#.vml'. "#EC NOTEXT
    constants c_xl_drawings_vml_rels type string value 'xl/drawings/_rels/vmlDrawing#.vml.rels'. "#EC NOTEXT
    data ixml type ref to cl_ixml_core.
    data control_characters type string.

    methods create_xl_sheet_sheet_data
      importing
        !io_document                   type ref to if_ixml_document
        !io_worksheet                  type ref to zcl_excel_worksheet
      returning
        value(rv_ixml_sheet_data_root) type ref to if_ixml_element
      raising
        zcx_excel .
    methods add_further_data_to_zip
      importing
        !io_zip type ref to cl_abap_zip .
    methods create
      returning
        value(ep_excel) type xstring
      raising
        zcx_excel .
    methods create_content_types
      returning
        value(ep_content) type xstring .
    methods create_docprops_app
      returning
        value(ep_content) type xstring .
    methods create_docprops_core
      returning
        value(ep_content) type xstring .
    methods create_dxf_style
      importing
        !iv_cell_style    type zif_excel_data_decl=>zexcel_cell_style
        !io_dxf_element   type ref to if_ixml_element
        !io_ixml_document type ref to if_ixml_document
        !it_cellxfs       type zif_excel_data_decl=>zexcel_t_cellxfs
        !it_fonts         type zif_excel_data_decl=>zexcel_t_style_font
        !it_fills         type zif_excel_data_decl=>zexcel_t_style_fill
      changing
        !cv_dfx_count     type i .
    methods create_relationships
      returning
        value(ep_content) type xstring .
    methods create_xl_charts
      importing
        !io_drawing       type ref to zcl_excel_drawing
      returning
        value(ep_content) type xstring .
    methods create_xl_comments
      importing
        !io_worksheet     type ref to zcl_excel_worksheet
      returning
        value(ep_content) type xstring .
    methods create_xl_drawings
      importing
        !io_worksheet     type ref to zcl_excel_worksheet
      returning
        value(ep_content) type xstring .
    methods create_xl_drawings_rels
      importing
        !io_worksheet     type ref to zcl_excel_worksheet
      returning
        value(ep_content) type xstring .
    methods create_xl_drawing_anchor
      importing
        !io_drawing      type ref to zcl_excel_drawing
        !io_document     type ref to if_ixml_document
        !ip_index        type i
      returning
        value(ep_anchor) type ref to if_ixml_element .
    methods create_xl_drawing_for_comments
      importing
        !io_worksheet     type ref to zcl_excel_worksheet
      returning
        value(ep_content) type xstring
      raising
        zcx_excel .
    methods create_xl_relationships
      returning
        value(ep_content) type xstring .
    methods create_xl_sharedstrings
      returning
        value(ep_content) type xstring .
    methods create_xl_sheet
      importing
        !io_worksheet     type ref to zcl_excel_worksheet
        !iv_active        type abap_boolean default ''
      returning
        value(ep_content) type xstring
      raising
        zcx_excel .
    methods create_xl_sheet_ignored_errors
      importing
        io_worksheet    type ref to zcl_excel_worksheet
        io_document     type ref to if_ixml_document
        io_element_root type ref to if_ixml_element.
    methods create_xl_sheet_pagebreaks
      importing
        !io_document  type ref to if_ixml_document
        !io_parent    type ref to if_ixml_element
        !io_worksheet type ref to zcl_excel_worksheet
      raising
        zcx_excel .
    methods create_xl_sheet_rels
      importing
        !io_worksheet     type ref to zcl_excel_worksheet
        !iv_drawing_index type i optional
        !iv_comment_index type i optional
        !iv_cmnt_vmlindex type i optional
        !iv_hdft_vmlindex type i optional
      returning
        value(ep_content) type xstring .
    methods create_xl_styles
      returning
        value(ep_content) type xstring .
    methods create_xl_styles_color_node
      importing
        !io_document        type ref to if_ixml_document
        !io_parent          type ref to if_ixml_element
        !iv_color_elem_name type string default 'color'
        !is_color           type zif_excel_data_decl=>zexcel_s_style_color .
    methods create_xl_styles_font_node
      importing
        !io_document type ref to if_ixml_document
        !io_parent   type ref to if_ixml_element
        !is_font     type zif_excel_data_decl=>zexcel_s_style_font
        !iv_use_rtf  type abap_bool default abap_false .
    methods create_xl_table
      importing
        !io_table         type ref to zcl_excel_table
      returning
        value(ep_content) type xstring
      raising
        zcx_excel .
    methods create_xl_theme
      returning
        value(ep_content) type xstring .
    methods create_xl_workbook
      returning
        value(ep_content) type xstring
      raising
        zcx_excel .
    methods get_shared_string_index
      importing
        !ip_cell_value  type zif_excel_data_decl=>zexcel_cell_value
        !it_rtf         type zif_excel_data_decl=>zexcel_t_rtf optional
      returning
        value(ep_index) type int4 .
    methods create_xl_drawings_vml
      returning
        value(ep_content) type xstring .
    methods set_vml_string
      returning
        value(ep_content) type string .
    methods create_xl_drawings_vml_rels
      returning
        value(ep_content) type xstring .
    methods escape_string_value
      importing
        !iv_value     type zif_excel_data_decl=>zexcel_cell_value
      returning
        value(result) type zif_excel_data_decl=>zexcel_cell_value.
    methods set_vml_shape_footer
      importing
        !is_footer        type zif_excel_data_decl=>zexcel_s_worksheet_head_foot
      returning
        value(ep_content) type string .
    methods set_vml_shape_header
      importing
        !is_header        type zif_excel_data_decl=>zexcel_s_worksheet_head_foot
      returning
        value(ep_content) type string .
    methods create_xl_drawing_for_hdft_im
      importing
        !io_worksheet     type ref to zcl_excel_worksheet
      returning
        value(ep_content) type xstring .
    methods create_xl_drawings_hdft_rels
      importing
        !io_worksheet     type ref to zcl_excel_worksheet
      returning
        value(ep_content) type xstring .
    methods create_xml_document
      returning
        value(ro_document) type ref to if_ixml_document.
    methods render_xml_document
      importing
        io_document       type ref to if_ixml_document
      returning
        value(ep_content) type xstring.
    methods create_xl_sheet_column_formula
      importing
        io_document             type ref to if_ixml_document
        it_column_formulas      type zcl_excel_worksheet=>mty_th_column_formula
        is_sheet_content        type zif_excel_data_decl=>zexcel_s_cell_data
      exporting
        eo_element              type ref to if_ixml_element
      changing
        ct_column_formulas_used type mty_column_formulas_used
        cv_si                   type i
      raising
        zcx_excel.
    methods is_formula_shareable
      importing
        ip_formula          type string
      returning
        value(ep_shareable) type abap_bool
      raising
        zcx_excel.
  private section.

*"* private components of class ZCL_EXCEL_WRITER_2007
*"* do not include other source files here!!!
    constants c_off type string value '0'.                  "#EC NOTEXT
    constants c_on type string value '1'.                   "#EC NOTEXT
    constants c_xl_printersettings type string value 'xl/printerSettings/printerSettings#.bin'. "#EC NOTEXT
    types: tv_charbool type c length 5.

    methods add_1_val_child_node
      importing
        io_document   type ref to if_ixml_document
        io_parent     type ref to if_ixml_element
        iv_elem_name  type string
        iv_attr_name  type string
        iv_attr_value type string.
    methods flag2bool
      importing
        !ip_flag          type abap_boolean
      returning
        value(ep_boolean) type tv_charbool  .
    methods number2string
      importing
        !ip_number       type numeric
      returning
        value(ep_string) type string.
endclass.



class zcl_excel_writer_2007 implementation.


  method add_1_val_child_node.

    data: lo_child type ref to if_ixml_element.

    lo_child = io_document->create_simple_element( name   = iv_elem_name
                                                   parent = io_document ).
    if iv_attr_name is not initial.
      lo_child->set_attribute_ns( name  = iv_attr_name
                                  value = iv_attr_value ).
    endif.
    io_parent->append_child( new_child = lo_child ).

  endmethod.


  method add_further_data_to_zip.
* Can be used by child classes like xlsm-writer to write additional data to zip archive
  endmethod.


  method constructor.
    data: lt_unicode_point_codes type table of string,
          lv_unicode_point_code  type i.

    me->ixml ?= cl_ixml_core=>create( ).

    split '0,1,2,3,4,5,6,7,8,' " U+0000 to U+0008
       && '11,12,'             " U+000B, U+000C
       && '14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,' " U+000E to U+001F
       && '65534,65535'        " U+FFFE, U+FFFF
      at ',' into table lt_unicode_point_codes.
    control_characters = ``.
    loop at lt_unicode_point_codes into lv_unicode_point_code.
      "@TODO: To be fixed by Juwin
      "control_characters = control_characters && cl_abap_conv_in_ce=>uccpi( lv_unicode_point_code ).
    endloop.

  endmethod.


  method create.

* Office 2007 file format is a cab of several xml files with extension .xlsx

    data: lo_zip              type ref to cl_abap_zip,
          lo_worksheet        type ref to zcl_excel_worksheet,
          lo_active_worksheet type ref to zcl_excel_worksheet,
          lo_iterator         type ref to zcl_excel_collection_iterator,
          lo_nested_iterator  type ref to zcl_excel_collection_iterator,
          lo_table            type ref to zcl_excel_table,
          lo_drawing          type ref to zcl_excel_drawing.

    data: lv_content                type xstring,
          lv_active                 type abap_boolean,
          lv_xl_sheet               type string,
          lv_xl_sheet_rels          type string,
          lv_xl_drawing_for_comment type string,   " (+) Issue #180
          lv_xl_comment             type string,   " (+) Issue #180
          lv_xl_drawing             type string,
          lv_xl_drawing_rels        type string,
          lv_index_str              type string,
          lv_value                  type string,
          lv_sheet_index            type i,
          lv_drawing_counter        type i,
          lv_comment_counter        type i,
          lv_vml_counter            type i,
          lv_drawing_index          type i,
          lv_comment_index          type i,        " (+) Issue #180
          lv_cmnt_vmlindex          type i,
          lv_hdft_vmlindex          type i.

**********************************************************************

**********************************************************************
* Start of insertion # issue 139 - Dateretention of cellstyles
    me->excel->add_static_styles( ).
* End of insertion # issue 139 - Dateretention of cellstyles

**********************************************************************
* STEP 1: Create archive object file (ZIP)
    create object lo_zip.

**********************************************************************
* STEP 2: Add [Content_Types].xml to zip
    lv_content = me->create_content_types( ).
    lo_zip->add( name    = me->c_content_types
                 content = lv_content ).

**********************************************************************
* STEP 3: Add _rels/.rels to zip
    lv_content = me->create_relationships( ).
    lo_zip->add( name    = me->c_relationships
                 content = lv_content ).

**********************************************************************
* STEP 4: Add docProps/app.xml to zip
    lv_content = me->create_docprops_app( ).
    lo_zip->add( name    = me->c_docprops_app
                 content = lv_content ).

**********************************************************************
* STEP 5: Add docProps/core.xml to zip
    lv_content = me->create_docprops_core( ).
    lo_zip->add( name    = me->c_docprops_core
                 content = lv_content ).

**********************************************************************
* STEP 6: Add xl/_rels/workbook.xml.rels to zip
    lv_content = me->create_xl_relationships( ).
    lo_zip->add( name    = me->c_xl_relationships
                 content = lv_content ).

**********************************************************************
* STEP 6: Add xl/_rels/workbook.xml.rels to zip
    lv_content = me->create_xl_theme( ).
    lo_zip->add( name    = me->c_xl_theme
                 content = lv_content ).

**********************************************************************
* STEP 7: Add xl/workbook.xml to zip
    lv_content = me->create_xl_workbook( ).
    lo_zip->add( name    = me->c_xl_workbook
                 content = lv_content ).

**********************************************************************
* STEP 8: Add xl/workbook.xml to zip
    lv_content = me->create_xl_styles( ).
    lo_zip->add( name    = me->c_xl_styles
                 content = lv_content ).

**********************************************************************
* STEP 9: Add sharedStrings.xml to zip
    lv_content = me->create_xl_sharedstrings( ).
    lo_zip->add( name    = me->c_xl_sharedstrings
                 content = lv_content ).

**********************************************************************
* STEP 10: Add sheet#.xml and drawing#.xml to zip
    lo_iterator = me->excel->get_worksheets_iterator( ).
    lo_active_worksheet = me->excel->get_active_worksheet( ).

    while lo_iterator->has_next( ) eq abap_true.
      lv_sheet_index = sy-index.

      lo_worksheet ?= lo_iterator->get_next( ).
      if lo_active_worksheet->get_guid( ) eq lo_worksheet->get_guid( ).
        lv_active = abap_true.
      else.
        lv_active = abap_false.
      endif.
      lv_content = me->create_xl_sheet( io_worksheet = lo_worksheet
                                        iv_active    = lv_active ).
      lv_xl_sheet = me->c_xl_sheet.

      lv_index_str = lv_sheet_index.
      condense lv_index_str no-gaps.
      replace all occurrences of '#' in lv_xl_sheet with lv_index_str.
      lo_zip->add( name    = lv_xl_sheet
                   content = lv_content ).

* Begin - Add - Issue #180
* Add comments **********************************
      if lo_worksheet->get_comments( )->is_empty( ) = abap_false.
        " Create comment itself
        lv_comment_counter += 1.
        lv_comment_index = lv_comment_counter.
        lv_index_str = lv_comment_index.
        condense lv_index_str no-gaps.

        lv_content = me->create_xl_comments( lo_worksheet ).
        lv_xl_comment = me->c_xl_comments.
        replace all occurrences of '#' in lv_xl_comment with lv_index_str.
        lo_zip->add( name    = lv_xl_comment
                     content = lv_content ).

        " Create vmlDrawing that will host the comment
        lv_vml_counter += 1.
        lv_cmnt_vmlindex = lv_vml_counter.
        lv_index_str = lv_cmnt_vmlindex.
        condense lv_index_str no-gaps.

        lv_content = me->create_xl_drawing_for_comments( lo_worksheet ).
        lv_xl_drawing_for_comment = me->cl_xl_drawing_for_comments.
        replace all occurrences of '#' in lv_xl_drawing_for_comment with lv_index_str.
        lo_zip->add( name    = lv_xl_drawing_for_comment
                     content = lv_content ).
      else.
        clear: lv_comment_index, lv_cmnt_vmlindex.
      endif.
* End   - Add - Issue #180

* Add drawings **********************************
      if lo_worksheet->get_drawings( )->is_empty( ) = abap_false.
        lv_drawing_counter += 1.
        lv_drawing_index = lv_drawing_counter.
        lv_index_str = lv_drawing_index.
        condense lv_index_str no-gaps.

        lv_content = me->create_xl_drawings( lo_worksheet ).
        lv_xl_drawing = me->c_xl_drawings.
        replace all occurrences of '#' in lv_xl_drawing with lv_index_str.
        lo_zip->add( name    = lv_xl_drawing
                     content = lv_content ).

        lv_content = me->create_xl_drawings_rels( lo_worksheet ).
        lv_xl_drawing_rels = me->c_xl_drawings_rels.
        replace all occurrences of '#' in lv_xl_drawing_rels with lv_index_str.
        lo_zip->add( name    = lv_xl_drawing_rels
                     content = lv_content ).
      else.
        clear lv_drawing_index.
      endif.

* Add Header/Footer image
      if lines( lo_worksheet->get_header_footer_drawings( ) ) > 0. "Header or footer image exist
        lv_vml_counter += 1.
        lv_hdft_vmlindex = lv_vml_counter.
        lv_index_str = lv_hdft_vmlindex.
        condense lv_index_str no-gaps.

        " Create vmlDrawing that will host the image
        lv_content = me->create_xl_drawing_for_hdft_im( lo_worksheet ).
        lv_xl_drawing_for_comment = me->cl_xl_drawing_for_comments.
        replace all occurrences of '#' in lv_xl_drawing_for_comment with lv_index_str.
        lo_zip->add( name    = lv_xl_drawing_for_comment
                     content = lv_content ).

        " Create vmlDrawing REL that will host the image
        lv_content = me->create_xl_drawings_hdft_rels( lo_worksheet ).
        lv_xl_drawing_rels = me->c_xl_drawings_vml_rels.
        replace all occurrences of '#' in lv_xl_drawing_rels with lv_index_str.
        lo_zip->add( name    = lv_xl_drawing_rels
                     content = lv_content ).
      else.
        clear lv_hdft_vmlindex.
      endif.


      lv_xl_sheet_rels = me->c_xl_sheet_rels.
      lv_content = me->create_xl_sheet_rels( io_worksheet = lo_worksheet
                                             iv_drawing_index = lv_drawing_index
                                             iv_comment_index = lv_comment_index         " (+) Issue #180
                                             iv_cmnt_vmlindex = lv_cmnt_vmlindex
                                             iv_hdft_vmlindex = lv_hdft_vmlindex ).

      lv_index_str = lv_sheet_index.
      condense lv_index_str no-gaps.
      replace all occurrences of '#' in lv_xl_sheet_rels with lv_index_str.
      lo_zip->add( name    = lv_xl_sheet_rels
                   content = lv_content ).

      lo_nested_iterator = lo_worksheet->get_tables_iterator( ).

      while lo_nested_iterator->has_next( ) eq abap_true.
        lo_table ?= lo_nested_iterator->get_next( ).
        lv_content = me->create_xl_table( lo_table ).

        lv_value = lo_table->get_name( ).
        concatenate 'xl/tables/' lv_value '.xml' into lv_value.
        lo_zip->add( name = lv_value
                      content = lv_content ).
      endwhile.



    endwhile.

**********************************************************************
* STEP 11: Add media
    lo_iterator = me->excel->get_drawings_iterator( zcl_excel_drawing=>type_image ).
    while lo_iterator->has_next( ) eq abap_true.
      lo_drawing ?= lo_iterator->get_next( ).

      lv_content = lo_drawing->get_media( ).
      lv_value = lo_drawing->get_media_name( ).
      concatenate 'xl/media/' lv_value into lv_value.
      lo_zip->add( name    = lv_value
                   content = lv_content ).
    endwhile.

**********************************************************************
* STEP 12: Add charts
    lo_iterator = me->excel->get_drawings_iterator( zcl_excel_drawing=>type_chart ).
    while lo_iterator->has_next( ) eq abap_true.
      lo_drawing ?= lo_iterator->get_next( ).

      lv_content = lo_drawing->get_media( ).

      "-------------Added by Alessandro Iannacci - Only if template exist
      if lv_content is not initial and me->excel->use_template eq abap_true.
        lv_value = lo_drawing->get_media_name( ).
        concatenate 'xl/charts/' lv_value into lv_value.
        lo_zip->add( name    = lv_value
                     content = lv_content ).
      else. "ADD CUSTOM CHART!!!!
        lv_content = me->create_xl_charts( lo_drawing ).
        lv_value = lo_drawing->get_media_name( ).
        concatenate 'xl/charts/' lv_value into lv_value.
        lo_zip->add( name    = lv_value
                     content = lv_content ).
      endif.
      "-------------------------------------------------
    endwhile.

* Second to last step: Allow further information put into the zip archive by child classes
    me->add_further_data_to_zip( lo_zip ).

**********************************************************************
* Last step: Create the final zip
    ep_excel = lo_zip->save( ).

  endmethod.


  method create_content_types.


** Constant node name
    data: lc_xml_node_types        type string value 'Types',
          lc_xml_node_override     type string value 'Override',
          lc_xml_node_default      type string value 'Default',
          " Node attributes
          lc_xml_attr_partname     type string value 'PartName',
          lc_xml_attr_extension    type string value 'Extension',
          lc_xml_attr_contenttype  type string value 'ContentType',
          " Node namespace
          lc_xml_node_types_ns     type string value 'http://schemas.openxmlformats.org/package/2006/content-types',
          " Node extension
          lc_xml_node_rels_ext     type string value 'rels',
          lc_xml_node_xml_ext      type string value 'xml',
          lc_xml_node_xml_vml      type string value 'vml',   " (+) GGAR
          " Node partnumber
          lc_xml_node_theme_pn     type string value '/xl/theme/theme1.xml',
          lc_xml_node_styles_pn    type string value '/xl/styles.xml',
          lc_xml_node_workb_pn     type string value '/xl/workbook.xml',
          lc_xml_node_props_pn     type string value '/docProps/app.xml',
          lc_xml_node_worksheet_pn type string value '/xl/worksheets/sheet#.xml',
          lc_xml_node_strings_pn   type string value '/xl/sharedStrings.xml',
          lc_xml_node_core_pn      type string value '/docProps/core.xml',
          lc_xml_node_chart_pn     type string value '/xl/charts/chart#.xml',
          " Node contentType
          lc_xml_node_theme_ct     type string value 'application/vnd.openxmlformats-officedocument.theme+xml',
          lc_xml_node_styles_ct    type string value 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
          lc_xml_node_workb_ct     type string value 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
          lc_xml_node_rels_ct      type string value 'application/vnd.openxmlformats-package.relationships+xml',
          lc_xml_node_vml_ct       type string value 'application/vnd.openxmlformats-officedocument.vmlDrawing',
          lc_xml_node_xml_ct       type string value 'application/xml',
          lc_xml_node_props_ct     type string value 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
          lc_xml_node_worksheet_ct type string value 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
          lc_xml_node_strings_ct   type string value 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
          lc_xml_node_core_ct      type string value 'application/vnd.openxmlformats-package.core-properties+xml',
          lc_xml_node_table_ct     type string value 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
          lc_xml_node_comments_ct  type string value 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',   " (+) GGAR
          lc_xml_node_drawings_ct  type string value 'application/vnd.openxmlformats-officedocument.drawing+xml',
          lc_xml_node_chart_ct     type string value 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'.

    data: lo_document        type ref to if_ixml_document,
          lo_element_root    type ref to if_ixml_element,
          lo_element         type ref to if_ixml_element,
          lo_worksheet       type ref to zcl_excel_worksheet,
          lo_iterator        type ref to zcl_excel_collection_iterator,
          lo_nested_iterator type ref to zcl_excel_collection_iterator,
          lo_table           type ref to zcl_excel_table.

    data: lv_worksheets_num        type i,
          lv_worksheets_numc       type n length 3,
          lv_xml_node_worksheet_pn type string,
          lv_value                 type string,
          lv_comment_index         type i value 1,  " (+) GGAR
          lv_drawing_index         type i value 1,
          lv_index_str             type string.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node types
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_types
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
    value = lc_xml_node_types_ns ).

**********************************************************************
* STEP 4: Create subnodes

    " rels node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_extension
                                  value = lc_xml_node_rels_ext ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_rels_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " extension node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_extension
                                  value = lc_xml_node_xml_ext ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_xml_ct ).
    lo_element_root->append_child( new_child = lo_element ).

* Begin - Add - GGAR
    " VML node (for comments)
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_extension
                                  value = lc_xml_node_xml_vml ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_vml_ct ).
    lo_element_root->append_child( new_child = lo_element ).
* End   - Add - GGAR

    " Theme node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_theme_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_theme_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " Styles node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_styles_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_styles_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " Workbook node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_workb_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_workb_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " Properties node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_props_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_props_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " Worksheet node
    lv_worksheets_num = excel->get_worksheets_size( ).
    do lv_worksheets_num times.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                       parent = lo_document ).

      lv_worksheets_numc = sy-index.
      shift lv_worksheets_numc left deleting leading '0'.
      lv_xml_node_worksheet_pn = lc_xml_node_worksheet_pn.
      replace all occurrences of '#' in lv_xml_node_worksheet_pn with lv_worksheets_numc.
      lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                    value = lv_xml_node_worksheet_pn ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                    value = lc_xml_node_worksheet_ct ).
      lo_element_root->append_child( new_child = lo_element ).
    enddo.

    lo_iterator = me->excel->get_worksheets_iterator( ).
    while lo_iterator->has_next( ) eq abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).

      lo_nested_iterator = lo_worksheet->get_tables_iterator( ).

      while lo_nested_iterator->has_next( ) eq abap_true.
        lo_table ?= lo_nested_iterator->get_next( ).

        lv_value = lo_table->get_name( ).
        concatenate '/xl/tables/' lv_value '.xml' into lv_value.

        lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lv_value ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_table_ct ).
        lo_element_root->append_child( new_child = lo_element ).
      endwhile.

* Begin - Add - GGAR
      " Comments
      data: lo_comments type ref to zcl_excel_comments.

      lo_comments = lo_worksheet->get_comments( ).
      if lo_comments->is_empty( ) = abap_false.
        lv_index_str = lv_comment_index.
        condense lv_index_str no-gaps.
        concatenate '/' me->c_xl_comments into lv_value.
        lv_value = replace( val = lv_value sub = '#' with = lv_index_str ).

        lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                         parent = lo_document ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                      value = lv_value ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                      value = lc_xml_node_comments_ct ).
        lo_element_root->append_child( new_child = lo_element ).

        lv_comment_index += 1.
      endif.
* End   - Add - GGAR

      " Drawings
      data: lo_drawings type ref to zcl_excel_drawings.

      lo_drawings = lo_worksheet->get_drawings( ).
      if lo_drawings->is_empty( ) = abap_false.
        lv_index_str = lv_drawing_index.
        condense lv_index_str no-gaps.
        concatenate '/' me->c_xl_drawings into lv_value.
        lv_value = replace( val = lv_value sub = '#' with = lv_index_str ).

        lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lv_value ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_drawings_ct ).
        lo_element_root->append_child( new_child = lo_element ).

        lv_drawing_index += 1.
      endif.
    endwhile.

    " media mimes
    data: lo_drawing    type ref to zcl_excel_drawing.
*          lt_media_type TYPE TABLE OF mimetypes-extension,
*          lv_media_type TYPE mimetypes-extension,
*          lv_mime_type  TYPE mimetypes-type.
*
*    lo_iterator = me->excel->get_drawings_iterator( zcl_excel_drawing=>type_image ).
*    WHILE lo_iterator->has_next( ) = abap_true.
*      lo_drawing ?= lo_iterator->get_next( ).
*
*      lv_media_type = lo_drawing->get_media_type( ).
*      COLLECT lv_media_type INTO lt_media_type.
*    ENDWHILE.
*
*    LOOP AT lt_media_type INTO lv_media_type.
*      CALL FUNCTION 'SDOK_MIMETYPE_GET'
*        EXPORTING
*          extension = lv_media_type
*        IMPORTING
*          mimetype  = lv_mime_type.
*
*      lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
*                                                       parent = lo_document ).
*      lv_value = lv_media_type.
*      lo_element->set_attribute_ns( name  = lc_xml_attr_extension
*                                    value = lv_value ).
*      lv_value = lv_mime_type.
*      lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
*                                    value = lv_value ).
*      lo_element_root->append_child( new_child = lo_element ).
*    ENDLOOP.

    " Charts
    lo_iterator = me->excel->get_drawings_iterator( zcl_excel_drawing=>type_chart ).
    while lo_iterator->has_next( ) = abap_true.
      lo_drawing ?= lo_iterator->get_next( ).

      lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                       parent = lo_document ).
      lv_index_str = lo_drawing->get_index( ).
      condense lv_index_str.
      lv_value = lc_xml_node_chart_pn.
      replace all occurrences of '#' in lv_value with lv_index_str.
      lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                    value = lc_xml_node_chart_ct ).
      lo_element_root->append_child( new_child = lo_element ).
    endwhile.

    " Strings node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_strings_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_strings_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " Strings node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_core_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_core_ct ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  endmethod.


  method create_docprops_app.


** Constant node name
    data: lc_xml_node_properties        type string value 'Properties',
          lc_xml_node_application       type string value 'Application',
          lc_xml_node_docsecurity       type string value 'DocSecurity',
          lc_xml_node_scalecrop         type string value 'ScaleCrop',
          lc_xml_node_headingpairs      type string value 'HeadingPairs',
          lc_xml_node_vector            type string value 'vector',
          lc_xml_node_variant           type string value 'variant',
          lc_xml_node_lpstr             type string value 'lpstr',
          lc_xml_node_i4                type string value 'i4',
          lc_xml_node_titlesofparts     type string value 'TitlesOfParts',
          lc_xml_node_company           type string value 'Company',
          lc_xml_node_linksuptodate     type string value 'LinksUpToDate',
          lc_xml_node_shareddoc         type string value 'SharedDoc',
          lc_xml_node_hyperlinkschanged type string value 'HyperlinksChanged',
          lc_xml_node_appversion        type string value 'AppVersion',
          " Namespace prefix
          lc_vt_ns                      type string value 'vt',
          lc_xml_node_props_ns          type string value 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
          lc_xml_node_props_vt_ns       type string value 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
          " Node attributes
          lc_xml_attr_size              type string value 'size',
          lc_xml_attr_basetype          type string value 'baseType'.

    data: lo_document            type ref to if_ixml_document,
          lo_element_root        type ref to if_ixml_element,
          lo_element             type ref to if_ixml_element,
          lo_sub_element_vector  type ref to if_ixml_element,
          lo_sub_element_variant type ref to if_ixml_element,
          lo_sub_element_lpstr   type ref to if_ixml_element,
          lo_sub_element_i4      type ref to if_ixml_element,
          lo_iterator            type ref to zcl_excel_collection_iterator,
          lo_worksheet           type ref to zcl_excel_worksheet.

    data: lv_value                type string.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node properties
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_properties
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_props_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:vt'
                                       value = lc_xml_node_props_vt_ns ).

**********************************************************************
* STEP 4: Create subnodes
    " Application
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_application
                                                     parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~application.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " DocSecurity
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_docsecurity
                                                              parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~docsecurity.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " ScaleCrop
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_scalecrop
                                                     parent = lo_document ).
    lv_value = me->flag2bool( excel->zif_excel_book_properties~scalecrop ).
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " HeadingPairs
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_headingpairs
                                                     parent = lo_document ).


    " * vector node
    lo_sub_element_vector = lo_document->create_simple_element_ns( name   = lc_xml_node_vector
                                                                   prefix = lc_vt_ns
                                                                   parent = lo_document ).
    lo_sub_element_vector->set_attribute_ns( name    = lc_xml_attr_size
                                             value   = '2' ).
    lo_sub_element_vector->set_attribute_ns( name    = lc_xml_attr_basetype
                                             value   = lc_xml_node_variant ).

    " ** variant node
    lo_sub_element_variant = lo_document->create_simple_element_ns( name   = lc_xml_node_variant
                                                                    prefix = lc_vt_ns
                                                                    parent = lo_document ).

    " *** lpstr node
    lo_sub_element_lpstr = lo_document->create_simple_element_ns( name   = lc_xml_node_lpstr
                                                                  prefix = lc_vt_ns
                                                                  parent = lo_document ).
    lv_value = excel->get_worksheets_name( ).
    lo_sub_element_lpstr->set_value( value = lv_value ).
    lo_sub_element_variant->append_child( new_child = lo_sub_element_lpstr ). " lpstr node

    lo_sub_element_vector->append_child( new_child = lo_sub_element_variant ). " variant node

    " ** variant node
    lo_sub_element_variant = lo_document->create_simple_element_ns( name   = lc_xml_node_variant
                                                                    prefix = lc_vt_ns
                                                                    parent = lo_document ).

    " *** i4 node
    lo_sub_element_i4 = lo_document->create_simple_element_ns( name   = lc_xml_node_i4
                                                               prefix = lc_vt_ns
                                                               parent = lo_document ).
    lv_value = excel->get_worksheets_size( ).
    shift lv_value right deleting trailing space.
    shift lv_value left deleting leading space.
    lo_sub_element_i4->set_value( value = lv_value ).
    lo_sub_element_variant->append_child( new_child = lo_sub_element_i4 ). " lpstr node

    lo_sub_element_vector->append_child( new_child = lo_sub_element_variant ). " variant node

    lo_element->append_child( new_child = lo_sub_element_vector ). " vector node

    lo_element_root->append_child( new_child = lo_element ). " HeadingPairs


    " TitlesOfParts
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_titlesofparts
                                                     parent = lo_document ).


    " * vector node
    lo_sub_element_vector = lo_document->create_simple_element_ns( name   = lc_xml_node_vector
                                                                   prefix = lc_vt_ns
                                                                   parent = lo_document ).
    lv_value = excel->get_worksheets_size( ).
    shift lv_value right deleting trailing space.
    shift lv_value left deleting leading space.
    lo_sub_element_vector->set_attribute_ns( name    = lc_xml_attr_size
                                             value   = lv_value ).
    lo_sub_element_vector->set_attribute_ns( name    = lc_xml_attr_basetype
                                             value   = lc_xml_node_lpstr ).

    lo_iterator = excel->get_worksheets_iterator( ).

    while lo_iterator->has_next( ) eq abap_true.
      " ** lpstr node
      lo_sub_element_lpstr = lo_document->create_simple_element_ns( name   = lc_xml_node_lpstr
                                                                    prefix = lc_vt_ns
                                                                    parent = lo_document ).
      lo_worksheet ?= lo_iterator->get_next( ).
      lv_value = lo_worksheet->get_title( ).
      lo_sub_element_lpstr->set_value( value = lv_value ).
      lo_sub_element_vector->append_child( new_child = lo_sub_element_lpstr ). " lpstr node
    endwhile.

    lo_element->append_child( new_child = lo_sub_element_vector ). " vector node

    lo_element_root->append_child( new_child = lo_element ). " TitlesOfParts



    " Company
    if excel->zif_excel_book_properties~company is not initial.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_company
                                                       parent = lo_document ).
      lv_value = excel->zif_excel_book_properties~company.
      lo_element->set_value( value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    endif.

    " LinksUpToDate
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_linksuptodate
                                                     parent = lo_document ).
    lv_value = me->flag2bool( excel->zif_excel_book_properties~linksuptodate ).
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " SharedDoc
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_shareddoc
                                                     parent = lo_document ).
    lv_value = me->flag2bool( excel->zif_excel_book_properties~shareddoc ).
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " HyperlinksChanged
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_hyperlinkschanged
                                                     parent = lo_document ).
    lv_value = me->flag2bool( excel->zif_excel_book_properties~hyperlinkschanged ).
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " AppVersion
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_appversion
                                                     parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~appversion.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  endmethod.


  method create_docprops_core.


** Constant node name
    data: lc_xml_node_coreproperties type string value 'coreProperties',
          lc_xml_node_creator        type string value 'creator',
          lc_xml_node_description    type string value 'description',
          lc_xml_node_lastmodifiedby type string value 'lastModifiedBy',
          lc_xml_node_created        type string value 'created',
          lc_xml_node_modified       type string value 'modified',
          " Node attributes
          lc_xml_attr_type           type string value 'type',
          lc_xml_attr_target         type string value 'dcterms:W3CDTF',
          " Node namespace
          lc_cp_ns                   type string value 'cp',
          lc_dc_ns                   type string value 'dc',
          lc_dcterms_ns              type string value 'dcterms',
*        lc_dcmitype_ns              TYPE string VALUE 'dcmitype',
          lc_xsi_ns                  type string value 'xsi',
          lc_xml_node_cp_ns          type string value 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
          lc_xml_node_dc_ns          type string value 'http://purl.org/dc/elements/1.1/',
          lc_xml_node_dcterms_ns     type string value 'http://purl.org/dc/terms/',
          lc_xml_node_dcmitype_ns    type string value 'http://purl.org/dc/dcmitype/',
          lc_xml_node_xsi_ns         type string value 'http://www.w3.org/2001/XMLSchema-instance'.

    data: lo_document     type ref to if_ixml_document,
          lo_element_root type ref to if_ixml_element,
          lo_element      type ref to if_ixml_element.

    data: lv_value type string,
          lv_date  type d,
          lv_time  type t.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node coreProperties
    lo_element_root  = lo_document->create_simple_element_ns( name   = lc_xml_node_coreproperties
                                                              prefix = lc_cp_ns
                                                              parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:cp'
                                       value = lc_xml_node_cp_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:dc'
                                       value = lc_xml_node_dc_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:dcterms'
                                       value = lc_xml_node_dcterms_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:dcmitype'
                                       value = lc_xml_node_dcmitype_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:xsi'
                                       value = lc_xml_node_xsi_ns ).

**********************************************************************
* STEP 4: Create subnodes
    " Creator node
    lo_element = lo_document->create_simple_element_ns( name   = lc_xml_node_creator
                                                        prefix = lc_dc_ns
                                                        parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~creator.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " Description node
    lo_element = lo_document->create_simple_element_ns( name   = lc_xml_node_description
                                                        prefix = lc_dc_ns
                                                        parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~description.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " lastModifiedBy node
    lo_element = lo_document->create_simple_element_ns( name   = lc_xml_node_lastmodifiedby
                                                        prefix = lc_cp_ns
                                                        parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~lastmodifiedby.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " Created node
    lo_element = lo_document->create_simple_element_ns( name   = lc_xml_node_created
                                                        prefix = lc_dcterms_ns
                                                        parent = lo_document ).
    lo_element->set_attribute_ns( name    = lc_xml_attr_type
                                  prefix  = lc_xsi_ns
                                  value   = lc_xml_attr_target ).
    try.
        convert time stamp excel->zif_excel_book_properties~created time zone cl_abap_context_info=>get_user_time_zone( ) into date lv_date time lv_time.
      catch cx_root.
    endtry.
    concatenate lv_date lv_time into lv_value respecting blanks.
    replace all occurrences of pcre  '([0-9]{4})([0-9]{2})([0-9]{2})([0-9]{2})([0-9]{2})([0-9]{2})' in lv_value with '$1-$2-$3T$4:$5:$6Z'.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " Modified node
    lo_element = lo_document->create_simple_element_ns( name   = lc_xml_node_modified
                                                        prefix = lc_dcterms_ns
                                                        parent = lo_document ).
    lo_element->set_attribute_ns( name    = lc_xml_attr_type
                                  prefix  = lc_xsi_ns
                                  value   = lc_xml_attr_target ).
    try.
        convert time stamp excel->zif_excel_book_properties~modified time zone cl_abap_context_info=>get_user_time_zone( ) into date lv_date time lv_time.
      catch cx_root.
    endtry.
    concatenate lv_date lv_time into lv_value respecting blanks.
    replace all occurrences of pcre  '([0-9]{4})([0-9]{2})([0-9]{2})([0-9]{2})([0-9]{2})([0-9]{2})' in lv_value with '$1-$2-$3T$4:$5:$6Z'.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  endmethod.


  method create_dxf_style.

    constants: lc_xml_node_dxf         type string value 'dxf',
               lc_xml_node_font        type string value 'font',
               lc_xml_node_b           type string value 'b',            "bold
               lc_xml_node_i           type string value 'i',            "italic
               lc_xml_node_u           type string value 'u',            "underline
               lc_xml_node_strike      type string value 'strike',       "strikethrough
               lc_xml_attr_val         type string value 'val',
               lc_xml_node_fill        type string value 'fill',
               lc_xml_node_patternfill type string value 'patternFill',
               lc_xml_attr_patterntype type string value 'patternType',
               lc_xml_node_fgcolor     type string value 'fgColor',
               lc_xml_node_bgcolor     type string value 'bgColor'.

    data: ls_styles_mapping     type zif_excel_data_decl=>zexcel_s_styles_mapping,
          ls_cellxfs            type zif_excel_data_decl=>zexcel_s_cellxfs,
          ls_style_cond_mapping type zif_excel_data_decl=>zexcel_s_styles_cond_mapping,
          lo_sub_element        type ref to if_ixml_element,
          lo_sub_element_2      type ref to if_ixml_element,
          lv_index              type i,
          ls_font               type zif_excel_data_decl=>zexcel_s_style_font,
          lo_element_font       type ref to if_ixml_element,
          lv_value              type string,
          ls_fill               type zif_excel_data_decl=>zexcel_s_style_fill,
          lo_element_fill       type ref to if_ixml_element.

    check iv_cell_style is not initial.

    "Don't insert guid twice or even more
    read table me->styles_cond_mapping transporting no fields with key guid = iv_cell_style.
    check sy-subrc ne 0.

    read table me->styles_mapping into ls_styles_mapping with key guid = iv_cell_style.

    read table me->styles_cond_mapping into ls_style_cond_mapping with key style = ls_styles_mapping-style.
    if sy-subrc eq 0.
      "The content of this style is equal to an existing one. Share its dxfid.
      ls_style_cond_mapping-guid  = iv_cell_style.
      append ls_style_cond_mapping to me->styles_cond_mapping.
    else.
      ls_style_cond_mapping-guid  = iv_cell_style.
      ls_style_cond_mapping-style = ls_styles_mapping-style.
      ls_style_cond_mapping-dxf   = cv_dfx_count.
      append ls_style_cond_mapping to me->styles_cond_mapping.
      cv_dfx_count += 1.

      " dxf node
      lo_sub_element = io_ixml_document->create_simple_element( name   = lc_xml_node_dxf
                                                                parent = io_ixml_document ).

      lv_index = ls_styles_mapping-style + 1.
      read table it_cellxfs into ls_cellxfs index lv_index.

      "Conditional formatting font style correction by Alessandro Iannacci START
      lv_index = ls_cellxfs-fontid + 1.
      read table it_fonts into ls_font index lv_index.
      if ls_font is not initial.
        lo_element_font = io_ixml_document->create_simple_element( name   = lc_xml_node_font
                                                              parent = io_ixml_document ).
        if ls_font-bold eq abap_true.
          lo_sub_element_2 = io_ixml_document->create_simple_element( name   = lc_xml_node_b
                                                               parent = io_ixml_document ).
          lo_element_font->append_child( new_child = lo_sub_element_2 ).
        endif.
        if ls_font-italic eq abap_true.
          lo_sub_element_2 = io_ixml_document->create_simple_element( name   = lc_xml_node_i
                                                               parent = io_ixml_document ).
          lo_element_font->append_child( new_child = lo_sub_element_2 ).
        endif.
        if ls_font-underline eq abap_true.
          lo_sub_element_2 = io_ixml_document->create_simple_element( name   = lc_xml_node_u
                                                               parent = io_ixml_document ).
          lv_value = ls_font-underline_mode.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_val
                                            value = lv_value ).
          lo_element_font->append_child( new_child = lo_sub_element_2 ).
        endif.
        if ls_font-strikethrough eq abap_true.
          lo_sub_element_2 = io_ixml_document->create_simple_element( name   = lc_xml_node_strike
                                                               parent = io_ixml_document ).
          lo_element_font->append_child( new_child = lo_sub_element_2 ).
        endif.
        "color
        create_xl_styles_color_node(
            io_document        = io_ixml_document
            io_parent          = lo_element_font
            is_color           = ls_font-color ).
        lo_sub_element->append_child( new_child = lo_element_font ).
      endif.
      "---Conditional formatting font style correction by Alessandro Iannacci END


      lv_index = ls_cellxfs-fillid + 1.
      read table it_fills into ls_fill index lv_index.
      if ls_fill is not initial.
        " fill properties
        lo_element_fill = io_ixml_document->create_simple_element( name   = lc_xml_node_fill
                                                                 parent = io_ixml_document ).
        "pattern
        lo_sub_element_2 = io_ixml_document->create_simple_element( name   = lc_xml_node_patternfill
                                                             parent = io_ixml_document ).
        lv_value = ls_fill-filltype.
        lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_patterntype
                                            value = lv_value ).
        " fgcolor
        create_xl_styles_color_node(
            io_document        = io_ixml_document
            io_parent          = lo_sub_element_2
            is_color           = ls_fill-fgcolor
            iv_color_elem_name = lc_xml_node_fgcolor ).

        if  ls_fill-fgcolor-rgb is initial and
          ls_fill-fgcolor-indexed eq zcl_excel_style_color=>c_indexed_not_set and
          ls_fill-fgcolor-theme eq zcl_excel_style_color=>c_theme_not_set and
          ls_fill-fgcolor-tint is initial and ls_fill-bgcolor-indexed eq zcl_excel_style_color=>c_indexed_sys_foreground.

          " bgcolor
          create_xl_styles_color_node(
              io_document        = io_ixml_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-bgcolor
              iv_color_elem_name = lc_xml_node_bgcolor ).

        endif.

        lo_element_fill->append_child( new_child = lo_sub_element_2 ). "pattern

        lo_sub_element->append_child( new_child = lo_element_fill ).
      endif.

      io_dxf_element->append_child( new_child = lo_sub_element ).
    endif.
  endmethod.


  method create_relationships.


** Constant node name
    data: lc_xml_node_relationships type string value 'Relationships',
          lc_xml_node_relationship  type string value 'Relationship',
          " Node attributes
          lc_xml_attr_id            type string value 'Id',
          lc_xml_attr_type          type string value 'Type',
          lc_xml_attr_target        type string value 'Target',
          " Node namespace
          lc_xml_node_rels_ns       type string value 'http://schemas.openxmlformats.org/package/2006/relationships',
          " Node id
          lc_xml_node_rid1_id       type string value 'rId1',
          lc_xml_node_rid2_id       type string value 'rId2',
          lc_xml_node_rid3_id       type string value 'rId3',
          " Node type
          lc_xml_node_rid1_tp       type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
          lc_xml_node_rid2_tp       type string value 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
          lc_xml_node_rid3_tp       type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
          " Node target
          lc_xml_node_rid1_tg       type string value 'xl/workbook.xml',
          lc_xml_node_rid2_tg       type string value 'docProps/core.xml',
          lc_xml_node_rid3_tg       type string value 'docProps/app.xml'.

    data: lo_document     type ref to if_ixml_document,
          lo_element_root type ref to if_ixml_element,
          lo_element      type ref to if_ixml_element.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_relationships
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_rels_ns ).

**********************************************************************
* STEP 4: Create subnodes
    " Theme node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lc_xml_node_rid3_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid3_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid3_tg ).
    lo_element_root->append_child( new_child = lo_element ).

    " Styles node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lc_xml_node_rid2_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid2_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid2_tg ).
    lo_element_root->append_child( new_child = lo_element ).

    " rels node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lc_xml_node_rid1_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid1_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid1_tg ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  endmethod.


  method create_xl_charts.


** Constant node name
    constants: lc_xml_node_chartspace         type string value 'c:chartSpace',
               lc_xml_node_ns_c               type string value 'http://schemas.openxmlformats.org/drawingml/2006/chart',
               lc_xml_node_ns_a               type string value 'http://schemas.openxmlformats.org/drawingml/2006/main',
               lc_xml_node_ns_r               type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
               lc_xml_node_date1904           type string value 'c:date1904',
               lc_xml_node_lang               type string value 'c:lang',
               lc_xml_node_roundedcorners     type string value 'c:roundedCorners',
               lc_xml_node_altcont            type string value 'mc:AlternateContent',
               lc_xml_node_altcont_ns_mc      type string value 'http://schemas.openxmlformats.org/markup-compatibility/2006',
               lc_xml_node_choice             type string value 'mc:Choice',
               lc_xml_node_choice_ns_requires type string value 'c14',
               lc_xml_node_choice_ns_c14      type string value 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart',
               lc_xml_node_style              type string value 'c14:style',
               lc_xml_node_fallback           type string value 'mc:Fallback',
               lc_xml_node_style2             type string value 'c:style',

               "---------------------------CHART
               lc_xml_node_chart              type string value 'c:chart',
               lc_xml_node_autotitledeleted   type string value 'c:autoTitleDeleted',
               "plotArea
               lc_xml_node_plotarea           type string value 'c:plotArea',
               lc_xml_node_layout             type string value 'c:layout',
               lc_xml_node_varycolors         type string value 'c:varyColors',
               lc_xml_node_ser                type string value 'c:ser',
               lc_xml_node_idx                type string value 'c:idx',
               lc_xml_node_order              type string value 'c:order',
               lc_xml_node_tx                 type string value 'c:tx',
               lc_xml_node_v                  type string value 'c:v',
               lc_xml_node_val                type string value 'c:val',
               lc_xml_node_cat                type string value 'c:cat',
               lc_xml_node_numref             type string value 'c:numRef',
               lc_xml_node_strref             type string value 'c:strRef',
               lc_xml_node_f                  type string value 'c:f', "this is the range
               lc_xml_node_overlap            type string value 'c:overlap',
               "note: numcache avoided
               lc_xml_node_dlbls              type string value 'c:dLbls',
               lc_xml_node_showlegendkey      type string value 'c:showLegendKey',
               lc_xml_node_showval            type string value 'c:showVal',
               lc_xml_node_showcatname        type string value 'c:showCatName',
               lc_xml_node_showsername        type string value 'c:showSerName',
               lc_xml_node_showpercent        type string value 'c:showPercent',
               lc_xml_node_showbubblesize     type string value 'c:showBubbleSize',
               "plotArea->pie
               lc_xml_node_piechart           type string value 'c:pieChart',
               lc_xml_node_showleaderlines    type string value 'c:showLeaderLines',
               lc_xml_node_firstsliceang      type string value 'c:firstSliceAng',
               "plotArea->line
               lc_xml_node_linechart          type string value 'c:lineChart',
               lc_xml_node_symbol             type string value 'c:symbol',
               lc_xml_node_marker             type string value 'c:marker',
               lc_xml_node_smooth             type string value 'c:smooth',
               "plotArea->bar
               lc_xml_node_invertifnegative   type string value 'c:invertIfNegative',
               lc_xml_node_barchart           type string value 'c:barChart',
               lc_xml_node_bardir             type string value 'c:barDir',
               lc_xml_node_gapwidth           type string value 'c:gapWidth',
               "plotArea->line + plotArea->bar
               lc_xml_node_grouping           type string value 'c:grouping',
               lc_xml_node_axid               type string value 'c:axId',
               lc_xml_node_catax              type string value 'c:catAx',
               lc_xml_node_valax              type string value 'c:valAx',
               lc_xml_node_scaling            type string value 'c:scaling',
               lc_xml_node_orientation        type string value 'c:orientation',
               lc_xml_node_delete             type string value 'c:delete',
               lc_xml_node_axpos              type string value 'c:axPos',
               lc_xml_node_numfmt             type string value 'c:numFmt',
               lc_xml_node_majorgridlines     type string value 'c:majorGridlines',
               lc_xml_node_majortickmark      type string value 'c:majorTickMark',
               lc_xml_node_minortickmark      type string value 'c:minorTickMark',
               lc_xml_node_ticklblpos         type string value 'c:tickLblPos',
               lc_xml_node_crossax            type string value 'c:crossAx',
               lc_xml_node_crosses            type string value 'c:crosses',
               lc_xml_node_auto               type string value 'c:auto',
               lc_xml_node_lblalgn            type string value 'c:lblAlgn',
               lc_xml_node_lbloffset          type string value 'c:lblOffset',
               lc_xml_node_nomultilvllbl      type string value 'c:noMultiLvlLbl',
               lc_xml_node_crossbetween       type string value 'c:crossBetween',
               "legend
               lc_xml_node_legend             type string value 'c:legend',
               "legend->pie
               lc_xml_node_legendpos          type string value 'c:legendPos',
*                  lc_xml_node_layout            TYPE string VALUE 'c:layout', "already exist
               lc_xml_node_overlay            type string value 'c:overlay',
               lc_xml_node_txpr               type string value 'c:txPr',
               lc_xml_node_bodypr             type string value 'a:bodyPr',
               lc_xml_node_lststyle           type string value 'a:lstStyle',
               lc_xml_node_p                  type string value 'a:p',
               lc_xml_node_ppr                type string value 'a:pPr',
               lc_xml_node_defrpr             type string value 'a:defRPr',
               lc_xml_node_endpararpr         type string value 'a:endParaRPr',
               "legend->bar + legend->line
               lc_xml_node_plotvisonly        type string value 'c:plotVisOnly',
               lc_xml_node_dispblanksas       type string value 'c:dispBlanksAs',
               lc_xml_node_showdlblsovermax   type string value 'c:showDLblsOverMax',
               "---------------------------END OF CHART

               lc_xml_node_printsettings      type string value 'c:printSettings',
               lc_xml_node_headerfooter       type string value 'c:headerFooter',
               lc_xml_node_pagemargins        type string value 'c:pageMargins',
               lc_xml_node_pagesetup          type string value 'c:pageSetup'.


    data: lo_document     type ref to if_ixml_document,
          lo_element_root type ref to if_ixml_element.


    data lo_element                               type ref to if_ixml_element.
    data lo_element2                              type ref to if_ixml_element.
    data lo_element3                              type ref to if_ixml_element.
    data lo_el_rootchart                           type ref to if_ixml_element.
    data lo_element4                              type ref to if_ixml_element.
    data lo_element5                              type ref to if_ixml_element.
    data lo_element6                              type ref to if_ixml_element.
    data lo_element7                              type ref to if_ixml_element.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

***********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_chartspace
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:c'
                                       value = lc_xml_node_ns_c ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:a'
                                       value = lc_xml_node_ns_a ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:r'
                                       value = lc_xml_node_ns_r ).

**********************************************************************
* STEP 4: Create chart

    data lo_chartb type ref to zcl_excel_graph_bars.
    data lo_chartp type ref to zcl_excel_graph_pie.
    data lo_chartl type ref to zcl_excel_graph_line.
    data lo_chart type ref to zcl_excel_graph.

    data ls_serie type zcl_excel_graph=>s_series.
    data ls_ax type zcl_excel_graph_bars=>s_ax.
    data lv_str type string.

    "Identify chart type
    case io_drawing->graph_type.
      when zcl_excel_drawing=>c_graph_bars.
        lo_chartb ?= io_drawing->graph.
      when zcl_excel_drawing=>c_graph_pie.
        lo_chartp ?= io_drawing->graph.
      when zcl_excel_drawing=>c_graph_line.
        lo_chartl ?= io_drawing->graph.
      when others.
    endcase.


    lo_chart = io_drawing->graph.

    lo_element = lo_document->create_simple_element( name = lc_xml_node_date1904
                                                         parent = lo_element_root ).
    lo_element->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_1904val ).

    lo_element = lo_document->create_simple_element( name = lc_xml_node_lang
                                                         parent = lo_element_root ).
    lo_element->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_langval ).

    lo_element = lo_document->create_simple_element( name = lc_xml_node_roundedcorners
                                                         parent = lo_element_root ).
    lo_element->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_roundedcornersval ).

    lo_element = lo_document->create_simple_element( name = lc_xml_node_altcont
                                                         parent = lo_element_root ).
    lo_element->set_attribute_ns( name  = 'xmlns:mc'
                                      value = lc_xml_node_altcont_ns_mc ).

    "Choice
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_choice
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'Requires'
                                      value = lc_xml_node_choice_ns_requires ).
    lo_element2->set_attribute_ns( name  = 'xmlns:c14'
                                      value = lc_xml_node_choice_ns_c14 ).

    "C14:style
    lo_element3 = lo_document->create_simple_element( name = lc_xml_node_style
                                                         parent = lo_element2 ).
    lo_element3->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_c14styleval ).

    "Fallback
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_fallback
                                                         parent = lo_element ).

    "C:style
    lo_element3 = lo_document->create_simple_element( name = lc_xml_node_style2
                                                         parent = lo_element2 ).
    lo_element3->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_styleval ).

    "---------------------------CHART
    lo_element = lo_document->create_simple_element( name = lc_xml_node_chart
                                                         parent = lo_element_root ).
    "Added
    if lo_chart->title is not initial.
      lo_element2 = lo_document->create_simple_element( name = 'c:title'
                                                           parent = lo_element ).
      lo_element3 = lo_document->create_simple_element( name = 'c:tx'
                                                           parent = lo_element2 ).
      lo_element4 = lo_document->create_simple_element( name = 'c:rich'
                                                           parent = lo_element3 ).
      lo_element5 = lo_document->create_simple_element( name = 'a:bodyPr'
                                                           parent = lo_element4 ).
      lo_element5 = lo_document->create_simple_element( name = 'a:lstStyle'
                                                           parent = lo_element4 ).
      lo_element5 = lo_document->create_simple_element( name = 'a:p'
                                                           parent = lo_element4 ).
      lo_element6 = lo_document->create_simple_element( name = 'a:pPr'
                                                           parent = lo_element5 ).
      lo_element7 = lo_document->create_simple_element( name = 'a:defRPr'
                                                           parent = lo_element6 ).
      lo_element6 = lo_document->create_simple_element( name = 'a:r'
                                                           parent = lo_element5 ).
      lo_element7 = lo_document->create_simple_element( name = 'a:rPr'
                                                           parent = lo_element6 ).
      lo_element7->set_attribute_ns( name  = 'lang'
                                        value = 'en-US' ).
      lo_element7 = lo_document->create_simple_element( name = 'a:t'
                                                           parent = lo_element6 ).
      lo_element7->set_value( value = lo_chart->title ).
    endif.
    "End
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_autotitledeleted
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_autotitledeletedval ).

    "plotArea
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_plotarea
                                                       parent = lo_element ).
    lo_element3 = lo_document->create_simple_element( name = lc_xml_node_layout
                                                       parent = lo_element2 ).
    case io_drawing->graph_type.
      when zcl_excel_drawing=>c_graph_bars.
        "----bar
        lo_element3 = lo_document->create_simple_element( name = lc_xml_node_barchart
                                                     parent = lo_element2 ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_bardir
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_bardirval ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_grouping
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_groupingval ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_varycolors
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_varycolorsval ).

        "series
        loop at lo_chartb->series into ls_serie.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ser
                                                     parent = lo_element3 ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_idx
                                                     parent = lo_element4 ).
          if ls_serie-idx is not initial.
            lv_str = ls_serie-idx.
          else.
            lv_str = sy-tabix - 1.
          endif.
          condense lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_order
                                                     parent = lo_element4 ).
          lv_str = ls_serie-order.
          condense lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          if ls_serie-sername is not initial.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_tx
                                                      parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_v
                                                      parent = lo_element5 ).
            lo_element6->set_value( value = ls_serie-sername ).
          endif.
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_invertifnegative
                                                     parent = lo_element4 ).
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = ls_serie-invertifnegative ).
          if ls_serie-lbl is not initial.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_cat
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_strref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-lbl ).
          endif.
          if ls_serie-ref is not initial.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_val
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_numref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-ref ).
          endif.
        endloop.
        "endseries
        if lo_chartb->ns_groupingval = zcl_excel_graph_bars=>c_groupingval_stacked.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_overlap
                                                            parent = lo_element3 ).
          lo_element4->set_attribute_ns( name  = 'val'
                                         value = '100' ).
        endif.

        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_dlbls
                                                     parent = lo_element3 ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showlegendkey
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showlegendkeyval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showval
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showvalval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showcatname
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showcatnameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showsername
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showsernameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showpercent
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showpercentval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showbubblesize
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showbubblesizeval ).

        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_gapwidth
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_gapwidthval ).

        "axes
        lo_el_rootchart = lo_element3.
        loop at lo_chartb->axes into ls_ax.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_el_rootchart ).
          lo_element4->set_attribute_ns( name  = 'val'
                                  value = ls_ax-axid ).
          case ls_ax-type.
            when zcl_excel_graph_bars=>c_catax.
              lo_element3 = lo_document->create_simple_element( name = lc_xml_node_catax
                                                     parent = lo_element2 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axid ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_scaling
                                                     parent = lo_element3 ).
              lo_element5 = lo_document->create_simple_element( name = lc_xml_node_orientation
                                                     parent = lo_element4 ).
              lo_element5->set_attribute_ns( name  = 'val'
                                             value = ls_ax-orientation ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_delete
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-delete ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_numfmt
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'formatCode'
                                             value = ls_ax-formatcode ).
              lo_element4->set_attribute_ns( name  = 'sourceLinked'
                                             value = ls_ax-sourcelinked ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-majortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_minortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-minortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ticklblpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-ticklblpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossax
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossax ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crosses
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crosses ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_auto
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-auto ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_lblalgn
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-lblalgn ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_lbloffset
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-lbloffset ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_nomultilvllbl
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-nomultilvllbl ).
            when zcl_excel_graph_bars=>c_valax.
              lo_element3 = lo_document->create_simple_element( name = lc_xml_node_valax
                                                     parent = lo_element2 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axid ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_scaling
                                                     parent = lo_element3 ).
              lo_element5 = lo_document->create_simple_element( name = lc_xml_node_orientation
                                                     parent = lo_element4 ).
              lo_element5->set_attribute_ns( name  = 'val'
                                             value = ls_ax-orientation ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_delete
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-delete ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majorgridlines
                                                     parent = lo_element3 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_numfmt
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'formatCode'
                                             value = ls_ax-formatcode ).
              lo_element4->set_attribute_ns( name  = 'sourceLinked'
                                             value = ls_ax-sourcelinked ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-majortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_minortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-minortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ticklblpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-ticklblpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossax
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossax ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crosses
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crosses ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossbetween
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossbetween ).
            when others.
          endcase.
        endloop.
        "endaxes

      when zcl_excel_drawing=>c_graph_pie.
        "----pie
        lo_element3 = lo_document->create_simple_element( name = lc_xml_node_piechart
                                                     parent = lo_element2 ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_varycolors
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_varycolorsval ).

        "series
        loop at lo_chartp->series into ls_serie.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ser
                                                     parent = lo_element3 ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_idx
                                                     parent = lo_element4 ).
          if ls_serie-idx is not initial.
            lv_str = ls_serie-idx.
          else.
            lv_str = sy-tabix - 1.
          endif.
          condense lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_order
                                                     parent = lo_element4 ).
          lv_str = ls_serie-order.
          condense lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          if ls_serie-sername is not initial.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_tx
                                                      parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_v
                                                      parent = lo_element5 ).
            lo_element6->set_value( value = ls_serie-sername ).
          endif.
          if ls_serie-lbl is not initial.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_cat
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_strref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-lbl ).
          endif.
          if ls_serie-ref is not initial.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_val
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_numref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-ref ).
          endif.
        endloop.
        "endseries

        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_dlbls
                                                     parent = lo_element3 ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showlegendkey
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showlegendkeyval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showval
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showvalval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showcatname
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showcatnameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showsername
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showsernameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showpercent
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showpercentval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showbubblesize
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showbubblesizeval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showleaderlines
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showleaderlinesval ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_firstsliceang
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_firstsliceangval ).
      when zcl_excel_drawing=>c_graph_line.
        "----line
        lo_element3 = lo_document->create_simple_element( name = lc_xml_node_linechart
                                                     parent = lo_element2 ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_grouping
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_groupingval ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_varycolors
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_varycolorsval ).

        "series
        loop at lo_chartl->series into ls_serie.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ser
                                                     parent = lo_element3 ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_idx
                                                     parent = lo_element4 ).
          if ls_serie-idx is not initial.
            lv_str = ls_serie-idx.
          else.
            lv_str = sy-tabix - 1.
          endif.
          condense lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_order
                                                     parent = lo_element4 ).
          lv_str = ls_serie-order.
          condense lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          if ls_serie-sername is not initial.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_tx
                                                      parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_v
                                                      parent = lo_element5 ).
            lo_element6->set_value( value = ls_serie-sername ).
          endif.
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_marker
                                                     parent = lo_element4 ).
          lo_element6 = lo_document->create_simple_element( name = lc_xml_node_symbol
                                                     parent = lo_element5 ).
          lo_element6->set_attribute_ns( name  = 'val'
                                  value = ls_serie-symbol ).
          if ls_serie-lbl is not initial.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_cat
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_strref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-lbl ).
          endif.
          if ls_serie-ref is not initial.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_val
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_numref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-ref ).
          endif.
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_smooth
                                                       parent = lo_element4 ).
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = ls_serie-smooth ).
        endloop.
        "endseries

        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_dlbls
                                                     parent = lo_element3 ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showlegendkey
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showlegendkeyval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showval
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showvalval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showcatname
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showcatnameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showsername
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showsernameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showpercent
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showpercentval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showbubblesize
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showbubblesizeval ).

        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_marker
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_markerval ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_smooth
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_smoothval ).

        "axes
        lo_el_rootchart = lo_element3.
        loop at lo_chartl->axes into ls_ax.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_el_rootchart ).
          lo_element4->set_attribute_ns( name  = 'val'
                                  value = ls_ax-axid ).
          case ls_ax-type.
            when zcl_excel_graph_line=>c_catax.
              lo_element3 = lo_document->create_simple_element( name = lc_xml_node_catax
                                                     parent = lo_element2 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axid ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_scaling
                                                     parent = lo_element3 ).
              lo_element5 = lo_document->create_simple_element( name = lc_xml_node_orientation
                                                     parent = lo_element4 ).
              lo_element5->set_attribute_ns( name  = 'val'
                                             value = ls_ax-orientation ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_delete
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-delete ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-majortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_minortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-minortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ticklblpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-ticklblpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossax
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossax ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crosses
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crosses ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_auto
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-auto ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_lblalgn
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-lblalgn ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_lbloffset
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-lbloffset ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_nomultilvllbl
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-nomultilvllbl ).
            when zcl_excel_graph_line=>c_valax.
              lo_element3 = lo_document->create_simple_element( name = lc_xml_node_valax
                                                     parent = lo_element2 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axid ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_scaling
                                                     parent = lo_element3 ).
              lo_element5 = lo_document->create_simple_element( name = lc_xml_node_orientation
                                                     parent = lo_element4 ).
              lo_element5->set_attribute_ns( name  = 'val'
                                             value = ls_ax-orientation ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_delete
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-delete ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majorgridlines
                                                     parent = lo_element3 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_numfmt
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'formatCode'
                                             value = ls_ax-formatcode ).
              lo_element4->set_attribute_ns( name  = 'sourceLinked'
                                             value = ls_ax-sourcelinked ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-majortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_minortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-minortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ticklblpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-ticklblpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossax
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossax ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crosses
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crosses ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossbetween
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossbetween ).
            when others.
          endcase.
        endloop.
        "endaxes

      when others.
    endcase.

    "legend
    if lo_chart->print_label eq abap_true.
      lo_element2 = lo_document->create_simple_element( name = lc_xml_node_legend
                                                         parent = lo_element ).
      case io_drawing->graph_type.
        when zcl_excel_drawing=>c_graph_bars.
          "----bar
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_legendpos
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartb->ns_legendposval ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_layout
                                                       parent = lo_element2 ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_overlay
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartb->ns_overlayval ).
        when zcl_excel_drawing=>c_graph_line.
          "----line
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_legendpos
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartl->ns_legendposval ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_layout
                                                       parent = lo_element2 ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_overlay
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartl->ns_overlayval ).
        when zcl_excel_drawing=>c_graph_pie.
          "----pie
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_legendpos
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartp->ns_legendposval ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_layout
                                                       parent = lo_element2 ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_overlay
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartp->ns_overlayval ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_txpr
                                                       parent = lo_element2 ).
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_bodypr
                                                       parent = lo_element3 ).
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_lststyle
                                                       parent = lo_element3 ).
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_p
                                                       parent = lo_element3 ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_ppr
                                                       parent = lo_element4 ).
          lo_element5->set_attribute_ns( name  = 'rtl'
                                    value = lo_chartp->ns_pprrtl ).
          lo_element6 = lo_document->create_simple_element( name = lc_xml_node_defrpr
                                                       parent = lo_element5 ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_endpararpr
                                                       parent = lo_element4 ).
          lo_element5->set_attribute_ns( name  = 'lang'
                                    value = lo_chartp->ns_endpararprlang ).
        when others.
      endcase.
    endif.

    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_plotvisonly
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_plotvisonlyval ).
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_dispblanksas
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_dispblanksasval ).
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_showdlblsovermax
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_showdlblsovermaxval ).
    "---------------------------END OF CHART

    "printSettings
    lo_element = lo_document->create_simple_element( name = lc_xml_node_printsettings
                                                         parent = lo_element_root ).
    "headerFooter
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_headerfooter
                                                         parent = lo_element ).
    "pageMargins
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_pagemargins
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'b'
                                      value = lo_chart->pagemargins-b ).
    lo_element2->set_attribute_ns( name  = 'l'
                                      value = lo_chart->pagemargins-l ).
    lo_element2->set_attribute_ns( name  = 'r'
                                      value = lo_chart->pagemargins-r ).
    lo_element2->set_attribute_ns( name  = 't'
                                      value = lo_chart->pagemargins-t ).
    lo_element2->set_attribute_ns( name  = 'header'
                                      value = lo_chart->pagemargins-header ).
    lo_element2->set_attribute_ns( name  = 'footer'
                                      value = lo_chart->pagemargins-footer ).
    "pageSetup
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_pagesetup
                                                         parent = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  endmethod.


  method create_xl_comments.
    data:
      lo_comment             type ref to zcl_excel_comment,
      lo_comments            type ref to zcl_excel_comments,
      lo_document            type ref to if_ixml_document,
      lo_element_author      type ref to if_ixml_element,
      lo_element_authors     type ref to if_ixml_element,
      lo_element_b           type ref to if_ixml_element,
      lo_element_comment     type ref to if_ixml_element,
      lo_element_commentlist type ref to if_ixml_element,
      lo_element_r           type ref to if_ixml_element,
      lo_element_root        type ref to if_ixml_element,
      lo_element_rpr         type ref to if_ixml_element,
      lo_element_t           type ref to if_ixml_element,
      lo_element_text        type ref to if_ixml_element,
      lo_iterator            type ref to zcl_excel_collection_iterator,
      lv_author              type string.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

***********************************************************************
* STEP 3: Create main node relationships
    lo_element_root = lo_document->create_simple_element( name   = `comments`
                                                          parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = `xmlns`
                                       value = `http://schemas.openxmlformats.org/spreadsheetml/2006/main` ).

**********************************************************************
* STEP 4: Create authors
* TO-DO: management of several authors
    lo_element_authors = lo_document->create_simple_element( name   = `authors`
                                                             parent = lo_element_root ).

    lo_element_author  = lo_document->create_simple_element( name   = `author`
                                                             parent = lo_element_authors ).
    lv_author = sy-uname.
    lo_element_author->set_value( lv_author ).

**********************************************************************
* STEP 5: Create comments

    lo_element_commentlist = lo_document->create_simple_element( name   = `commentList`
                                                                 parent = lo_element_root ).

    lo_comments = io_worksheet->get_comments( ).

    lo_iterator = lo_comments->get_iterator( ).
    while lo_iterator->has_next( ) eq abap_true.
      lo_comment ?= lo_iterator->get_next( ).

      lo_element_comment = lo_document->create_simple_element( name   = `comment`
                                                               parent = lo_element_commentlist ).
      lo_element_comment->set_attribute_ns( name  = `ref`
                                            value = lo_comment->get_ref( ) ).
      lo_element_comment->set_attribute_ns( name  = `authorId`
                                            value = `0` ).  " TO-DO

      lo_element_text = lo_document->create_simple_element( name   = `text`
                                                            parent = lo_element_comment ).
      lo_element_r    = lo_document->create_simple_element( name   = `r`
                                                            parent = lo_element_text ).
      lo_element_rpr  = lo_document->create_simple_element( name   = `rPr`
                                                            parent = lo_element_r ).

      lo_element_b    = lo_document->create_simple_element( name   = `b`
                                                            parent = lo_element_rpr ).

      add_1_val_child_node( io_document   = lo_document
                            io_parent     = lo_element_rpr
                            iv_elem_name  = `sz`
                            iv_attr_name  = `val`
                            iv_attr_value = `9` ).
      add_1_val_child_node( io_document   = lo_document
                            io_parent     = lo_element_rpr
                            iv_elem_name  = `color`
                            iv_attr_name  = `indexed`
                            iv_attr_value = `81` ).
      add_1_val_child_node( io_document   = lo_document
                            io_parent     = lo_element_rpr
                            iv_elem_name  = `rFont`
                            iv_attr_name  = `val`
                            iv_attr_value = `Tahoma` ).
      add_1_val_child_node( io_document   = lo_document
                            io_parent     = lo_element_rpr
                            iv_elem_name  = `family`
                            iv_attr_name  = `val`
                            iv_attr_value = `2` ).

      lo_element_t    = lo_document->create_simple_element( name   = `t`
                                                            parent = lo_element_r ).
      lo_element_t->set_attribute_ns( name  = `xml:space`
                                      value = `preserve` ).
      lo_element_t->set_value( lo_comment->get_text( ) ).
    endwhile.

    lo_element_root->append_child( new_child = lo_element_commentlist ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  endmethod.


  method create_xl_drawings.


** Constant node name
    constants: lc_xml_node_wsdr   type string value 'xdr:wsDr',
               lc_xml_node_ns_xdr type string value 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
               lc_xml_node_ns_a   type string value 'http://schemas.openxmlformats.org/drawingml/2006/main'.

    data: lo_document           type ref to if_ixml_document,
          lo_element_root       type ref to if_ixml_element,
          lo_element_cellanchor type ref to if_ixml_element,
          lo_iterator           type ref to zcl_excel_collection_iterator,
          lo_drawings           type ref to zcl_excel_drawings,
          lo_drawing            type ref to zcl_excel_drawing.
    data: lv_rel_id            type i.



**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

***********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_wsdr
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:xdr'
                                       value = lc_xml_node_ns_xdr ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:a'
                                       value = lc_xml_node_ns_a ).

**********************************************************************
* STEP 4: Create drawings

    clear: lv_rel_id.

    lo_drawings = io_worksheet->get_drawings( ).

    lo_iterator = lo_drawings->get_iterator( ).
    while lo_iterator->has_next( ) eq abap_true.
      lo_drawing ?= lo_iterator->get_next( ).

      lv_rel_id += 1.
      lo_element_cellanchor = me->create_xl_drawing_anchor(
              io_drawing    = lo_drawing
              io_document   = lo_document
              ip_index      = lv_rel_id ).

      lo_element_root->append_child( new_child = lo_element_cellanchor ).

    endwhile.

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  endmethod.


  method create_xl_drawings_hdft_rels.

** Constant node name
    data: lc_xml_node_relationships type string value 'Relationships',
          lc_xml_node_relationship  type string value 'Relationship',
          " Node attributes
          lc_xml_attr_id            type string value 'Id',
          lc_xml_attr_type          type string value 'Type',
          lc_xml_attr_target        type string value 'Target',
          " Node namespace
          lc_xml_node_rels_ns       type string value 'http://schemas.openxmlformats.org/package/2006/relationships',
          lc_xml_node_rid_image_tp  type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          lc_xml_node_rid_chart_tp  type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'.

    types: begin of ty_temp,
             row_index type i,
             str       type string,
           end of ty_temp.

    data: lo_drawing      type ref to zcl_excel_drawing,
          lo_document     type ref to if_ixml_document,
          lo_element_root type ref to if_ixml_element,
          lo_element      type ref to if_ixml_element,
          lv_value        type string,
          lv_relation_id  type i,
          lt_temp         type standard table of ty_temp with default key,
          lt_drawings     type zif_excel_data_decl=>zexcel_t_drawings.

    field-symbols: <fs_temp>     like line of lt_temp,
                   <fs_drawings> type zif_excel_data_decl=>zexcel_s_drawings.


* BODY
**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_relationships
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_rels_ns ).

**********************************************************************
* STEP 4: Create subnodes

**********************************************************************


    lt_drawings = io_worksheet->get_header_footer_drawings( ).
    loop at lt_drawings assigning <fs_drawings>. "Header or footer image exist
      lv_relation_id += 1.
      lv_value = <fs_drawings>-drawing->get_index( ).
      read table lt_temp with key str = lv_value transporting no fields.
      if sy-subrc ne 0.
        append initial line to lt_temp assigning <fs_temp>.
        <fs_temp>-row_index = sy-tabix.
        <fs_temp>-str = lv_value.
        condense lv_value.
        concatenate 'rId' lv_value into lv_value.
        lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                           parent = lo_document ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                      value = lv_value ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                      value = lc_xml_node_rid_image_tp ).

        lv_value = '../media/#'.
        replace '#' in lv_value with <fs_drawings>-drawing->get_media_name( ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                      value = lv_value ).
        lo_element_root->append_child( new_child = lo_element ).
      endif.
    endloop.

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  endmethod.                    "create_xl_drawings_hdft_rels


  method create_xl_drawings_rels.

** Constant node name
    data: lc_xml_node_relationships type string value 'Relationships',
          lc_xml_node_relationship  type string value 'Relationship',
          " Node attributes
          lc_xml_attr_id            type string value 'Id',
          lc_xml_attr_type          type string value 'Type',
          lc_xml_attr_target        type string value 'Target',
          " Node namespace
          lc_xml_node_rels_ns       type string value 'http://schemas.openxmlformats.org/package/2006/relationships',
          lc_xml_node_rid_image_tp  type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          lc_xml_node_rid_chart_tp  type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'.

    data: lo_document     type ref to if_ixml_document,
          lo_element_root type ref to if_ixml_element,
          lo_element      type ref to if_ixml_element,
          lo_iterator     type ref to zcl_excel_collection_iterator,
          lo_drawings     type ref to zcl_excel_drawings,
          lo_drawing      type ref to zcl_excel_drawing.

    data: lv_value   type string,
          lv_counter type i.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_relationships
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_rels_ns ).

**********************************************************************
* STEP 4: Create subnodes

    " Add sheet Relationship nodes here
    lv_counter = 0.
    lo_drawings = io_worksheet->get_drawings( ).
    lo_iterator = lo_drawings->get_iterator( ).
    while lo_iterator->has_next( ) eq abap_true.
      lo_drawing ?= lo_iterator->get_next( ).
      lv_counter += 1.

      lv_value = lv_counter.
      condense lv_value.
      concatenate 'rId' lv_value into lv_value.

      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                   parent = lo_document ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).

      lv_value = lo_drawing->get_media_name( ).
      case lo_drawing->get_type( ).
        when zcl_excel_drawing=>type_image.
          concatenate '../media/' lv_value into lv_value.
          lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                        value = lc_xml_node_rid_image_tp ).

        when zcl_excel_drawing=>type_chart.
          concatenate '../charts/' lv_value into lv_value.
          lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                        value = lc_xml_node_rid_chart_tp ).

      endcase.
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                    value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    endwhile.


**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  endmethod.


  method create_xl_drawings_vml.

    data:
      ld_stream       type string.


* INIT_RESULT
    clear ep_content.


* BODY
    ld_stream = set_vml_string( ).
    ep_content = xco_cp=>string( ld_stream )->as_xstring( xco_cp_character=>code_page->utf_8 )->value.
*    call function 'SCMS_STRING_TO_XSTRING'
*      exporting
*        text   = ld_stream
*      importing
*        buffer = ep_content
*      exceptions
*        failed = 1
*        others = 2.
*    if sy-subrc <> 0.
*      clear ep_content.
*    endif.


  endmethod.


  method create_xl_drawings_vml_rels.

** Constant node name
    data: lc_xml_node_relationships type string value 'Relationships',
          lc_xml_node_relationship  type string value 'Relationship',
          " Node attributes
          lc_xml_attr_id            type string value 'Id',
          lc_xml_attr_type          type string value 'Type',
          lc_xml_attr_target        type string value 'Target',
          " Node namespace
          lc_xml_node_rels_ns       type string value 'http://schemas.openxmlformats.org/package/2006/relationships',
          lc_xml_node_rid_image_tp  type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          lc_xml_node_rid_chart_tp  type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'.

    data: lo_iterator     type ref to zcl_excel_collection_iterator,
          lo_drawing      type ref to zcl_excel_drawing,
          lo_document     type ref to if_ixml_document,
          lo_element_root type ref to if_ixml_element,
          lo_element      type ref to if_ixml_element,
          lv_value        type string,
          lv_relation_id  type i.


* BODY
**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_relationships
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_rels_ns ).

**********************************************************************
* STEP 4: Create subnodes
    lv_relation_id = 0.
    lo_iterator = me->excel->get_drawings_iterator( zcl_excel_drawing=>type_image ).
    while lo_iterator->has_next( ) eq abap_true.
      lo_drawing ?= lo_iterator->get_next( ).
      if lo_drawing->get_type( ) = zcl_excel_drawing=>type_image_header_footer.
        lv_relation_id += 1.
        lv_value = lv_relation_id.
        condense lv_value.
        concatenate 'rId' lv_value into lv_value.
        lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                           parent = lo_document ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_id
*                                    value = 'LOGO' ).
                                      value = lv_value ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                      value = lc_xml_node_rid_image_tp ).

        lv_value = '../media/#'.
        replace '#' in lv_value with lo_drawing->get_media_name( ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_target
*                                    value = '../media/LOGO.png' ).
                                      value = lv_value ).
        lo_element_root->append_child( new_child = lo_element ).
      endif.

    endwhile.



**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  endmethod.


  method create_xl_drawing_anchor.

** Constant node name
    constants: lc_xml_node_onecellanchor     type string value 'xdr:oneCellAnchor',
               lc_xml_node_twocellanchor     type string value 'xdr:twoCellAnchor',
               lc_xml_node_from              type string value 'xdr:from',
               lc_xml_node_to                type string value 'xdr:to',
               lc_xml_node_pic               type string value 'xdr:pic',
               lc_xml_node_ext               type string value 'xdr:ext',
               lc_xml_node_clientdata        type string value 'xdr:clientData',

               lc_xml_node_col               type string value 'xdr:col',
               lc_xml_node_coloff            type string value 'xdr:colOff',
               lc_xml_node_row               type string value 'xdr:row',
               lc_xml_node_rowoff            type string value 'xdr:rowOff',

               lc_xml_node_nvpicpr           type string value 'xdr:nvPicPr',
               lc_xml_node_cnvpr             type string value 'xdr:cNvPr',
               lc_xml_node_cnvpicpr          type string value 'xdr:cNvPicPr',
               lc_xml_node_piclocks          type string value 'a:picLocks',

               lc_xml_node_sppr              type string value 'xdr:spPr',
               lc_xml_node_apgeom            type string value 'a:prstGeom',
               lc_xml_node_aavlst            type string value 'a:avLst',

               lc_xml_node_graphicframe      type string value 'xdr:graphicFrame',
               lc_xml_node_nvgraphicframepr  type string value 'xdr:nvGraphicFramePr',
               lc_xml_node_cnvgraphicframepr type string value 'xdr:cNvGraphicFramePr',
               lc_xml_node_graphicframelocks type string value 'a:graphicFrameLocks',
               lc_xml_node_xfrm              type string value 'xdr:xfrm',
               lc_xml_node_aoff              type string value 'a:off',
               lc_xml_node_aext              type string value 'a:ext',
               lc_xml_node_agraphic          type string value 'a:graphic',
               lc_xml_node_agraphicdata      type string value 'a:graphicData',

               lc_xml_node_ns_c              type string value 'http://schemas.openxmlformats.org/drawingml/2006/chart',
               lc_xml_node_cchart            type string value 'c:chart',

               lc_xml_node_blipfill          type string value 'xdr:blipFill',
               lc_xml_node_ablip             type string value 'a:blip',
               lc_xml_node_astretch          type string value 'a:stretch',
               lc_xml_node_ns_r              type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'.

    data: lo_element_graphicframe type ref to if_ixml_element,
          lo_element              type ref to if_ixml_element,
          lo_element2             type ref to if_ixml_element,
          lo_element3             type ref to if_ixml_element,
          lo_element_from         type ref to if_ixml_element,
          lo_element_to           type ref to if_ixml_element,
          lo_element_ext          type ref to if_ixml_element,
          lo_element_pic          type ref to if_ixml_element,
          lo_element_clientdata   type ref to if_ixml_element,
          ls_position             type zif_excel_data_decl=>zexcel_drawing_position,
          lv_col                  type string, " zexcel_cell_column,
          lv_row                  type string, " zexcel_cell_row.
          lv_col_offset           type string,
          lv_row_offset           type string,
          lv_value                type string.

    ls_position = io_drawing->get_position( ).

    if ls_position-anchor = 'ONE'.
      ep_anchor = io_document->create_simple_element( name   = lc_xml_node_onecellanchor
                                                                  parent = io_document ).
    else.
      ep_anchor = io_document->create_simple_element( name   = lc_xml_node_twocellanchor
                                                                  parent = io_document ).
    endif.

*   from cell ******************************
    lo_element_from = io_document->create_simple_element( name   = lc_xml_node_from
                                                          parent = io_document ).

    lv_col = ls_position-from-col.
    lv_row = ls_position-from-row.
    lv_col_offset = ls_position-from-col_offset.
    lv_row_offset = ls_position-from-row_offset.
    condense lv_col no-gaps.
    condense lv_row no-gaps.
    condense lv_col_offset no-gaps.
    condense lv_row_offset no-gaps.

    lo_element = io_document->create_simple_element( name = lc_xml_node_col
                                                     parent = io_document ).
    lo_element->set_value( value = lv_col ).
    lo_element_from->append_child( new_child = lo_element ).

    lo_element = io_document->create_simple_element( name = lc_xml_node_coloff
                                                     parent = io_document ).
    lo_element->set_value( value = lv_col_offset ).
    lo_element_from->append_child( new_child = lo_element ).

    lo_element = io_document->create_simple_element( name = lc_xml_node_row
                                                     parent = io_document ).
    lo_element->set_value( value = lv_row ).
    lo_element_from->append_child( new_child = lo_element ).

    lo_element = io_document->create_simple_element( name = lc_xml_node_rowoff
                                                     parent = io_document ).
    lo_element->set_value( value = lv_row_offset ).
    lo_element_from->append_child( new_child = lo_element ).
    ep_anchor->append_child( new_child = lo_element_from ).

    if ls_position-anchor = 'ONE'.

*   ext ******************************
      lo_element_ext = io_document->create_simple_element( name   = lc_xml_node_ext
                                                           parent = io_document ).

      lv_value = io_drawing->get_width_emu_str( ).
      lo_element_ext->set_attribute_ns( name  = 'cx'
                                     value = lv_value ).
      lv_value = io_drawing->get_height_emu_str( ).
      lo_element_ext->set_attribute_ns( name  = 'cy'
                                     value = lv_value ).
      ep_anchor->append_child( new_child = lo_element_ext ).

    elseif ls_position-anchor = 'TWO'.

*   to cell ******************************
      lo_element_to = io_document->create_simple_element( name   = lc_xml_node_to
                                                          parent = io_document ).

      lv_col = ls_position-to-col.
      lv_row = ls_position-to-row.
      lv_col_offset = ls_position-to-col_offset.
      lv_row_offset = ls_position-to-row_offset.
      condense lv_col no-gaps.
      condense lv_row no-gaps.
      condense lv_col_offset no-gaps.
      condense lv_row_offset no-gaps.

      lo_element = io_document->create_simple_element( name = lc_xml_node_col
                                                       parent = io_document ).
      lo_element->set_value( value = lv_col ).
      lo_element_to->append_child( new_child = lo_element ).

      lo_element = io_document->create_simple_element( name = lc_xml_node_coloff
                                                       parent = io_document ).
      lo_element->set_value( value = lv_col_offset ).
      lo_element_to->append_child( new_child = lo_element ).

      lo_element = io_document->create_simple_element( name = lc_xml_node_row
                                                       parent = io_document ).
      lo_element->set_value( value = lv_row ).
      lo_element_to->append_child( new_child = lo_element ).

      lo_element = io_document->create_simple_element( name = lc_xml_node_rowoff
                                                       parent = io_document ).
      lo_element->set_value( value = lv_row_offset ).
      lo_element_to->append_child( new_child = lo_element ).
      ep_anchor->append_child( new_child = lo_element_to ).

    endif.

    case io_drawing->get_type( ).
      when zcl_excel_drawing=>type_image.
*     pic **********************************
        lo_element_pic = io_document->create_simple_element( name   = lc_xml_node_pic
                                                             parent = io_document ).
*     nvPicPr
        lo_element  = io_document->create_simple_element( name = lc_xml_node_nvpicpr
                                                          parent = io_document ).
*     cNvPr
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_cnvpr
                                                          parent = io_document ).
        lv_value = sy-index.
        condense lv_value.
        lo_element2->set_attribute_ns( name  = 'id'
                                       value = lv_value ).
        lo_element2->set_attribute_ns( name  = 'name'
                                       value = io_drawing->title ).
        lo_element->append_child( new_child = lo_element2 ).

*     cNvPicPr
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_cnvpicpr
                                                          parent = io_document ).

*     picLocks
        lo_element3 = io_document->create_simple_element( name = lc_xml_node_piclocks
                                                          parent = io_document ).
        lo_element3->set_attribute_ns( name  = 'noChangeAspect'
                                       value = '1' ).

        lo_element2->append_child( new_child = lo_element3 ).
        lo_element->append_child( new_child = lo_element2 ).
        lo_element_pic->append_child( new_child = lo_element ).

*     blipFill
        lv_value = ip_index.
        condense lv_value.
        concatenate 'rId' lv_value into lv_value.

        lo_element  = io_document->create_simple_element( name = lc_xml_node_blipfill
                                                          parent = io_document ).
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_ablip
                                                          parent = io_document ).
        lo_element2->set_attribute_ns( name  = 'xmlns:r'
                                       value = lc_xml_node_ns_r ).
        lo_element2->set_attribute_ns( name  = 'r:embed'
                                       value = lv_value ).
        lo_element->append_child( new_child = lo_element2 ).

        lo_element2  = io_document->create_simple_element( name = lc_xml_node_astretch
                                                          parent = io_document ).
        lo_element->append_child( new_child = lo_element2 ).

        lo_element_pic->append_child( new_child = lo_element ).

*     spPr
        lo_element  = io_document->create_simple_element( name = lc_xml_node_sppr
                                                          parent = io_document ).

        lo_element2 = io_document->create_simple_element( name = lc_xml_node_apgeom
                                                          parent = io_document ).
        lo_element2->set_attribute_ns( name  = 'prst'
                                       value = 'rect' ).
        lo_element3 = io_document->create_simple_element( name = lc_xml_node_aavlst
                                                          parent = io_document ).
        lo_element2->append_child( new_child = lo_element3 ).
        lo_element->append_child( new_child = lo_element2 ).

        lo_element_pic->append_child( new_child = lo_element ).
        ep_anchor->append_child( new_child = lo_element_pic ).
      when zcl_excel_drawing=>type_chart.
*     graphicFrame **********************************
        lo_element_graphicframe = io_document->create_simple_element( name   = lc_xml_node_graphicframe
                                                             parent = io_document ).
*     nvGraphicFramePr
        lo_element  = io_document->create_simple_element( name = lc_xml_node_nvgraphicframepr
                                                          parent = io_document ).
*     cNvPr
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_cnvpr
                                                          parent = io_document ).
        lv_value = sy-index.
        condense lv_value.
        lo_element2->set_attribute_ns( name  = 'id'
                                       value = lv_value ).
        lo_element2->set_attribute_ns( name  = 'name'
                                       value = io_drawing->title ).
        lo_element->append_child( new_child = lo_element2 ).
*     cNvGraphicFramePr
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_cnvgraphicframepr
                                                          parent = io_document ).
        lo_element3 = io_document->create_simple_element( name = lc_xml_node_graphicframelocks
                                                          parent = io_document ).
        lo_element2->append_child( new_child = lo_element3 ).
        lo_element->append_child( new_child = lo_element2 ).
        lo_element_graphicframe->append_child( new_child = lo_element ).

*     xfrm
        lo_element  = io_document->create_simple_element( name = lc_xml_node_xfrm
                                                          parent = io_document ).
*     off
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_aoff
                                                          parent = io_document ).
        lo_element2->set_attribute_ns( name  = 'y' value = '0' ).
        lo_element2->set_attribute_ns( name  = 'x' value = '0' ).
        lo_element->append_child( new_child = lo_element2 ).
*     ext
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_aext
                                                          parent = io_document ).
        lo_element2->set_attribute_ns( name  = 'cy' value = '0' ).
        lo_element2->set_attribute_ns( name  = 'cx' value = '0' ).
        lo_element->append_child( new_child = lo_element2 ).
        lo_element_graphicframe->append_child( new_child = lo_element ).

*     graphic
        lo_element  = io_document->create_simple_element( name = lc_xml_node_agraphic
                                                          parent = io_document ).
*     graphicData
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_agraphicdata
                                                          parent = io_document ).
        lo_element2->set_attribute_ns( name  = 'uri' value = lc_xml_node_ns_c ).

*     chart
        lo_element3 = io_document->create_simple_element( name = lc_xml_node_cchart
                                                          parent = io_document ).

        lo_element3->set_attribute_ns( name  = 'xmlns:r'
                                       value = lc_xml_node_ns_r ).
        lo_element3->set_attribute_ns( name  = 'xmlns:c'
                                       value = lc_xml_node_ns_c ).

        lv_value = ip_index.
        condense lv_value.
        concatenate 'rId' lv_value into lv_value.
        lo_element3->set_attribute_ns( name  = 'r:id'
                                       value = lv_value ).
        lo_element2->append_child( new_child = lo_element3 ).
        lo_element->append_child( new_child = lo_element2 ).
        lo_element_graphicframe->append_child( new_child = lo_element ).
        ep_anchor->append_child( new_child = lo_element_graphicframe ).

    endcase.

*   client data ***************************
    lo_element_clientdata = io_document->create_simple_element( name   = lc_xml_node_clientdata
                                                                parent = io_document ).
    ep_anchor->append_child( new_child = lo_element_clientdata ).

  endmethod.


  method create_xl_drawing_for_comments.
** Constant node name
    constants: lc_xml_node_xml             type string value 'xml',
               lc_xml_node_ns_v            type string value 'urn:schemas-microsoft-com:vml',
               lc_xml_node_ns_o            type string value 'urn:schemas-microsoft-com:office:office',
               lc_xml_node_ns_x            type string value 'urn:schemas-microsoft-com:office:excel',
               " shapelayout
               lc_xml_node_shapelayout     type string value 'o:shapelayout',
               lc_xml_node_idmap           type string value 'o:idmap',
               " shapetype
               lc_xml_node_shapetype       type string value 'v:shapetype',
               lc_xml_node_stroke          type string value 'v:stroke',
               lc_xml_node_path            type string value 'v:path',
               " shape
               lc_xml_node_shape           type string value 'v:shape',
               lc_xml_node_fill            type string value 'v:fill',
               lc_xml_node_shadow          type string value 'v:shadow',
               lc_xml_node_textbox         type string value 'v:textbox',
               lc_xml_node_div             type string value 'div',
               lc_xml_node_clientdata      type string value 'x:ClientData',
               lc_xml_node_movewithcells   type string value 'x:MoveWithCells',
               lc_xml_node_sizewithcells   type string value 'x:SizeWithCells',
               lc_xml_node_anchor          type string value 'x:Anchor',
               lc_xml_node_autofill        type string value 'x:AutoFill',
               lc_xml_node_row             type string value 'x:Row',
               lc_xml_node_column          type string value 'x:Column',
               " attributes,
               lc_xml_attr_vext            type string value 'v:ext',
               lc_xml_attr_data            type string value 'data',
               lc_xml_attr_id              type string value 'id',
               lc_xml_attr_coordsize       type string value 'coordsize',
               lc_xml_attr_ospt            type string value 'o:spt',
               lc_xml_attr_joinstyle       type string value 'joinstyle',
               lc_xml_attr_path            type string value 'path',
               lc_xml_attr_gradientshapeok type string value 'gradientshapeok',
               lc_xml_attr_oconnecttype    type string value 'o:connecttype',
               lc_xml_attr_type            type string value 'type',
               lc_xml_attr_style           type string value 'style',
               lc_xml_attr_fillcolor       type string value 'fillcolor',
               lc_xml_attr_oinsetmode      type string value 'o:insetmode',
               lc_xml_attr_color           type string value 'color',
               lc_xml_attr_color2          type string value 'color2',
               lc_xml_attr_on              type string value 'on',
               lc_xml_attr_obscured        type string value 'obscured',
               lc_xml_attr_objecttype      type string value 'ObjectType',
               " attributes values
               lc_xml_attr_val_edit        type string value 'edit',
               lc_xml_attr_val_rect        type string value 'rect',
               lc_xml_attr_val_t           type string value 't',
               lc_xml_attr_val_miter       type string value 'miter',
               lc_xml_attr_val_auto        type string value 'auto',
               lc_xml_attr_val_black       type string value 'black',
               lc_xml_attr_val_none        type string value 'none',
               lc_xml_attr_val_msodir      type string value 'mso-direction-alt:auto',
               lc_xml_attr_val_note        type string value 'Note'.


    data: lo_document              type ref to if_ixml_document,
          lo_element_root          type ref to if_ixml_element,
          "shapelayout
          lo_element_shapelayout   type ref to if_ixml_element,
          lo_element_idmap         type ref to if_ixml_element,
          "shapetype
          lo_element_shapetype     type ref to if_ixml_element,
          lo_element_stroke        type ref to if_ixml_element,
          lo_element_path          type ref to if_ixml_element,
          "shape
          lo_element_shape         type ref to if_ixml_element,
          lo_element_fill          type ref to if_ixml_element,
          lo_element_shadow        type ref to if_ixml_element,
          lo_element_textbox       type ref to if_ixml_element,
          lo_element_div           type ref to if_ixml_element,
          lo_element_clientdata    type ref to if_ixml_element,
          lo_element_movewithcells type ref to if_ixml_element,
          lo_element_sizewithcells type ref to if_ixml_element,
          lo_element_anchor        type ref to if_ixml_element,
          lo_element_autofill      type ref to if_ixml_element,
          lo_element_row           type ref to if_ixml_element,
          lo_element_column        type ref to if_ixml_element,
          lo_iterator              type ref to zcl_excel_collection_iterator,
          lo_comments              type ref to zcl_excel_comments,
          lo_comment               type ref to zcl_excel_comment,
          lv_row                   type zif_excel_data_decl=>zexcel_cell_row,
          lv_str_column            type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_column                type zif_excel_data_decl=>zexcel_cell_column,
          lv_index                 type i,
          lv_attr_id_index         type i,
          lv_attr_id               type string,
          lv_int_value             type i,
          lv_int_value_string      type string.
    data: lv_rel_id            type i.
    data lv_anchor         type string.
    data lv_bottom_row     type i.
    data lv_right_column   type i.
    data lv_bottom_row_str type string.
    data lv_right_column_str  type string.
    data lv_top_row         type i.
    data lv_left_column     type i.
    data lv_top_row_str     type string.
    data lv_left_column_str type string.


**********************************************************************
* STEP 1: Create XML document
    lo_document = ixml->if_ixml_core~create_document( ).

***********************************************************************
* STEP 2: Create main node relationships
    lo_element_root = lo_document->create_simple_element( name   = lc_xml_node_xml
                                                          parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:v'  value = lc_xml_node_ns_v ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:o'  value = lc_xml_node_ns_o ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:x'  value = lc_xml_node_ns_x ).

**********************************************************************
* STEP 3: Create o:shapeLayout
* TO-DO: management of several authors
    lo_element_shapelayout = lo_document->create_simple_element( name   = lc_xml_node_shapelayout
                                                                 parent = lo_document ).

    lo_element_shapelayout->set_attribute_ns( name  = lc_xml_attr_vext
                                              value = lc_xml_attr_val_edit ).

    lo_element_idmap = lo_document->create_simple_element( name   = lc_xml_node_idmap
                                                           parent = lo_document ).
    lo_element_idmap->set_attribute_ns( name  = lc_xml_attr_vext  value = lc_xml_attr_val_edit ).
    lo_element_idmap->set_attribute_ns( name  = lc_xml_attr_data  value = '1' ).

    lo_element_shapelayout->append_child( new_child = lo_element_idmap ).

    lo_element_root->append_child( new_child = lo_element_shapelayout ).

**********************************************************************
* STEP 4: Create v:shapetype

    lo_element_shapetype = lo_document->create_simple_element( name   = lc_xml_node_shapetype
                                                               parent = lo_document ).

    lo_element_shapetype->set_attribute_ns( name  = lc_xml_attr_id         value = '_x0000_t202' ).
    lo_element_shapetype->set_attribute_ns( name  = lc_xml_attr_coordsize  value = '21600,21600' ).
    lo_element_shapetype->set_attribute_ns( name  = lc_xml_attr_ospt       value = '202' ).
    lo_element_shapetype->set_attribute_ns( name  = lc_xml_attr_path       value = 'm,l,21600r21600,l21600,xe' ).

    lo_element_stroke = lo_document->create_simple_element( name   = lc_xml_node_stroke
                                                            parent = lo_document ).
    lo_element_stroke->set_attribute_ns( name  = lc_xml_attr_joinstyle       value = lc_xml_attr_val_miter ).

    lo_element_path   = lo_document->create_simple_element( name   = lc_xml_node_path
                                                            parent = lo_document ).
    lo_element_path->set_attribute_ns( name  = lc_xml_attr_gradientshapeok value = lc_xml_attr_val_t ).
    lo_element_path->set_attribute_ns( name  = lc_xml_attr_oconnecttype    value = lc_xml_attr_val_rect ).

    lo_element_shapetype->append_child( new_child = lo_element_stroke ).
    lo_element_shapetype->append_child( new_child = lo_element_path ).

    lo_element_root->append_child( new_child = lo_element_shapetype ).

**********************************************************************
* STEP 4: Create v:shapetype

    lo_comments = io_worksheet->get_comments( ).

    lo_iterator = lo_comments->get_iterator( ).
    while lo_iterator->has_next( ) eq abap_true.
      lv_index = sy-index.
      lo_comment ?= lo_iterator->get_next( ).

      zcl_excel_common=>convert_columnrow2column_a_row( exporting i_columnrow = lo_comment->get_ref( )
                                                        importing e_column = lv_str_column
                                                                  e_row    = lv_row ).
      lv_column = zcl_excel_common=>convert_column2int( lv_str_column ).

      lo_element_shape = lo_document->create_simple_element( name   = lc_xml_node_shape
                                                             parent = lo_document ).

      lv_attr_id_index = 1024 + lv_index.
      lv_attr_id = lv_attr_id_index.
      concatenate '_x0000_s' lv_attr_id into lv_attr_id.
      lo_element_shape->set_attribute_ns( name  = lc_xml_attr_id          value = lv_attr_id ).
      lo_element_shape->set_attribute_ns( name  = lc_xml_attr_type        value = '#_x0000_t202' ).
      lo_element_shape->set_attribute_ns( name  = lc_xml_attr_style       value = 'size:auto;width:auto;height:auto;position:absolute;margin-left:117pt;margin-top:172.5pt;z-index:1;visibility:hidden' ).
      lo_element_shape->set_attribute_ns( name  = lc_xml_attr_fillcolor   value = '#ffffe1' ).
      lo_element_shape->set_attribute_ns( name  = lc_xml_attr_oinsetmode  value = lc_xml_attr_val_auto ).

      " Fill
      lo_element_fill = lo_document->create_simple_element( name   = lc_xml_node_fill
                                                            parent = lo_document ).
      lo_element_fill->set_attribute_ns( name = lc_xml_attr_color2  value = '#ffffe1' ).
      lo_element_shape->append_child( new_child = lo_element_fill ).
      " Shadow
      lo_element_shadow = lo_document->create_simple_element( name   = lc_xml_node_shadow
                                                              parent = lo_document ).
      lo_element_shadow->set_attribute_ns( name = lc_xml_attr_on        value = lc_xml_attr_val_t ).
      lo_element_shadow->set_attribute_ns( name = lc_xml_attr_color     value = lc_xml_attr_val_black ).
      lo_element_shadow->set_attribute_ns( name = lc_xml_attr_obscured  value = lc_xml_attr_val_t ).
      lo_element_shape->append_child( new_child = lo_element_shadow ).
      " Path
      lo_element_path = lo_document->create_simple_element( name   = lc_xml_node_path
                                                            parent = lo_document ).
      lo_element_path->set_attribute_ns( name = lc_xml_attr_oconnecttype  value = lc_xml_attr_val_none ).
      lo_element_shape->append_child( new_child = lo_element_path ).
      " Textbox
      lo_element_textbox = lo_document->create_simple_element( name   = lc_xml_node_textbox
                                                               parent = lo_document ).
      lo_element_textbox->set_attribute_ns( name = lc_xml_attr_style  value = lc_xml_attr_val_msodir ).
      lo_element_div = lo_document->create_simple_element( name   = lc_xml_node_div
                                                           parent = lo_document ).
      lo_element_div->set_attribute_ns( name = lc_xml_attr_style  value = 'text-align:left' ).
      lo_element_textbox->append_child( new_child = lo_element_div ).
      lo_element_shape->append_child( new_child = lo_element_textbox ).
      " ClientData
      lo_element_clientdata = lo_document->create_simple_element( name   = lc_xml_node_clientdata
                                                                  parent = lo_document ).
      lo_element_clientdata->set_attribute_ns( name = lc_xml_attr_objecttype  value = lc_xml_attr_val_note ).
      lo_element_movewithcells = lo_document->create_simple_element( name   = lc_xml_node_movewithcells
                                                                     parent = lo_document ).
      lo_element_clientdata->append_child( new_child = lo_element_movewithcells ).
      lo_element_sizewithcells = lo_document->create_simple_element( name   = lc_xml_node_sizewithcells
                                                                     parent = lo_document ).
      lo_element_clientdata->append_child( new_child = lo_element_sizewithcells ).
      lo_element_anchor = lo_document->create_simple_element( name   = lc_xml_node_anchor
                                                              parent = lo_document ).

      " Anchor represents 4 pairs of numbers:
      "   ( left column, left offset ), ( top row, top offset ),
      "   ( right column, right offset ), ( bottom row, botton offset )
      " Offsets are a number of pixels.
      " Reference: Anchor Class at
      "   https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.spreadsheet.anchor?view=openxml-3.0.1
      lv_anchor = number2string( lo_comment->get_left_column( ) )
       && `, ` && number2string( lo_comment->get_left_offset( ) )
       && `, ` && number2string( lo_comment->get_top_row( ) )
       && `, ` && number2string( lo_comment->get_top_offset( ) )
       && `, ` && number2string( lo_comment->get_right_column( ) )
       && `, ` && number2string( lo_comment->get_right_offset( ) )
       && `, ` && number2string( lo_comment->get_bottom_row( ) )
       && `, ` && number2string( lo_comment->get_bottom_offset( ) ).
      lo_element_anchor->set_value( lv_anchor ).

      lo_element_clientdata->append_child( new_child = lo_element_anchor ).
      lo_element_autofill = lo_document->create_simple_element( name   = lc_xml_node_autofill
                                                                parent = lo_document ).
      lo_element_autofill->set_value( 'False' ).
      lo_element_clientdata->append_child( new_child = lo_element_autofill ).
      lo_element_row = lo_document->create_simple_element( name   = lc_xml_node_row
                                                           parent = lo_document ).
      lv_int_value = lv_row - 1.
      lv_int_value_string = lv_int_value.
      lo_element_row->set_value( lv_int_value_string ).
      lo_element_clientdata->append_child( new_child = lo_element_row ).
      lo_element_column = lo_document->create_simple_element( name   = lc_xml_node_column
                                                                parent = lo_document ).
      lv_int_value = lv_column - 1.
      lv_int_value_string = lv_int_value.
      lo_element_column->set_value( lv_int_value_string ).
      lo_element_clientdata->append_child( new_child = lo_element_column ).

      lo_element_shape->append_child( new_child = lo_element_clientdata ).

      lo_element_root->append_child( new_child = lo_element_shape ).
    endwhile.

**********************************************************************
* STEP 6: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  endmethod.


  method create_xl_drawing_for_hdft_im.


    data:
      ld_1           type string,
      ld_2           type string,
      ld_3           type string,
      ld_4           type string,
      ld_5           type string,
      ld_7           type string,

      ls_odd_header  type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
      ls_odd_footer  type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
      ls_even_header type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
      ls_even_footer type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
      lv_content     type string.


* INIT_RESULT
    clear ep_content.


* BODY
    ld_1 = '<xml xmlns:v="urn:schemas-microsoft-com:vml"  xmlns:o="urn:schemas-microsoft-com:office:office"  xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'.
    ld_2 = '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/>'.
    ld_3 = '<v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/>'.
    ld_4 = '<v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>'.


    concatenate ld_1
                ld_2
                ld_3
                ld_4
         into lv_content.

    io_worksheet->sheet_setup->get_header_footer( importing ep_odd_header = ls_odd_header
                                                            ep_odd_footer = ls_odd_footer
                                                            ep_even_header = ls_even_header
                                                            ep_even_footer = ls_even_footer ).

    ld_5 = me->set_vml_shape_header( ls_odd_header ).
    concatenate lv_content
                ld_5
           into lv_content.
    ld_5 = me->set_vml_shape_header( ls_even_header ).
    concatenate lv_content
                ld_5
           into lv_content.
    ld_5 = me->set_vml_shape_footer( ls_odd_footer ).
    concatenate lv_content
                ld_5
           into lv_content.
    ld_5 = me->set_vml_shape_footer( ls_even_footer ).
    concatenate lv_content
                ld_5
           into lv_content.

    ld_7 = '</xml>'.

    concatenate lv_content
                ld_7
           into lv_content.

    ep_content = xco_cp=>string( lv_content )->as_xstring( xco_cp_character=>code_page->utf_8 )->value.
*    call function 'SCMS_STRING_TO_XSTRING'
*      exporting
*        text   = lv_content
*      importing
*        buffer = ep_content
*      exceptions
*        failed = 1
*        others = 2.
*    if sy-subrc <> 0.
*      clear ep_content.
*    endif.

  endmethod.


  method create_xl_relationships.


** Constant node name
    data: lc_xml_node_relationships type string value 'Relationships',
          lc_xml_node_relationship  type string value 'Relationship',
          " Node attributes
          lc_xml_attr_id            type string value 'Id',
          lc_xml_attr_type          type string value 'Type',
          lc_xml_attr_target        type string value 'Target',
          " Node namespace
          lc_xml_node_rels_ns       type string value 'http://schemas.openxmlformats.org/package/2006/relationships',
          " Node id
          lc_xml_node_ridx_id       type string value 'rId#',
          " Node type
          lc_xml_node_rid_sheet_tp  type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
          lc_xml_node_rid_theme_tp  type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
          lc_xml_node_rid_styles_tp type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
          lc_xml_node_rid_shared_tp type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
          " Node target
          lc_xml_node_ridx_tg       type string value 'worksheets/sheet#.xml',
          lc_xml_node_rid_shared_tg type string value 'sharedStrings.xml',
          lc_xml_node_rid_styles_tg type string value 'styles.xml',
          lc_xml_node_rid_theme_tg  type string value 'theme/theme1.xml'.

    data: lo_document     type ref to if_ixml_document,
          lo_element_root type ref to if_ixml_element,
          lo_element      type ref to if_ixml_element.

    data: lv_xml_node_ridx_tg type string,
          lv_xml_node_ridx_id type string,
          lv_size             type i,
          lv_syindex          type string.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_relationships
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_rels_ns ).

**********************************************************************
* STEP 4: Create subnodes

    lv_size = excel->get_worksheets_size( ).


    " Relationship node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
    parent = lo_document ).
    lv_size = lv_size + 1.
    lv_syindex = lv_size.
    shift lv_syindex right deleting trailing space.
    shift lv_syindex left deleting leading space.
    lv_xml_node_ridx_id = lc_xml_node_ridx_id.
    replace all occurrences of '#' in lv_xml_node_ridx_id with lv_syindex.
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
    value = lv_xml_node_ridx_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
    value = lc_xml_node_rid_theme_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
    value = lc_xml_node_rid_theme_tg ).
    lo_element_root->append_child( new_child = lo_element ).


    " Relationship node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    lv_size = lv_size + 1.
    lv_syindex = lv_size.
    shift lv_syindex right deleting trailing space.
    shift lv_syindex left deleting leading space.
    lv_xml_node_ridx_id = lc_xml_node_ridx_id.
    replace all occurrences of '#' in lv_xml_node_ridx_id with lv_syindex.
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lv_xml_node_ridx_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid_styles_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid_styles_tg ).
    lo_element_root->append_child( new_child = lo_element ).



    lv_size = excel->get_worksheets_size( ).

    do lv_size times.
      " Relationship node
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
      parent = lo_document ).
      lv_xml_node_ridx_id = lc_xml_node_ridx_id.
      lv_xml_node_ridx_tg = lc_xml_node_ridx_tg.
      lv_syindex = sy-index.
      shift lv_syindex right deleting trailing space.
      shift lv_syindex left deleting leading space.
      replace all occurrences of '#' in lv_xml_node_ridx_id with lv_syindex.
      replace all occurrences of '#' in lv_xml_node_ridx_tg with lv_syindex.
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
      value = lv_xml_node_ridx_id ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
      value = lc_xml_node_rid_sheet_tp ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
      value = lv_xml_node_ridx_tg ).
      lo_element_root->append_child( new_child = lo_element ).
    enddo.

    " Relationship node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    lv_size += 3.
    lv_syindex = lv_size.
    shift lv_syindex right deleting trailing space.
    shift lv_syindex left deleting leading space.
    lv_xml_node_ridx_id = lc_xml_node_ridx_id.
    replace all occurrences of '#' in lv_xml_node_ridx_id with lv_syindex.
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lv_xml_node_ridx_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid_shared_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid_shared_tg ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  endmethod.


  method create_xl_sharedstrings.


** Constant node name
    data: lc_xml_node_sst         type string value 'sst',
          lc_xml_node_si          type string value 'si',
          lc_xml_node_t           type string value 't',
          lc_xml_node_r           type string value 'r',
          lc_xml_node_rpr         type string value 'rPr',
          " Node attributes
          lc_xml_attr_count       type string value 'count',
          lc_xml_attr_uniquecount type string value 'uniqueCount',
          " Node namespace
          lc_xml_node_ns          type string value 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'.

    data: lo_document     type ref to if_ixml_document,
          lo_element_root type ref to if_ixml_element,
          lo_element      type ref to if_ixml_element,
          lo_sub_element  type ref to if_ixml_element,
          lo_sub2_element type ref to if_ixml_element,
          lo_font_element type ref to if_ixml_element,
          lo_iterator     type ref to zcl_excel_collection_iterator,
          lo_worksheet    type ref to zcl_excel_worksheet.

    data: lt_cell_data       type zif_excel_data_decl=>zexcel_t_cell_data_unsorted,
          lt_cell_data_rtf   type zif_excel_data_decl=>zexcel_t_cell_data_unsorted,
          lv_value           type string,
          ls_shared_string   type zif_excel_data_decl=>zexcel_s_shared_string,
          lv_count_str       type string,
          lv_uniquecount_str type string,
          lv_sytabix         type i,
          lv_count           type i,
          lv_uniquecount     type i.

    field-symbols: <fs_sheet_content> type zif_excel_data_decl=>zexcel_s_cell_data,
                   <fs_rtf>           type zif_excel_data_decl=>zexcel_s_rtf,
                   <fs_sheet_string>  type zif_excel_data_decl=>zexcel_s_shared_string.

**********************************************************************
* STEP 1: Collect strings from each worksheet
    lo_iterator = excel->get_worksheets_iterator( ).

    while lo_iterator->has_next( ) eq abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).
      append lines of lo_worksheet->sheet_content to lt_cell_data.
    endwhile.

    delete lt_cell_data where cell_formula is not initial. " delete formula content

    lv_count = lines( lt_cell_data ).
    lv_count_str = lv_count.

    " separating plain and rich text format strings
    lt_cell_data_rtf = lt_cell_data.
    delete lt_cell_data where rtf_tab is not initial.
    delete lt_cell_data_rtf where rtf_tab is initial.

    shift lv_count_str right deleting trailing space.
    shift lv_count_str left deleting leading space.

    sort lt_cell_data by cell_value data_type.
    delete adjacent duplicates from lt_cell_data comparing cell_value data_type.

    " leave unique rich text format strings
    sort lt_cell_data_rtf by cell_value rtf_tab.
    delete adjacent duplicates from lt_cell_data_rtf comparing cell_value rtf_tab.
    " merge into single list
    append lines of lt_cell_data_rtf to lt_cell_data.
    sort lt_cell_data by cell_value rtf_tab.
    free lt_cell_data_rtf.

    lv_uniquecount = lines( lt_cell_data ).
    lv_uniquecount_str = lv_uniquecount.

    shift lv_uniquecount_str right deleting trailing space.
    shift lv_uniquecount_str left deleting leading space.

    clear lv_count.
    loop at lt_cell_data assigning <fs_sheet_content> where data_type = 's'.
      lv_sytabix = lv_count.
      ls_shared_string-string_no = lv_sytabix.
      ls_shared_string-string_value = <fs_sheet_content>-cell_value.
      ls_shared_string-string_type = <fs_sheet_content>-data_type.
      ls_shared_string-rtf_tab = <fs_sheet_content>-rtf_tab.
      insert ls_shared_string into table shared_strings.
      lv_count += 1.
    endloop.


**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_sst
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_ns ).
    lo_element_root->set_attribute_ns( name  = lc_xml_attr_count
                                       value = lv_count_str ).
    lo_element_root->set_attribute_ns( name  = lc_xml_attr_uniquecount
                                       value = lv_uniquecount_str ).

**********************************************************************
* STEP 4: Create subnode
    loop at shared_strings assigning <fs_sheet_string>.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_si
                                                       parent = lo_document ).
      if <fs_sheet_string>-rtf_tab is initial.
        lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_t
                                                             parent = lo_document ).
        if boolc( contains( val = <fs_sheet_string>-string_value start = ` ` ) ) = abap_true
              or boolc( contains( val = <fs_sheet_string>-string_value end = ` ` ) ) = abap_true.
          lo_sub_element->set_attribute( name = 'space' namespace = 'xml' value = 'preserve' ).
        endif.
        lv_value = escape_string_value( <fs_sheet_string>-string_value ).
        lo_sub_element->set_value( value = lv_value ).
      else.
        loop at <fs_sheet_string>-rtf_tab assigning <fs_rtf>.
          lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_r
                                                               parent = lo_element ).
          try.
              lv_value = substring( val = <fs_sheet_string>-string_value
                                    off = <fs_rtf>-offset
                                    len = <fs_rtf>-length ).
            catch cx_sy_range_out_of_bounds.
              exit.
          endtry.
          lv_value = escape_string_value( lv_value ).
          if <fs_rtf>-font is not initial.
            lo_font_element = lo_document->create_simple_element( name   = lc_xml_node_rpr
                                                                  parent = lo_sub_element ).
            create_xl_styles_font_node( io_document = lo_document
                                        io_parent   = lo_font_element
                                        is_font     = <fs_rtf>-font
                                        iv_use_rtf  = abap_true ).
          endif.
          lo_sub2_element = lo_document->create_simple_element( name   = lc_xml_node_t
                                                              parent = lo_sub_element ).
          if boolc( contains( val = lv_value start = ` ` ) ) = abap_true
                or boolc( contains( val = lv_value end = ` ` ) ) = abap_true.
            lo_sub2_element->set_attribute( name = 'space' namespace = 'xml' value = 'preserve' ).
          endif.
          lo_sub2_element->set_value( lv_value ).
        endloop.
      endif.
      lo_element->append_child( new_child = lo_sub_element ).
      lo_element_root->append_child( new_child = lo_element ).
    endloop.

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  endmethod.


  method create_xl_sheet.

** Constant node name
    data: lc_xml_node_worksheet type string value 'worksheet',
          " Node namespace
          lc_xml_node_ns        type string value 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          lc_xml_node_r_ns      type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          lc_xml_node_comp_ns   type string value 'http://schemas.openxmlformats.org/markup-compatibility/2006',
          lc_xml_node_comp_pref type string value 'x14ac',
          lc_xml_node_ig_ns     type string value 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'.

    data: lo_document        type ref to if_ixml_document,
          lo_element_root    type ref to if_ixml_element,
          lo_create_xl_sheet type ref to lcl_create_xl_sheet.



**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

***********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_worksheet
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:r'
                                       value = lc_xml_node_r_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:mc'
                                       value = lc_xml_node_comp_ns ).
    lo_element_root->set_attribute_ns( name  = 'mc:Ignorable'
                                       value = lc_xml_node_comp_pref ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:x14ac'
                                       value = lc_xml_node_ig_ns ).


**********************************************************************
* STEP 4: Create subnodes

    create object lo_create_xl_sheet.
    lo_create_xl_sheet->create( io_worksheet         = io_worksheet
                                io_document          = lo_document
                                iv_active            = iv_active
                                io_excel_writer_2007 = me ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  endmethod.


  method create_xl_sheet_column_formula.

    types: ls_column_formula_used     type mty_column_formula_used,
           lv_column_alpha            type zif_excel_data_decl=>zexcel_cell_column_alpha,
           lv_top_cell_coords         type zif_excel_data_decl=>zexcel_cell_coords,
           lv_bottom_cell_coords      type zif_excel_data_decl=>zexcel_cell_coords,
           lv_cell_coords             type zif_excel_data_decl=>zexcel_cell_coords,
           lv_ref_value               type string,
           lv_test_shared             type string,
           lv_si                      type i,
           lv_1st_line_shared_formula type abap_bool.
    data: lv_value                   type string,
          ls_column_formula_used     type mty_column_formula_used,
          lv_column_alpha            type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_top_cell_coords         type zif_excel_data_decl=>zexcel_cell_coords,
          lv_bottom_cell_coords      type zif_excel_data_decl=>zexcel_cell_coords,
          lv_cell_coords             type zif_excel_data_decl=>zexcel_cell_coords,
          lv_ref_value               type string,
          lv_1st_line_shared_formula type abap_bool.
    field-symbols: <ls_column_formula>      type zcl_excel_worksheet=>mty_s_column_formula,
                   <ls_column_formula_used> type mty_column_formula_used.


    read table it_column_formulas with table key id = is_sheet_content-column_formula_id assigning <ls_column_formula>.
    assert sy-subrc = 0.

    lv_value = <ls_column_formula>-formula.
    lv_1st_line_shared_formula = abap_false.
    eo_element = io_document->create_simple_element( name   = 'f'
                                                     parent = io_document ).
    read table ct_column_formulas_used with table key id = is_sheet_content-column_formula_id assigning <ls_column_formula_used>.
    if sy-subrc <> 0.
      clear ls_column_formula_used.
      ls_column_formula_used-id = is_sheet_content-column_formula_id.
      if is_formula_shareable( ip_formula = lv_value ) = abap_true.
        ls_column_formula_used-t = 'shared'.
        ls_column_formula_used-si = cv_si.
        condense ls_column_formula_used-si.
        cv_si = cv_si + 1.
        lv_1st_line_shared_formula = abap_true.
      endif.
      insert ls_column_formula_used into table ct_column_formulas_used assigning <ls_column_formula_used>.
    endif.

    if lv_1st_line_shared_formula = abap_true or <ls_column_formula_used>-t <> 'shared'.
      lv_column_alpha = zcl_excel_common=>convert_column2alpha( ip_column = is_sheet_content-cell_column ).
      lv_top_cell_coords = |{ lv_column_alpha }{ <ls_column_formula>-table_top_left_row + 1 }|.
      lv_bottom_cell_coords = |{ lv_column_alpha }{ <ls_column_formula>-table_bottom_right_row + 1 }|.
      lv_cell_coords = |{ lv_column_alpha }{ is_sheet_content-cell_row }|.
      if lv_top_cell_coords = lv_cell_coords.
        lv_ref_value = |{ lv_top_cell_coords }:{ lv_bottom_cell_coords }|.
      else.
        lv_ref_value = |{ lv_cell_coords }:{ lv_bottom_cell_coords }|.
        lv_value = zcl_excel_common=>shift_formula(
            iv_reference_formula = lv_value
            iv_shift_cols        = 0
            iv_shift_rows        = is_sheet_content-cell_row - <ls_column_formula>-table_top_left_row - 1 ).
      endif.
    endif.

    if <ls_column_formula_used>-t = 'shared'.
      eo_element->set_attribute( name  = 't'
                                 value = <ls_column_formula_used>-t ).
      eo_element->set_attribute( name  = 'si'
                                 value = <ls_column_formula_used>-si ).
      if lv_1st_line_shared_formula = abap_true.
        eo_element->set_attribute( name  = 'ref'
                                   value = lv_ref_value ).
        eo_element->set_value( value = lv_value ).
      endif.
    else.
      eo_element->set_value( value = lv_value ).
    endif.

  endmethod.


  method create_xl_sheet_ignored_errors.
    data: lo_element        type ref to if_ixml_element,
          lo_element2       type ref to if_ixml_element,
          lt_ignored_errors type zcl_excel_worksheet=>mty_th_ignored_errors.
    field-symbols: <ls_ignored_errors> type zcl_excel_worksheet=>mty_s_ignored_errors.

    lt_ignored_errors = io_worksheet->get_ignored_errors( ).

    if lt_ignored_errors is not initial.
      lo_element = io_document->create_simple_element( name   = 'ignoredErrors'
                                                       parent = io_document ).


      loop at lt_ignored_errors assigning <ls_ignored_errors>.

        lo_element2 = io_document->create_simple_element( name   = 'ignoredError'
                                                          parent = io_document ).

        lo_element2->set_attribute_ns( name  = 'sqref'
                                       value = <ls_ignored_errors>-cell_coords ).

        if <ls_ignored_errors>-eval_error = abap_true.
          lo_element2->set_attribute_ns( name  = 'evalError'
                                         value = '1' ).
        endif.
        if <ls_ignored_errors>-two_digit_text_year = abap_true.
          lo_element2->set_attribute_ns( name  = 'twoDigitTextYear'
                                         value = '1' ).
        endif.
        if <ls_ignored_errors>-number_stored_as_text = abap_true.
          lo_element2->set_attribute_ns( name  = 'numberStoredAsText'
                                         value = '1' ).
        endif.
        if <ls_ignored_errors>-formula = abap_true.
          lo_element2->set_attribute_ns( name  = 'formula'
                                         value = '1' ).
        endif.
        if <ls_ignored_errors>-formula_range = abap_true.
          lo_element2->set_attribute_ns( name  = 'formulaRange'
                                         value = '1' ).
        endif.
        if <ls_ignored_errors>-unlocked_formula = abap_true.
          lo_element2->set_attribute_ns( name  = 'unlockedFormula'
                                         value = '1' ).
        endif.
        if <ls_ignored_errors>-empty_cell_reference = abap_true.
          lo_element2->set_attribute_ns( name  = 'emptyCellReference'
                                         value = '1' ).
        endif.
        if <ls_ignored_errors>-list_data_validation = abap_true.
          lo_element2->set_attribute_ns( name  = 'listDataValidation'
                                         value = '1' ).
        endif.
        if <ls_ignored_errors>-calculated_column = abap_true.
          lo_element2->set_attribute_ns( name  = 'calculatedColumn'
                                         value = '1' ).
        endif.

        lo_element->append_child( lo_element2 ).

      endloop.

      io_element_root->append_child( lo_element ).

    endif.

  endmethod.


  method create_xl_sheet_pagebreaks.
    data: lo_pagebreaks     type ref to zcl_excel_worksheet_pagebreaks,
          lt_pagebreaks     type zcl_excel_worksheet_pagebreaks=>tt_pagebreak_at,
          lt_rows           type hashed table of int4 with unique key table_line,
          lt_columns        type hashed table of int4 with unique key table_line,

          lo_node_rowbreaks type ref to if_ixml_element,
          lo_node_colbreaks type ref to if_ixml_element,
          lo_node_break     type ref to if_ixml_element,

          lv_value          type string.


    field-symbols: <ls_pagebreak> like line of lt_pagebreaks.

    lo_pagebreaks = io_worksheet->get_pagebreaks( ).
    check lo_pagebreaks is bound.

    lt_pagebreaks = lo_pagebreaks->get_all_pagebreaks( ).
    check lt_pagebreaks is not initial.  " No need to proceed if don't have any pagebreaks.

    lo_node_rowbreaks = io_document->create_simple_element( name   = 'rowBreaks'
                                                            parent = io_document ).

    lo_node_colbreaks = io_document->create_simple_element( name   = 'colBreaks'
                                                            parent = io_document ).


    loop at lt_pagebreaks assigning <ls_pagebreak>.

* Count how many rows and columns need to be broken
      insert <ls_pagebreak>-cell_row    into table lt_rows.
      if sy-subrc = 0. " New
        lv_value = <ls_pagebreak>-cell_row.
        condense lv_value.

        lo_node_break = io_document->create_simple_element( name   = 'brk'
                                                            parent = io_document ).
        lo_node_break->set_attribute( name = 'id'  value = lv_value ).
        lo_node_break->set_attribute( name = 'man' value = '1' ).      " Manual break
        lo_node_break->set_attribute( name = 'max' value = '16383' ).  " Max columns

        lo_node_rowbreaks->append_child( new_child = lo_node_break ).
      endif.

      insert <ls_pagebreak>-cell_column into table lt_columns.
      if sy-subrc = 0. " New
        lv_value = <ls_pagebreak>-cell_column.
        condense lv_value.

        lo_node_break = io_document->create_simple_element( name   = 'brk'
                                                            parent = io_document ).
        lo_node_break->set_attribute( name = 'id'  value = lv_value ).
        lo_node_break->set_attribute( name = 'man' value = '1' ).        " Manual break
        lo_node_break->set_attribute( name = 'max' value = '1048575' ).  " Max rows

        lo_node_colbreaks->append_child( new_child = lo_node_break ).
      endif.


    endloop.

    lv_value = lines( lt_rows ).
    condense lv_value.
    lo_node_rowbreaks->set_attribute( name = 'count'             value = lv_value ).
    lo_node_rowbreaks->set_attribute( name = 'manualBreakCount'  value = lv_value ).

    lv_value = lines( lt_rows ).
    condense lv_value.
    lo_node_colbreaks->set_attribute( name = 'count'             value = lv_value ).
    lo_node_colbreaks->set_attribute( name = 'manualBreakCount'  value = lv_value ).

    io_parent->append_child( new_child = lo_node_rowbreaks ).
    io_parent->append_child( new_child = lo_node_colbreaks ).

  endmethod.


  method create_xl_sheet_rels.


** Constant node name
    data: lc_xml_node_relationships      type string value 'Relationships',
          lc_xml_node_relationship       type string value 'Relationship',
          " Node attributes
          lc_xml_attr_id                 type string value 'Id',
          lc_xml_attr_type               type string value 'Type',
          lc_xml_attr_target             type string value 'Target',
          lc_xml_attr_target_mode        type string value 'TargetMode',
          lc_xml_val_external            type string value 'External',
          " Node namespace
          lc_xml_node_rels_ns            type string value 'http://schemas.openxmlformats.org/package/2006/relationships',
          lc_xml_node_rid_table_tp       type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
          lc_xml_node_rid_printer_tp     type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings',
          lc_xml_node_rid_drawing_tp     type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
          lc_xml_node_rid_comment_tp     type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',        " (+) Issue #180
          lc_xml_node_rid_drawing_cmt_tp type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',      " (+) Issue #180
          lc_xml_node_rid_link_tp        type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'.

    data: lo_document     type ref to if_ixml_document,
          lo_element_root type ref to if_ixml_element,
          lo_element      type ref to if_ixml_element,
          lo_iterator     type ref to zcl_excel_collection_iterator,
          lo_table        type ref to zcl_excel_table,
          lo_link         type ref to zcl_excel_hyperlink.

    data: lv_value       type string,
          lv_relation_id type i,
          lv_index_str   type string.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_relationships
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_rels_ns ).

**********************************************************************
* STEP 4: Create subnodes

    " Add sheet Relationship nodes here
    lv_relation_id = 0.
    lo_iterator = io_worksheet->get_hyperlinks_iterator( ).
    while lo_iterator->has_next( ) eq abap_true.
      lo_link ?= lo_iterator->get_next( ).
      check lo_link->is_internal( ) = abap_false.  " issue #340 - don't put internal links here
      lv_relation_id += 1.

      lv_value = lv_relation_id.
      condense lv_value.
      concatenate 'rId' lv_value into lv_value.

      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_link_tp ).

      lv_value = lo_link->get_url( ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_target_mode
                                    value = lc_xml_val_external ).
      lo_element_root->append_child( new_child = lo_element ).
    endwhile.

* drawing
    if iv_drawing_index > 0.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      lv_relation_id += 1.

      lv_value = lv_relation_id.
      condense lv_value.
      concatenate 'rId' lv_value into lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_drawing_tp ).

      lv_index_str = iv_drawing_index.
      condense lv_index_str no-gaps.
      lv_value = me->c_xl_drawings.
      lv_value = replace( val = lv_value sub = 'xl' with = '..' ).
      lv_value = replace( val = lv_value sub = '#' with = lv_index_str ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    endif.

* Begin - Add - Issue #180
    if iv_cmnt_vmlindex > 0 and iv_comment_index > 0.
      " Drawing for comment
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      lv_relation_id += 1.

      lv_value = lv_relation_id.
      condense lv_value.
      concatenate 'rId' lv_value into lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_drawing_cmt_tp ).

      lv_index_str = iv_cmnt_vmlindex.
      condense lv_index_str no-gaps.
      lv_value = me->cl_xl_drawing_for_comments.
      lv_value = replace( val = lv_value sub = 'xl' with = '..' ).
      lv_value = replace( val = lv_value sub = '#' with = lv_index_str ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                    value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).

      " Comment
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      lv_relation_id += 1.

      lv_value = lv_relation_id.
      condense lv_value.
      concatenate 'rId' lv_value into lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_comment_tp ).

      lv_index_str = iv_comment_index.
      condense lv_index_str no-gaps.
      lv_value = me->c_xl_comments.
      lv_value = replace( val = lv_value sub = 'xl' with = '..' ).
      lv_value = replace( val = lv_value sub = '#' with = lv_index_str ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    endif.
* End   - Add - Issue #180

**********************************************************************
* header footer image
    if iv_hdft_vmlindex > 0. "Header or footer image exist
      " Drawing for comment/header/footer
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      lv_relation_id += 1.

      lv_value = lv_relation_id.
      condense lv_value.
      concatenate 'rId' lv_value into lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_drawing_cmt_tp ).

      lv_index_str = iv_hdft_vmlindex.
      condense lv_index_str no-gaps.
      lv_value = me->cl_xl_drawing_for_comments.
      lv_value = replace( val = lv_value sub = 'xl' with = '..' ).
      lv_value = replace( val = lv_value sub = '#' with = lv_index_str ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                    value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    endif.
*** End Header Footer
**********************************************************************


    lo_iterator = io_worksheet->get_tables_iterator( ).
    while lo_iterator->has_next( ) eq abap_true.
      lo_table ?= lo_iterator->get_next( ).
      lv_relation_id += 1.

      lv_value = lv_relation_id.
      condense lv_value.
      concatenate 'rId' lv_value into lv_value.

      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_table_tp ).

      lv_value = lo_table->get_name( ).
      concatenate '../tables/' lv_value '.xml' into lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    endwhile.

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  endmethod.


  method create_xl_sheet_sheet_data.

    types: begin of lty_table_area,
             left   type i,
             right  type i,
             top    type i,
             bottom type i,
           end of lty_table_area.

    constants: lc_dummy_cell_content       type zif_excel_data_decl=>zexcel_s_cell_data-cell_value value '})~~~ This is a dummy value for ABAP2XLSX and you should never find this in a real excelsheet Ihope'.

    constants: lc_xml_node_sheetdata type string value 'sheetData',   " SheetData tag
               lc_xml_node_row       type string value 'row',         " Row tag
               lc_xml_attr_r         type string value 'r',           " Cell:  row-attribute
               lc_xml_attr_spans     type string value 'spans',       " Cell: spans-attribute
               lc_xml_node_c         type string value 'c',           " Cell tag
               lc_xml_node_v         type string value 'v',           " Cell: value
               lc_xml_node_f         type string value 'f',           " Cell: formula
               lc_xml_attr_s         type string value 's',           " Cell: style
               lc_xml_attr_t         type string value 't'.           " Cell: type

    data: col_count              type int4,
          lo_autofilters         type ref to zcl_excel_autofilters,
          lo_autofilter          type ref to zcl_excel_autofilter,

          lo_iterator            type ref to zcl_excel_collection_iterator,
          lo_table               type ref to zcl_excel_table,
          lt_table_areas         type sorted table of lty_table_area with non-unique key left right top bottom,
          ls_table_area          like line of lt_table_areas,
          lo_column              type ref to zcl_excel_column,

          ls_sheet_content       like line of io_worksheet->sheet_content,
          ls_sheet_content_empty like line of io_worksheet->sheet_content,
          lv_current_row         type i,
          lv_next_row            type i,
          lv_last_row            type i,

*        lts_row_dimensions     type zif_excel_data_decl=>zexcel_t_worksheet_rowdimensio,
          lo_row_iterator        type ref to zcl_excel_collection_iterator,
          lo_row                 type ref to zcl_excel_row,
          lo_row_empty           type ref to zcl_excel_row,
          lts_row_outlines       type zcl_excel_worksheet=>mty_ts_outlines_row,

          ls_last_row            type zif_excel_data_decl=>zexcel_s_cell_data,
          ls_style_mapping       type zif_excel_data_decl=>zexcel_s_styles_mapping,

          lo_element_2           type ref to if_ixml_element,
          lo_element_3           type ref to if_ixml_element,
          lo_element_4           type ref to if_ixml_element,

          lv_value               type string,
          lv_style_guid          type zif_excel_data_decl=>zexcel_cell_style.
    data: lt_column_formulas_used type mty_column_formulas_used,
          lv_si                   type i.

    field-symbols: <ls_sheet_content> type zif_excel_data_decl=>zexcel_s_cell_data,
                   <ls_row_outline>   like line of lts_row_outlines.


    " sheetData node
    rv_ixml_sheet_data_root = io_document->create_simple_element( name   = lc_xml_node_sheetdata
                                                                  parent = io_document ).

    " Get column count
    col_count      = io_worksheet->get_highest_column( ).
    " Get autofilter
    lo_autofilters = excel->get_autofilters_reference( ).
    lo_autofilter  = lo_autofilters->get( io_worksheet = io_worksheet ) .
    if lo_autofilter is bound.
*     Area not used here, but makes the validation for lo_autofilter->is_row_hidden
      lo_autofilter->get_filter_area( ) .
    endif.
*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 1 - start
*--------------------------------------------------------------------*
*Build table to hold all table-areas attached to this sheet
    lo_iterator = io_worksheet->get_tables_iterator( ).
    while lo_iterator->has_next( ) eq abap_true.
      lo_table ?= lo_iterator->get_next( ).
      ls_table_area-left   = zcl_excel_common=>convert_column2int( lo_table->settings-top_left_column ).
      ls_table_area-right  = lo_table->get_right_column_integer( ).
      ls_table_area-top    = lo_table->settings-top_left_row.
      ls_table_area-bottom = lo_table->get_bottom_row_integer( ).
      insert ls_table_area into table lt_table_areas.
    endwhile.
*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 1 - end
*--------------------------------------------------------------------*
*We have problems when the first rows or trailing rows are not set but we have rowinformation
*to solve this we add dummycontent into first and last line that will not be set
*Set first line if necessary
    read table io_worksheet->sheet_content transporting no fields with key cell_row = 1.
    if sy-subrc <> 0.
      ls_sheet_content_empty-cell_row      = 1.
      ls_sheet_content_empty-cell_column   = 1.
      ls_sheet_content_empty-cell_value    = lc_dummy_cell_content.
      insert ls_sheet_content_empty into table io_worksheet->sheet_content.
    endif.
*Set last line if necessary
*Last row with cell content
    lv_last_row = io_worksheet->get_highest_row( ).
*Last line with row-information set directly ( like line height, hidden-status ... )

    lo_row_iterator = io_worksheet->get_rows_iterator( ).
    while lo_row_iterator->has_next( ) = abap_true.
      lo_row ?= lo_row_iterator->get_next( ).
      if lo_row->get_row_index( ) > lv_last_row.
        lv_last_row = lo_row->get_row_index( ).
      endif.
    endwhile.

*Last line with row-information set indirectly by row outline
    lts_row_outlines = io_worksheet->get_row_outlines( ).
    loop at lts_row_outlines assigning <ls_row_outline>.
      if <ls_row_outline>-collapsed = 'X'.
        lv_current_row = <ls_row_outline>-row_to + 1.  " collapsed-status may be set on following row
      else.
        lv_current_row = <ls_row_outline>-row_to.  " collapsed-status may be set on following row
      endif.
      if lv_current_row > lv_last_row.
        lv_last_row = lv_current_row.
      endif.
    endloop.
    read table io_worksheet->sheet_content transporting no fields with key cell_row = lv_last_row.
    if sy-subrc <> 0.
      ls_sheet_content_empty-cell_row      = lv_last_row.
      ls_sheet_content_empty-cell_column   = 1.
      ls_sheet_content_empty-cell_value    = lc_dummy_cell_content.
      insert ls_sheet_content_empty into table io_worksheet->sheet_content.
    endif.

    clear ls_sheet_content.
    loop at io_worksheet->sheet_content into ls_sheet_content.
      clear ls_style_mapping.
*Create row element
*issues #346,#154, #195  - problems when we have information in row_dimension but no cell content in that row
*Get next line that may have to be added.  If we have empty lines this is the next line after previous cell content
*Otherwise it is the line of the current cell content
      lv_current_row = ls_last_row-cell_row + 1.
      if lv_current_row > ls_sheet_content-cell_row.
        lv_current_row = ls_sheet_content-cell_row.
      endif.
*Fill in empty lines if necessary - assign an emtpy sheet content
      lv_next_row = lv_current_row.
      while lv_next_row <= ls_sheet_content-cell_row.
        lv_current_row = lv_next_row.
        lv_next_row = lv_current_row + 1.
        if lv_current_row = ls_sheet_content-cell_row. " cell value found in this row
          assign ls_sheet_content to <ls_sheet_content>.
        else.
*Check if empty row is really necessary - this is basically the case when we have information in row_dimension
          lo_row_empty = io_worksheet->get_row( lv_current_row ).
          check lo_row_empty->get_row_height( )                 >= 0          or
                lo_row_empty->get_collapsed( io_worksheet )      = abap_true  or
                lo_row_empty->get_outline_level( io_worksheet )  > 0          or
                lo_row_empty->get_xf_index( )                   <> 0.
          " Dummyentry A1
          ls_sheet_content_empty-cell_row      = lv_current_row.
          ls_sheet_content_empty-cell_column   = 1.
          assign ls_sheet_content_empty to <ls_sheet_content>.
        endif.

        if ls_last_row-cell_row ne <ls_sheet_content>-cell_row.
          if ls_last_row-cell_row is not initial.
            " Row visibility of previos row.
            if lo_row->get_visible( io_worksheet ) = abap_false or
               ( lo_autofilter is bound and
                 lo_autofilter->is_row_hidden( ls_last_row-cell_row ) = abap_true ).
              lo_element_2->set_attribute_ns( name  = 'hidden' value = 'true' ).
            endif.
            rv_ixml_sheet_data_root->append_child( new_child = lo_element_2 ). " row node
          endif.
          " Add new row
          lo_element_2 = io_document->create_simple_element( name   = lc_xml_node_row
                                                             parent = io_document ).
          " r
          lv_value = <ls_sheet_content>-cell_row.
          shift lv_value right deleting trailing space.
          shift lv_value left deleting leading space.

          lo_element_2->set_attribute_ns( name  = lc_xml_attr_r
                                          value = lv_value ).
          " Spans
          lv_value = col_count.
          concatenate '1:' lv_value into lv_value.
          shift lv_value right deleting trailing space.
          shift lv_value left deleting leading space.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_spans
                                          value = lv_value ).
          lo_row = io_worksheet->get_row( <ls_sheet_content>-cell_row ).
          " Row dimensions
          if lo_row->get_custom_height( ) = abap_true.
            lo_element_2->set_attribute_ns( name  = 'customHeight' value = '1' ).
          endif.
          if lo_row->get_row_height( ) > 0.
            lv_value = lo_row->get_row_height( ).
            lo_element_2->set_attribute_ns( name  = 'ht' value = lv_value ).
          endif.
          " Collapsed
          if lo_row->get_collapsed( io_worksheet ) = abap_true.
            lo_element_2->set_attribute_ns( name  = 'collapsed' value = 'true' ).
          endif.
          " Outline level
          if lo_row->get_outline_level( io_worksheet ) > 0.
            lv_value = lo_row->get_outline_level( io_worksheet ).
            shift lv_value right deleting trailing space.
            shift lv_value left deleting leading space.
            lo_element_2->set_attribute_ns( name  = 'outlineLevel' value = lv_value ).
          endif.
          " Style
          if lo_row->get_xf_index( ) <> 0.
            lv_value = lo_row->get_xf_index( ).
            lo_element_2->set_attribute_ns( name  = 's' value = lv_value ).
            lo_element_2->set_attribute_ns( name  = 'customFormat'  value = '1' ).
          endif.
        else.

        endif.
      endwhile.

      lo_element_3 = io_document->create_simple_element( name   = lc_xml_node_c
                                                         parent = io_document ).

      lo_element_3->set_attribute_ns( name  = lc_xml_attr_r
                                      value = <ls_sheet_content>-cell_coords ).

*begin of change issue #157 - allow column cellstyle
*if no cellstyle is set, look into column, then into sheet
      if <ls_sheet_content>-cell_style is not initial.
        lv_style_guid = <ls_sheet_content>-cell_style.
      else.
*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 2 - start
*--------------------------------------------------------------------*
*Check if cell in any of the table areas
        loop at lt_table_areas transporting no fields where top    <= <ls_sheet_content>-cell_row
                                                        and bottom >= <ls_sheet_content>-cell_row
                                                        and left   <= <ls_sheet_content>-cell_column
                                                        and right  >= <ls_sheet_content>-cell_column. "#EC CI_SORTSEQ
          exit.
        endloop.
        if sy-subrc = 0.
          clear lv_style_guid.     " No style --> EXCEL will use built-in-styles as declared in the tables-section
        else.
*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 2 - end
*--------------------------------------------------------------------*
          lv_style_guid = io_worksheet->zif_excel_sheet_properties~get_style( ).
          lo_column ?= io_worksheet->get_column( <ls_sheet_content>-cell_column ).
          if lo_column->get_column_index( ) = <ls_sheet_content>-cell_column.
            lv_style_guid = lo_column->get_column_style_guid( ).
            if lv_style_guid is initial.
              lv_style_guid = io_worksheet->zif_excel_sheet_properties~get_style( ).
            endif.
          endif.

*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 3 - start
*--------------------------------------------------------------------*
        endif.
*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 3 - end
*--------------------------------------------------------------------*
      endif.
      if lv_style_guid is not initial.
        read table styles_mapping into ls_style_mapping with key guid = lv_style_guid.
*end of change issue #157 - allow column cellstyles
        lv_value = ls_style_mapping-style.
        shift lv_value right deleting trailing space.
        shift lv_value left deleting leading space.
        lo_element_3->set_attribute_ns( name  = lc_xml_attr_s
                                        value = lv_value ).
      endif.

      " For cells with formula ignore the value - Excel will calculate it
      if <ls_sheet_content>-cell_formula is not initial.
        " fomula node
        lo_element_4 = io_document->create_simple_element( name   = lc_xml_node_f
                                                           parent = io_document ).
        lo_element_4->set_value( value = <ls_sheet_content>-cell_formula ).
        lo_element_3->append_child( new_child = lo_element_4 ). " formula node
      elseif <ls_sheet_content>-column_formula_id <> 0.
        create_xl_sheet_column_formula(
          exporting
            io_document             = io_document
            it_column_formulas      = io_worksheet->column_formulas
            is_sheet_content        = <ls_sheet_content>
          importing
            eo_element              = lo_element_4
          changing
            ct_column_formulas_used = lt_column_formulas_used
            cv_si                   = lv_si ).
        lo_element_3->append_child( new_child = lo_element_4 ).
      elseif <ls_sheet_content>-cell_value is not initial           "cell can have just style or formula
         and <ls_sheet_content>-cell_value <> lc_dummy_cell_content.
        if <ls_sheet_content>-data_type is not initial.
          if <ls_sheet_content>-data_type eq 's_leading_blanks'.
            lo_element_3->set_attribute_ns( name  = lc_xml_attr_t
                                            value = 's' ).
          else.
            lo_element_3->set_attribute_ns( name  = lc_xml_attr_t
                                            value = <ls_sheet_content>-data_type ).
          endif.
        endif.

        " value node
        lo_element_4 = io_document->create_simple_element( name   = lc_xml_node_v
                                                           parent = io_document ).

        if <ls_sheet_content>-data_type eq 's' or <ls_sheet_content>-data_type eq 's_leading_blanks'.
          lv_value = me->get_shared_string_index( ip_cell_value = <ls_sheet_content>-cell_value
                                                  it_rtf        = <ls_sheet_content>-rtf_tab ).
          condense lv_value.
          lo_element_4->set_value( value = lv_value ).
        else.
          lv_value = <ls_sheet_content>-cell_value.
          condense lv_value.
          lo_element_4->set_value( value = lv_value ).
        endif.

        lo_element_3->append_child( new_child = lo_element_4 ). " value node
      endif.

      lo_element_2->append_child( new_child = lo_element_3 ). " column node
      ls_last_row = <ls_sheet_content>.
    endloop.
    if sy-subrc = 0.
      " Row visibility of previos row.
      if lo_row->get_visible( ) = abap_false or
         ( lo_autofilter is bound and
           lo_autofilter->is_row_hidden( ls_last_row-cell_row ) = abap_true ).
        lo_element_2->set_attribute_ns( name  = 'hidden' value = 'true' ).
      endif.
      rv_ixml_sheet_data_root->append_child( new_child = lo_element_2 ). " row node
    endif.
    delete io_worksheet->sheet_content where cell_value = lc_dummy_cell_content. "#EC CI_SORTSEQ " Get rid of dummyentries
  endmethod.


  method create_xl_styles.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   dxfs-cellstyles are used in conditional formats:
*                CellIs, Expression, top10 ( forthcoming above average as well )
*                create own method to write dsfx-cellstyle to be reuseable by all these
*--------------------------------------------------------------------*


** Constant node name
    constants: lc_xml_node_stylesheet        type string value 'styleSheet',
               " font
               lc_xml_node_fonts             type string value 'fonts',
               lc_xml_node_font              type string value 'font',
               lc_xml_node_color             type string value 'color',
               " fill
               lc_xml_node_fills             type string value 'fills',
               lc_xml_node_fill              type string value 'fill',
               lc_xml_node_patternfill       type string value 'patternFill',
               lc_xml_node_fgcolor           type string value 'fgColor',
               lc_xml_node_bgcolor           type string value 'bgColor',
               lc_xml_node_gradientfill      type string value 'gradientFill',
               lc_xml_node_stop              type string value 'stop',
               " borders
               lc_xml_node_borders           type string value 'borders',
               lc_xml_node_border            type string value 'border',
               lc_xml_node_left              type string value 'left',
               lc_xml_node_right             type string value 'right',
               lc_xml_node_top               type string value 'top',
               lc_xml_node_bottom            type string value 'bottom',
               lc_xml_node_diagonal          type string value 'diagonal',
               " numfmt
               lc_xml_node_numfmts           type string value 'numFmts',
               lc_xml_node_numfmt            type string value 'numFmt',
               " Styles
               lc_xml_node_cellstylexfs      type string value 'cellStyleXfs',
               lc_xml_node_xf                type string value 'xf',
               lc_xml_node_cellxfs           type string value 'cellXfs',
               lc_xml_node_cellstyles        type string value 'cellStyles',
               lc_xml_node_cellstyle         type string value 'cellStyle',
               lc_xml_node_dxfs              type string value 'dxfs',
               lc_xml_node_tablestyles       type string value 'tableStyles',
               " Colors
               lc_xml_node_colors            type string value 'colors',
               lc_xml_node_indexedcolors     type string value 'indexedColors',
               lc_xml_node_rgbcolor          type string value 'rgbColor',
               lc_xml_node_mrucolors         type string value 'mruColors',
               " Alignment
               lc_xml_node_alignment         type string value 'alignment',
               " Protection
               lc_xml_node_protection        type string value 'protection',
               " Node attributes
               lc_xml_attr_count             type string value 'count',
               lc_xml_attr_val               type string value 'val',
               lc_xml_attr_theme             type string value 'theme',
               lc_xml_attr_rgb               type string value 'rgb',
               lc_xml_attr_indexed           type string value 'indexed',
               lc_xml_attr_tint              type string value 'tint',
               lc_xml_attr_style             type string value 'style',
               lc_xml_attr_position          type string value 'position',
               lc_xml_attr_degree            type string value 'degree',
               lc_xml_attr_patterntype       type string value 'patternType',
               lc_xml_attr_numfmtid          type string value 'numFmtId',
               lc_xml_attr_fontid            type string value 'fontId',
               lc_xml_attr_fillid            type string value 'fillId',
               lc_xml_attr_borderid          type string value 'borderId',
               lc_xml_attr_xfid              type string value 'xfId',
               lc_xml_attr_applynumberformat type string value 'applyNumberFormat',
               lc_xml_attr_applyprotection   type string value 'applyProtection',
               lc_xml_attr_applyfont         type string value 'applyFont',
               lc_xml_attr_applyfill         type string value 'applyFill',
               lc_xml_attr_applyborder       type string value 'applyBorder',
               lc_xml_attr_name              type string value 'name',
               lc_xml_attr_builtinid         type string value 'builtinId',
               lc_xml_attr_defaulttablestyle type string value 'defaultTableStyle',
               lc_xml_attr_defaultpivotstyle type string value 'defaultPivotStyle',
               lc_xml_attr_applyalignment    type string value 'applyAlignment',
               lc_xml_attr_horizontal        type string value 'horizontal',
               lc_xml_attr_formatcode        type string value 'formatCode',
               lc_xml_attr_vertical          type string value 'vertical',
               lc_xml_attr_wraptext          type string value 'wrapText',
               lc_xml_attr_textrotation      type string value 'textRotation',
               lc_xml_attr_shrinktofit       type string value 'shrinkToFit',
               lc_xml_attr_indent            type string value 'indent',
               lc_xml_attr_locked            type string value 'locked',
               lc_xml_attr_hidden            type string value 'hidden',
               lc_xml_attr_diagonalup        type string value 'diagonalUp',
               lc_xml_attr_diagonaldown      type string value 'diagonalDown',
               " Node namespace
               lc_xml_node_ns                type string value 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
               lc_xml_attr_type              type string value 'type',
               lc_xml_attr_bottom            type string value 'bottom',
               lc_xml_attr_top               type string value 'top',
               lc_xml_attr_right             type string value 'right',
               lc_xml_attr_left              type string value 'left'.

    data: lo_document        type ref to if_ixml_document,
          lo_element_root    type ref to if_ixml_element,
          lo_element_fonts   type ref to if_ixml_element,
          lo_element_font    type ref to if_ixml_element,
          lo_element_fills   type ref to if_ixml_element,
          lo_element_fill    type ref to if_ixml_element,
          lo_element_borders type ref to if_ixml_element,
          lo_element_border  type ref to if_ixml_element,
          lo_element_numfmts type ref to if_ixml_element,
          lo_element_numfmt  type ref to if_ixml_element,
          lo_element_cellxfs type ref to if_ixml_element,
          lo_element         type ref to if_ixml_element,
          lo_sub_element     type ref to if_ixml_element,
          lo_sub_element_2   type ref to if_ixml_element,
          lo_iterator        type ref to zcl_excel_collection_iterator,
          lo_iterator2       type ref to zcl_excel_collection_iterator,
          lo_worksheet       type ref to zcl_excel_worksheet,
          lo_style_cond      type ref to zcl_excel_style_cond,
          lo_style           type ref to zcl_excel_style.


    data: lt_fonts          type zif_excel_data_decl=>zexcel_t_style_font,
          ls_font           type zif_excel_data_decl=>zexcel_s_style_font,
          lt_fills          type zif_excel_data_decl=>zexcel_t_style_fill,
          ls_fill           type zif_excel_data_decl=>zexcel_s_style_fill,
          lt_borders        type zif_excel_data_decl=>zexcel_t_style_border,
          ls_border         type zif_excel_data_decl=>zexcel_s_style_border,
          lt_numfmts        type zif_excel_data_decl=>zexcel_t_style_numfmt,
          ls_numfmt         type zif_excel_data_decl=>zexcel_s_style_numfmt,
          lt_protections    type zif_excel_data_decl=>zexcel_t_style_protection,
          ls_protection     type zif_excel_data_decl=>zexcel_s_style_protection,
          lt_alignments     type zif_excel_data_decl=>zexcel_t_style_alignment,
          ls_alignment      type zif_excel_data_decl=>zexcel_s_style_alignment,
          lt_cellxfs        type zif_excel_data_decl=>zexcel_t_cellxfs,
          ls_cellxfs        type zif_excel_data_decl=>zexcel_s_cellxfs,
          ls_styles_mapping type zif_excel_data_decl=>zexcel_s_styles_mapping,
          lt_colors         type zif_excel_data_decl=>zexcel_t_style_color_argb,
          ls_color          like line of lt_colors.

    data: lv_value         type string,
          lv_dfx_count     type i,
          lv_fonts_count   type i,
          lv_fills_count   type i,
          lv_borders_count type i,
          lv_cellxfs_count type i.

    types: begin of ts_built_in_format,
             num_format type zif_excel_data_decl=>zexcel_number_format,
             id         type i,
           end of ts_built_in_format.

    data: lt_built_in_num_formats type hashed table of ts_built_in_format with unique key num_format,
          ls_built_in_num_format  like line of lt_built_in_num_formats.
    field-symbols: <ls_built_in_format> like line of lt_built_in_num_formats,
                   <ls_reader_built_in> like line of zcl_excel_style_number_format=>mt_built_in_num_formats.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

***********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_stylesheet
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_ns ).

**********************************************************************
* STEP 4: Create subnodes

    lo_element_fonts = lo_document->create_simple_element( name   = lc_xml_node_fonts
                                                           parent = lo_document ).

    lo_element_fills = lo_document->create_simple_element( name   = lc_xml_node_fills
                                                           parent = lo_document ).

    lo_element_borders = lo_document->create_simple_element( name   = lc_xml_node_borders
                                                             parent = lo_document ).

    lo_element_cellxfs = lo_document->create_simple_element( name   = lc_xml_node_cellxfs
                                                             parent = lo_document ).

    lo_element_numfmts = lo_document->create_simple_element( name   = lc_xml_node_numfmts
                                                             parent = lo_document ).

* Prepare built-in number formats.
    loop at zcl_excel_style_number_format=>mt_built_in_num_formats assigning <ls_reader_built_in>.
      ls_built_in_num_format-id         = <ls_reader_built_in>-id.
      ls_built_in_num_format-num_format = <ls_reader_built_in>-format->format_code.
      insert ls_built_in_num_format into table lt_built_in_num_formats.
    endloop.
* Compress styles
    lo_iterator = excel->get_styles_iterator( ).
    while lo_iterator->has_next( ) eq abap_true.
      lo_style ?= lo_iterator->get_next( ).
      ls_font       = lo_style->font->get_structure( ).
      ls_fill       = lo_style->fill->get_structure( ).
      ls_border     = lo_style->borders->get_structure( ).
      ls_alignment  = lo_style->alignment->get_structure( ).
      ls_protection = lo_style->protection->get_structure( ).
      ls_numfmt     = lo_style->number_format->get_structure( ).

      clear ls_cellxfs.


* Compress fonts
      read table lt_fonts from ls_font transporting no fields.
      if sy-subrc eq 0.
        ls_cellxfs-fontid = sy-tabix.
      else.
        append ls_font to lt_fonts.
        ls_cellxfs-fontid = lines( lt_fonts ).
      endif.
      ls_cellxfs-fontid -= 1.

* Compress alignment
      read table lt_alignments from ls_alignment transporting no fields.
      if sy-subrc eq 0.
        ls_cellxfs-alignmentid = sy-tabix.
      else.
        append ls_alignment to lt_alignments.
        ls_cellxfs-alignmentid = lines( lt_alignments ).
      endif.
      ls_cellxfs-alignmentid -= 1.

* Compress fills
      read table lt_fills from ls_fill transporting no fields.
      if sy-subrc eq 0.
        ls_cellxfs-fillid = sy-tabix.
      else.
        append ls_fill to lt_fills.
        ls_cellxfs-fillid = lines( lt_fills ).
      endif.
      ls_cellxfs-fillid -= 1.

* Compress borders
      read table lt_borders from ls_border transporting no fields.
      if sy-subrc eq 0.
        ls_cellxfs-borderid = sy-tabix.
      else.
        append ls_border to lt_borders.
        ls_cellxfs-borderid = lines( lt_borders ).
      endif.
      ls_cellxfs-borderid -= 1.

* Compress protection
      if ls_protection-locked eq c_on and ls_protection-hidden eq c_off.
        ls_cellxfs-applyprotection    = 0.
      else.
        read table lt_protections from ls_protection transporting no fields.
        if sy-subrc eq 0.
          ls_cellxfs-protectionid = sy-tabix.
        else.
          append ls_protection to lt_protections.
          ls_cellxfs-protectionid = lines( lt_protections ).
        endif.
        ls_cellxfs-applyprotection    = 1.
      endif.
      ls_cellxfs-protectionid -= 1.

* Compress number formats

      "-----------
      if ls_numfmt-numfmt ne zcl_excel_style_number_format=>c_format_date_std." and ls_numfmt-NUMFMT ne 'STD_NDEC'. " ALE Changes on going
        "---
        if ls_numfmt is not initial.
* issue  #389 - Problem with built-in format ( those are not being taken account of )
* There are some internal number formats built-in into EXCEL
* Use these instead of duplicating the entries here, since they seem to be language-dependant and adjust to user settings in excel
          read table lt_built_in_num_formats assigning <ls_built_in_format> with table key num_format = ls_numfmt-numfmt.
          if sy-subrc = 0.
            ls_cellxfs-numfmtid = <ls_built_in_format>-id.
          else.
            read table lt_numfmts from ls_numfmt transporting no fields.
            if sy-subrc eq 0.
              ls_cellxfs-numfmtid = sy-tabix.
            else.
              append ls_numfmt to lt_numfmts.
              ls_cellxfs-numfmtid = lines( lt_numfmts ).
            endif.
            ls_cellxfs-numfmtid += zcl_excel_common=>c_excel_numfmt_offset. " Add OXML offset for custom styles
          endif.
          ls_cellxfs-applynumberformat    = 1.
        else.
          ls_cellxfs-applynumberformat    = 0.
        endif.
        "----------- " ALE changes on going
      else.
        ls_cellxfs-applynumberformat    = 1.
        if ls_numfmt-numfmt eq zcl_excel_style_number_format=>c_format_date_std.
          ls_cellxfs-numfmtid = 14.
        endif.
      endif.
      "---

      if ls_cellxfs-fontid ne 0.
        ls_cellxfs-applyfont    = 1.
      else.
        ls_cellxfs-applyfont    = 0.
      endif.
      if ls_cellxfs-alignmentid ne 0.
        ls_cellxfs-applyalignment = 1.
      else.
        ls_cellxfs-applyalignment = 0.
      endif.
      if ls_cellxfs-fillid ne 0.
        ls_cellxfs-applyfill    = 1.
      else.
        ls_cellxfs-applyfill    = 0.
      endif.
      if ls_cellxfs-borderid ne 0.
        ls_cellxfs-applyborder    = 1.
      else.
        ls_cellxfs-applyborder    = 0.
      endif.

* Remap styles
      read table lt_cellxfs from ls_cellxfs transporting no fields.
      if sy-subrc eq 0.
        ls_styles_mapping-style = sy-tabix.
      else.
        append ls_cellxfs to lt_cellxfs.
        ls_styles_mapping-style = lines( lt_cellxfs ).
      endif.
      ls_styles_mapping-style -= 1.
      ls_styles_mapping-guid = lo_style->get_guid( ).
      append ls_styles_mapping to me->styles_mapping.
    endwhile.

    " create numfmt elements
    loop at lt_numfmts into ls_numfmt.
      lo_element_numfmt = lo_document->create_simple_element( name   = lc_xml_node_numfmt
                                                              parent = lo_document ).
      lv_value = sy-tabix + zcl_excel_common=>c_excel_numfmt_offset.
      condense lv_value.
      lo_element_numfmt->set_attribute_ns( name  = lc_xml_attr_numfmtid
                                        value = lv_value ).
      lv_value = ls_numfmt-numfmt.
      lo_element_numfmt->set_attribute_ns( name  = lc_xml_attr_formatcode
                                           value = lv_value ).
      lo_element_numfmts->append_child( new_child = lo_element_numfmt ).
    endloop.

    " create font elements
    loop at lt_fonts into ls_font.
      lo_element_font = lo_document->create_simple_element( name   = lc_xml_node_font
                                                            parent = lo_document ).
      create_xl_styles_font_node( io_document = lo_document
                                  io_parent   = lo_element_font
                                  is_font     = ls_font ).
      lo_element_fonts->append_child( new_child = lo_element_font ).
    endloop.

    " create fill elements
    loop at lt_fills into ls_fill.
      lo_element_fill = lo_document->create_simple_element( name   = lc_xml_node_fill
                                                            parent = lo_document ).

      if ls_fill-gradtype is not initial.
        "gradient

        lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_gradientfill
                                                            parent = lo_document ).
        if ls_fill-gradtype-degree is not initial.
          lv_value = ls_fill-gradtype-degree.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_degree  value = lv_value ).
        endif.
        if ls_fill-gradtype-type is not initial.
          lv_value = ls_fill-gradtype-type.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_type  value = lv_value ).
        endif.
        if ls_fill-gradtype-bottom is not initial.
          lv_value = ls_fill-gradtype-bottom.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_bottom  value = lv_value ).
        endif.
        if ls_fill-gradtype-top is not initial.
          lv_value = ls_fill-gradtype-top.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_top  value = lv_value ).
        endif.
        if ls_fill-gradtype-right is not initial.
          lv_value = ls_fill-gradtype-right.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_right  value = lv_value ).
        endif.
        if ls_fill-gradtype-left is not initial.
          lv_value = ls_fill-gradtype-left.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_left  value = lv_value ).
        endif.

        if ls_fill-gradtype-position3 is not initial.
          "create <stop> elements for gradients, we can have 2 or 3 stops in each gradient
          lo_sub_element_2 =  lo_document->create_simple_element( name   = lc_xml_node_stop
                                                                  parent = lo_sub_element ).
          lv_value = ls_fill-gradtype-position1.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_position value = lv_value ).

          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-bgcolor
              iv_color_elem_name = lc_xml_node_color ).
          lo_sub_element->append_child( new_child = lo_sub_element_2 ).

          lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_stop
                                                                 parent = lo_sub_element ).

          lv_value = ls_fill-gradtype-position2.

          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_position
                                              value = lv_value ).

          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-fgcolor
              iv_color_elem_name = lc_xml_node_color ).
          lo_sub_element->append_child( new_child = lo_sub_element_2 ).

          lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_stop
                                                                 parent = lo_sub_element ).

          lv_value = ls_fill-gradtype-position3.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_position
                                              value = lv_value ).

          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-bgcolor
              iv_color_elem_name = lc_xml_node_color ).
          lo_sub_element->append_child( new_child = lo_sub_element_2 ).

        else.
          "create <stop> elements for gradients, we can have 2 or 3 stops in each gradient
          lo_sub_element_2 =  lo_document->create_simple_element( name   = lc_xml_node_stop
                                                                  parent = lo_sub_element ).
          lv_value = ls_fill-gradtype-position1.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_position value = lv_value ).

          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-bgcolor
              iv_color_elem_name = lc_xml_node_color ).
          lo_sub_element->append_child( new_child = lo_sub_element_2 ).

          lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_stop
                                                                 parent = lo_sub_element ).

          lv_value = ls_fill-gradtype-position2.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_position
                                              value = lv_value ).

          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-fgcolor
              iv_color_elem_name = lc_xml_node_color ).
          lo_sub_element->append_child( new_child = lo_sub_element_2 ).
        endif.

      else.
        "pattern
        lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_patternfill
                                                             parent = lo_document ).
        lv_value = ls_fill-filltype.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_patterntype
                                          value = lv_value ).
        " fgcolor
        create_xl_styles_color_node(
            io_document        = lo_document
            io_parent          = lo_sub_element
            is_color           = ls_fill-fgcolor
            iv_color_elem_name = lc_xml_node_fgcolor ).

        if  ls_fill-fgcolor-rgb is initial and
            ls_fill-fgcolor-indexed eq zcl_excel_style_color=>c_indexed_not_set and
            ls_fill-fgcolor-theme eq zcl_excel_style_color=>c_theme_not_set and
            ls_fill-fgcolor-tint is initial and ls_fill-bgcolor-indexed eq zcl_excel_style_color=>c_indexed_sys_foreground.

          " bgcolor
          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element
              is_color           = ls_fill-bgcolor
              iv_color_elem_name = lc_xml_node_bgcolor ).

        endif.
      endif.

      lo_element_fill->append_child( new_child = lo_sub_element )."pattern
      lo_element_fills->append_child( new_child = lo_element_fill ).
    endloop.

    " create border elements
    loop at lt_borders into ls_border.
      lo_element_border = lo_document->create_simple_element( name   = lc_xml_node_border
                                                              parent = lo_document ).

      if ls_border-diagonalup is not initial.
        lv_value = ls_border-diagonalup.
        condense lv_value.
        lo_element_border->set_attribute_ns( name  = lc_xml_attr_diagonalup
                                          value = lv_value ).
      endif.

      if ls_border-diagonaldown is not initial.
        lv_value = ls_border-diagonaldown.
        condense lv_value.
        lo_element_border->set_attribute_ns( name  = lc_xml_attr_diagonaldown
                                          value = lv_value ).
      endif.

      "left
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_left
                                                           parent = lo_document ).
      if ls_border-left_style is not initial.
        lv_value = ls_border-left_style.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
      endif.

      create_xl_styles_color_node(
          io_document        = lo_document
          io_parent          = lo_sub_element
          is_color           = ls_border-left_color ).

      lo_element_border->append_child( new_child = lo_sub_element ).

      "right
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_right
                                                           parent = lo_document ).
      if ls_border-right_style is not initial.
        lv_value = ls_border-right_style.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
      endif.

      create_xl_styles_color_node(
          io_document        = lo_document
          io_parent          = lo_sub_element
          is_color           = ls_border-right_color ).

      lo_element_border->append_child( new_child = lo_sub_element ).

      "top
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_top
                                                           parent = lo_document ).
      if ls_border-top_style is not initial.
        lv_value = ls_border-top_style.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
      endif.

      create_xl_styles_color_node(
          io_document        = lo_document
          io_parent          = lo_sub_element
          is_color           = ls_border-top_color ).

      lo_element_border->append_child( new_child = lo_sub_element ).

      "bottom
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_bottom
                                                           parent = lo_document ).
      if ls_border-bottom_style is not initial.
        lv_value = ls_border-bottom_style.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
      endif.

      create_xl_styles_color_node(
          io_document        = lo_document
          io_parent          = lo_sub_element
          is_color           = ls_border-bottom_color ).

      lo_element_border->append_child( new_child = lo_sub_element ).

      "diagonal
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_diagonal
                                                           parent = lo_document ).
      if ls_border-diagonal_style is not initial.
        lv_value = ls_border-diagonal_style.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
      endif.

      create_xl_styles_color_node(
          io_document        = lo_document
          io_parent          = lo_sub_element
          is_color           = ls_border-diagonal_color ).

      lo_element_border->append_child( new_child = lo_sub_element ).
      lo_element_borders->append_child( new_child = lo_element_border ).
    endloop.

    " update attribute "count"
    lv_fonts_count = lines( lt_fonts ).
    lv_value = lv_fonts_count.
    shift lv_value right deleting trailing space.
    shift lv_value left deleting leading space.
    lo_element_fonts->set_attribute_ns( name  = lc_xml_attr_count
                                        value = lv_value ).
    lv_fills_count = lines( lt_fills ).
    lv_value = lv_fills_count.
    shift lv_value right deleting trailing space.
    shift lv_value left deleting leading space.
    lo_element_fills->set_attribute_ns( name  = lc_xml_attr_count
                                        value = lv_value ).
    lv_borders_count = lines( lt_borders ).
    lv_value = lv_borders_count.
    shift lv_value right deleting trailing space.
    shift lv_value left deleting leading space.
    lo_element_borders->set_attribute_ns( name  = lc_xml_attr_count
                                          value = lv_value ).
    lv_cellxfs_count = lines( lt_cellxfs ).
    lv_value = lv_cellxfs_count.
    shift lv_value right deleting trailing space.
    shift lv_value left deleting leading space.
    lo_element_cellxfs->set_attribute_ns( name  = lc_xml_attr_count
                                          value = lv_value ).

    " Append to root node
    lo_element_root->append_child( new_child = lo_element_numfmts ).
    lo_element_root->append_child( new_child = lo_element_fonts ).
    lo_element_root->append_child( new_child = lo_element_fills ).
    lo_element_root->append_child( new_child = lo_element_borders ).

    " cellstylexfs node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_cellstylexfs
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_count
                                  value = '1' ).
    lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_xf
                                                         parent = lo_document ).

    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_numfmtid
                                      value = c_off ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_fontid
                                      value = c_off ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_fillid
                                      value = c_off ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_borderid
                                      value = c_off ).

    lo_element->append_child( new_child = lo_sub_element ).
    lo_element_root->append_child( new_child = lo_element ).

    loop at lt_cellxfs into ls_cellxfs.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_xf
                                                          parent = lo_document ).
      lv_value = ls_cellxfs-numfmtid.
      shift lv_value right deleting trailing space.
      shift lv_value left deleting leading space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_numfmtid
                                    value = lv_value ).
      lv_value = ls_cellxfs-fontid.
      shift lv_value right deleting trailing space.
      shift lv_value left deleting leading space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_fontid
                                    value = lv_value ).
      lv_value = ls_cellxfs-fillid.
      shift lv_value right deleting trailing space.
      shift lv_value left deleting leading space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_fillid
                                    value = lv_value ).
      lv_value = ls_cellxfs-borderid.
      shift lv_value right deleting trailing space.
      shift lv_value left deleting leading space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_borderid
                                    value = lv_value ).
      lv_value = ls_cellxfs-xfid.
      shift lv_value right deleting trailing space.
      shift lv_value left deleting leading space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_xfid
                                    value = lv_value ).
      if ls_cellxfs-applynumberformat eq 1.
        lv_value = ls_cellxfs-applynumberformat.
        shift lv_value right deleting trailing space.
        shift lv_value left deleting leading space.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applynumberformat
                                      value = lv_value ).
      endif.
      if ls_cellxfs-applyfont eq 1.
        lv_value = ls_cellxfs-applyfont.
        shift lv_value right deleting trailing space.
        shift lv_value left deleting leading space.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applyfont
                                      value = lv_value ).
      endif.
      if ls_cellxfs-applyfill eq 1.
        lv_value = ls_cellxfs-applyfill.
        shift lv_value right deleting trailing space.
        shift lv_value left deleting leading space.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applyfill
                                      value = lv_value ).
      endif.
      if ls_cellxfs-applyborder eq 1.
        lv_value = ls_cellxfs-applyborder.
        shift lv_value right deleting trailing space.
        shift lv_value left deleting leading space.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applyborder
                                      value = lv_value ).
      endif.
      if ls_cellxfs-applyalignment eq 1. " depends on each style not for all the sheet
        lv_value = ls_cellxfs-applyalignment.
        shift lv_value right deleting trailing space.
        shift lv_value left deleting leading space.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applyalignment
                                      value = lv_value ).
        lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_alignment
                                                               parent = lo_document ).
        ls_cellxfs-alignmentid += 1. "Table index starts from 1
        read table lt_alignments into ls_alignment index ls_cellxfs-alignmentid.
        ls_cellxfs-alignmentid -= 1.
        if ls_alignment-horizontal is not initial.
          lv_value = ls_alignment-horizontal.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_horizontal
                                              value = lv_value ).
        endif.
        if ls_alignment-vertical is not initial.
          lv_value = ls_alignment-vertical.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_vertical
                                              value = lv_value ).
        endif.
        if ls_alignment-wraptext eq abap_true.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_wraptext
                                              value = c_on ).
        endif.
        if ls_alignment-textrotation is not initial.
          lv_value = ls_alignment-textrotation.
          shift lv_value right deleting trailing space.
          shift lv_value left deleting leading space.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_textrotation
                                              value = lv_value ).
        endif.
        if ls_alignment-shrinktofit eq abap_true.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_shrinktofit
                                              value = c_on ).
        endif.
        if ls_alignment-indent is not initial.
          lv_value = ls_alignment-indent.
          shift lv_value right deleting trailing space.
          shift lv_value left deleting leading space.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_indent
                                              value = lv_value ).
        endif.

        lo_element->append_child( new_child = lo_sub_element_2 ).
      endif.
      if ls_cellxfs-applyprotection eq 1.
        lv_value = ls_cellxfs-applyprotection.
        condense lv_value no-gaps.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applyprotection
                                      value = lv_value ).
        lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_protection
                                                               parent = lo_document ).
        ls_cellxfs-protectionid += 1. "Table index starts from 1
        read table lt_protections into ls_protection index ls_cellxfs-protectionid.
        ls_cellxfs-protectionid -= 1.
        if ls_protection-locked is not initial.
          lv_value = ls_protection-locked.
          condense lv_value.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_locked
                                              value = lv_value ).
        endif.
        if ls_protection-hidden is not initial.
          lv_value = ls_protection-hidden.
          condense lv_value.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_hidden
                                              value = lv_value ).
        endif.
        lo_element->append_child( new_child = lo_sub_element_2 ).
      endif.
      lo_element_cellxfs->append_child( new_child = lo_element ).
    endloop.

    lo_element_root->append_child( new_child = lo_element_cellxfs ).

    " cellStyles node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_cellstyles
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_count
                                  value = '1' ).
    lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_cellstyle
                                                         parent = lo_document ).

    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_name
                                      value = 'Normal' ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_xfid
                                      value = c_off ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_builtinid
                                      value = c_off ).

    lo_element->append_child( new_child = lo_sub_element ).
    lo_element_root->append_child( new_child = lo_element ).

    " dxfs node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_dxfs
                                                     parent = lo_document ).

    lo_iterator = me->excel->get_worksheets_iterator( ).
    " get sheets
    while lo_iterator->has_next( ) eq abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).
      " Conditional formatting styles into exch sheet
      lo_iterator2 = lo_worksheet->get_style_cond_iterator( ).
      while lo_iterator2->has_next( ) eq abap_true.
        lo_style_cond ?= lo_iterator2->get_next( ).
        case lo_style_cond->rule.
* begin of change issue #366 - missing conditional rules: top10, move dfx-styles to own method
          when zcl_excel_style_cond=>c_rule_cellis.
            me->create_dxf_style( exporting
                                    iv_cell_style    = lo_style_cond->mode_cellis-cell_style
                                    io_dxf_element   = lo_element
                                    io_ixml_document = lo_document
                                    it_cellxfs       = lt_cellxfs
                                    it_fonts         = lt_fonts
                                    it_fills         = lt_fills
                                  changing
                                    cv_dfx_count     = lv_dfx_count ).

          when zcl_excel_style_cond=>c_rule_expression.
            me->create_dxf_style( exporting
                          iv_cell_style    = lo_style_cond->mode_expression-cell_style
                          io_dxf_element   = lo_element
                          io_ixml_document = lo_document
                          it_cellxfs       = lt_cellxfs
                          it_fonts         = lt_fonts
                          it_fills         = lt_fills
                        changing
                          cv_dfx_count     = lv_dfx_count ).



          when zcl_excel_style_cond=>c_rule_top10.
            me->create_dxf_style( exporting
                                    iv_cell_style    = lo_style_cond->mode_top10-cell_style
                                    io_dxf_element   = lo_element
                                    io_ixml_document = lo_document
                                    it_cellxfs       = lt_cellxfs
                                    it_fonts         = lt_fonts
                                    it_fills         = lt_fills
                                  changing
                                    cv_dfx_count     = lv_dfx_count ).

          when zcl_excel_style_cond=>c_rule_above_average.
            me->create_dxf_style( exporting
                                    iv_cell_style    = lo_style_cond->mode_above_average-cell_style
                                    io_dxf_element   = lo_element
                                    io_ixml_document = lo_document
                                    it_cellxfs       = lt_cellxfs
                                    it_fonts         = lt_fonts
                                    it_fills         = lt_fills
                                  changing
                                    cv_dfx_count     = lv_dfx_count ).
* begin of change issue #366 - missing conditional rules: top10, move dfx-styles to own method

          when zcl_excel_style_cond=>c_rule_textfunction.
            me->create_dxf_style( exporting
                                    iv_cell_style    = lo_style_cond->mode_textfunction-cell_style
                                    io_dxf_element   = lo_element
                                    io_ixml_document = lo_document
                                    it_cellxfs       = lt_cellxfs
                                    it_fonts         = lt_fonts
                                    it_fills         = lt_fills
                                  changing
                                    cv_dfx_count     = lv_dfx_count ).

          when others.
            continue.
        endcase.
      endwhile.
    endwhile.

    lv_value = lv_dfx_count.
    condense lv_value.
    lo_element->set_attribute_ns( name  = lc_xml_attr_count
                                  value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " tableStyles node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_tablestyles
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_count
                                  value = '0' ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_defaulttablestyle
                                  value = zcl_excel_table=>builtinstyle_medium9 ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_defaultpivotstyle
                                  value = zcl_excel_table=>builtinstyle_pivot_light16 ).
    lo_element_root->append_child( new_child = lo_element ).

    "write legacy color palette in case any indexed color was changed
    if excel->legacy_palette->is_modified( ) = abap_true.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_colors
                                                     parent   = lo_document ).
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_indexedcolors
                                                         parent   = lo_document ).
      lo_element->append_child( new_child = lo_sub_element ).

      lt_colors = excel->legacy_palette->get_colors( ).
      loop at lt_colors into ls_color.
        lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_rgbcolor
                                                               parent = lo_document ).
        lv_value = ls_color.
        lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_rgb
                                            value = lv_value ).
        lo_sub_element->append_child( new_child = lo_sub_element_2 ).
      endloop.

      lo_element_root->append_child( new_child = lo_element ).
    endif.

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  endmethod.


  method create_xl_styles_color_node.
    data: lo_sub_element type ref to if_ixml_element,
          lv_value       type string.

    constants: lc_xml_attr_theme   type string value 'theme',
               lc_xml_attr_rgb     type string value 'rgb',
               lc_xml_attr_indexed type string value 'indexed',
               lc_xml_attr_tint    type string value 'tint'.

    "add node only if at least one attribute is set
    check is_color-rgb is not initial or
          is_color-indexed <> zcl_excel_style_color=>c_indexed_not_set or
          is_color-theme <> zcl_excel_style_color=>c_theme_not_set or
          is_color-tint is not initial.

    lo_sub_element = io_document->create_simple_element(
        name      = iv_color_elem_name
        parent    = io_parent ).

    if is_color-rgb is not initial.
      lv_value = is_color-rgb.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_rgb
                                        value = lv_value ).
    endif.

    if is_color-indexed <> zcl_excel_style_color=>c_indexed_not_set.
      lv_value = zcl_excel_common=>number_to_excel_string( is_color-indexed ).
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_indexed
                                        value = lv_value ).
    endif.

    if is_color-theme <> zcl_excel_style_color=>c_theme_not_set.
      lv_value = zcl_excel_common=>number_to_excel_string( is_color-theme ).
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_theme
                                        value = lv_value ).
    endif.

    if is_color-tint is not initial.
      lv_value = zcl_excel_common=>number_to_excel_string( is_color-tint ).
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_tint
                                        value = lv_value ).
    endif.

    io_parent->append_child( new_child = lo_sub_element ).
  endmethod.


  method create_xl_styles_font_node.

    constants: lc_xml_node_b      type string value 'b',            "bold
               lc_xml_node_i      type string value 'i',            "italic
               lc_xml_node_u      type string value 'u',            "underline
               lc_xml_node_strike type string value 'strike',       "strikethrough
               lc_xml_node_sz     type string value 'sz',
               lc_xml_node_name   type string value 'name',
               lc_xml_node_rfont  type string value 'rFont',
               lc_xml_node_family type string value 'family',
               lc_xml_node_scheme type string value 'scheme',
               lc_xml_attr_val    type string value 'val'.

    data: lo_document     type ref to if_ixml_document,
          lo_element_font type ref to if_ixml_element,
          ls_font         type zif_excel_data_decl=>zexcel_s_style_font,
          lo_sub_element  type ref to if_ixml_element,
          lv_value        type string.

    lo_document = io_document.
    lo_element_font = io_parent.
    ls_font = is_font.

    if ls_font-bold eq abap_true.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_b
                                                           parent = lo_document ).
      lo_element_font->append_child( new_child = lo_sub_element ).
    endif.
    if ls_font-italic eq abap_true.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_i
                                                           parent = lo_document ).
      lo_element_font->append_child( new_child = lo_sub_element ).
    endif.
    if ls_font-underline eq abap_true.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_u
                                                           parent = lo_document ).
      lv_value = ls_font-underline_mode.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_val
                                        value = lv_value ).
      lo_element_font->append_child( new_child = lo_sub_element ).
    endif.
    if ls_font-strikethrough eq abap_true.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_strike
                                                           parent = lo_document ).
      lo_element_font->append_child( new_child = lo_sub_element ).
    endif.
    "size
    lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_sz
                                                         parent = lo_document ).
    lv_value = ls_font-size.
    shift lv_value right deleting trailing space.
    shift lv_value left deleting leading space.
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_val
                                      value = lv_value ).
    lo_element_font->append_child( new_child = lo_sub_element ).
    "color
    create_xl_styles_color_node(
        io_document        = lo_document
        io_parent          = lo_element_font
        is_color           = ls_font-color ).

    "name
    if iv_use_rtf = abap_false.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_name
                                                           parent = lo_document ).
    else.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_rfont
                                                           parent = lo_document ).
    endif.
    lv_value = ls_font-name.
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_val
                                      value = lv_value ).
    lo_element_font->append_child( new_child = lo_sub_element ).
    "family
    lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_family
                                                         parent = lo_document ).
    lv_value = ls_font-family.
    shift lv_value right deleting trailing space.
    shift lv_value left deleting leading space.
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_val
                                      value = lv_value ).
    lo_element_font->append_child( new_child = lo_sub_element ).
    "scheme
    if ls_font-scheme is not initial.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_scheme
                                                           parent = lo_document ).
      lv_value = ls_font-scheme.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_val
                                        value = lv_value ).
      lo_element_font->append_child( new_child = lo_sub_element ).
    endif.

  endmethod.

  method create_xl_table.
    data lc_xml_node_table        type string                                     value 'table'.
    " Node attributes
    data lc_xml_attr_id           type string                                     value 'id'.
    data lc_xml_attr_name         type string                                     value 'name'.
    data lc_xml_attr_display_name type string                                     value 'displayName'.
    data lc_xml_attr_ref          type string                                     value 'ref'.
    data lc_xml_attr_totals       type string                                     value 'totalsRowShown'.
    " Node namespace
    data lc_xml_node_table_ns     type string
                                  value 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'.
    " Node id

    data lo_document              type ref to if_ixml_document.
    data lo_element_root          type ref to if_ixml_element.
    data lo_element               type ref to if_ixml_element.
    data lo_element2              type ref to if_ixml_element.
    data lo_element3              type ref to if_ixml_element.
    data lv_table_name            type string.
    data lv_id                    type i.
    data lv_match                 type i.
    data lv_ref                   type string.
    data lv_value                 type string.
    data lv_num_columns           type i.
    data ls_fieldcat              type zif_excel_data_decl=>zexcel_s_fieldcatalog.

    " ---------------------------------------------------------------------
    " STEP 1: Create xml
    lo_document = create_xml_document( ).

    " ---------------------------------------------------------------------
    " STEP 3: Create main node table
    lo_element_root = lo_document->create_simple_element( name   = lc_xml_node_table
                                                          parent = lo_document ).

    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_table_ns  ).

    lv_id = io_table->get_id( ).
    lv_value = zcl_excel_common=>number_to_excel_string( ip_value = lv_id ).
    lo_element_root->set_attribute_ns( name  = lc_xml_attr_id
                                       value = lv_value ).

    find all occurrences of pcre '[^_a-zA-Z0-9]' in io_table->settings-table_name ignoring case match count lv_match.
    if io_table->settings-table_name is not initial and lv_match = 0.
      " Name rules (https://support.microsoft.com/en-us/office/rename-an-excel-table-fbf49a4f-82a3-43eb-8ba2-44d21233b114)
      "   - You can't use "C", "c", "R", or "r" for the name, because they're already designated as a shortcut for selecting the column or row for the active cell when you enter them in the Name or Go To box.
      "   - Don't use cell references — Names can't be the same as a cell reference, such as Z$100 or R1C1
      if    ( strlen( io_table->settings-table_name ) = 1 and io_table->settings-table_name co 'CcRr' )
         or zcl_excel_common=>shift_formula( iv_reference_formula = io_table->settings-table_name
                                             iv_shift_cols        = 0
                                             iv_shift_rows        = 1 ) <> io_table->settings-table_name.
        lv_table_name = io_table->get_name( ).
      else.
        lv_table_name = io_table->settings-table_name.
      endif.
    else.
      lv_table_name = io_table->get_name( ).
    endif.
    lo_element_root->set_attribute_ns( name  = lc_xml_attr_name
                                       value = lv_table_name ).

    lo_element_root->set_attribute_ns( name  = lc_xml_attr_display_name
                                       value = lv_table_name ).

    lv_ref = io_table->get_reference( ).
    lo_element_root->set_attribute_ns( name  = lc_xml_attr_ref
                                       value = lv_ref ).
    if io_table->has_totals( ) = abap_true.
      lo_element_root->set_attribute_ns( name  = 'totalsRowCount'
                                         value = '1' ).
    else.
      lo_element_root->set_attribute_ns( name  = lc_xml_attr_totals
                                         value = '0' ).
    endif.

    " ---------------------------------------------------------------------
    " STEP 4: Create subnodes

    " autoFilter
    if io_table->settings-nofilters = abap_false.
      lo_element = lo_document->create_simple_element( name   = 'autoFilter'
                                                       parent = lo_document ).

      lv_ref = io_table->get_reference( ip_include_totals_row = abap_false ).
      lo_element->set_attribute_ns( name  = 'ref'
                                    value = lv_ref ).

      lo_element_root->append_child( new_child = lo_element ).
    endif.

    " columns
    lo_element = lo_document->create_simple_element( name   = 'tableColumns'
                                                     parent = lo_document ).

    loop at io_table->fieldcat into ls_fieldcat where dynpfld = abap_true.
      lv_num_columns += 1.
    endloop.

    lv_value = lv_num_columns.
    lv_value = condense( lv_value ).
    lo_element->set_attribute_ns( name  = 'count'
                                  value = lv_value ).

    lo_element_root->append_child( new_child = lo_element ).

    loop at io_table->fieldcat into ls_fieldcat where dynpfld = abap_true.
      lo_element2 = lo_document->create_simple_element_ns( name   = 'tableColumn'
                                                           parent = lo_element ).

      lv_value = ls_fieldcat-position.
      shift lv_value left deleting leading '0'.
      lo_element2->set_attribute_ns( name  = 'id'
                                     value = lv_value ).

      lv_value = ls_fieldcat-column_name.

      " The text "_x...._", with "_x" not "_X", with exactly 4 ".", each being 0-9 a-f or A-F (case insensitive), is interpreted
      " like Unicode character U+.... (e.g. "_x0041_" is rendered like "A") is for characters.
      " To not interpret it, Excel replaces the first "_" is to be replaced with "_x005f_".
      if lv_value cs '_x'.
        replace all occurrences of pcre '_(x[0-9a-fA-F]{4}_)' in lv_value with '_x005f_$1' respecting case.
      endif.

      " XML chapter 2.2: Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
      " NB: although Excel supports _x0009_, it's not rendered except if you edit the text.
      " Excel considers _x000d_ as being an error (_x000a_ is sufficient and rendered).
      replace all occurrences of cl_abap_char_utilities=>newline in lv_value with '_x000a_'.
      replace all occurrences of cl_abap_char_utilities=>cr_lf(1) in lv_value with ``.
      replace all occurrences of cl_abap_char_utilities=>horizontal_tab in lv_value with '_x0009_'.

      lo_element2->set_attribute_ns( name  = 'name'
                                     value = lv_value ).

      if ls_fieldcat-totals_function is not initial.
        lo_element2->set_attribute_ns( name  = 'totalsRowFunction'
                                       value = ls_fieldcat-totals_function ).
      endif.

      if ls_fieldcat-column_formula is not initial.
        lv_value = ls_fieldcat-column_formula.
        lv_value = condense( lv_value ).
        lo_element3 = lo_document->create_simple_element_ns( name   = 'calculatedColumnFormula'
                                                             parent = lo_element2 ).
        lo_element3->set_value( lv_value ).
        lo_element2->append_child( new_child = lo_element3 ).
      endif.

      lo_element->append_child( new_child = lo_element2 ).
    endloop.

    lo_element = lo_document->create_simple_element( name   = 'tableStyleInfo'
                                                     parent = lo_element_root ).

    lo_element->set_attribute_ns( name  = 'name'
                                  value = io_table->settings-table_style  ).

    lo_element->set_attribute_ns( name  = 'showFirstColumn'
                                  value = '0' ).

    lo_element->set_attribute_ns( name  = 'showLastColumn'
                                  value = '0' ).

    if io_table->settings-show_row_stripes = abap_true.
      lv_value = '1'.
    else.
      lv_value = '0'.
    endif.

    lo_element->set_attribute_ns( name  = 'showRowStripes'
                                  value = lv_value ).

    if io_table->settings-show_column_stripes = abap_true.
      lv_value = '1'.
    else.
      lv_value = '0'.
    endif.

    lo_element->set_attribute_ns( name  = 'showColumnStripes'
                                  value = lv_value ).

    lo_element_root->append_child( new_child = lo_element ).
    " ---------------------------------------------------------------------
    " STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  endmethod.


  method create_xl_theme.
    data: lo_theme type ref to zcl_excel_theme.

    excel->get_theme(
    importing
      eo_theme = lo_theme
      ).
    if lo_theme is initial.
      create object lo_theme.
    endif.
    ep_content = lo_theme->write_theme( ).

  endmethod.


  method create_xl_workbook.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-07
*              - ...
* changes: aligning code
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmoecker,                             2012-12-01
* changes:  correction of pointer to localSheetId
*--------------------------------------------------------------------*

** Constant node name
    data: lc_xml_node_workbook           type string value 'workbook',
          lc_xml_node_fileversion        type string value 'fileVersion',
          lc_xml_node_workbookpr         type string value 'workbookPr',
          lc_xml_node_bookviews          type string value 'bookViews',
          lc_xml_node_workbookview       type string value 'workbookView',
          lc_xml_node_sheets             type string value 'sheets',
          lc_xml_node_sheet              type string value 'sheet',
          lc_xml_node_calcpr             type string value 'calcPr',
          lc_xml_node_workbookprotection type string value 'workbookProtection',
          lc_xml_node_definednames       type string value 'definedNames',
          lc_xml_node_definedname        type string value 'definedName',
          " Node attributes
          lc_xml_attr_appname            type string value 'appName',
          lc_xml_attr_lastedited         type string value 'lastEdited',
          lc_xml_attr_lowestedited       type string value 'lowestEdited',
          lc_xml_attr_rupbuild           type string value 'rupBuild',
          lc_xml_attr_xwindow            type string value 'xWindow',
          lc_xml_attr_ywindow            type string value 'yWindow',
          lc_xml_attr_windowwidth        type string value 'windowWidth',
          lc_xml_attr_windowheight       type string value 'windowHeight',
          lc_xml_attr_activetab          type string value 'activeTab',
          lc_xml_attr_name               type string value 'name',
          lc_xml_attr_sheetid            type string value 'sheetId',
          lc_xml_attr_state              type string value 'state',
          lc_xml_attr_id                 type string value 'id',
          lc_xml_attr_calcid             type string value 'calcId',
          lc_xml_attr_lockrevision       type string value 'lockRevision',
          lc_xml_attr_lockstructure      type string value 'lockStructure',
          lc_xml_attr_lockwindows        type string value 'lockWindows',
          lc_xml_attr_revisionspassword  type string value 'revisionsPassword',
          lc_xml_attr_workbookpassword   type string value 'workbookPassword',
          lc_xml_attr_hidden             type string value 'hidden',
          lc_xml_attr_localsheetid       type string value 'localSheetId',
          " Node namespace
          lc_r_ns                        type string value 'r',
          lc_xml_node_ns                 type string value 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          lc_xml_node_r_ns               type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          " Node id
          lc_xml_node_ridx_id            type string value 'rId#'.

    data: lo_document       type ref to if_ixml_document,
          lo_element_root   type ref to if_ixml_element,
          lo_element        type ref to if_ixml_element,
          lo_element_range  type ref to if_ixml_element,
          lo_sub_element    type ref to if_ixml_element,
          lo_iterator       type ref to zcl_excel_collection_iterator,
          lo_iterator_range type ref to zcl_excel_collection_iterator,
          lo_worksheet      type ref to zcl_excel_worksheet,
          lo_range          type ref to zcl_excel_range,
          lo_autofilters    type ref to zcl_excel_autofilters,
          lo_autofilter     type ref to zcl_excel_autofilter.

    data: lv_xml_node_ridx_id type string,
          lv_value            type string,
          lv_syindex          type string,
          lv_active_sheet     type zif_excel_data_decl=>zexcel_active_worksheet.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_workbook
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:r'
                                       value = lc_xml_node_r_ns ).

**********************************************************************
* STEP 4: Create subnode
    " fileVersion node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_fileversion
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_appname
                                  value = 'xl' ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_lastedited
                                  value = '4' ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_lowestedited
                                  value = '4' ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_rupbuild
                                  value = '4506' ).
    lo_element_root->append_child( new_child = lo_element ).

    " fileVersion node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_workbookpr
                                                     parent = lo_document ).
    lo_element_root->append_child( new_child = lo_element ).

    " workbookProtection node
    if me->excel->zif_excel_book_protection~protected eq abap_true.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_workbookprotection
                                                       parent = lo_document ).
      lv_value = me->excel->zif_excel_book_protection~workbookpassword.
      if lv_value is not initial.
        lo_element->set_attribute_ns( name  = lc_xml_attr_workbookpassword
                                      value = lv_value ).
      endif.
      lv_value = me->excel->zif_excel_book_protection~revisionspassword.
      if lv_value is not initial.
        lo_element->set_attribute_ns( name  = lc_xml_attr_revisionspassword
                                      value = lv_value ).
      endif.
      lv_value = me->excel->zif_excel_book_protection~lockrevision.
      condense lv_value no-gaps.
      lo_element->set_attribute_ns( name  = lc_xml_attr_lockrevision
                                    value = lv_value ).
      lv_value = me->excel->zif_excel_book_protection~lockstructure.
      condense lv_value no-gaps.
      lo_element->set_attribute_ns( name  = lc_xml_attr_lockstructure
                                    value = lv_value ).
      lv_value = me->excel->zif_excel_book_protection~lockwindows.
      condense lv_value no-gaps.
      lo_element->set_attribute_ns( name  = lc_xml_attr_lockwindows
                                    value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    endif.

    " bookviews node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_bookviews
                                                     parent = lo_document ).
    " bookview node
    lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_workbookview
                                                         parent = lo_document ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_xwindow
                                      value = '120' ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_ywindow
                                      value = '120' ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_windowwidth
                                      value = '19035' ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_windowheight
                                      value = '8445' ).
    " Set Active Sheet
    lv_active_sheet = excel->get_active_sheet_index( ).
* issue #365 - test if sheet exists - otherwise set active worksheet to 1
    lo_worksheet = excel->get_worksheet_by_index( lv_active_sheet ).
    if lo_worksheet is not bound.
      lv_active_sheet = 1.
      excel->set_active_sheet_index( lv_active_sheet ).
    endif.
    if lv_active_sheet > 1.
      lv_active_sheet = lv_active_sheet - 1.
      lv_value = lv_active_sheet.
      condense lv_value.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_activetab
                                        value = lv_value ).
    endif.
    lo_element->append_child( new_child = lo_sub_element )." bookview node
    lo_element_root->append_child( new_child = lo_element )." bookviews node

    " sheets node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_sheets
                                                     parent = lo_document ).
    lo_iterator = excel->get_worksheets_iterator( ).

    " ranges node
    lo_element_range = lo_document->create_simple_element( name   = lc_xml_node_definednames " issue 163 +
                                                           parent = lo_document ).           " issue 163 +

    while lo_iterator->has_next( ) eq abap_true.
      " sheet node
      lo_sub_element = lo_document->create_simple_element_ns( name   = lc_xml_node_sheet
                                                              parent = lo_document ).
      lo_worksheet ?= lo_iterator->get_next( ).
      lv_syindex = sy-index.                                                                  " question by Stefan SchmÃ¶cker 2012-12-02:  sy-index seems to do the job - but is it proven to work or purely coincedence
      lv_value = lo_worksheet->get_title( ).
      shift lv_syindex right deleting trailing space.
      shift lv_syindex left deleting leading space.
      lv_xml_node_ridx_id = lc_xml_node_ridx_id.
      replace all occurrences of '#' in lv_xml_node_ridx_id with lv_syindex.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_name
                                        value = lv_value ).
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_sheetid
                                        value = lv_syindex ).
      if lo_worksheet->zif_excel_sheet_properties~hidden eq zif_excel_sheet_properties=>c_hidden.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_state
                                          value = 'hidden' ).
      elseif lo_worksheet->zif_excel_sheet_properties~hidden eq zif_excel_sheet_properties=>c_veryhidden.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_state
                                          value = 'veryHidden' ).
      endif.
      lo_sub_element->set_attribute_ns( name    = lc_xml_attr_id
                                        prefix  = lc_r_ns
                                        value   = lv_xml_node_ridx_id ).
      lo_element->append_child( new_child = lo_sub_element ). " sheet node

      " issue 163 >>>
      lo_iterator_range = lo_worksheet->get_ranges_iterator( ).

*--------------------------------------------------------------------*
* Defined names sheetlocal:  Ranges, Repeat rows and columns
*--------------------------------------------------------------------*
      while lo_iterator_range->has_next( ) eq abap_true.
        " range node
        lo_sub_element = lo_document->create_simple_element_ns( name   = lc_xml_node_definedname
                                                                parent = lo_document ).
        lo_range ?= lo_iterator_range->get_next( ).
        lv_value = lo_range->name.

        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_name
                                          value = lv_value ).

*      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_localsheetid           "del #235 Repeat rows/cols - EXCEL starts couting from zero
*                                        value = lv_xml_node_ridx_id ).             "del #235 Repeat rows/cols - and needs absolute referencing to localSheetId
        lv_value   = lv_syindex - 1.                                                  "ins #235 Repeat rows/cols
        condense lv_value no-gaps.                                                    "ins #235 Repeat rows/cols
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_localsheetid
                                          value = lv_value ).

        lv_value = lo_range->get_value( ).
        lo_sub_element->set_value( value = lv_value ).
        lo_element_range->append_child( new_child = lo_sub_element ). " range node

      endwhile.
      " issue 163 <<<

    endwhile.
    lo_element_root->append_child( new_child = lo_element )." sheets node


*--------------------------------------------------------------------*
* Defined names workbookgolbal:  Ranges
*--------------------------------------------------------------------*
*  " ranges node
*  lo_element = lo_document->create_simple_element( name   = lc_xml_node_definednames " issue 163 -
*                                                   parent = lo_document ).           " issue 163 -
    lo_iterator = excel->get_ranges_iterator( ).

    while lo_iterator->has_next( ) eq abap_true.
      " range node
      lo_sub_element = lo_document->create_simple_element_ns( name   = lc_xml_node_definedname
                                                              parent = lo_document ).
      lo_range ?= lo_iterator->get_next( ).
      lv_value = lo_range->name.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_name
                                        value = lv_value ).
      lv_value = lo_range->get_value( ).
      lo_sub_element->set_value( value = lv_value ).
      lo_element_range->append_child( new_child = lo_sub_element ). " range node

    endwhile.

*--------------------------------------------------------------------*
* Defined names - Autofilters ( also sheetlocal )
*--------------------------------------------------------------------*
    lo_autofilters = excel->get_autofilters_reference( ).
    if lo_autofilters->is_empty( ) = abap_false.
      lo_iterator = excel->get_worksheets_iterator( ).
      while lo_iterator->has_next( ) eq abap_true.

        lo_worksheet ?= lo_iterator->get_next( ).
        lv_syindex = sy-index - 1 .
        lo_autofilter = lo_autofilters->get( io_worksheet = lo_worksheet ).
        if lo_autofilter is bound.
          lo_sub_element = lo_document->create_simple_element_ns( name   = lc_xml_node_definedname
                                                                  parent = lo_document ).
          lv_value = lo_autofilters->c_autofilter.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_name
                                            value = lv_value ).
          lv_value = lv_syindex.
          condense lv_value no-gaps.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_localsheetid
                                            value = lv_value ).
          lv_value = '1'. " Always hidden
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_hidden
                                            value = lv_value ).
          lv_value = lo_autofilter->get_filter_reference( ).
          lo_sub_element->set_value( value = lv_value ).
          lo_element_range->append_child( new_child = lo_sub_element ). " range node
        endif.

      endwhile.
    endif.
    lo_element_root->append_child( new_child = lo_element_range ).                      " ranges node


    " calcPr node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_calcpr
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_calcid
                                  value = '125725' ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  endmethod.


  method create_xml_document.
    data lo_encoding type ref to if_ixml_encoding.
    lo_encoding = me->ixml->if_ixml_core~create_encoding( byte_order = if_ixml_encoding=>co_platform_endian
                                             character_set = 'utf-8' ).
    ro_document = me->ixml->if_ixml_core~create_document( ).
    ro_document->set_encoding( lo_encoding ).
    ro_document->set_standalone( abap_true ).
  endmethod.


  method escape_string_value.

    data: lt_character_positions type table of i,
          lv_character_position  type i,
          lv_escaped_value       type string.

    result = iv_value.
    if result ca control_characters.

      clear lt_character_positions.
      append sy-fdpos to lt_character_positions.
      lv_character_position = sy-fdpos + 1.
      while result+lv_character_position ca control_characters.
        lv_character_position += sy-fdpos.
        append lv_character_position to lt_character_positions.
        lv_character_position += 1.
      endwhile.
      sort lt_character_positions by table_line descending.

      loop at lt_character_positions into lv_character_position.
        "@TODO: to be fixed by Juwin
        "lv_escaped_value = |_x{ cl_abap_conv_out_ce=>uccp( substring( val = result off = lv_character_position len = 1 ) ) }_|.
        replace section offset lv_character_position length 1 of result with lv_escaped_value.
      endloop.
    endif.

  endmethod.


  method flag2bool.


    if ip_flag eq abap_true.
      ep_boolean = 'true'.
    else.
      ep_boolean = 'false'.
    endif.
  endmethod.


  method get_shared_string_index.


    data ls_shared_string type zif_excel_data_decl=>zexcel_s_shared_string.

    if it_rtf is initial.
      read table shared_strings into ls_shared_string with table key string_value = ip_cell_value.
      ep_index = ls_shared_string-string_no.
    else.
      loop at shared_strings into ls_shared_string where string_value = ip_cell_value
                                                     and rtf_tab = it_rtf.

        ep_index = ls_shared_string-string_no.
        exit.
      endloop.
    endif.

  endmethod.


  method is_formula_shareable.
    data: lv_test_shared type string.

    ep_shareable = abap_false.
    if ip_formula na '!'.
      lv_test_shared = zcl_excel_common=>shift_formula(
          iv_reference_formula = ip_formula
          iv_shift_cols        = 1
          iv_shift_rows        = 1 ).
      if lv_test_shared <> ip_formula.
        ep_shareable = abap_true.
      endif.
    endif.
  endmethod.


  method number2string.
    ep_string = ip_number.
    condense ep_string.
  endmethod.


  method render_xml_document.
    data lo_streamfactory type ref to if_ixml_stream_factory_core.
    data lo_ostream       type ref to if_ixml_ostream_core.
    data lo_renderer      type ref to if_ixml_renderer_core.
    data lv_xstring type xstring.
    data lv_string        type string.

    " So that the rendering of io_document to a XML text in UTF-8 XSTRING works for all Unicode characters (Chinese,
    " emoticons, etc.) the method CREATE_OSTREAM_CSTRING must be used instead of CREATE_OSTREAM_XSTRING as explained
    " in note 2922674 below (original there: https://launchpad.support.sap.com/#/notes/2922674), and then the STRING
    " variable can be converted into UTF-8.
    "
    " Excerpt from Note 2922674 - Support for Unicode Characters U+10000 to U+10FFFF in the iXML kernel library / ABAP package SIXML.
    "
    "   You are running a unicode system with SAP Netweaver / SAP_BASIS release equal or lower than 7.51.
    "
    "   Some functions in the iXML kernel library / ABAP package SIXML does not fully or incorrectly support unicode
    "   characters of the supplementary planes. This is caused by using UCS-2 in codepage conversion functions.
    "   Therefore, when reading from iXML input steams, the characters from the supplementary planes, that are not
    "   supported by UCS-2, might be replaced by the character #. When writing to iXML output streams, UTF-16 surrogate
    "   pairs, representing characters from the supplementary planes, might be incorrectly encoded in UTF-8.
    "
    "   The characters incorrectly encoded in UTF-8, might be accepted as input for the iXML parser or external parsers,
    "   but might also be rejected.
    "
    "   Support for unicode characters of the supplementary planes was introduced for SAP_BASIS 7.51 or lower with note
    "   2220720, but later withdrawn with note 2346627 for functional issues.
    "
    "   Characters of the supplementary planes are supported with ABAP Platform 1709 / SAP_BASIS 7.52 and higher.
    "
    "   Please note, that the iXML runtime behaves like the ABAP runtime concerning the handling of unicode characters of
    "   the supplementary planes. In iXML and ABAP, these characters have length 2 (as returned by ABAP build-in function
    "   STRLEN), and string processing functions like SUBSTRING might split these characters into 2 invalid characters
    "   with length 1. These invalid characters are commonly referred to as broken surrogate pairs.
    "
    "   A workaround for the incorrect UTF-8 encoding in SAP_BASIS 7.51 or lower is to render the document to an ABAP
    "   variable with type STRING using a output stream created with factory method IF_IXML_STREAM_FACTORY=>CREATE_OSTREAM_CSTRING
    "   and then to convert the STRING variable to UTF-8 using method CL_ABAP_CODEPAGE=>CONVERT_TO.

    " 1) RENDER TO XML STRING
    lo_streamfactory = me->ixml->if_ixml_core~create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = lv_xstring ).
    lo_renderer = me->ixml->if_ixml_core~create_renderer( ostream  = lo_ostream document = io_document ).
    lo_renderer->render( ).

    " 2) CONVERT IT TO UTF-8
    "-----------------
    " The beginning of the XML string has these 57 characters:
    "   X<?xml version="1.0" encoding="utf-16" standalone="yes"?>
    "   (where "X" is the special character corresponding to the utf-16 BOM, hexadecimal FFFE or FEFF,
    "   but there's no "X" in non-Unicode SAP systems)
    " The encoding must be removed otherwise Excel would fail to decode correctly the UTF-8 XML.
    " For a better performance, it's assumed that "encoding" is in the first 100 characters.
    lv_string = xco_cp=>xstring( lv_xstring )->as_string( xco_cp_character=>code_page->utf_8 )->value.
    if strlen( lv_string ) < 100.
      replace pcre  'encoding="[^"]+"' in lv_string with ``.
    else.
      replace pcre  'encoding="[^"]+"' in section length 100 of lv_string with ``.
    endif.
    " Convert XML text to UTF-8 (NB: if 2 first bytes are the UTF-16 BOM, they are converted into 3 bytes of UTF-8 BOM)
    ep_content = xco_cp=>string( lv_string )->as_xstring( xco_cp_character=>code_page->utf_8 )->value.
    " Add the UTF-8 Byte Order Mark if missing (NB: that serves as substitute of "encoding")
    if xstrlen( ep_content ) >= 3 and ep_content(3) <> cl_abap_char_utilities=>byte_order_mark_utf8.
      concatenate cl_abap_char_utilities=>byte_order_mark_utf8 ep_content into ep_content in byte mode.
    endif.

  endmethod.


  method set_vml_shape_footer.

    constants: lc_shape               type string value '<v:shape id="{ID}" o:spid="_x0000_s1025" type="#_x0000_t75" style=''position:absolute;margin-left:0;margin-top:0;width:{WIDTH}pt;height:{HEIGHT}pt; z-index:1''>',
               lc_shape_image         type string value '<v:imagedata o:relid="{RID}" o:title="Logo Title"/><o:lock v:ext="edit" rotation="t"/></v:shape>',
               lc_shape_header_center type string value 'CH',
               lc_shape_header_left   type string value 'LH',
               lc_shape_header_right  type string value 'RH',
               lc_shape_footer_center type string value 'CF',
               lc_shape_footer_left   type string value 'LF',
               lc_shape_footer_right  type string value 'RF'.

    data: lv_content_left         type string,
          lv_content_center       type string,
          lv_content_right        type string,
          lv_content_image_left   type string,
          lv_content_image_center type string,
          lv_content_image_right  type string,
          lv_value                type string,
          ls_drawing_position     type zif_excel_data_decl=>zexcel_drawing_position.

    if is_footer-left_image is not initial.
      lv_content_left = lc_shape.
      replace '{ID}' in lv_content_left with lc_shape_footer_left.
      ls_drawing_position = is_footer-left_image->get_position( ).
      if ls_drawing_position-size-height is not initial.
        lv_value = ls_drawing_position-size-height.
      else.
        lv_value = '100'.
      endif.
      condense lv_value.
      replace '{HEIGHT}' in lv_content_left with lv_value.
      if ls_drawing_position-size-width is not initial.
        lv_value = ls_drawing_position-size-width.
      else.
        lv_value = '100'.
      endif.
      condense lv_value.
      replace '{WIDTH}' in lv_content_left with lv_value.
      lv_content_image_left = lc_shape_image.
      lv_value = is_footer-left_image->get_index( ).
      condense lv_value.
      concatenate 'rId' lv_value into lv_value.
      replace '{RID}' in lv_content_image_left with lv_value.
    endif.
    if is_footer-center_image is not initial.
      lv_content_center = lc_shape.
      replace '{ID}' in lv_content_center with lc_shape_footer_center.
      ls_drawing_position = is_footer-left_image->get_position( ).
      if ls_drawing_position-size-height is not initial.
        lv_value = ls_drawing_position-size-height.
      else.
        lv_value = '100'.
      endif.
      condense lv_value.
      replace '{HEIGHT}' in lv_content_center with lv_value.
      if ls_drawing_position-size-width is not initial.
        lv_value = ls_drawing_position-size-width.
      else.
        lv_value = '100'.
      endif.
      condense lv_value.
      replace '{WIDTH}' in lv_content_center with lv_value.
      lv_content_image_center = lc_shape_image.
      lv_value = is_footer-center_image->get_index( ).
      condense lv_value.
      concatenate 'rId' lv_value into lv_value.
      replace '{RID}' in lv_content_image_center with lv_value.
    endif.
    if is_footer-right_image is not initial.
      lv_content_right = lc_shape.
      replace '{ID}' in lv_content_right with lc_shape_footer_right.
      ls_drawing_position = is_footer-left_image->get_position( ).
      if ls_drawing_position-size-height is not initial.
        lv_value = ls_drawing_position-size-height.
      else.
        lv_value = '100'.
      endif.
      condense lv_value.
      replace '{HEIGHT}' in lv_content_right with lv_value.
      if ls_drawing_position-size-width is not initial.
        lv_value = ls_drawing_position-size-width.
      else.
        lv_value = '100'.
      endif.
      condense lv_value.
      replace '{WIDTH}' in lv_content_right with lv_value.
      lv_content_image_right = lc_shape_image.
      lv_value = is_footer-right_image->get_index( ).
      condense lv_value.
      concatenate 'rId' lv_value into lv_value.
      replace '{RID}' in lv_content_image_right with lv_value.
    endif.

    concatenate lv_content_left
                lv_content_image_left
                lv_content_center
                lv_content_image_center
                lv_content_right
                lv_content_image_right
           into ep_content.

  endmethod.


  method set_vml_shape_header.

*  CONSTANTS: lc_shape TYPE string VALUE '<v:shape id="{ID}" o:spid="_x0000_s1025" type="#_x0000_t75" style=''position:absolute;margin-left:0;margin-top:0;width:198.75pt;height:48.75pt; z-index:1''>',
    constants: lc_shape               type string value '<v:shape id="{ID}" o:spid="_x0000_s1025" type="#_x0000_t75" style=''position:absolute;margin-left:0;margin-top:0;width:{WIDTH}pt;height:{HEIGHT}pt; z-index:1''>',
               lc_shape_image         type string value '<v:imagedata o:relid="{RID}" o:title="Logo Title"/><o:lock v:ext="edit" rotation="t"/></v:shape>',
               lc_shape_header_center type string value 'CH',
               lc_shape_header_left   type string value 'LH',
               lc_shape_header_right  type string value 'RH',
               lc_shape_footer_center type string value 'CF',
               lc_shape_footer_left   type string value 'LF',
               lc_shape_footer_right  type string value 'RF'.

    data: lv_content_left         type string,
          lv_content_center       type string,
          lv_content_right        type string,
          lv_content_image_left   type string,
          lv_content_image_center type string,
          lv_content_image_right  type string,
          lv_value                type string,
          ls_drawing_position     type zif_excel_data_decl=>zexcel_drawing_position.

    clear ep_content.

    if is_header-left_image is not initial.
      lv_content_left = lc_shape.
      replace '{ID}' in lv_content_left with lc_shape_header_left.
      ls_drawing_position = is_header-left_image->get_position( ).
      if ls_drawing_position-size-height is not initial.
        lv_value = ls_drawing_position-size-height.
      else.
        lv_value = '100'.
      endif.
      condense lv_value.
      replace '{HEIGHT}' in lv_content_left with lv_value.
      if ls_drawing_position-size-width is not initial.
        lv_value = ls_drawing_position-size-width.
      else.
        lv_value = '100'.
      endif.
      condense lv_value.
      replace '{WIDTH}' in lv_content_left with lv_value.
      lv_content_image_left = lc_shape_image.
      lv_value = is_header-left_image->get_index( ).
      condense lv_value.
      concatenate 'rId' lv_value into lv_value.
      replace '{RID}' in lv_content_image_left with lv_value.
    endif.
    if is_header-center_image is not initial.
      lv_content_center = lc_shape.
      replace '{ID}' in lv_content_center with lc_shape_header_center.
      ls_drawing_position = is_header-center_image->get_position( ).
      if ls_drawing_position-size-height is not initial.
        lv_value = ls_drawing_position-size-height.
      else.
        lv_value = '100'.
      endif.
      condense lv_value.
      replace '{HEIGHT}' in lv_content_center with lv_value.
      if ls_drawing_position-size-width is not initial.
        lv_value = ls_drawing_position-size-width.
      else.
        lv_value = '100'.
      endif.
      condense lv_value.
      replace '{WIDTH}' in lv_content_center with lv_value.
      lv_content_image_center = lc_shape_image.
      lv_value = is_header-center_image->get_index( ).
      condense lv_value.
      concatenate 'rId' lv_value into lv_value.
      replace '{RID}' in lv_content_image_center with lv_value.
    endif.
    if is_header-right_image is not initial.
      lv_content_right = lc_shape.
      replace '{ID}' in lv_content_right with lc_shape_header_right.
      ls_drawing_position = is_header-right_image->get_position( ).
      if ls_drawing_position-size-height is not initial.
        lv_value = ls_drawing_position-size-height.
      else.
        lv_value = '100'.
      endif.
      condense lv_value.
      replace '{HEIGHT}' in lv_content_right with lv_value.
      if ls_drawing_position-size-width is not initial.
        lv_value = ls_drawing_position-size-width.
      else.
        lv_value = '100'.
      endif.
      condense lv_value.
      replace '{WIDTH}' in lv_content_right with lv_value.
      lv_content_image_right = lc_shape_image.
      lv_value = is_header-right_image->get_index( ).
      condense lv_value.
      concatenate 'rId' lv_value into lv_value.
      replace '{RID}' in lv_content_image_right with lv_value.
    endif.

    concatenate lv_content_left
                lv_content_image_left
                lv_content_center
                lv_content_image_center
                lv_content_right
                lv_content_image_right
           into ep_content.

  endmethod.


  method set_vml_string.

    data:
      ld_1           type string,
      ld_2           type string,
      ld_3           type string,
      ld_4           type string,
      ld_5           type string,
      ld_7           type string,

      lv_relation_id type i,
      lo_iterator    type ref to zcl_excel_collection_iterator,
      lo_worksheet   type ref to zcl_excel_worksheet,
      ls_odd_header  type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
      ls_odd_footer  type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
      ls_even_header type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
      ls_even_footer type zif_excel_data_decl=>zexcel_s_worksheet_head_foot.


* INIT_RESULT
    clear ep_content.


* BODY
    ld_1 = '<xml xmlns:v="urn:schemas-microsoft-com:vml"  xmlns:o="urn:schemas-microsoft-com:office:office"  xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'.
    ld_2 = '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/>'.
    ld_3 = '<v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/>'.
    ld_4 = '<v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>'.


    concatenate ld_1
                ld_2
                ld_3
                ld_4
         into ep_content.

    lv_relation_id = 0.
    lo_iterator = me->excel->get_worksheets_iterator( ).
    while lo_iterator->has_next( ) eq abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).

      lo_worksheet->sheet_setup->get_header_footer( importing ep_odd_header = ls_odd_header
                                                              ep_odd_footer = ls_odd_footer
                                                              ep_even_header = ls_even_header
                                                              ep_even_footer = ls_even_footer ).

      ld_5 = me->set_vml_shape_header( ls_odd_header ).
      concatenate ep_content
                  ld_5
             into ep_content.
      ld_5 = me->set_vml_shape_header( ls_even_header ).
      concatenate ep_content
                  ld_5
             into ep_content.
      ld_5 = me->set_vml_shape_footer( ls_odd_footer ).
      concatenate ep_content
                  ld_5
             into ep_content.
      ld_5 = me->set_vml_shape_footer( ls_even_footer ).
      concatenate ep_content
                  ld_5
             into ep_content.
    endwhile.

    ld_7 = '</xml>'.

    concatenate ep_content
                ld_7
           into ep_content.

  endmethod.


  method zif_excel_writer~write_file.
    me->excel = io_excel.

    ep_file = me->create( ).
  endmethod.
endclass.
